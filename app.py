import csv
import io
import re
import shutil
import zipfile
from pathlib import Path
import streamlit as st
import requests
from playwright.sync_api import sync_playwright

st.set_page_config(
    page_title="Bulk Bill Downloader",
    page_icon="🧾",
    layout="centered"
)

# Custom Styling
st.markdown("""
    <style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-text {
        font-size: 1rem;
        color: #4B5563;
        margin-bottom: 1.5rem;
    }
    .stProgress > div > div > div > div {
        background-color: #2563EB;
    }
    </style>
""", unsafe_allow_html=True)

TEMP_DIR = Path("downloaded_bills_session")
LINK_HEADERS = ["invoice link", "bill link", "link", "url"]
NAME_HEADERS = ["invoice no", "invoice no.", "invoice number", "bill no", "invoice"]

REAL_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

def find_column(headers, candidates):
    for i, h in enumerate(headers):
        if h is None:
            continue
        h_clean = str(h).strip().lower()
        for c in candidates:
            if c in h_clean:
                return i
    return None

def sanitize_filename(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r'[\\/*?:"<>|]', "_", name)
    return name or "bill"

def parse_uploaded_file(uploaded_file):
    filename = uploaded_file.name.lower()
    rows = []

    if filename.endswith(".csv"):
        content = uploaded_file.getvalue().decode("utf-8-sig", errors="replace")
        reader = csv.reader(io.StringIO(content))
        rows = [r for r in reader if any(field.strip() for field in r)]
    elif filename.endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        wb = load_workbook(uploaded_file, data_only=True)
        ws = wb.active
        rows = [r for r in list(ws.iter_rows(values_only=True)) if any(r)]
    else:
        raise ValueError("Invalid format. Please upload .csv or .xlsx")

    if not rows:
        raise ValueError("File bilkul khaali hai.")

    header_row_idx, link_idx, name_idx = None, None, None
    for idx, row in enumerate(rows):
        l_idx = find_column(row, LINK_HEADERS)
        n_idx = find_column(row, NAME_HEADERS)
        if l_idx is not None and n_idx is not None:
            header_row_idx = idx
            link_idx = l_idx
            name_idx = n_idx
            break

    if header_row_idx is None:
        raise ValueError("Header matching failed. File mein 'Invoice No' aur 'Invoice Link' columns nahi mile.")

    entries = []
    for row in rows[header_row_idx + 1:]:
        if len(row) > max(link_idx, name_idx):
            link = row[link_idx]
            name = row[name_idx]
            if link and str(link).strip():
                entries.append((sanitize_filename(name), str(link).strip()))

    return entries, rows[header_row_idx]

def download_single_invoice(context, name, link, total, progress_tracker, log_messages, log_area):
    out_path = TEMP_DIR / f"{name}.pdf"
    pdf_bytes = []
    downloaded = False
    page = context.new_page()

    # Block heavy tracking and media to speed up load time
    page.route("**/*", lambda route: route.abort() 
               if route.request.resource_type in ["image", "media", "font"] 
               else route.continue_())

    def intercept_pdf(response):
        try:
            ctype = response.headers.get("content-type", "").lower()
            if "application/pdf" in ctype or ".pdf" in response.url.lower():
                b = response.body()
                if len(b) > 2000:
                    pdf_bytes.append(b)
        except Exception:
            pass

    page.on("response", intercept_pdf)

    try:
        page.goto(link, wait_until="domcontentloaded", timeout=15000)
        page.wait_for_timeout(1000)

        # Strategy A: Captured raw PDF stream
        if pdf_bytes:
            with open(out_path, "wb") as f:
                f.write(pdf_bytes[-1])
            downloaded = True

        # Strategy B: Fast DOM PDF render
        if not downloaded:
            page.evaluate("""() => {
                const hideList = ['nav', 'header', '.sidebar', '[class*="sidebar"]', '[class*="banner"]', '[class*="drawer"]'];
                hideList.forEach(sel => {
                    document.querySelectorAll(sel).forEach(el => el.style.display = 'none');
                });
            }""")
            page.pdf(
                path=str(out_path),
                format="A4",
                print_background=True,
                margin={"top": "8mm", "bottom": "8mm", "left": "8mm", "right": "8mm"}
            )
            downloaded = True

        if downloaded and out_path.exists() and out_path.stat().st_size > 1000:
            msg = f"✅ [{progress_tracker['count']}/{total}] Done: {name}.pdf"
            error = None
        else:
            raise Exception("File empty or not generated")

    except Exception as e:
        msg = f"❌ [{progress_tracker['count']}/{total}] Failed: {name}"
        error = f"{name} | {link} | {e}"
    finally:
        page.close()

    progress_tracker['count'] += 1
    log_messages.append(msg)
    log_area.text_area("Download Terminal Logs", value="\n".join(log_messages[-15:]), height=220)
    return error

def execute_downloads(entries, progress_bar, status_text, log_area):
    if TEMP_DIR.exists():
        shutil.rmtree(TEMP_DIR)
    TEMP_DIR.mkdir(exist_ok=True)

    failed = []
    log_messages = []
    total = len(entries)
    progress_tracker = {'count': 1}

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-gpu", "--disable-dev-shm-usage"]
        )
        context = browser.new_context(user_agent=REAL_UA, viewport={"width": 1280, "height": 800})

        # Process 4 tabs concurrently
        max_workers = 4
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [
                executor.submit(
                    download_single_invoice,
                    context, name, link, total, progress_tracker, log_messages, log_area
                )
                for name, link in entries
            ]

            for idx, future in enumerate(as_completed(futures), 1):
                err = future.result()
                if err:
                    failed.append(err)
                progress_bar.progress(idx / total)
                status_text.markdown(f"**Processed:** `{idx}/{total}` invoices")

        browser.close()

    # Package into ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in TEMP_DIR.glob("*.pdf"):
            zf.write(file, arcname=file.name)
        if failed:
            zf.writestr("failed_bills.txt", "\n".join(failed))

    zip_buffer.seek(0)
    return zip_buffer, len(failed)
    # Create ZIP archive
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in TEMP_DIR.glob("*.pdf"):
            zf.write(file, arcname=file.name)
        if failed:
            zf.writestr("failed_bills.txt", "\n".join(failed))

    zip_buffer.seek(0)
    return zip_buffer, len(failed)

# --- Streamlit UI ---

st.markdown('<div class="main-header">Bulk Bill Downloader</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-text">Upload your MyBillBook Excel/CSV report to batch-download all PDF invoices into a single ZIP archive.</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Invoice Report", type=["csv", "xlsx"])

if uploaded_file is not None:
    try:
        entries, detected_headers = parse_uploaded_file(uploaded_file)
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric(label="Total Invoices Found", value=len(entries))
        with col2:
            st.metric(label="File Format", value=uploaded_file.name.split(".")[-1].upper())

        st.info(f"Headers auto-detected: `{', '.join([str(h) for h in detected_headers if h])}`")

        # --- Action Controls in UI ---
        test_mode = st.checkbox("🧪 Test Mode (Download only the first bill to verify)", value=False)

        if st.button("Start Download", type="primary", use_container_width=True):
            # Slice list to 1 item if test mode is enabled
            entries_to_process = entries[:1] if test_mode else entries

            progress_bar = st.progress(0)
            status_text = st.empty()
            log_area = st.empty()

            with st.spinner("Processing..."):
                zip_data, fail_count = execute_downloads(
                    entries_to_process, progress_bar, status_text, log_area
                )

            status_text.empty()
            if fail_count == 0:
                st.success(f"Processed {len(entries_to_process)} invoice(s) successfully!")
            else:
                st.warning(f"Completed with {fail_count} failed items. Details saved in 'failed_bills.txt'.")

            st.download_button(
                label="📦 Download Bills (ZIP File)",
                data=zip_data,
                file_name="test_invoice.zip" if test_mode else "all_downloaded_bills.zip",
                mime="application/zip",
                use_container_width=True,
            )

    except Exception as e:
        st.error(f"Error reading file: {e}")
