import os, re, time, zipfile, tempfile
from urllib.parse import urlparse
from io import BytesIO

import pandas as pd
import requests
import streamlit as st


# ====== CONFIG ======
SHEET_NAME = "Sheet1"
SKIP_ROWS = 0
TIMEOUT = 30
RETRIES = 3
SLEEP_BETWEEN = 0.3
DEDUPE_SAME_URL_IN_THIS_RUN = True
# ====================


def is_url(x: str) -> bool:
    return isinstance(x, str) and x.strip().lower().startswith(("http://", "https://"))


def guess_ext_from_headers_or_url(url: str, content_type: str) -> str:
    if content_type:
        ct = content_type.split(";")[0].strip().lower()
        mapping = {
            "image/jpeg": ".jpg",
            "image/jpg": ".jpg",
            "image/png": ".png",
            "image/webp": ".webp",
            "image/gif": ".gif",
            "image/bmp": ".bmp",
            "image/tiff": ".tif",
        }
        if ct in mapping:
            return mapping[ct]

    path = urlparse(url).path.lower()
    for ext in [".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".tif", ".tiff"]:
        if path.endswith(ext):
            return ".jpg" if ext == ".jpeg" else ".tif" if ext == ".tiff" else ext
    return ".jpg"


def parse_sku_from_url(url: str) -> str:
    m = re.search(r"/(\d+)(?:r\d*)?(?:_[0-9]+)?\.[a-zA-Z0-9]+", str(url))
    return m.group(1) if m else ""


def download_images(uploaded_file):
    with tempfile.TemporaryDirectory() as temp_dir:
        output_dir = os.path.join(temp_dir, "images")
        os.makedirs(output_dir, exist_ok=True)

        df = pd.read_excel(
            uploaded_file,
            sheet_name=SHEET_NAME,
            usecols="A",
            skiprows=SKIP_ROWS,
            engine="openpyxl"
        )

        urls = df.iloc[:, 0].dropna().astype(str).map(str.strip)
        urls = [u for u in urls if is_url(u)]

        session = requests.Session()
        session.headers.update({"User-Agent": "Mozilla/5.0"})

        run_cache = set() if DEDUPE_SAME_URL_IN_THIS_RUN else None
        sku_index_map = {}
        results = []

        progress = st.progress(0)
        status_text = st.empty()

        for row_no, url in enumerate(urls, start=1):
            status_text.text(f"Downloading {row_no}/{len(urls)}")

            if run_cache is not None:
                if url in run_cache:
                    results.append((row_no, url, "SKIP (dup url in run)", ""))
                    progress.progress(row_no / len(urls))
                    continue
                run_cache.add(url)

            sku = parse_sku_from_url(url)
            if not sku:
                results.append((row_no, url, "SKIP (no sku parsed)", ""))
                progress.progress(row_no / len(urls))
                continue

            ok = False
            last_err = ""

            for attempt in range(1, RETRIES + 1):
                try:
                    r = session.get(
                        url,
                        timeout=TIMEOUT,
                        stream=True,
                        allow_redirects=True
                    )
                    r.raise_for_status()

                    ct = r.headers.get("Content-Type", "")
                    if not ct.lower().startswith("image/"):
                        results.append((row_no, url, "SKIP (not image)", ct))
                        ok = True
                        break

                    ext = guess_ext_from_headers_or_url(url, ct)

                    sku_index_map[sku] = sku_index_map.get(sku, 0) + 1
                    filename = f"{sku}_{sku_index_map[sku]}{ext}"
                    filepath = os.path.join(output_dir, filename)

                    with open(filepath, "wb") as f:
                        for chunk in r.iter_content(chunk_size=1024 * 256):
                            if chunk:
                                f.write(chunk)

                    results.append((row_no, url, "OK", filename))
                    ok = True
                    break

                except Exception as e:
                    last_err = f"{type(e).__name__}: {e}"
                    time.sleep(0.8)

            if not ok:
                results.append((row_no, url, "ERROR", last_err))

            progress.progress(row_no / len(urls))
            time.sleep(SLEEP_BETWEEN)

        res_df = pd.DataFrame(results, columns=["row", "url", "status", "info"])

        result_excel_path = os.path.join(temp_dir, "result.xlsx")
        res_df.to_excel(result_excel_path, index=False)

        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for filename in os.listdir(output_dir):
                file_path = os.path.join(output_dir, filename)
                zipf.write(file_path, arcname=f"images/{filename}")

            zipf.write(result_excel_path, arcname="result.xlsx")

        zip_buffer.seek(0)

        return zip_buffer, res_df, len(urls)


st.set_page_config(page_title="Image Downloader", layout="centered")

st.title("Image Downloader from Excel")
st.write("อัปโหลดไฟล์ Excel ที่มี URL รูปภาพอยู่ในคอลัมน์ A")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    if st.button("Download Images"):
        with st.spinner("กำลังดาวน์โหลดรูปภาพ..."):
            zip_file, result_df, total_urls = download_images(uploaded_file)

        st.success(f"เสร็จแล้ว พบ URL ทั้งหมด {total_urls} รายการ")

        st.dataframe(result_df)

        st.download_button(
            label="Download ZIP File",
            data=zip_file,
            file_name="downloaded_images.zip",
            mime="application/zip"
        )
