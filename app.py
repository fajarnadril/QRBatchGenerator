import streamlit as st
import pandas as pd
import qrcode
import os
import shutil
from PIL import Image, ImageDraw, ImageFont
import zipfile

# --- Konfigurasi Awal ---
output_folder = 'qr_codes'

# Coba beberapa font umum
try:
    # Font Arial ukuran 30
    font = ImageFont.truetype("arial.ttf", 30) 
except IOError:
    try:
        # Font umum di Linux
        font = ImageFont.truetype("LiberationSans-Regular.ttf", 30) 
    except IOError:
        # Fallback ke font default
        font = ImageFont.load_default() 

# Ukuran gambar QR Code
qr_size = 500
# Tinggi tambahan untuk teks di bawah QR Code
text_height_padding = 60 
# Total tinggi gambar dengan teks
total_image_height = qr_size + text_height_padding 

# --- Streamlit UI ---
st.title("QR Code Generator from Excel")
st.markdown("""
    Untuk memulai, Anda bisa mengunduh template Excel dengan format yang benar dari link berikut:
    [Download Template Excel](https://github.com/fajarnadril/QRBatchGenerator/blob/main/template.xlsx)
""")
st.write("Upload file Excel yang berisi nama file dan URL.")

# File uploader untuk Excel
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    # Baca file Excel
    df = pd.read_excel(uploaded_file)

    if len(df.columns) < 2:
        st.error("File Excel harus memiliki minimal 2 kolom (Nama File dan URL).")
    else:
        nama_kolom = df.columns[0]
        url_kolom = df.columns[1]
        
        st.write(f"Kolom yang tersedia: {df.columns.tolist()}")
        st.write(f"Nama file diambil dari kolom '{nama_kolom}' dan URL dari kolom '{url_kolom}'.")

        # Buat folder untuk menyimpan QR Code
        if os.path.exists(output_folder):
            shutil.rmtree(output_folder)
        os.makedirs(output_folder)

        # Proses pembuatan QR Code
        success_count = 0
        failed_count = 0
        failed_urls = []

        for index, row in df.iterrows():
            nama_file = str(row[nama_kolom])
            url = str(row[url_kolom]).strip()

            # Bersihkan nama file
            nama_file_clean = "".join(c for c in nama_file if c.isalnum() or c in (' ', '-', '_')).strip()
            nama_teks = nama_file_clean

            # Skip jika URL kosong atau NaN
            if pd.isna(url) or url == 'nan' or url.strip() == '':
                failed_count += 1
                continue

            try:
                qr = qrcode.QRCode(
                    version=1,
                    error_correction=qrcode.constants.ERROR_CORRECT_H,
                    box_size=10,
                    border=4,
                )
                qr.add_data(url)
                qr.make(fit=True)

                img_qr = qr.make_image(fill_color="black", back_color="white")
                img_qr_resized = img_qr.resize((qr_size, qr_size), Image.Resampling.LANCZOS)

                final_img = Image.new('RGB', (qr_size, total_image_height), color='white')
                final_img.paste(img_qr_resized, (0, 0))

                draw = ImageDraw.Draw(final_img)
                text_width = draw.textlength(nama_teks, font=font)
                x_text = (qr_size - text_width) / 2
                y_text = qr_size + (text_height_padding - 30) / 2 
                draw.text((x_text, y_text), nama_teks, font=font, fill=(0, 0, 0))

                output_path = os.path.join(output_folder, f"{nama_file_clean}.jpg")
                counter = 1
                while os.path.exists(output_path):
                    output_path = os.path.join(output_folder, f"{nama_file_clean}_{counter}.jpg")
                    counter += 1

                final_img.save(output_path, 'JPEG', quality=95)
                success_count += 1

            except Exception as e:
                failed_count += 1
                failed_urls.append((nama_file_clean, url, str(e)))

        # Menampilkan ringkasan
        st.write(f"Total data: {len(df)}")
        st.write(f"Berhasil: {success_count}")
        st.write(f"Gagal: {failed_count}")

        if failed_urls:
            st.write("📋 Daftar URL yang gagal:")
            for nama, url, error in failed_urls:
                st.write(f"- {nama}: {error}")

        # Membuat ZIP dan memberikan link untuk di-download
        if success_count > 0:
            zip_filename = "qr_codes_hasil.zip"
            shutil.make_archive(zip_filename.replace(".zip", ""), 'zip', output_folder)
            with open(f"{zip_filename}", "rb") as f:
                st.download_button("Download ZIP", f, file_name=zip_filename)
        else:
            st.warning("Tidak ada QR Code yang berhasil dibuat.")
