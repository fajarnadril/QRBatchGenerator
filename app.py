import streamlit as st
import pandas as pd
import qrcode
import os
import shutil
from PIL import Image, ImageDraw, ImageFont

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

# --- Streamlit Setup ---
st.title("QR Code Generator from Excel")
st.markdown("""
    Untuk memulai, Anda bisa mengunduh template Excel dengan format yang benar dari link berikut:
    [Download Template Excel](https://github.com/fajarnadril/QRBatchGenerator/blob/main/template.xlsx)
""")
st.markdown("Upload an Excel file with two columns: file name and URL to generate QR codes.")

# --- Langkah 1: Upload File Excel ---
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    st.write(f"Data loaded with {len(df)} rows.")

    # Asumsi kolom pertama adalah nama file, kolom kedua adalah URL
    if len(df.columns) < 2:
        st.error("❌ ERROR: File Excel harus memiliki minimal 2 kolom (Nama File dan URL).")
    else:
        nama_kolom = df.columns[0]
        url_kolom = df.columns[1]
        
        st.write(f"Kolom yang tersedia: {df.columns.tolist()}")
        st.write(f"Using '{nama_kolom}' for file names and '{url_kolom}' for URLs")

        # --- Langkah 2: Proses Pembuatan QR Code ---
        st.subheader("QR Code Generation Progress")

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        success_count = 0
        failed_count = 0
        failed_urls = []

        for index, row in df.iterrows():
            nama_file = str(row[nama_kolom])
            url = str(row[url_kolom]).strip()

            # Bersihkan nama file dari karakter yang tidak valid (untuk NAMA FILE JPG)
            nama_file_clean = "".join(c for c in nama_file if c.isalnum() or c in (' ', '-', '_')).strip()
            nama_teks = nama_file_clean  # Teks di bawah QR tanpa .jpg

            # Skip jika URL kosong atau NaN
            if pd.isna(url) or url == 'nan' or url.strip() == '':
                st.warning(f"⚠️  Row {index + 1}: Empty URL for '{nama_file_clean}'")
                failed_count += 1
                continue

            try:
                # Buat objek QR Code
                qr = qrcode.QRCode(
                    version=1,
                    error_correction=qrcode.constants.ERROR_CORRECT_H,
                    box_size=10,
                    border=4,
                )
                qr.add_data(url)
                qr.make(fit=True)

                # Buat gambar QR code dan Resize menjadi 500x500
                img_qr = qr.make_image(fill_color="black", back_color="white")
                img_qr_resized = img_qr.resize((qr_size, qr_size), Image.Resampling.LANCZOS)

                # Buat kanvas putih baru dengan ruang untuk teks
                final_img = Image.new('RGB', (qr_size, total_image_height), color='white')
                final_img.paste(img_qr_resized, (0, 0))  # Tempel QR Code

                draw = ImageDraw.Draw(final_img)

                # Tambahkan teks di bawah QR Code
                text_width = draw.textlength(nama_teks, font=font)
                x_text = (qr_size - text_width) / 2
                y_text = qr_size + (text_height_padding - 30) / 2 

                draw.text((x_text, y_text), nama_teks, font=font, fill=(0, 0, 0))

                # Tentukan path penyimpanan
                output_path = os.path.join(output_folder, f"{nama_file_clean}.jpg")

                # Jika file sudah ada, tambahkan nomor (penanganan duplikasi)
                counter = 1
                while os.path.exists(output_path):
                    output_path = os.path.join(output_folder, f"{nama_file_clean}_{counter}.jpg")
                    counter += 1

                final_img.save(output_path, 'JPEG', quality=95)

                success_count += 1
                st.success(f"✓ Berhasil: {nama_file_clean}.jpg")

            except Exception as e:
                failed_count += 1
                failed_urls.append((nama_file_clean, url, str(e)))
                st.error(f"✗ Gagal: {nama_file_clean} - {str(e)}")

        # --- Langkah 3: Ringkasan ---
        st.subheader("Summary of QR Code Generation")

        st.write(f"Total data: {len(df)}")
        st.write(f"Success: {success_count}")
        st.write(f"Failed: {failed_count}")

        if failed_urls:
            st.write("📋 List of Failed URLs:")
            for nama, url, error in failed_urls:
                st.write(f"  - {nama}: {error}")

        # --- Download ZIP ---
        if success_count > 0:
            zip_filename = 'qr_codes_hasil.zip'
            shutil.make_archive('qr_codes_hasil', 'zip', output_folder)  # Create zip file

            st.download_button(
                label="Download QR Codes as ZIP",
                data=open('qr_codes_hasil.zip', 'rb'),
                file_name=zip_filename,
                mime='application/zip'
            )
        else:
            st.warning("⚠️ No QR codes were successfully created.")
