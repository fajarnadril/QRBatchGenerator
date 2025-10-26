import streamlit as st
import pandas as pd
import qrcode
import os
import shutil
import zipfile
from PIL import Image, ImageDraw, ImageFont
import io
import base64

# Konfigurasi halaman
st.set_page_config(
  page_title="QR Code Generator",
  page_icon="📱",
  layout="wide"
)

# Fungsi untuk membuat font
def get_font():
  try:
      font = ImageFont.truetype("arial.ttf", 30)
  except IOError:
      try:
          font = ImageFont.truetype("LiberationSans-Regular.ttf", 30)
      except IOError:
          font = ImageFont.load_default()
  return font

# Fungsi untuk membuat QR code
def create_qr_code(url, nama_file, font):
  qr_size = 500
  text_height_padding = 60
  total_image_height = qr_size + text_height_padding
  
  # Buat objek QR Code
  qr = qrcode.QRCode(
      version=1,
      error_correction=qrcode.constants.ERROR_CORRECT_H,
      box_size=10,
      border=4,
  )
  qr.add_data(url)
  qr.make(fit=True)

  # Buat gambar QR code dan resize
  img_qr = qr.make_image(fill_color="black", back_color="white")
  img_qr_resized = img_qr.resize((qr_size, qr_size), Image.Resampling.LANCZOS)

  # Buat kanvas putih baru dengan ruang untuk teks
  final_img = Image.new('RGB', (qr_size, total_image_height), color='white')
  final_img.paste(img_qr_resized, (0, 0))
  
  draw = ImageDraw.Draw(final_img)

  # Tambahkan teks di bawah QR Code
  text_width = draw.textlength(nama_file, font=font)
  x_text = (qr_size - text_width) / 2
  y_text = qr_size + (text_height_padding - 30) / 2

  draw.text((x_text, y_text), nama_file, font=font, fill=(0, 0, 0))
  
  return final_img

# Fungsi untuk membuat file ZIP
def create_zip_file(images_dict):
  zip_buffer = io.BytesIO()
  with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
      for filename, image in images_dict.items():
          img_buffer = io.BytesIO()
          image.save(img_buffer, format='JPEG', quality=95)
          zip_file.writestr(f"{filename}.jpg", img_buffer.getvalue())
  zip_buffer.seek(0)
  return zip_buffer.getvalue()

# Header aplikasi
st.title("📱 QR Code Generator Excel")
st.markdown("---")

# Sidebar untuk instruksi
with st.sidebar:
  st.header("📋 Instruksi Penggunaan")
  st.markdown("""
  1. **Upload file Excel** dengan format:
     - Kolom 1: Nama file QR Code
     - Kolom 2: URL yang akan di-encode
  
  2. **Klik tombol Generate** untuk membuat QR Code
  
  3. **Download hasil** dalam format ZIP
  
  **Format Excel yang didukung:**
  - .xlsx
  - .xls
  """)

# Upload file
uploaded_file = st.file_uploader(
  "Upload file Excel Anda",
  type=['xlsx', 'xls'],
  help="File harus memiliki minimal 2 kolom: Nama File dan URL"
)

if uploaded_file is not None:
  try:
      # Baca file Excel
      df = pd.read_excel(uploaded_file)
      
      st.success(f"✅ File berhasil dibaca! Ditemukan {len(df)} baris data.")
      
      # Tampilkan preview data
      st.subheader("📊 Preview Data")
      st.dataframe(df.head(10))
      
      # Validasi kolom
      if len(df.columns) < 2:
          st.error("❌ File Excel harus memiliki minimal 2 kolom (Nama File dan URL).")
      else:
          nama_kolom = df.columns[0]
          url_kolom = df.columns[1]
          
          st.info(f"📝 Menggunakan kolom '{nama_kolom}' untuk nama file dan '{url_kolom}' untuk URL")
          
          # Tombol generate
          if st.button("🚀 Generate QR Codes", type="primary"):
              
              # Progress bar
              progress_bar = st.progress(0)
              status_text = st.empty()
              
              # Container untuk hasil
              results_container = st.container()
              
              font = get_font()
              images_dict = {}
              success_count = 0
              failed_count = 0
              failed_urls = []
              
              # Proses setiap baris
              for index, row in df.iterrows():
                  progress = (index + 1) / len(df)
                  progress_bar.progress(progress)
                  status_text.text(f"Memproses {index + 1}/{len(df)}: {row[nama_kolom]}")
                  
                  nama_file = str(row[nama_kolom])
                  url = str(row[url_kolom]).strip()
                  
                  # Bersihkan nama file
                  nama_file_clean = "".join(c for c in nama_file if c.isalnum() or c in (' ', '-', '_')).strip()
                  
                  # Skip jika URL kosong
                  if pd.isna(url) or url == 'nan' or url.strip() == '':
                      failed_count += 1
                      failed_urls.append((nama_file_clean, "URL kosong", "URL tidak valid"))
                      continue
                  
                  try:
                      # Buat QR code
                      qr_image = create_qr_code(url, nama_file_clean, font)
                      
                      # Handle duplikasi nama file
                      original_name = nama_file_clean
                      counter = 1
                      while nama_file_clean in images_dict:
                          nama_file_clean = f"{original_name}_{counter}"
                          counter += 1
                      
                      images_dict[nama_file_clean] = qr_image
                      success_count += 1
                      
                  except Exception as e:
                      failed_count += 1
                      failed_urls.append((nama_file_clean, url, str(e)))
              
              # Selesai processing
              progress_bar.progress(1.0)
              status_text.text("✅ Selesai!")
              
              # Tampilkan hasil
              with results_container:
                  st.markdown("---")
                  st.subheader("📊 Hasil Pembuatan QR Code")
                  
                  col1, col2, col3 = st.columns(3)
                  with col1:
                      st.metric("Total Data", len(df))
                  with col2:
                      st.metric("Berhasil", success_count)
                  with col3:
                      st.metric("Gagal", failed_count)
                  
                  # Tampilkan error jika ada
                  if failed_urls:
                      st.warning("⚠️ Beberapa QR Code gagal dibuat:")
                      error_df = pd.DataFrame(failed_urls, columns=['Nama File', 'URL', 'Error'])
                      st.dataframe(error_df)
                  
                  # Download ZIP jika ada yang berhasil
                  if success_count > 0:
                      st.success(f"🎉 Berhasil membuat {success_count} QR Code!")
                      
                      # Buat file ZIP
                      zip_data = create_zip_file(images_dict)
                      
                      # Tombol download
                      st.download_button(
                          label="📥 Download QR Codes (ZIP)",
                          data=zip_data,
                          file_name="qr_codes_hasil.zip",
                          mime="application/zip"
                      )
                      
                      # Preview beberapa QR code
                      st.subheader("👀 Preview QR Codes")
                      preview_count = min(6, len(images_dict))
                      cols = st.columns(3)
                      
                      for i, (filename, image) in enumerate(list(images_dict.items())[:preview_count]):
                          with cols[i % 3]:
                              st.image(image, caption=filename, width=200)
                  
                  else:
                      st.error("❌ Tidak ada QR Code yang berhasil dibuat.")
                      
  except Exception as e:
      st.error(f"❌ Error membaca file: {str(e)}")
      st.info("Pastikan file Excel Anda dalam format yang benar dan tidak corrupt.")

else:
  st.info("👆 Silakan upload file Excel untuk memulai.")

# Footer
st.markdown("---")
st.markdown("*Dibuat dengan ❤️ menggunakan Streamlit*")