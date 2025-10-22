import streamlit as st
import pandas as pd
import os
import qrcode
from PIL import Image
import shutil

# Streamlit app
st.title("QR Code Generator from Excel")

st.markdown("""
    Untuk memulai, Anda bisa mengunduh template Excel dengan format yang benar dari link berikut:
    [Download Template Excel](https://github.com/fajarnadril/QRBatchGenerator/blob/main/template.xlsx)
""")

st.write("Upload an Excel file containing URLs and names, and the app will generate QR codes.")

# Upload Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    # Read the file
    df = pd.read_excel(uploaded_file)
    
    # Show the dataframe to the user for inspection
    st.write(f"Preview of your data:")
    st.write(df.head())

    # Create folder for storing QR codes
    output_folder = 'qr_codes'
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Get column names
    nama_kolom = df.columns[0]
    url_kolom = df.columns[1]

    st.write(f"Using '{nama_kolom}' for filenames and '{url_kolom}' for URLs")

    # Counter for tracking success and failure
    success_count = 0
    failed_count = 0
    failed_urls = []

    # Process each row to generate QR codes
    for index, row in df.iterrows():
        nama_file = str(row[nama_kolom])
        url = str(row[url_kolom]).strip()
        
        if pd.isna(url) or url == 'nan' or url.strip() == '':
            st.warning(f"⚠️ Row {index + 1}: Empty URL for '{nama_file}'")
            failed_count += 1
            continue
        
        try:
            # Generate QR code
            qr = qrcode.QRCode(
                version=1,  # QR code size
                error_correction=qrcode.constants.ERROR_CORRECT_H,  # High error correction
                box_size=10,  # Size of each box
                border=4,  # Border thickness
            )
            qr.add_data(url)
            qr.make(fit=True)
            
            # Create the QR code image
            img = qr.make_image(fill_color="black", back_color="white")
            
            # Resize to 500x500 pixels
            img_resized = img.resize((500, 500), Image.Resampling.LANCZOS)
            
            # Convert to RGB for saving as JPG
            if img_resized.mode != 'RGB':
                img_resized = img_resized.convert('RGB')
            
            # Clean file name
            nama_file_clean = "".join(c for c in nama_file if c.isalnum() or c in (' ', '-', '_')).strip()
            
            # Save image as JPG
            output_path = os.path.join(output_folder, f"{nama_file_clean}.jpg")
            
            # Handle filename conflict by adding a counter
            counter = 1
            while os.path.exists(output_path):
                output_path = os.path.join(output_folder, f"{nama_file_clean}_{counter}.jpg")
                counter += 1
            
            img_resized.save(output_path, 'JPEG', quality=95)
            
            success_count += 1
            st.success(f"✓ Success: {nama_file_clean}.jpg")
        
        except Exception as e:
            failed_count += 1
            failed_urls.append((nama_file, url, str(e)))
            st.error(f"✗ Failed: {nama_file} - {str(e)}")

    # Display the summary
    st.write("="*60)
    st.write("SUMMARY OF QR CODE CREATION:")
    st.write("="*60)
    st.write(f"Total rows: {len(df)}")
    st.write(f"Successfully created QR codes: {success_count}")
    st.write(f"Failed attempts: {failed_count}")

    if failed_urls:
        st.write("\n📋 List of failed URLs:")
        for nama, url, error in failed_urls:
            st.write(f"  - {nama}: {error}")

    st.write(f"\n📁 QR codes saved in folder: {output_folder}")

    # Download as ZIP
    if success_count > 0:
        st.write("\n📦 Creating ZIP file...")
        shutil.make_archive('qr_codes', 'zip', output_folder)
        
        st.write("⬇️  Download the ZIP file below:")
        st.download_button(
            label="Download QR Codes ZIP",
            data=open('qr_codes.zip', 'rb').read(),
            file_name="qr_codes.zip"
        )
        st.write("✅ Done!")
    else:
        st.warning("\n⚠️ No QR codes were generated.")
