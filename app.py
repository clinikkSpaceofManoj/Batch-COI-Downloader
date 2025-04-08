import streamlit as st
import pandas as pd
import os
import requests
from zipfile import ZipFile
import tempfile

st.title("📄 COI Downloader!")

uploaded_file = st.file_uploader("Upload a Excel file with PDF URLs", type=["xlsx"])

if uploaded_file:
    data = pd.read_excel(uploaded_file)
    st.write("Data Uploaded")

    if st.button("Download and Zip PDFs"):
        with st.spinner("Downloading PDFs and zipping..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                pdf_folder = os.path.join(tmpdir, "pdfs")
                os.makedirs(pdf_folder, exist_ok=True)

                data["Downloaded Status"] = ""

                for idx, row in data.iterrows():
                    url = row[data.columns[1]]
                    try:
                        response = requests.get(url)
                        response.raise_for_status()

                        filename = f"file_{row[data.columns[0]]}.pdf"
                        filepath = os.path.join(pdf_folder, filename)

                        with open(filepath, 'wb') as f:
                            f.write(response.content)

                        data.at[idx, "Downloaded Status"] = "Downloaded"
                    except Exception as e:
                        data.at[idx, "Downloaded Status"] = f"Failed: {e}"

                zip_path = os.path.join(tmpdir, "Coi.zip")
                with ZipFile(zip_path, 'w') as zipf:
                    for file_name in os.listdir(pdf_folder):
                        zipf.write(os.path.join(pdf_folder, file_name), arcname=file_name)

                st.success("All PDFs downloaded and zipped!")

                with open(zip_path, 'rb') as f:
                    st.download_button("📥 Download ZIP", f, file_name="Coi.zip")

        st.write("Download Status:")
        missingData = data[data["Downloaded Status"] != "Downloaded"]
        st.dataframe(missingData)
