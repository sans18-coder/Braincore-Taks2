import streamlit as st
import pandas as pd
import zipfile
import os
from generate import execute
import shutil
def remove_file(path):
    if os.path.exists(path):
        for item in os.listdir(path):
            item_path = os.path.join(path, item)
            if os.path.isfile(item_path):
                os.remove(item_path)
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)
def extract_zip(zip_file):
    with zipfile.ZipFile(zip_file, 'r') as z:
        z.extractall('extracted_files')
        return z.namelist()
    
remove_file('./extracted_files/')
remove_file('./output_file/')
remove_file('./temp/')
st.title('Sistem Rekapitulasi Suara Pemilu')
st.markdown('**Note:** zip harus berisi image, jangan sampai image terbalik.')
if 'file_processed' not in st.session_state:
    st.session_state.file_processed = False
    
uploaded_file = st.file_uploader("Choose a ZIP file", type="zip")

if uploaded_file is not None:
    filename = uploaded_file.name
    basename = os.path.splitext(filename)[0]
    extract_zip(uploaded_file)
    st.markdown('**Note:** Silahkan tunggu beberapa menit lama proses tergantung banyaknya gambar')
    execute(data_path=os.path.join('extracted_files',basename,'*'))
    st.session_state.file_processed = True
    if os.path.exists('./output_file/hasil_ocr.xlsx'):
        df = pd.read_excel('./output_file/hasil_ocr.xlsx')
        st.markdown('**Note:** periksa kembali hasilnya.')
        st.dataframe(df)
        with open('./output_file/hasil_ocr.xlsx', 'rb') as file:
            file_data = file.read()
        st.download_button(
                            label='Unduh Excel',
                            data=file_data,
                            file_name='Result.xlsx',)
        uploaded_file = None

if st.button('Reset'):
    st.session_state.file_processed = False
    st.experimental_rerun()
