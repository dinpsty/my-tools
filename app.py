import io
import os
import re
import zipfile

import streamlit as st
from docx import Document
from openpyxl import Workbook
from reportlab.pdfgen import canvas


st.set_page_config(
    page_title="My Tools",
    page_icon="🛠️",
    layout="centered"
)

st.title("🛠️ MY TOOLS")

pilihan = st.selectbox(
    "Pilih Tools",
    [
        "Folder Maker",
        "PDF Blank",
        "DOC Blank",
        "XCL Blank"
    ]
)


karakter_terlarang = r'[<>:"/\\|?*]'


def bersihkan_nama(nama):
    nama = re.sub(karakter_terlarang, "", nama)
    return nama.strip().rstrip(". ")


def daftar_nama_bersih(teks, ekstensi):
    hasil = []
    sudah_ada = set()

    for nama in teks.splitlines():
        nama = nama.strip()

        if nama.lower().endswith(ekstensi):
            nama = nama[:-len(ekstensi)]

        nama = bersihkan_nama(nama)

        if nama.lower().endswith(" (blank)"):
            nama = nama[:-8].strip()

        if nama and nama.lower() not in sudah_ada:
            sudah_ada.add(nama.lower())
            hasil.append(nama)

    return hasil


def buat_zip(daftar_file):
    buffer = io.BytesIO()

    with zipfile.ZipFile(
        buffer, "w", zipfile.ZIP_DEFLATED
    ) as zip_file:

        for nama, data in daftar_file:
            zip_file.writestr(nama, data)

    return buffer.getvalue()


def nama_unik(nama, digunakan):
    nama_awal = nama
    nomor = 2

    while nama.lower() in digunakan:
        nama = f"{nama_awal} ({nomor})"
        nomor += 1

    digunakan.add(nama.lower())

    return nama


# ==========================================
# FOLDER MAKER
# ==========================================

if pilihan == "Folder Maker":

    st.subheader("📁 Folder Maker")

    st.write("Masukkan nama folder satu per baris.")

    input_nama_folder = st.text_area(
        "Nama Folder",
        height=250,
        placeholder="Contoh:\nAndi\nBudi\nCitra",
        key="nama_folder"
    )

    if st.button("Buat Folder"):

        nama_folder = []
        nama_sudah_ada = set()

        for nama in input_nama_folder.splitlines():

            nama = bersihkan_nama(nama.strip())

            if not nama or nama in (".", ".."):
                continue

            if nama.lower() in nama_sudah_ada:
                continue

            nama_sudah_ada.add(nama.lower())
            nama_folder.append(nama)

        if not nama_folder:

            st.error("Masukkan minimal satu nama folder.")

        else:

            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(
                zip_buffer,
                "w",
                zipfile.ZIP_DEFLATED
            ) as zip_file:

                for nama in nama_folder:
                    zip_file.writestr(f"{nama}/", "")

            st.success(
                f"{len(nama_folder)} folder berhasil dibuat."
            )

            st.download_button(
                label="⬇️ Download Folder",
                data=zip_buffer.getvalue(),
                file_name="folder_hasil.zip",
                mime="application/zip"
            )


# ==========================================
# PDF BLANK, DOC BLANK, XCL BLANK
# ==========================================

elif pilihan in [
    "PDF Blank",
    "DOC Blank",
    "XCL Blank"
]:

    st.subheader(pilihan)

    st.write("Masukkan nama file satu per baris.")

    input_nama = st.text_area(
        "Nama File",
        height=250,
        placeholder="Contoh:\nAndi\nBudi\nCitra"
    )

    if pilihan == "PDF Blank":
        ekstensi = ".pdf"

    elif pilihan == "DOC Blank":
        ekstensi = ".docx"

    else:
        ekstensi = ".xlsx"

    if st.button("Buat File"):

        daftar_nama = daftar_nama_bersih(
            input_nama,
            ekstensi
        )

        if not daftar_nama:

            st.error("Masukkan minimal satu nama file.")

        else:

            daftar_file = []

            for nama in daftar_nama:

                buffer = io.BytesIO()

                if pilihan == "PDF Blank":

                    pdf = canvas.Canvas(buffer)
                    pdf.showPage()
                    pdf.save()

                elif pilihan == "DOC Blank":

                    document = Document()
                    document.save(buffer)

                else:

                    workbook = Workbook()
                    workbook.save(buffer)
                    workbook.close()

                nama_file = f"{nama} (BLANK){ekstensi}"

                daftar_file.append(
                    (nama_file, buffer.getvalue())
                )

            st.success(
                f"{len(daftar_file)} file berhasil dibuat."
            )

            st.download_button(
                "⬇️ Download File",
                data=buat_zip(daftar_file),
                file_name="blank_files.zip",
                mime="application/zip"
            )