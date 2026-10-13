import streamlit as st
import io
import zipfile
import re
import os
import tempfile
import subprocess
from openpyxl import load_workbook

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
        "PDF Maker" ,
        "Excel Maker" ,
        "Word Maker" , 
        "Excel Sheet to PDF"
    ]
)

karakter_terlarang = r'[<>:"/\\|?*]'

if pilihan == "Folder Maker":

    st.subheader("📁 Folder Maker")

    st.write(
        "Masukkan nama folder satu per baris."
    )

    input_nama = st.text_area(
        "Nama Folder",
        height=250,
        placeholder="Contoh:\nAndi\nBudi\nCitra"
    )

    if st.button("Buat Folder"):

        daftar_nama = input_nama.splitlines()

        nama_bersih = []
        nama_sudah_ada = set()

        for nama in daftar_nama:

            nama = nama.strip()

            if not nama:
                continue

            nama = re.sub(karakter_terlarang, "", nama).strip()

            if not nama:
                continue

            if nama.lower() in nama_sudah_ada:
                continue

            nama_sudah_ada.add(nama.lower())
            nama_bersih.append(nama)

        if len(nama_bersih) == 0:

            st.error("Masukkan minimal satu nama folder.")

        else:

            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(
                zip_buffer,
                "w",
                zipfile.ZIP_DEFLATED
            ) as zip_file:

                for nama in nama_bersih:
                    zip_file.writestr(f"{nama}/", "")

            st.success(
                f"{len(nama_bersih)} folder berhasil dibuat."
            )

            st.download_button(
                label="⬇️ Download Folder",
                data=zip_buffer.getvalue(),
                file_name="folder_hasil.zip",
                mime="application/zip"
            )


elif pilihan == "PDF Maker":

    from reportlab.pdfgen import canvas

    st.subheader("📄 PDF Maker")

    st.write("Masukkan nama PDF satu per baris.")

    input_nama = st.text_area(
        "Nama PDF",
        height=250,
        placeholder="Contoh:\nAndi\nBudi\nCitra"
    )

    if st.button("Buat PDF"):

        daftar_nama = input_nama.splitlines()

        nama_bersih = []
        nama_sudah_ada = set()

        for nama in daftar_nama:

            nama = nama.strip()

            if not nama:
                continue

            if nama.lower().endswith(".pdf"):
                nama = nama[:-4]

            nama = re.sub(karakter_terlarang, "", nama).strip()

            if not nama:
                continue

            if nama.lower() in nama_sudah_ada:
                continue

            nama_sudah_ada.add(nama.lower())
            nama_bersih.append(nama)

        if len(nama_bersih) == 0:

            st.error("Masukkan minimal satu nama PDF.")

        else:

            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(
                zip_buffer,
                "w",
                zipfile.ZIP_DEFLATED
            ) as zip_file:

                for nama in nama_bersih:

                    pdf_buffer = io.BytesIO()

                    pdf = canvas.Canvas(pdf_buffer)
                    pdf.showPage()
                    pdf.save()

                    zip_file.writestr(
                        f"{nama}.pdf",
                        pdf_buffer.getvalue()
                    )

            st.success(
                f"{len(nama_bersih)} PDF berhasil dibuat."
            )

            st.download_button(
                label="⬇️ Download PDF",
                data=zip_buffer.getvalue(),
                file_name="pdf_hasil.zip",
                mime="application/zip"
            )

elif pilihan == "Excel Maker":

    from openpyxl import Workbook

    st.subheader("📊 Excel Maker")

    st.write("Masukkan nama Excel satu per baris.")

    input_nama = st.text_area(
        "Nama Excel",
        height=250,
        placeholder="Contoh:\nLaporan Januari\nLaporan Februari\nLaporan Maret"
    )

    if st.button("Buat Excel"):

        daftar_nama = input_nama.splitlines()

        nama_bersih = []
        nama_sudah_ada = set()

        for nama in daftar_nama:

            nama = nama.strip()

            if not nama:
                continue

            if nama.lower().endswith(".xlsx"):
                nama = nama[:-5]

            nama = re.sub(karakter_terlarang, "", nama).strip()

            if not nama:
                continue

            if nama.lower() in nama_sudah_ada:
                continue

            nama_sudah_ada.add(nama.lower())
            nama_bersih.append(nama)

        if len(nama_bersih) == 0:

            st.error("Masukkan minimal satu nama Excel.")

        else:

            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(
                zip_buffer,
                "w",
                zipfile.ZIP_DEFLATED
            ) as zip_file:

                for nama in nama_bersih:

                    excel_buffer = io.BytesIO()

                    workbook = Workbook()
                    workbook.save(excel_buffer)

                    zip_file.writestr(
                        f"{nama}.xlsx",
                        excel_buffer.getvalue()
                    )

            st.success(
                f"{len(nama_bersih)} file Excel berhasil dibuat."
            )

            st.download_button(
                label="⬇️ Download Excel",
                data=zip_buffer.getvalue(),
                file_name="excel_hasil.zip",
                mime="application/zip"
            )
elif pilihan == "Word Maker":

    from docx import Document

    st.subheader("📝 Word Maker")

    st.write("Masukkan nama Word satu per baris.")

    input_nama = st.text_area(
        "Nama Word",
        height=250,
        placeholder="Contoh:\nLaporan Januari\nLaporan Februari\nLaporan Maret"
    )

    if st.button("Buat Word"):

        daftar_nama = input_nama.splitlines()

        nama_bersih = []
        nama_sudah_ada = set()

        for nama in daftar_nama:

            nama = nama.strip()

            if not nama:
                continue

            if nama.lower().endswith(".docx"):
                nama = nama[:-5]

            nama = re.sub(karakter_terlarang, "", nama).strip()

            if not nama:
                continue

            if nama.lower() in nama_sudah_ada:
                continue

            nama_sudah_ada.add(nama.lower())
            nama_bersih.append(nama)

        if len(nama_bersih) == 0:

            st.error("Masukkan minimal satu nama Word.")

        else:

            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(
                zip_buffer,
                "w",
                zipfile.ZIP_DEFLATED
            ) as zip_file:

                for nama in nama_bersih:

                    word_buffer = io.BytesIO()

                    document = Document()
                    document.save(word_buffer)

                    zip_file.writestr(
                        f"{nama}.docx",
                        word_buffer.getvalue()
                    )

            st.success(
                f"{len(nama_bersih)} file Word berhasil dibuat."
            )

            st.download_button(
                label="⬇️ Download Word",
                data=zip_buffer.getvalue(),
                file_name="word_hasil.zip",
                mime="application/zip"
            )


elif pilihan == "Excel Sheet to PDF":

    st.subheader("📊 Excel Sheet to PDF")

    st.write("Upload file Excel dan pilih sheet yang ingin dijadikan PDF.")

    file_excel = st.file_uploader(
        "Pilih File Excel",
        type=["xlsx"]
    )

    if file_excel is not None:

        data_excel = file_excel.getvalue()

        workbook = load_workbook(
            io.BytesIO(data_excel),
            read_only=True
        )

        daftar_sheet = workbook.sheetnames
        workbook.close()

        sheet_terpilih = st.multiselect(
            "Pilih Sheet",
            daftar_sheet
        )

        if st.button("Convert ke PDF"):

            if not sheet_terpilih:
                st.warning("Pilih minimal satu sheet.")

            else:
                with tempfile.TemporaryDirectory() as folder:

                    file_input = os.path.join(folder, "input.xlsx")

                    with open(file_input, "wb") as f:
                        f.write(data_excel)

                    folder_pdf = os.path.join(folder, "hasil")
                    os.makedirs(folder_pdf)

                    zip_buffer = io.BytesIO()
                    berhasil = []
                    gagal = []

                    with zipfile.ZipFile(
                        zip_buffer,
                        "w",
                        zipfile.ZIP_DEFLATED
                    ) as zip_file:

                        for nomor, nama_sheet in enumerate(
                            sheet_terpilih, start=1
                        ):

                            try:
                                wb = load_workbook(
                                    io.BytesIO(data_excel)
                                )

                                for sheet in list(wb.sheetnames):
                                    if sheet != nama_sheet:
                                        del wb[sheet]

                                nama_aman = re.sub(
                                    karakter_terlarang,
                                    "",
                                    nama_sheet
                                ).strip() or f"Sheet_{nomor}"

                                nama_input = f"sheet_{nomor}.xlsx"

                                path_input = os.path.join(
                                    folder,
                                    nama_input
                                )

                                wb.save(path_input)
                                wb.close()

                                hasil = subprocess.run(
                                    [
                                        "libreoffice",
                                        "-env:UserInstallation=file://" + folder.replace("\\", "/") + f"/lo_{nomor}",
                                        "--headless",
                                        "--convert-to",
                                        "pdf",
                                        "--outdir",
                                        folder_pdf,
                                        path_input
                                    ],
                                    capture_output=True,
                                    text=True,
                                    timeout=120
                                )

                                path_pdf = os.path.join(
                                    folder_pdf,
                                    f"sheet_{nomor}.pdf"
                                )

                                if hasil.returncode != 0 or not os.path.exists(path_pdf):
                                    raise RuntimeError(
                                        hasil.stderr or hasil.stdout or "Konversi gagal."
                                    )

                                with open(path_pdf, "rb") as f:
                                    zip_file.writestr(
                                        f"{nama_aman}.pdf",
                                        f.read()
                                    )

                                berhasil.append(nama_sheet)

                            except Exception as e:
                                gagal.append(f"{nama_sheet}: {e}")

                    if berhasil:

                        st.success(
                            f"{len(berhasil)} sheet berhasil dikonversi."
                        )

                        st.download_button(
                            "⬇️ Download PDF",
                            data=zip_buffer.getvalue(),
                            file_name="excel_sheet_pdf.zip",
                            mime="application/zip"
                        )

                    if gagal:
                        st.error("Beberapa sheet gagal dikonversi.")
                        for pesan in gagal:
                            st.write(pesan)