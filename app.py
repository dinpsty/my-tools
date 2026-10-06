import streamlit as st
import io
import zipfile
import re

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
        "PDF Maker"
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