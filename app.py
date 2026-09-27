import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import sqlite3
import streamlit as st


# --- KONEKSI DATABASE SQLITE ---
def init_db():
  conn = sqlite3.connect("sekolah.db", check_same_thread=False)
  cursor = conn.cursor()
  # Tabel Master Siswa
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS master_siswa (
            nis TEXT PRIMARY KEY,
            nama TEXT,
            kelas TEXT
        )
    """)
  # Tabel Hasil Ujian Mentah
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS hasil_ujian (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama_tes TEXT,
            nis TEXT,
            nama TEXT,
            kelas TEXT,
            poin REAL
        )
    """)
  conn.commit()
  return conn


conn = init_db()

st.set_page_config(page_title="Aplikasi Rekap Nilai Ujian", layout="wide")
st.title("📚 Aplikasi Rekap Nilai Ujian & Pelacak Siswa")
st.markdown("---")

menu = st.sidebar.selectbox(
    "Pilih Menu:",
    [
        "1. Operator: Upload & Kelola Data",
        "2. Guru: Download Rekap Nilai",
        "3. Cek Siswa Belum Ujian",
    ],
)

# --- MENU 1: OPERATOR UPLOAD & KELOLA ---
if menu == "1. Operator: Upload & Kelola Data":
  st.header("Panel Operator: Upload & Kelola Database")

  tab_up1, tab_up2, tab_up3 = st.tabs([
      "Upload Data Master Siswa",
      "Upload Hasil Ujian CBT",
      "Kelola / Hapus Data",
  ])

  with tab_up1:
    st.subheader("1. Upload Data Master Siswa (Excel/CSV)")
    st.markdown("Format kolom wajib: `NIS`, `Nama`, `Kelas`")
    file_master = st.file_uploader(
        "Pilih file Master Siswa", type=["xlsx", "csv"], key="master"
    )
    if file_master:
      if file_master.name.endswith(".csv"):
        df_m = pd.read_csv(file_master)
      else:
        df_m = pd.read_excel(file_master)

      st.write("Pratinjau Data Master:", df_m.head())
      if st.button("Simpan Master Siswa ke Database"):
        cursor = conn.cursor()
        for _, row in df_m.iterrows():
          cursor.execute(
              "INSERT OR REPLACE INTO master_siswa (nis, nama, kelas) VALUES"
              " (?, ?, ?)",
              (str(row["NIS"]), str(row["Nama"]), str(row["Kelas"])),
          )
        conn.commit()
        st.success("Data Master Siswa berhasil disimpan ke database!")

  with tab_up2:
    st.subheader("2. Upload Hasil Ujian Mentah dari CBT")
    st.markdown(
        "Format kolom dari CBT: `Nama Tes`, `Username` (sebagai NIS), `Nama`,"
        " `Group` (Kelas), `Poin`"
    )
    file_ujian = st.file_uploader(
        "Pilih file Hasil Ujian", type=["xlsx", "csv"], key="ujian"
    )
    if file_ujian:
      if file_ujian.name.endswith(".csv"):
        df_u = pd.read_csv(file_ujian)
      else:
        df_u = pd.read_excel(file_ujian)

      st.write("Pratinjau Data Ujian:", df_u.head())
      if st.button("Simpan Hasil Ujian ke Database"):
        cursor = conn.cursor()
        for _, row in df_u.iterrows():
          cursor.execute(
              "INSERT INTO hasil_ujian (nama_tes, nis, nama, kelas, poin)"
              " VALUES (?, ?, ?, ?, ?)",
              (
                  str(row["Nama Tes"]),
                  str(row["Username"]),
                  str(row["Nama"]),
                  str(row["Group"]),
                  float(row["Poin"]),
              ),
          )
        conn.commit()
        st.success("Data Hasil Ujian berhasil di-upload dan masuk database!")

  with tab_up3:
    st.subheader("3. Kelola & Hapus Data di Database")
    st.markdown(
        "Gunakan fitur ini jika ada kesalahan upload atau ingin menghapus data"
        " ujian tertentu."
    )

    # Hapus berdasarkan Nama Tes (Mapel) tertentu
    df_mapel_del = pd.read_sql(
        "SELECT DISTINCT nama_tes FROM hasil_ujian", conn
    )
    if df_mapel_del.empty:
        st.info("Belum ada data hasil ujian di dalam database.")
    else:
      pilih_hapus_tes = st.selectbox(
          "Pilih Nama Tes / Mata Pelajaran yang ingin dihapus datanya:",
          df_mapel_del["nama_tes"].tolist(),
      )
      if st.button("🗑️ Hapus Data Tes Ini Saja", type="primary"):
        cursor = conn.cursor()
        cursor.execute(
            "DELETE FROM hasil_ujian WHERE nama_tes = ?", (pilih_hapus_tes,)
        )
        conn.commit()
        st.success(
            f"Data ujian untuk '{pilih_hapus_tes}' berhasil dihapus dari"
            " database!"
        )
        st.rerun()

    st.markdown("---")
    # Reset total
    col_r1, col_r2 = st.columns(2)
    with col_r1:
      if st.button("⚠️ Kosongkan SEMUA Hasil Ujian"):
        cursor = conn.cursor()
        cursor.execute("DELETE FROM hasil_ujian")
        conn.commit()
        st.warning("Semua data hasil ujian telah dikosongkan!")
        st.rerun()
    with col_r2:
      if st.button("⚠️ Kosongkan SEMUA Data Master Siswa"):
        cursor = conn.cursor()
        cursor.execute("DELETE FROM master_siswa")
        conn.commit()
        st.warning("Semua data master siswa telah dikosongkan!")
        st.rerun()

# --- MENU 2: GURU DOWNLOAD REKAP ---
elif menu == "2. Guru: Download Rekap Nilai":
  st.header("Dashboard Guru: Download Rekap Nilai Per Kelas")

  df_mapel = pd.read_sql("SELECT DISTINCT nama_tes FROM hasil_ujian", conn)
  df_kelas = pd.read_sql("SELECT DISTINCT kelas FROM master_siswa", conn)

  if df_mapel.empty or df_kelas.empty:
    st.warning(
        "Data belum lengkap. Operator harus meng-upload Data Master Siswa dan"
        " Hasil Ujian terlebih dahulu."
    )
  else:
    pilih_mapel = st.selectbox(
        "Pilih Mata Pelajaran / Nama Tes:", df_mapel["nama_tes"].tolist()
    )

    list_kelas_tersedia = df_kelas["kelas"].tolist()
    pilih_kelas = st.multiselect(
        "Pilih Kelas (Bisa pilih lebih dari 1 kelas sekaligus):",
        list_kelas_tersedia,
    )

    if st.button("Generate & Download Excel (Multi-Sheet)") and pilih_kelas:
      output = io.BytesIO()
      wb = openpyxl.Workbook()
      default_sheet = wb.active
      wb.remove(default_sheet)

      font_title = Font(name="Calibri", size=12, bold=True)
      font_bold = Font(name="Calibri", size=11, bold=True)
      fill_header = PatternFill(
          start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"
      )
      thin_border = Border(
          left=Side(style="thin", color="000000"),
          right=Side(style="thin", color="000000"),
          top=Side(style="thin", color="000000"),
          bottom=Side(style="thin", color="000000"),
      )

      for kelas in pilih_kelas:
        ws = wb.create_sheet(title=str(kelas))

        ws["A1"] = "MATA PELAJARAN:"
        ws["B1"] = pilih_mapel
        ws["A1"].font = font_title
        ws["B1"].font = font_title

        ws["A2"] = "KELAS:"
        ws["B2"] = kelas
        ws["A2"].font = font_title
        ws["B2"].font = font_title

        headers = ["NO", "NAMA", "NILAI PG"]
        for col_num, h_title in enumerate(headers, 1):
          cell = ws.cell(row=4, column=col_num, value=h_title)
          cell.font = font_bold
          cell.fill = fill_header
          cell.alignment = Alignment(horizontal="center", vertical="center")
          cell.border = thin_border

        query_siswa = f"SELECT nis, nama FROM master_siswa WHERE kelas = '{kelas}'"
        df_m_kelas = pd.read_sql(query_siswa, conn)

        query_nilai = (
            f"SELECT nis, poin FROM hasil_ujian WHERE nama_tes ="
            f" '{pilih_mapel}'"
        )
        df_n_mapel = pd.read_sql(query_nilai, conn)

        dict_nilai = dict(
            zip(df_n_mapel["nis"].astype(str), df_n_mapel["poin"])
        )

        for idx, (_, s_row) in enumerate(df_m_kelas.iterrows(), 1):
          row_num = 4 + idx
          nis_siswa = str(s_row["nis"])
          nama_siswa = s_row["nama"]
          nilai_pg = dict_nilai.get(nis_siswa, 0.0)

          c_no = ws.cell(row=row_num, column=1, value=idx)
          c_nama = ws.cell(row=row_num, column=2, value=nama_siswa)
          c_nilai = ws.cell(row=row_num, column=3, value=nilai_pg)

          c_no.alignment = Alignment(horizontal="center")
          c_nilai.alignment = Alignment(horizontal="center")

          c_no.border = thin_border
          c_nama.border = thin_border
          c_nilai.border = thin_border

        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 35
        ws.column_dimensions["C"].width = 15

      wb.save(output)
      output.seek(0)

      st.success("File rekap multi-kelas berhasil dibuat!")
      st.download_button(
          label="📥 Download File Excel Rekap (Multi-Sheet)",
          data=output,
          file_name=f"Rekap_Nilai_{pilih_mapel}.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )

# --- MENU 3: CEK SISWA BELUM UJIAN ---
elif menu == "3. Cek Siswa Belum Ujian":
  st.header("Pelacak Siswa Belum Mengikuti Ujian")

  df_mapel = pd.read_sql("SELECT DISTINCT nama_tes FROM hasil_ujian", conn)
  if df_mapel.empty:
    st.warning("Belum ada data ujian yang di-upload.")
  else:
    pilih_mapel = st.selectbox(
        "Pilih Mata Pelajaran:",
        df_mapel["nama_tes"].tolist(),
        key="cek_mapel",
    )

    query_belum = f"""
            SELECT m.kelas, m.nis, m.nama 
            FROM master_siswa m
            WHERE m.nis NOT IN (
                SELECT h.nis FROM hasil_ujian h WHERE h.nama_tes = '{pilih_mapel}'
            )
            ORDER BY m.kelas, m.nama
        """
    df_belum = pd.read_sql(query_belum, conn)

    if df_belum.empty:
      st.info("Hebat! Semua siswa sudah mengikuti ujian untuk mata pelajaran ini.")
    else:
      st.warning(
          f"Ditemukan {len(df_belum)} siswa yang belum mengikuti ujian."
      )
      st.dataframe(df_belum, use_container_width=True)
