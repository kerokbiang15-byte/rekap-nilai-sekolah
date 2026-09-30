import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import sqlite3
import streamlit as st


# --- KONEKSI DATABASE SQLITE & MIGRASI OTOMATIS ---
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
            poin REAL,
            status TEXT DEFAULT 'normal'
        )
    """)
  # Pastikan kolom status ada (migrasi otomatis jika database sudah ada sebelumnya)
  cursor.execute("PRAGMA table_info(hasil_ujian)")
  columns = [info[1] for info in cursor.fetchall()]
  if "status" not in columns:
    cursor.execute(
        "ALTER TABLE hasil_ujian ADD COLUMN status TEXT DEFAULT 'normal'"
    )

  # Tabel Data Remedial Siswa
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS remedial_siswa (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama_tes TEXT,
            nis TEXT,
            nama TEXT,
            kelas TEXT
        )
    """)
  conn.commit()
  return conn


conn = init_db()

st.set_page_config(page_title="APLIKASI REKAP NILAI UJIAN", layout="wide")

# --- KELOLA SESI LOGIN ---
if "logged_in" not in st.session_state:
  st.session_state.logged_in = False
  st.session_state.role = None

# Jika belum login, tampilkan halaman login
if not st.session_state.logged_in:
  st.title("🔐 Silakan Login ke Sistem")
  st.markdown("Masukkan akun akses sesuai peran Anda.")

  with st.form("login_form"):
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    submit = st.form_submit_button("Login")

    if submit:
      if username == "admin" and password == "123":
        st.session_state.logged_in = True
        st.session_state.role = "admin"
        st.success("Login berhasil sebagai Admin!")
        st.rerun()
      elif username == "guru" and password == "123":
        st.session_state.logged_in = True
        st.session_state.role = "guru"
        st.success("Login berhasil sebagai Guru Mapel!")
        st.rerun()
      else:
        st.error("Username atau Password salah!")

else:
  # Jika sudah login, tampilkan aplikasi utama
  st.sidebar.write(
      f"👤 Login sebagai: **{st.session_state.role.upper()}**"
  )
  if st.sidebar.button("🚪 Logout"):
    st.session_state.logged_in = False
    st.session_state.role = None
    st.rerun()

  st.title("APLIKASI REKAP NILAI UJIAN")
  st.markdown("---")

  # Atur menu berdasarkan role
  if st.session_state.role == "admin":
    menu = st.sidebar.selectbox(
        "Pilih Menu:",
        [
            "1. Operator: Upload & Kelola Data",
            "2. Guru: Download Rekap Nilai",
            "3. Cek Siswa Belum Ujian",
            "4. Cek & Download Peserta Remedial",
        ],
    )
  else:
    menu = st.sidebar.selectbox(
        "Pilih Menu:",
        [
            "2. Guru: Download Rekap Nilai",
            "3. Cek Siswa Belum Ujian",
            "4. Cek & Download Peserta Remedial",
        ],
    )

  # --- MENU 1: OPERATOR UPLOAD & KELOLA (Hanya Admin) ---
  if menu == "1. Operator: Upload & Kelola Data":
    st.header("Panel Operator: Upload & Kelola Database")

    tab_up1, tab_up2, tab_up3, tab_up4, tab_up5 = st.tabs([
        "Upload Master Siswa",
        "Upload Hasil Ujian CBT",
        "Upload Data Remedial",
        "Upload Update Nilai Remedial",
        "Kelola / Hapus Data",
    ])

    with tab_up1:
      st.subheader("1. Upload Data Master Siswa (Excel/CSV)")
      st.markdown("Pastikan file memiliki kolom: `NIS`, `Nama`, `Kelas`")
      file_master = st.file_uploader(
          "Pilih file Master Siswa", type=["xlsx", "csv"], key="master"
      )
      if file_master:
        if file_master.name.endswith(".csv"):
          df_m = pd.read_csv(file_master)
        else:
          df_m = pd.read_excel(file_master)

        df_m.columns = [str(col).strip() for col in df_m.columns]
        col_map = {c.lower(): c for c in df_m.columns}

        nis_col = col_map.get("nis") or col_map.get("no induk")
        nama_col = col_map.get("nama") or col_map.get("nama lengkap")
        kelas_col = col_map.get("kelas") or col_map.get("group")

        if not nis_col or not nama_col or not kelas_col:
          st.error(
              f"Kolom tidak lengkap! Kolom terbaca: {list(df_m.columns)}"
          )
        else:
          st.write("Pratinjau Data Master:", df_m.head())
          if st.button("Simpan Master Siswa ke Database"):
            cursor = conn.cursor()
            for _, row in df_m.iterrows():
              cursor.execute(
                  "INSERT OR REPLACE INTO master_siswa (nis, nama, kelas)"
                  " VALUES (?, ?, ?)",
                  (
                      str(row[nis_col]),
                      str(row[nama_col]),
                      str(row[kelas_col]),
                  ),
              )
            conn.commit()
            st.success("Data Master Siswa berhasil disimpan ke database!")

    with tab_up2:
      st.subheader("2. Upload Hasil Ujian Mentah dari CBT")
      st.markdown(
          "Format standar CBT: `Nama Tes`, `Username` (NIS), `Nama`, `Group`"
          " (Kelas), `Poin`"
      )
      file_ujian = st.file_uploader(
          "Pilih file Hasil Ujian", type=["xlsx", "csv"], key="ujian"
      )
      if file_ujian:
        if file_ujian.name.endswith(".csv"):
          df_u = pd.read_csv(file_ujian)
        else:
          df_u = pd.read_excel(file_ujian)

        df_u.columns = [str(col).strip() for col in df_u.columns]
        col_map_u = {c.lower(): c for c in df_u.columns}

        tes_c = col_map_u.get("nama tes") or col_map_u.get("namates")
        nis_c = col_map_u.get("username") or col_map_u.get("nis")
        nama_c = col_map_u.get("nama")
        kelas_c = col_map_u.get("group") or col_map_u.get("kelas")
        poin_c = col_map_u.get("poin") or col_map_u.get("nilai")

        if not tes_c or not nis_c or not nama_c or not kelas_c or not poin_c:
          st.error(
              f"Kolom file ujian tidak sesuai! Kolom terbaca: {list(df_u.columns)}"
          )
        else:
          st.write("Pratinjau Data Ujian:", df_u.head())
          if st.button("Simpan Hasil Ujian ke Database"):
            cursor = conn.cursor()
            for _, row in df_u.iterrows():
              cursor.execute(
                  "INSERT INTO hasil_ujian (nama_tes, nis, nama, kelas, poin,"
                  " status) VALUES (?, ?, ?, ?, ?, 'normal')",
                  (
                      str(row[tes_c]),
                      str(row[nis_c]),
                      str(row[nama_c]),
                      str(row[kelas_c]),
                      float(row[poin_c]),
                  ),
              )
            conn.commit()
            st.success("Data Hasil Ujian berhasil di-upload dan masuk database!")

    with tab_up3:
      st.subheader("3. Upload Data Siswa Remedial (Wali Kelas)")
      st.markdown(
          "Format kolom: `Nama Tes`, `Username` (NIS), `Nama`, `Group` (Kelas)"
      )
      file_rem = st.file_uploader(
          "Pilih file Data Remedial", type=["xlsx", "csv"], key="remedial_file"
      )
      if file_rem:
        if file_rem.name.endswith(".csv"):
          df_r = pd.read_csv(file_rem)
        else:
          df_r = pd.read_excel(file_rem)

        df_r.columns = [str(col).strip() for col in df_r.columns]
        col_map_r = {c.lower(): c for c in df_r.columns}

        tes_rc = col_map_r.get("nama tes") or col_map_r.get("namates")
        nis_rc = col_map_r.get("username") or col_map_r.get("nis")
        nama_rc = col_map_r.get("nama")
        kelas_rc = col_map_r.get("group") or col_map_r.get("kelas")

        if not tes_rc or not nis_rc or not nama_rc or not kelas_rc:
          st.error(
              f"Kolom file remedial tidak sesuai! Kolom terbaca:"
              f" {list(df_r.columns)}"
          )
        else:
          st.write("Pratinjau Data Remedial:", df_r.head())
          if st.button("Simpan Data Remedial ke Database"):
            cursor = conn.cursor()
            for _, row in df_r.iterrows():
              cursor.execute(
                  "INSERT INTO remedial_siswa (nama_tes, nis, nama, kelas)"
                  " VALUES (?, ?, ?, ?)",
                  (
                      str(row[tes_rc]),
                      str(row[nis_rc]),
                      str(row[nama_rc]),
                      str(row[kelas_rc]),
                  ),
              )
            conn.commit()
            st.success("Data Remedial berhasil di-upload ke database!")

    with tab_up4:
      st.subheader("4. Upload Update Nilai Remedial (Perbarui Nilai)")
      st.markdown(
          "Format kolom: `Nama Tes`, `Username` (NIS), `Poin` (Nilai Remedial"
          " Baru)"
      )
      file_up_rem = st.file_uploader(
          "Pilih file Update Nilai Remedial",
          type=["xlsx", "csv"],
          key="up_rem_score",
      )
      if file_up_rem:
        if file_up_rem.name.endswith(".csv"):
          df_up_r = pd.read_csv(file_up_rem)
        else:
          df_up_r = pd.read_excel(file_up_rem)

        df_up_r.columns = [str(col).strip() for col in df_up_r.columns]
        col_map_ur = {c.lower(): c for c in df_up_r.columns}

        tes_ur = col_map_ur.get("nama tes") or col_map_ur.get("namates")
        nis_ur = col_map_ur.get("username") or col_map_ur.get("nis")
        poin_ur = col_map_ur.get("poin") or col_map_ur.get("nilai")

        if not tes_ur or not nis_ur or not poin_ur:
          st.error(
              f"Kolom file tidak sesuai! Kolom terbaca: {list(df_up_r.columns)}"
          )
        else:
          st.write("Pratinjau Update Nilai:", df_up_r.head())
          if st.button("Perbarui Nilai Remedial"):
            cursor = conn.cursor()
            updated_count = 0
            for _, row in df_up_r.iterrows():
              cursor.execute(
                  """
                                UPDATE hasil_ujian 
                                SET poin = ?, status = 'remedial' 
                                WHERE nis = ? AND nama_tes = ?
                            """,
                  (float(row[poin_ur]), str(row[nis_ur]), str(row[tes_ur])),
              )
              updated_count += cursor.rowcount
            conn.commit()
            st.success(
                f"Berhasil memperbarui {updated_count} nilai siswa remedial!"
                " Nilai yang diperbarui akan ditandai font hijau saat"
                " didownload."
            )

    with tab_up5:
      st.subheader("5. Kelola & Hapus Data di Database (Multiple Choice)")

      st.markdown("##### 🗑️ Hapus Data Hasil Ujian")
      df_mapel_del = pd.read_sql(
          "SELECT DISTINCT nama_tes FROM hasil_ujian", conn
      )
      if df_mapel_del.empty:
        st.info("Belum ada data hasil ujian di dalam database.")
      else:
        pilih_hapus_tes_list = st.multiselect(
            "Pilih Mata Pelajaran Ujian yang ingin dihapus:",
            df_mapel_del["nama_tes"].tolist(),
            key="del_ujian_multiselect",
        )
        if st.button("🗑️ Hapus Data Ujian yang Dipilih", type="primary"):
          if pilih_hapus_tes_list:
            cursor = conn.cursor()
            for tes in pilih_hapus_tes_list:
              cursor.execute(
                  "DELETE FROM hasil_ujian WHERE nama_tes = ?", (tes,)
              )
            conn.commit()
            st.success(
                "Data ujian untuk mata pelajaran yang dipilih berhasil"
                " dihapus!"
            )
            st.rerun()
          else:
            st.warning("Pilih minimal satu mata pelajaran ujian.")

      st.markdown("---")
      st.markdown("##### 🗑️️ Hapus Data Remedial")
      df_mapel_rem_del = pd.read_sql(
          "SELECT DISTINCT nama_tes FROM remedial_siswa", conn
      )
      if df_mapel_rem_del.empty:
        st.info("Belum ada data remedial di dalam database.")
      else:
        pilih_hapus_rem_list = st.multiselect(
            "Pilih Mata Pelajaran Remedial yang ingin dihapus:",
            df_mapel_rem_del["nama_tes"].tolist(),
            key="del_remedial_multiselect",
        )
        if st.button("🗑️ Hapus Data Remedial yang Dipilih", type="primary"):
          if pilih_hapus_rem_list:
            cursor = conn.cursor()
            for tes in pilih_hapus_rem_list:
              cursor.execute(
                  "DELETE FROM remedial_siswa WHERE nama_tes = ?", (tes,)
              )
            conn.commit()
            st.success(
                "Data remedial untuk mata pelajaran yang dipilih berhasil"
                " dihapus!"
            )
            st.rerun()
          else:
            st.warning("Pilih minimal satu mata pelajaran remedial.")

      st.markdown("---")
      col_r1, col_r2, col_r3 = st.columns(3)
      with col_r1:
        if st.button("⚠️ Kosongkan SEMUA Hasil Ujian"):
          cursor = conn.cursor()
          cursor.execute("DELETE FROM hasil_ujian")
          conn.commit()
          st.warning("Semua data hasil ujian dikosongkan!")
          st.rerun()
      with col_r2:
        if st.button("⚠️ Kosongkan SEMUA Data Master"):
          cursor = conn.cursor()
          cursor.execute("DELETE FROM master_siswa")
          conn.commit()
          st.warning("Semua data master siswa dikosongkan!")
          st.rerun()
      with col_r3:
        if st.button("⚠️ Kosongkan SEMUA Data Remedial"):
          cursor = conn.cursor()
          cursor.execute("DELETE FROM remedial_siswa")
          conn.commit()
          st.warning("Semua data remedial dikosongkan!")
          st.rerun()

  # --- MENU 2: GURU PREVIEW & DOWNLOAD REKAP ---
  elif menu == "2. Guru: Download Rekap Nilai":
    st.header("Dashboard Guru: Pratinjau & Download Rekap Nilai Per Kelas")

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

      if pilih_kelas:
        st.markdown("---")
        st.subheader("👀 Pratinjau (Preview) Data Nilai")

        preview_data = {}

        for kelas in pilih_kelas:
          st.markdown(f"**Kelas: {kelas}**")
          df_m_kelas = pd.read_sql(
              "SELECT nis, nama FROM master_siswa WHERE kelas = ?",
              conn,
              params=(kelas,),
          )
          df_n_mapel = pd.read_sql(
              "SELECT nis, poin, status FROM hasil_ujian WHERE nama_tes = ?",
              conn,
              params=(pilih_mapel,),
          )

          dict_nilai = dict(
              zip(
                  df_n_mapel["nis"].astype(str),
                  zip(df_n_mapel["poin"], df_n_mapel["status"]),
              )
          )

          rows = []
          for idx, (_, s_row) in enumerate(df_m_kelas.iterrows(), 1):
            nis_siswa = str(s_row["nis"])
            nama_siswa = s_row["nama"]
            data_n = dict_nilai.get(nis_siswa, (0.0, "normal"))
            nilai_pg = data_n[0]
            status_n = data_n[1]

            ket = "Remedial" if status_n == "remedial" else "-"
            rows.append({
                "NO": idx,
                "NAMA": nama_siswa,
                "NILAI PG": nilai_pg,
                "KET": ket,
            })

          df_preview = pd.DataFrame(rows)
          if df_preview.empty:
            st.info(f"Tidak ada data siswa untuk kelas {kelas}.")
          else:
            st.dataframe(df_preview, use_container_width=True, hide_index=True)

          preview_data[kelas] = df_preview

        st.markdown("---")
        if st.button("📥 Generate & Download Excel (Multi-Sheet)"):
          output = io.BytesIO()
          wb = openpyxl.Workbook()
          default_sheet = wb.active
          wb.remove(default_sheet)

          font_title = Font(name="Calibri", size=11, bold=True)
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

          df_n_mapel_all = pd.read_sql(
              "SELECT nis, status FROM hasil_ujian WHERE nama_tes = ?",
              conn,
              params=(pilih_mapel,),
          )
          dict_status = dict(
              zip(
                  df_n_mapel_all["nis"].astype(str),
                  df_n_mapel_all["status"],
              )
          )

          for kelas, df_prev in preview_data.items():
            safe_sheet_name = re.sub(r"[\\/?:*\[\]]", "-", str(kelas))[:31]
            ws = wb.create_sheet(title=safe_sheet_name)

            ws.merge_cells("A1:C1")
            ws["A1"] = f"MATA PELAJARAN : {pilih_mapel}"
            ws["A1"].font = font_title
            ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

            ws.merge_cells("A2:C2")
            ws["A2"] = f"KELAS : {kelas}"
            ws["A2"].font = font_title
            ws["A2"].alignment = Alignment(horizontal="left", vertical="center")

            headers = ["NO", "NAMA", "NILAI PG"]
            for col_num, h_title in enumerate(headers, 1):
              cell = ws.cell(row=4, column=col_num, value=h_title)
              cell.font = font_bold
              cell.fill = fill_header
              cell.alignment = Alignment(horizontal="center", vertical="center")
              cell.border = thin_border

            df_m_kls_full = pd.read_sql(
                "SELECT nis FROM master_siswa WHERE kelas = ?",
                conn,
                params=(kelas,),
            )
            dict_nis_by_idx = dict(
                enumerate(df_m_kls_full["nis"].astype(str), 1)
            )

            for _, r_data in df_prev.iterrows():
              r_idx = int(r_data["NO"])
              row_num = 4 + r_idx

              c_no = ws.cell(row=row_num, column=1, value=r_idx)
              c_nama = ws.cell(row=row_num, column=2, value=r_data["NAMA"])
              c_nilai = ws.cell(row=row_num, column=3, value=r_data["NILAI PG"])

              nis_s = dict_nis_by_idx.get(r_idx, "")
              is_remed = dict_status.get(nis_s) == "remedial"

              if is_remed:
                c_nilai.font = Font(
                    name="Calibri", size=11, color="008000", bold=True
                )
              else:
                c_nilai.font = Font(name="Calibri", size=11)

              c_no.alignment = Alignment(
                  horizontal="center", vertical="center"
              )
              c_nama.alignment = Alignment(horizontal="left", vertical="center")
              c_nilai.alignment = Alignment(
                  horizontal="center", vertical="center"
              )

              c_no.border = thin_border
              c_nama.border = thin_border
              c_nilai.border = thin_border

            ws.column_dimensions["A"].width = 8
            ws.column_dimensions["B"].width = 38
            ws.column_dimensions["C"].width = 16

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

      match_grade = re.search(r"\b([789])\b", pilih_mapel)
      if match_grade:
        grade_num = match_grade.group(1)
        query_belum = f"""
                    SELECT m.kelas, m.nis, m.nama 
                    FROM master_siswa m
                    WHERE m.kelas LIKE ? AND m.nis NOT IN (
                        SELECT h.nis FROM hasil_ujian h WHERE h.nama_tes = ?
                    )
                    ORDER BY m.kelas, m.nama
                """
        df_belum = pd.read_sql(
            query_belum, conn, params=(f"%{grade_num}%", pilih_mapel)
        )
      else:
        query_belum = """
                    SELECT m.kelas, m.nis, m.nama 
                    FROM master_siswa m
                    WHERE m.nis NOT IN (
                        SELECT h.nis FROM hasil_ujian h WHERE h.nama_tes = ?
                    )
                    ORDER BY m.kelas, m.nama
                """
        df_belum = pd.read_sql(query_belum, conn, params=(pilih_mapel,))

      if df_belum.empty:
        st.info("Hebat! Semua siswa sudah mengikuti ujian untuk mata pelajaran ini.")
      else:
        st.warning(
            f"Ditemukan {len(df_belum)} siswa yang belum mengikuti ujian."
        )
        st.dataframe(df_belum, use_container_width=True)

  # --- MENU 4: CEK & DOWNLOAD PESERTA REMEDIAL ---
  elif menu == "4. Cek & Download Peserta Remedial":
    st.header("Pelacak & Rekap Peserta Didik Remedial")

    df_mapel_rem = pd.read_sql(
        "SELECT DISTINCT nama_tes FROM remedial_siswa", conn
    )
    if df_mapel_rem.empty:
      st.warning("Belum ada data remedial yang di-upload oleh operator.")
    else:
      pilih_mapel_rem = st.selectbox(
          "Pilih Mata Pelajaran Remedial:", df_mapel_rem["nama_tes"].tolist()
      )

      df_kelas_rem = pd.read_sql(
          "SELECT DISTINCT kelas FROM remedial_siswa WHERE nama_tes = ?",
          conn,
          params=(pilih_mapel_rem,),
      )
      list_kelas_rem = df_kelas_rem["kelas"].tolist()

      select_all_rem = st.checkbox(
          "Pilih Semua Kelas Remedial", key="select_all_rem_chk"
      )
      if select_all_rem:
        pilih_kelas_rem = st.multiselect(
            "Pilih Kelas untuk melihat/download daftar remedial:",
            list_kelas_rem,
            default=list_kelas_rem,
            key="multiselect_kelas_rem",
        )
      else:
        pilih_kelas_rem = st.multiselect(
            "Pilih Kelas untuk melihat/download daftar remedial:",
            list_kelas_rem,
            key="multiselect_kelas_rem",
        )

      if pilih_kelas_rem:
        placeholders = ",".join(["?"] * len(pilih_kelas_rem))
        query_r = f"""
                SELECT kelas, nis, nama 
                FROM remedial_siswa 
                WHERE nama_tes = ? AND kelas IN ({placeholders})
                ORDER BY kelas, nama
            """
        params = [pilih_mapel_rem] + pilih_kelas_rem
        df_hasil_rem = pd.read_sql(query_r, conn, params=params)

        if df_hasil_rem.empty:
          st.info(
              "Tidak ada peserta remedial untuk kelas yang dipilih pada mata"
              " pelajaran ini."
          )
        else:
          st.warning(
              f"Ditemukan {len(df_hasil_rem)} siswa yang harus mengikuti"
              " remedial."
          )
          st.dataframe(df_hasil_rem, use_container_width=True)

          output_rem = io.BytesIO()
          with pd.ExcelWriter(output_rem, engine="openpyxl") as writer:
            df_hasil_rem.to_excel(
                writer, sheet_name="Peserta Remedial", index=False
            )
          output_rem.seek(0)

          st.download_button(
              label="📥 Download Daftar Remedial ke Excel",
              data=output_rem,
              file_name=(
                  f"Daftar_Remedial_{pilih_mapel_rem.replace('/', '-')}.xlsx"
              ),
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
          )
