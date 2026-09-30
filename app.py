import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import sqlite3
import streamlit as st


# --- FUNGSI BANTU PENCOCOKAN & PARSING ---
def clean_tes_name(name):
  if not name:
    return ""
  cleaned = re.sub(
      r"\b(sts|pas|pts|uh|us|tp|uas|uts)\b", "", str(name), flags=re.IGNORECASE
  )
  return re.sub(r"\s+", " ", cleaned).strip().lower()


def parse_poin(val):
  if pd.isna(val):
    return 0.0
  if isinstance(val, (int, float)):
    return float(val)
  val_str = str(val).strip().replace(",", ".")
  try:
    return float(val_str)
  except:
    return 0.0


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
  st.sidebar.write(
      f"👤 Login sebagai: **{st.session_state.role.upper()}**"
  )
  if st.sidebar.button("🚪 Logout"):
    st.session_state.logged_in = False
    st.session_state.role = None
    st.rerun()

  st.title("APLIKASI REKAP NILAI UJIAN")
  st.markdown("---")

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

  # --- MENU 1: OPERATOR UPLOAD & KELOLA ---
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
      file_master = st.file_uploader(
          "Pilih file Master Siswa", type=["xlsx", "csv"], key="master"
      )
      if file_master:
        df_m = (
            pd.read_csv(file_master)
            if file_master.name.endswith(".csv")
            else pd.read_excel(file_master)
        )
        df_m.columns = [str(col).strip() for col in df_m.columns]
        col_map = {c.lower(): c for c in df_m.columns}
        nis_col = col_map.get("nis") or col_map.get("no induk")
        nama_col = col_map.get("nama") or col_map.get("nama lengkap")
        kelas_col = col_map.get("kelas") or col_map.get("group")

        if nis_col and nama_col and kelas_col:
          st.markdown("**Pratinjau Data Master:**")
          st.dataframe(
              df_m.head(), hide_index=True, use_container_width=True
          )
          if st.button("Simpan Master Siswa ke Database"):
            cursor = conn.cursor()
            for _, row in df_m.iterrows():
              cursor.execute(
                  "INSERT OR REPLACE INTO master_siswa (nis, nama, kelas)"
                  " VALUES (?, ?, ?)",
                  (
                      str(row[nis_col]).strip(),
                      str(row[nama_col]).strip(),
                      str(row[kelas_col]).strip(),
                  ),
              )
            conn.commit()
            st.success("Data Master Siswa berhasil disimpan!")

    with tab_up2:
      st.subheader("2. Upload Hasil Ujian CBT (Utama / Susulan)")
      jenis_upload = st.radio(
          "Pilih Jenis Upload File Ujian:",
          [
              "Ujian Utama CBT (Normal)",
              "Ujian Susulan / Susulan & Remedial",
          ],
      )

      file_ujian = st.file_uploader(
          "Pilih file Hasil Ujian", type=["xlsx", "csv"], key="ujian"
      )
      if file_ujian:
        df_u = (
            pd.read_csv(file_ujian)
            if file_ujian.name.endswith(".csv")
            else pd.read_excel(file_ujian)
        )
        df_u.columns = [str(col).strip() for col in df_u.columns]
        col_map_u = {c.lower(): c for c in df_u.columns}
        tes_c = col_map_u.get("nama tes") or col_map_u.get("namates")
        nis_c = col_map_u.get("username") or col_map_u.get("nis")
        nama_c = col_map_u.get("nama")
        kelas_c = col_map_u.get("group") or col_map_u.get("kelas")
        poin_c = col_map_u.get("poin") or col_map_u.get("nilai")

        if tes_c and nis_c and nama_c and kelas_c and poin_c:
          st.markdown("**Pratinjau Data Ujian:**")
          st.dataframe(
              df_u.head(), hide_index=True, use_container_width=True
          )
          if st.button("Simpan Hasil Ujian ke Database"):
            cursor = conn.cursor()
            for _, row in df_u.iterrows():
              nis_val = str(row[nis_c]).strip()
              nama_tes_val = str(row[tes_c]).strip()
              clean_upload_tes = clean_tes_name(nama_tes_val)
              poin_val = parse_poin(row[poin_c])

              cursor.execute(
                  "SELECT kelas, nama FROM master_siswa WHERE nis = ?",
                  (nis_val,),
              )
              res_m = cursor.fetchone()
              kelas_val = res_m[0] if res_m else str(row[kelas_c]).strip()
              nama_val = res_m[1] if res_m else str(row[nama_c]).strip()

              if "Utama" in jenis_upload:
                status_ujian = "normal"
              else:
                # Cek apakah siswa terdaftar di remedial
                cursor.execute(
                    "SELECT nama_tes FROM remedial_siswa WHERE nis = ?",
                    (nis_val,),
                )
                all_rem = cursor.fetchall()
                is_remedial = False
                for (r_tes,) in all_rem:
                  clean_r_tes = clean_tes_name(r_tes)
                  if (
                      clean_r_tes == clean_upload_tes
                      or clean_r_tes in clean_upload_tes
                      or clean_upload_tes in clean_r_tes
                  ):
                    is_remedial = True
                    break

                if is_remedial:
                  status_ujian = "remedial"
                else:
                  status_ujian = "susulan"

              cursor.execute(
                  """
                                INSERT INTO hasil_ujian (nama_tes, nis, nama, kelas, poin, status)
                                VALUES (?, ?, ?, ?, ?, ?)
                            """,
                  (
                      nama_tes_val,
                      nis_val,
                      nama_val,
                      kelas_val,
                      poin_val,
                      status_ujian,
                  ),
              )
            conn.commit()
            st.success("Data Hasil Ujian berhasil di-upload dan disimpan!")

    with tab_up3:
      st.subheader("3. Upload Data Siswa Remedial dari Guru")
      file_rem = st.file_uploader(
          "Pilih file Data Remedial", type=["xlsx", "csv"], key="remedial_file"
      )
      if file_rem:
        df_r = (
            pd.read_csv(file_rem)
            if file_rem.name.endswith(".csv")
            else pd.read_excel(file_rem)
        )
        df_r.columns = [str(col).strip() for col in df_r.columns]
        col_map_r = {c.lower(): c for c in df_r.columns}
        tes_rc = col_map_r.get("nama tes") or col_map_r.get("namates")
        nis_rc = col_map_r.get("username") or col_map_r.get("nis")
        nama_rc = col_map_r.get("nama")
        kelas_rc = col_map_r.get("group") or col_map_r.get("kelas")

        if tes_rc and nis_rc and nama_rc:
          st.markdown("**Pratinjau Data Remedial:**")
          st.dataframe(
              df_r.head(), hide_index=True, use_container_width=True
          )
          if st.button("Simpan Data Remedial ke Database"):
            cursor = conn.cursor()
            for _, row in df_r.iterrows():
              nis_val = str(row[nis_rc]).strip()
              cursor.execute(
                  "SELECT kelas, nama FROM master_siswa WHERE nis = ?",
                  (nis_val,),
              )
              res_m = cursor.fetchone()
              if res_m:
                kelas_val, nama_val = res_m[0], res_m[1]
              else:
                kelas_val = (
                    str(row[kelas_rc]).strip()
                    if kelas_rc and pd.notna(row[kelas_rc])
                    else "-"
                )
                nama_val = str(row[nama_rc]).strip()

              cursor.execute(
                  "INSERT INTO remedial_siswa (nama_tes, nis, nama, kelas)"
                  " VALUES (?, ?, ?, ?)",
                  (
                      str(row[tes_rc]).strip(),
                      nis_val,
                      nama_val,
                      kelas_val,
                  ),
              )
            conn.commit()
            st.success("Data Remedial berhasil disimpan!")

    with tab_up4:
      st.subheader("4. Upload Update Nilai Remedial")
      file_up_rem = st.file_uploader(
          "Pilih file Update Nilai Remedial",
          type=["xlsx", "csv"],
          key="up_rem_score",
      )
      if file_up_rem:
        df_up_r = (
            pd.read_csv(file_up_rem)
            if file_up_rem.name.endswith(".csv")
            else pd.read_excel(file_up_rem)
        )
        df_up_r.columns = [str(col).strip() for col in df_up_r.columns]
        col_map_ur = {c.lower(): c for c in df_up_r.columns}
        tes_ur = col_map_ur.get("nama tes") or col_map_ur.get("namates")
        nis_ur = col_map_ur.get("username") or col_map_ur.get("nis")
        poin_ur = col_map_ur.get("poin") or col_map_ur.get("nilai")

        if tes_ur and nis_ur and poin_ur:
          st.markdown("**Pratinjau Update Nilai:**")
          st.dataframe(
              df_up_r.head(), hide_index=True, use_container_width=True
          )
          if st.button("Perbarui Nilai Remedial"):
            cursor = conn.cursor()
            updated_count = 0
            for _, row in df_up_r.iterrows():
              poin_val = parse_poin(row[poin_ur])
              cursor.execute(
                  """
                                UPDATE hasil_ujian 
                                SET poin = ?, status = 'remedial' 
                                WHERE nis = ? 
                            """,
                  (poin_val, str(row[nis_ur]).strip()),
              )
              updated_count += cursor.rowcount
            conn.commit()
            st.success(
                f"Berhasil memperbarui {updated_count} nilai siswa remedial"
                " menjadi 'Sudah Remedial'!"
            )

    with tab_up5:
      st.subheader("5. Kelola & Hapus Data di Database (Multiple Choice & Edit)")

      st.markdown("##### 🗑️ Hapus Data Hasil Ujian (Berdasarkan Mata Pelajaran)")
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
      st.markdown(
          "##### 🗑️ Hapus Data Ujian Berdasarkan Kelas Tertentu (Fitur Edit"
          " Kelas)"
      )
      if df_mapel_del.empty:
        st.info("Belum ada data ujian.")
      else:
        pilih_tes_edit_kelas = st.selectbox(
            "Pilih Mata Pelajaran untuk Edit/Hapus Kelas:",
            df_mapel_del["nama_tes"].tolist(),
            key="edit_kelas_tes_sel",
        )
        df_kelas_edit = pd.read_sql(
            "SELECT DISTINCT kelas FROM hasil_ujian WHERE nama_tes = ?",
            conn,
            params=(pilih_tes_edit_kelas,),
        )
        list_kelas_edit = df_kelas_edit["kelas"].tolist()

        pilih_kelas_hapus_list = st.multiselect(
            f"Pilih Kelas pada '{pilih_tes_edit_kelas}' yang ingin dihapus"
            " datanya:",
            list_kelas_edit,
            key="edit_kelas_multiselect",
        )
        if st.button(
            "🗑️️ Hapus Kelas Terpilih dari Ujian Ini", type="primary"
        ):
          if pilih_kelas_hapus_list:
            cursor = conn.cursor()
            for kls in pilih_kelas_hapus_list:
              cursor.execute(
                  "DELETE FROM hasil_ujian WHERE nama_tes = ? AND kelas = ?",
                  (pilih_tes_edit_kelas, kls),
              )
            conn.commit()
            st.success(
                "Data kelas yang dipilih berhasil dihapus dari mata pelajaran"
                f" '{pilih_tes_edit_kelas}'!"
            )
            st.rerun()
          else:
            st.warning("Pilih minimal satu kelas yang ingin dihapus.")

      st.markdown("---")
      st.markdown(
          "##### 🗑️ Hapus Peserta Remedial / Susulan Secara Spesifik (Fitur"
          " Edit Siswa)"
      )
      df_mapel_rem_del = pd.read_sql(
          "SELECT DISTINCT nama_tes FROM remedial_siswa", conn
      )
      if df_mapel_rem_del.empty:
        st.info("Belum ada data remedial/susulan di dalam database.")
      else:
        pilih_tes_rem_edit = st.selectbox(
            "Pilih Mata Pelajaran Remedial untuk Edit/Hapus Peserta:",
            df_mapel_rem_del["nama_tes"].tolist(),
            key="edit_rem_tes_sel",
        )
        df_rem_siswa_list = pd.read_sql(
            "SELECT id, nis, nama, kelas FROM remedial_siswa WHERE nama_tes = ?"
            " ORDER BY kelas, nama",
            conn,
            params=(pilih_tes_rem_edit,),
        )

        if df_rem_siswa_list.empty:
          st.info(
              "Tidak ada data peserta remedial untuk mata pelajaran ini."
          )
        else:
          st.dataframe(
              df_rem_siswa_list, hide_index=True, use_container_width=True
          )
          pilih_id_hapus = st.multiselect(
              "Pilih ID atau Nama Siswa yang ingin dihapus dari daftar"
              " remedial:",
              options=df_rem_siswa_list["id"].tolist(),
              format_func=lambda x: f"ID {x} - "
              + f"{df_rem_siswa_list.loc[df_rem_siswa_list['id'] == x, 'nama'].values[0]}"
              + f" ({df_rem_siswa_list.loc[df_rem_siswa_list['id'] == x, 'kelas'].values[0]})",
              key="multiselect_id_hapus_rem",
          )
          if st.button(
              "🗑️ Hapus Siswa Terpilih dari Daftar Remedial", type="primary"
          ):
            if pilih_id_hapus:
              cursor = conn.cursor()
              for sid in pilih_id_hapus:
                cursor.execute(
                    "DELETE FROM remedial_siswa WHERE id = ?", (sid,)
                )
              conn.commit()
              st.success(
                  "Peserta remedial yang dipilih berhasil dihapus dari"
                  " database!"
              )
              st.rerun()
            else:
              st.warning("Pilih minimal satu siswa yang ingin dihapus.")

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
        cleaned_selected_tes = clean_tes_name(pilih_mapel)

        df_all_ujian = pd.read_sql(
            "SELECT nis, poin, status, nama_tes FROM hasil_ujian", conn
        )
        df_matched_ujian = df_all_ujian[
            df_all_ujian["nama_tes"].apply(clean_tes_name)
            == cleaned_selected_tes
        ]

        dict_nilai = {}
        for _, n_row in df_matched_ujian.iterrows():
          dict_nilai[str(n_row["nis"])] = (n_row["poin"], n_row["status"])

        for kelas in pilih_kelas:
          st.markdown(f"**Kelas: {kelas}**")
          df_m_kelas = pd.read_sql(
              "SELECT nis, nama FROM master_siswa WHERE kelas = ?",
              conn,
              params=(kelas,),
          )

          rows = []
          for idx, (_, s_row) in enumerate(df_m_kelas.iterrows(), 1):
            nis_siswa = str(s_row["nis"])
            nama_siswa = s_row["nama"]
            data_n = dict_nilai.get(nis_siswa, (0.0, "normal"))
            nilai_pg = data_n[0]
            status_n = data_n[1]

            if status_n == "remedial":
              ket = "Sudah Remedial"
            elif status_n == "susulan":
              ket = "Sudah Susulan"
            else:
              ket = "-"

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
            st.dataframe(
                df_preview, hide_index=True, use_container_width=True
            )

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

          for kelas, df_prev in preview_data.items():
            safe_sheet_name = re.sub(r"[\\/?:*\[\]]", "-", str(kelas))[:31]
            ws = wb.create_sheet(title=safe_sheet_name)

            ws.merge_cells("A1:D1")
            ws["A1"] = f"MATA PELAJARAN : {pilih_mapel}"
            ws["A1"].font = font_title
            ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

            ws.merge_cells("A2:D2")
            ws["A2"] = f"KELAS : {kelas}"
            ws["A2"].font = font_title
            ws["A2"].alignment = Alignment(horizontal="left", vertical="center")

            headers = ["NO", "NAMA", "NILAI PG", "KET"]
            for col_num, h_title in enumerate(headers, 1):
              cell = ws.cell(row=4, column=col_num, value=h_title)
              cell.font = font_bold
              cell.fill = fill_header
              cell.alignment = Alignment(horizontal="center", vertical="center")
              cell.border = thin_border

            for _, r_data in df_prev.iterrows():
              r_idx = int(r_data["NO"])
              row_num = 4 + r_idx

              c_no = ws.cell(row=row_num, column=1, value=r_idx)
              c_nama = ws.cell(row=row_num, column=2, value=r_data["NAMA"])
              c_nilai = ws.cell(row=row_num, column=3, value=r_data["NILAI PG"])
              c_ket = ws.cell(row=row_num, column=4, value=r_data["KET"])

              nis_s = str(
                  pd.read_sql(
                      "SELECT nis FROM master_siswa WHERE kelas = ? AND nama ="
                      " ?",
                      conn,
                      params=(kelas, r_data["NAMA"]),
                  ).iloc[0]["nis"]
              )
              status_s = dict_nilai.get(nis_s, (0.0, "normal"))[1]

              if status_s == "remedial":
                c_nilai.font = Font(
                    name="Calibri", size=11, color="008000", bold=True
                )
              elif status_s == "susulan":
                c_nilai.font = Font(
                    name="Calibri", size=11, color="0000FF", bold=True
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
              c_ket.alignment = Alignment(
                  horizontal="center", vertical="center"
              )

              c_no.border = thin_border
              c_nama.border = thin_border
              c_nilai.border = thin_border
              c_ket.border = thin_border

            ws.column_dimensions["A"].width = 8
            ws.column_dimensions["B"].width = 38
            ws.column_dimensions["C"].width = 16
            ws.column_dimensions["D"].width = 20

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

      cleaned_cek_tes = clean_tes_name(pilih_mapel)
      df_all_ujian = pd.read_sql(
          "SELECT nis, nama_tes FROM hasil_ujian", conn
      )
      matched_nis = df_all_ujian[
          df_all_ujian["nama_tes"].apply(clean_tes_name) == cleaned_cek_tes
      ]["nis"].astype(str).tolist()

      if matched_nis:
        placeholders = ",".join(["?"] * len(matched_nis))
        query_belum = f"""
                    SELECT kelas, nis, nama 
                    FROM master_siswa
                    WHERE nis NOT IN ({placeholders})
                    ORDER BY kelas, nama
                """
        df_belum = pd.read_sql(query_belum, conn, params=matched_nis)
      else:
        df_belum = pd.read_sql(
            "SELECT kelas, nis, nama FROM master_siswa ORDER BY kelas, nama",
            conn,
        )

      if df_belum.empty:
        st.info("Hebat! Semua siswa sudah mengikuti ujian untuk mata pelajaran ini.")
      else:
        list_kelas_belum = df_belum["kelas"].unique().tolist()

        select_all_kls_blm = st.checkbox(
            "Pilih Semua Kelas", value=True, key="select_all_kls_blm_chk"
        )
        if select_all_kls_blm:
          pilih_kelas_blm = st.multiselect(
              "Pilih Kelas:",
              list_kelas_belum,
              default=list_kelas_belum,
              key="multiselect_kelas_blm",
          )
        else:
          pilih_kelas_blm = st.multiselect(
              "Pilih Kelas:",
              list_kelas_belum,
              key="multiselect_kelas_blm",
          )

        if pilih_kelas_blm:
          df_filtered_belum = df_belum[df_belum["kelas"].isin(pilih_kelas_blm)]
        else:
          df_filtered_belum = pd.DataFrame(columns=df_belum.columns)

        if df_filtered_belum.empty:
          st.info("Tidak ada siswa belum ujian untuk kelas yang dipilih.")
        else:
          st.warning(
              f"Ditemukan {len(df_filtered_belum)} siswa yang belum mengikuti"
              " ujian."
          )
          st.dataframe(
              df_filtered_belum, hide_index=True, use_container_width=True
          )

          output_blm = io.BytesIO()
          with pd.ExcelWriter(output_blm, engine="openpyxl") as writer:
            df_filtered_belum.to_excel(
                writer, sheet_name="Siswa Belum Ujian", index=False
            )
          output_blm.seek(0)

          st.download_button(
              label="📥 Download Daftar Siswa Belum Ujian ke Excel",
              data=output_blm,
              file_name=(
                  f"Siswa_Belum_Ujian_{pilih_mapel.replace('/', '-')}.xlsx"
              ),
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
          )

  # --- MENU 4: CEK & DOWNLOAD PESERTA REMEDIAL ---
  elif menu == "4. Cek & Download Peserta Remedial":
    st.header("Pelacak & Rekap Peserta Didik Remedial")

    df_mapel_rem = pd.read_sql(
        "SELECT DISTINCT nama_tes FROM remedial_siswa", conn
    )
    df_kelas_rem = pd.read_sql(
        "SELECT DISTINCT kelas FROM remedial_siswa", conn
    )

    if df_mapel_rem.empty or df_kelas_rem.empty:
      st.warning("Belum ada data remedial yang di-upload oleh operator.")
    else:
      list_mapel_rem = df_mapel_rem["nama_tes"].tolist()
      list_kelas_rem = df_kelas_rem["kelas"].tolist()

      pilih_mapel_rem = st.multiselect(
          "Pilih Mata Pelajaran Remedial:",
          list_mapel_rem,
          key="multiselect_mapel_rem",
      )

      if "multiselect_kelas_rem" not in st.session_state:
        st.session_state.multiselect_kelas_rem = []


      def toggle_all_kelas():
        if st.session_state.get("select_all_kls_rem_chk", False):
          st.session_state.multiselect_kelas_rem = list_kelas_rem
        else:
          st.session_state.multiselect_kelas_rem = []


      st.checkbox(
          "Pilih Semua Kelas Remedial",
          key="select_all_kls_rem_chk",
          on_change=toggle_all_kelas,
      )

      pilih_kelas_rem = st.multiselect(
          "Pilih Kelas:", list_kelas_rem, key="multiselect_kelas_rem"
      )

      if pilih_mapel_rem or pilih_kelas_rem:
        query_r = "SELECT nama_tes, kelas, nis, nama FROM remedial_siswa WHERE 1=1"
        params = []

        if pilih_mapel_rem:
          placeholders_m = ",".join(["?"] * len(pilih_mapel_rem))
          query_r += f" AND nama_tes IN ({placeholders_m})"
          params.extend(pilih_mapel_rem)

        if pilih_kelas_rem:
          placeholders_k = ",".join(["?"] * len(pilih_kelas_rem))
          query_r += f" AND kelas IN ({placeholders_k})"
          params.extend(pilih_kelas_rem)

        query_r += " ORDER BY nama_tes, kelas, nama"
        df_hasil_rem = pd.read_sql(query_r, conn, params=params)

        if df_hasil_rem.empty:
          st.info("Tidak ada data peserta remedial untuk filter yang dipilih.")
        else:
          st.warning(
              f"Ditemukan {len(df_hasil_rem)} data siswa peserta remedial."
          )
          st.dataframe(
              df_hasil_rem, hide_index=True, use_container_width=True
          )

          output_rem = io.BytesIO()
          with pd.ExcelWriter(output_rem, engine="openpyxl") as writer:
            df_hasil_rem.to_excel(
                writer, sheet_name="Peserta Remedial", index=False
            )
          output_rem.seek(0)

          st.download_button(
              label="📥 Download Daftar Remedial ke Excel",
              data=output_rem,
              file_name="Daftar_Remedial_Pilihan.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
          )
