import streamlit as st
import pandas as pd
import random
import os
import time

st.set_page_config(page_title="Simulasi CAT UKP ANT II", layout="wide", page_icon="🚢")

# --- CUSTOM CSS DENGAN PENINGKATAN KONTRAS DAN PENAMBALAN TOMBOL BLANK ---
st.markdown("""
    <style>
    /* Background Utama Modern */
    .stApp {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        color: #f8fafc;
    }
    
    /* Teks Pembacaan Jelas */
    div[data-testid="stMarkdownContainer"] p, 
    div[data-testid="stMarkdownContainer"] h1,
    div[data-testid="stMarkdownContainer"] h2,
    div[data-testid="stMarkdownContainer"] h3,
    div[data-testid="stWidgetLabel"] p,
    div[role="radiogroup"] label p {
        color: #f8fafc !important;
    }

    /* FIX TOMBOL UMUM & NAVIGASI BAWAH (SEBELUMNYA / SELANJUTNYA) */
    div.stButton > button {
        background-color: #334155 !important;
        color: #ffffff !important;
        border: 1px solid #64748b !important;
        font-weight: 600 !important;
        padding: 8px 16px !important;
        border-radius: 8px !important;
    }

    div.stButton > button:hover {
        background-color: #38bdf8 !important;
        color: #0f172a !important;
        border-color: #38bdf8 !important;
    }

    /* Tombol Utama / Primary (Login, Mulai Ujian, Kirim Ujian) */
    button[kind="primary"] {
        background-color: #0284c7 !important;
        color: #ffffff !important;
        border: none !important;
    }

    button[kind="primary"]:hover {
        background-color: #0369a1 !important;
    }

    /* Card Pilihan Jawaban Radio Button */
    div[role="radiogroup"] label {
        background-color: rgba(255, 255, 255, 0.08) !important;
        padding: 10px 15px !important;
        border-radius: 8px !important;
        margin-bottom: 8px !important;
        border: 1px solid rgba(255, 255, 255, 0.15) !important;
    }
    
    div[role="radiogroup"] label:hover {
        background-color: rgba(255, 255, 255, 0.18) !important;
        border-color: #38bdf8 !important;
    }
    
    /* Container Login Rapi Tanpa Kotak Kosong */
    .login-box {
        background: rgba(255, 255, 255, 0.05);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border-radius: 16px;
        padding: 30px;
        border: 1px solid rgba(255, 255, 255, 0.18);
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
    }
    
    /* Tombol Grid Nomor Soal di Sidebar */
    div[data-testid="stSidebar"] div[data-testid="stHorizontalBlock"] {
        gap: 4px !important;
    }
    div[data-testid="stSidebar"] button {
        padding: 2px 1px !important;
        font-size: 12px !important;
        height: 36px !important;
        min-height: 36px !important;
        white-space: nowrap !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        border-radius: 6px !important;
    }

    section[data-testid="stSidebar"] {
        background-color: rgba(15, 23, 42, 0.95) !important;
        border-right: 1px solid rgba(255, 255, 255, 0.1);
    }
    </style>
""", unsafe_allow_html=True)

# --- SISTEM LOGIN ---
USER_CREDENTIALS = {
    "admin": "ant2pass",
    "kru1": "12345",
    "taruna": "pip2026"
}

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    col_l1, col_l2, col_l3 = st.columns([1, 2, 1])
    
    with col_l2:
        # Tampilan Header Login Bersih
        st.markdown("""
            <div style='text-align: center; padding-top: 20px; padding-bottom: 15px;'>
                <h1 style='font-size: 55px; margin-bottom: 0;'>🚢</h1>
                <h1 style='font-size: 36px; font-weight: 800; color: #ffffff; margin-top: 5px; margin-bottom: 5px; text-transform: uppercase;'>SIMULASI CAT UKP ANT II</h1>
                <p style='color: #94a3b8; font-size: 15px; margin-top: 0;'>Sistem Ujian Keahlian Pelaut - Tingkat ANT II</p>
            </div>
        """, unsafe_allow_html=True)
        
        # Form Login tanpa container ganda/kotak kosong
        username = st.text_input("Username", placeholder="Masukkan username...")
        password = st.text_input("Password", type="password", placeholder="Masukkan password...")
        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚀 LOGIN UJIAN", type="primary", use_container_width=True):
            if username in USER_CREDENTIALS and USER_CREDENTIALS[username] == password:
                st.session_state["logged_in"] = True
                st.session_state["username"] = username
                st.success(f"Selamat datang, {username}!")
                st.rerun()
            else:
                st.error("Username atau Password salah!")
            
    st.stop()  # Menahan agar konten di bawahnya tidak muncul jika belum login

# --- TOMBOL LOGOUT DI SIDEBAR ---
with st.sidebar:
    st.write(f"👤 Login sebagai: **{st.session_state.get('username', 'User')}**")
    if st.button("Logout", use_container_width=True):
        st.session_state["logged_in"] = False
        st.rerun()

# JUDUL HALAMAN UTAMA
st.markdown("<h1 style='font-size: 34px; font-weight: 800; color: #ffffff;'>🚢 SIMULASI CAT UKP ANT II</h1>", unsafe_allow_html=True)
st.markdown("---")

# --- LOAD DATA EXCEL MULTI-SHEET & SUPPORT GAMBAR ---
@st.cache_data
def load_data():
    file_candidates = [
        "Bank_Soal_UKP_ANT_II.xlsx",
        "bank_soal_updated.xlsx",
        "bank_soal.xlsx"
    ]
    
    file_path = None
    for f in file_candidates:
        if os.path.exists(f):
            file_path = f
            break
            
    if not file_path:
        raise FileNotFoundError("Tidak ada file bank soal Excel yang ditemukan di direktori aplikasi.")

    xls = pd.ExcelFile(file_path)
    all_sheets = []

    for sheet_name in xls.sheet_names:
        df_sheet = pd.read_excel(xls, sheet_name=sheet_name)
        df_sheet.columns = [str(c).strip() for c in df_sheet.columns]
        
        if 'Fungsi' not in df_sheet.columns and 'Function' not in df_sheet.columns:
            df_sheet['Fungsi'] = sheet_name.strip()
            
        all_sheets.append(df_sheet)

    df_full = pd.concat(all_sheets, ignore_index=True)

    # Standardisasi nama kolom (Termasuk dukungan Gambar)
    rename_map = {
        'Function': 'Fungsi',
        'Competency': 'Competency',
        'Pertanyaan': 'Soal',
        'Opsi A': 'Opsi_A',
        'Opsi B': 'Opsi_B',
        'Opsi C': 'Opsi_C',
        'Opsi D': 'Opsi_D',
        'Kunci Jawaban': 'Jawaban_Benar',
        'Kunci': 'Jawaban_Benar',
        'Jawaban Benar': 'Jawaban_Benar',
        'Gambar': 'Gambar_Soal',
        'Image': 'Gambar_Soal',
        'URL_Gambar': 'Gambar_Soal'
    }
    df_full.rename(columns=rename_map, inplace=True)

    df_full = df_full.dropna(subset=['Soal']).reset_index(drop=True)

    if 'Fungsi' not in df_full.columns:
        df_full['Fungsi'] = 'Umum'
    if 'Competency' not in df_full.columns:
        df_full['Competency'] = 'Umum'
    if 'Penjelasan' not in df_full.columns:
        df_full['Penjelasan'] = 'Belum ada penjelasan/dasar aturan khusus untuk soal ini.'
    if 'Gambar_Soal' not in df_full.columns:
        df_full['Gambar_Soal'] = None

    if 'Jawaban_Benar' in df_full.columns:
        df_full['Jawaban_Benar'] = df_full['Jawaban_Benar'].fillna('').astype(str).str.strip().str.upper()
        df_full['Jawaban_Benar'] = df_full['Jawaban_Benar'].apply(
            lambda x: str(x)[0] if len(str(x)) > 0 and str(x)[0] in ['A', 'B', 'C', 'D'] else ''
        )
    else:
        df_full['Jawaban_Benar'] = ''

    return df_full

try:
    df_raw = load_data()
except Exception as e:
    st.error(f"Gagal membaca file bank_soal: {e}")
    st.stop()

# Inisialisasi Session State
if 'started' not in st.session_state:
    st.session_state.started = False
if 'user_answers' not in st.session_state:
    st.session_state.user_answers = {}
if 'finished' not in st.session_state:
    st.session_state.finished = False
if 'current_q' not in st.session_state:
    st.session_state.current_q = 0
if 'start_time' not in st.session_state:
    st.session_state.start_time = None
if 'duration_seconds' not in st.session_state:
    st.session_state.duration_seconds = 3600
if 'filtered_df' not in st.session_state:
    st.session_state.filtered_df = df_raw.copy()

# Tampilan Menu Awal / Dashboard Ujian
if not st.session_state.started:
    st.subheader("📋 Dashboard & Pengaturan Ujian")
    
    col_a, col_b = st.columns(2)
    
    with col_a:
        list_fungsi = ["Semua Fungsi"] + sorted(list(df_raw['Fungsi'].dropna().unique()))
        selected_fungsi = st.selectbox("🎯 Pilih Fungsi Ujian:", list_fungsi)
        
        if selected_fungsi == "Semua Fungsi":
            df_pool = df_raw.copy()
        else:
            df_pool = df_raw[df_raw['Fungsi'] == selected_fungsi].copy()
            
        max_soal = len(df_pool)
        st.info(f"Total Bank Soal Tersedia: **{max_soal} Soal**")

    with col_b:
        default_target = min(60, max_soal)
        target_soal = st.number_input(
            "🎲 Jumlah Soal yang Akan Diuji (Acak):", 
            min_value=5, 
            max_value=max_soal if max_soal > 5 else 5, 
            value=default_target,
            step=5
        )
        
        durasi_menit = st.number_input("⏱️ Set Durasi Waktu Ujian (Menit):", min_value=1, max_value=180, value=60)

    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        st.write("**Petunjuk Ujian:**")
        st.write("1. Soal akan **diacak otomatis** dari bank soal pilihan Anda.")
        st.write("2. Bebas berpindah nomor soal melalui grid tombol di sidebar.")
        st.write("3. Jawaban tersimpan otomatis saat Anda memilih opsi.")
        st.write("4. Apabila waktu habis, ujian akan otomatis dikirim.")
    
    if st.button("🚀 MULAI SIMULASI CAT", type="primary", use_container_width=True):
        st.session_state.filtered_df = df_pool.sample(n=int(target_soal), random_state=random.randint(1, 10000)).reset_index(drop=True)
        st.session_state.started = True
        st.session_state.finished = False
        st.session_state.current_q = 0
        st.session_state.user_answers = {}
        st.session_state.duration_seconds = durasi_menit * 60
        st.session_state.start_time = time.time()
        st.rerun()

# Tampilan Halaman Ujian
elif st.session_state.started and not st.session_state.finished:
    df = st.session_state.filtered_df
    total_questions = len(df)
    
    elapsed_time = time.time() - st.session_state.start_time
    remaining_time = st.session_state.duration_seconds - elapsed_time
    
    if remaining_time <= 0:
        st.session_state.finished = True
        st.warning("⏰ Waktu ujian telah habis!")
        st.rerun()
        
    mins, secs = divmod(int(remaining_time), 60)
    timer_text = f"{mins:02d}:{secs:02d}"
    
    st.sidebar.title("⏱️ TIMER & NAVIGASI")
    
    if remaining_time < 180:
        st.sidebar.error(f"⏳ Sisa Waktu: **{timer_text}**")
    else:
        st.sidebar.warning(f"⏳ Sisa Waktu: **{timer_text}**")
        
    st.sidebar.markdown("---")
    st.sidebar.write(f"**Navigasi Soal ({total_questions} Soal)**")
    
    cols = st.sidebar.columns(4)
    for i in range(total_questions):
        col = cols[i % 4]
        answered = i in st.session_state.user_answers
        
        btn_label = f"✓{i+1}" if answered else f"{i+1}"
        is_current = (i == st.session_state.current_q)
        btn_type = "primary" if is_current else "secondary"
        
        if col.button(btn_label, key=f"nav_{i}", use_container_width=True, type=btn_type):
            st.session_state.current_q = i
            st.rerun()
            
    st.sidebar.markdown("---")
    if st.sidebar.button("🔴 Selesai & Kirim Ujian", type="primary", use_container_width=True):
        st.session_state.finished = True
        st.rerun()

    # Tampilan Pertanyaan
    q_idx = st.session_state.current_q
    row = df.iloc[q_idx]
    
    head_col1, head_col2 = st.columns([3, 1])
    with head_col1:
        st.subheader(f"Soal No. {q_idx + 1} dari {total_questions}")
        st.caption(f"Fungsi: {row.get('Fungsi', '-')} | Sub-Topik: {row.get('Competency', '-')}")
    with head_col2:
        st.metric(label="⏱️ Sisa Waktu", value=timer_text)

    st.write(f"### {row['Soal']}")
    
    # DUKUNGAN MENAMPILKAN GAMBAR SOAL JIKA ADA DI EXCEL
    img_val = row.get('Gambar_Soal', None)
    if pd.notna(img_val) and str(img_val).strip() != "":
        img_path = str(img_val).strip()
        if os.path.exists(img_path) or img_path.startswith(('http://', 'https://')):
            st.image(img_path, caption="Gambar Pendukung Soal", use_container_width=True)

    options = {
        'A': f"A. {row.get('Opsi_A', '')}",
        'B': f"B. {row.get('Opsi_B', '')}",
        'C': f"C. {row.get('Opsi_C', '')}",
        'D': f"D. {row.get('Opsi_D', '')}"
    }
    
    current_answer = st.session_state.user_answers.get(q_idx, None)
    
    selected_option = st.radio(
        "Pilih Jawaban Anda:",
        options=list(options.keys()),
        format_func=lambda x: options[x],
        index=list(options.keys()).index(current_answer) if current_answer in options else None,
        key=f"radio_{q_idx}"
    )
    
    if selected_option:
        st.session_state.user_answers[q_idx] = selected_option

    st.markdown("<br>", unsafe_allow_html=True)
    
    # TOMBOL NAVIGASI BAWAH (SEBELUMNYA / SELANJUTNYA) - WARNA SUDAH TEGAS & JELAS
    col_prev, col_next = st.columns([1, 1])
    with col_prev:
        if q_idx > 0:
            if st.button("⬅️ Soal Sebelumnya", use_container_width=True):
                st.session_state.current_q -= 1
                st.rerun()
    with col_next:
        if q_idx < total_questions - 1:
            if st.button("Soal Selanjutnya ➡️", use_container_width=True):
                st.session_state.current_q += 1
                st.rerun()

# Tampilan Hasil & Analisis
elif st.session_state.finished:
    df = st.session_state.filtered_df
    st.header("📊 HASIL & ANALISIS UJIAN CAT")
    st.markdown("---")
    
    total = len(df)
    correct_count = 0
    
    for i in range(total):
        user_ans = st.session_state.user_answers.get(i, None)
        correct_ans = str(df.iloc[i].get('Jawaban_Benar', '')).strip().upper()
        if user_ans == correct_ans:
            correct_count += 1
            
    score_percentage = (correct_count / total) * 100 if total > 0 else 0
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Soal", f"{total}")
    c2.metric("Jawaban Benar", f"{correct_count}")
    c3.metric("Skor Akhir", f"{score_percentage:.1f}%")
    
    if score_percentage >= 70:
        st.success("🎉 **LULUS!** Performa Anda sangat baik.")
    else:
        st.error("⚠️ **BELUM LULUS.** Silakan tinjau kembali pembahasan di bawah.")
        
    st.markdown("---")
    st.subheader("📝 Pembahasan Detail & Evaluasi Regulasinya")
    
    for i in range(total):
        row = df.iloc[i]
        user_ans = st.session_state.user_answers.get(i, "Tidak Dijawab")
        correct_ans = str(row.get('Jawaban_Benar', '')).strip().upper()
        is_correct = (user_ans == correct_ans)
        
        status_icon = "✅" if is_correct else "❌"
        with st.expander(f"{status_icon} Soal No. {i+1}: {str(row['Soal'])[:60]}..."):
            st.write(f"**Soal Lengkap:** {row['Soal']}")
            
            # Tampilkan Gambar jika ada
            img_val = row.get('Gambar_Soal', None)
            if pd.notna(img_val) and str(img_val).strip() != "":
                img_path = str(img_val).strip()
                if os.path.exists(img_path) or img_path.startswith(('http://', 'https://')):
                    st.image(img_path, width=300)

            st.write(f"- A: {row.get('Opsi_A', '')}")
            st.write(f"- B: {row.get('Opsi_B', '')}")
            st.write(f"- C: {row.get('Opsi_C', '')}")
            st.write(f"- D: {row.get('Opsi_D', '')}")
            
            st.write(f"**Jawaban Anda:** `{user_ans}`")
            st.write(f"**Jawaban Benar:** `{correct_ans}`")
            st.info(f"**Penjelasan / Dasar Aturan:**\n{row.get('Penjelasan', 'Belum ada penjelasan.')}")

    if st.button("🔄 Kembali ke Menu Utama / Ulangi", type="primary"):
        st.session_state.started = False
        st.session_state.finished = False
        st.session_state.current_q = 0
        st.session_state.user_answers = {}
        st.rerun()