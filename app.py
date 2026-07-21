import streamlit as st
import pandas as pd
import random
import os
import time

st.set_page_config(page_title="Simulasi CAT UKP ANT II", layout="wide", page_icon="🚢")

# --- SISTEM LOGIN SEDERHANA ---
USER_CREDENTIALS = {
    "admin": "ant2pass",
    "kru1": "12345",
    "taruna": "pip2026"
}

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    st.title("🔐 Login Simulasi CAT ANT-II")
    
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    
    if st.button("Login"):
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
    if st.button("Logout"):
        st.session_state["logged_in"] = False
        st.rerun()
        
st.title("🚢 SIMULASI CAT UKP ANT II")
st.markdown("---")

# Load Data Excel
@st.cache_data
def load_data():
    file_path = "bank_soal_updated.xlsx"
    if not os.path.exists(file_path):
        file_path = "bank_soal.xlsx"
    df = pd.read_excel(file_path)
    return df

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
    st.session_state.duration_seconds = 600 # Default 10 menit
if 'filtered_df' not in st.session_state:
    st.session_state.filtered_df = df_raw.copy()

# Tampilan Menu Awal
if not st.session_state.started:
    st.subheader("📋 Dashboard & Pengaturan Ujian")
    
    # Filter Fungsi
    list_fungsi = ["Semua Fungsi"] + sorted(list(df_raw['Fungsi'].dropna().unique()))
    selected_fungsi = st.selectbox("🎯 Pilih Fungsi Ujian:", list_fungsi)
    
    # Pengaturan Waktu
    durasi_menit = st.number_input("⏱️ Set Durasi Waktu Ujian (Menit):", min_value=1, max_value=180, value=15)
    
    # Filter Data berdasarkan Pilihan
    if selected_fungsi == "Semua Fungsi":
        df_selected = df_raw.copy()
    else:
        df_selected = df_raw[df_raw['Fungsi'] == selected_fungsi].copy()
        
    st.info(f"Total Soal Tersedia untuk Kategori Ini: **{len(df_selected)} Soal**")
    
    col1, col2 = st.columns(2)
    with col1:
        st.write("**Petunjuk Ujian:**")
        st.write("1. Timer akan otomatis berjalan mundur begitu Anda menekan tombol Mulai.")
        st.write("2. Bebas berpindah nomor soal melalui sidebar navigasi.")
        st.write("3. Apabila waktu habis, ujian akan otomatis dikirim.")
    
    if st.button("🚀 MULAI SIMULASI CAT", type="primary", use_container_width=True):
        st.session_state.filtered_df = df_selected.reset_index(drop=True)
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
    
    # Hitung Sisa Waktu
    elapsed_time = time.time() - st.session_state.start_time
    remaining_time = st.session_state.duration_seconds - elapsed_time
    
    if remaining_time <= 0:
        st.session_state.finished = True
        st.warning("⏰ Waktu ujian telah habis!")
        st.rerun()
        
    mins, secs = divmod(int(remaining_time), 60)
    timer_text = f"{mins:02d}:{secs:02d}"
    
    # Sidebar Navigasi Soal & Timer
    st.sidebar.title("⏱️ TIMER & NAVIGASI")
    
    # Widget Waktu di Sidebar
    if remaining_time < 180: # Merah jika kurang dari 3 menit
        st.sidebar.error(f"⏳ Sisa Waktu: **{timer_text}**")
    else:
        st.sidebar.warning(f"⏳ Sisa Waktu: **{timer_text}**")
        
    st.sidebar.markdown("---")
    
    # Grid Navigasi Nomor Soal
    cols = st.sidebar.columns(5)
    for i in range(total_questions):
        col = cols[i % 5]
        answered = i in st.session_state.user_answers
        btn_label = f"✓ {i+1}" if answered else f"{i+1}"
        if col.button(btn_label, key=f"nav_{i}", use_container_width=True):
            st.session_state.current_q = i
            st.rerun()
            
    st.sidebar.markdown("---")
    if st.sidebar.button("🔴 Selesai & Kirim Ujian", type="primary", use_container_width=True):
        st.session_state.finished = True
        st.rerun()

    # Tampilan Pertanyaan
    q_idx = st.session_state.current_q
    row = df.iloc[q_idx]
    
    # Header Soal + Live Timer di kanan
    head_col1, head_col2 = st.columns([3, 1])
    with head_col1:
        st.subheader(f"Soal No. {q_idx + 1} dari {total_questions}")
        st.caption(f"Fungsi: {row['Fungsi']} | Sub-Topik: {row['Competency']}")
    with head_col2:
        st.metric(label="⏱️ Sisa Waktu", value=timer_text)

    st.write(f"### {row['Soal']}")
    
    options = {
        'A': f"A. {row['Opsi_A']}",
        'B': f"B. {row['Opsi_B']}",
        'C': f"C. {row['Opsi_C']}",
        'D': f"D. {row['Opsi_D']}"
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

    # Navigasi Bawah
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
        correct_ans = str(df.iloc[i]['Jawaban_Benar']).strip().upper()
        if user_ans == correct_ans:
            correct_count += 1
            
    score_percentage = (correct_count / total) * 100
    
    # Ringkasan Nilai
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
        correct_ans = str(row['Jawaban_Benar']).strip().upper()
        is_correct = (user_ans == correct_ans)
        
        status_icon = "✅" if is_correct else "❌"
        with st.expander(f"{status_icon} Soal No. {i+1}: {row['Soal'][:60]}..."):
            st.write(f"**Soal Lengkap:** {row['Soal']}")
            st.write(f"- A: {row['Opsi_A']}")
            st.write(f"- B: {row['Opsi_B']}")
            st.write(f"- C: {row['Opsi_C']}")
            st.write(f"- D: {row['Opsi_D']}")
            
            st.write(f"**Jawaban Anda:** `{user_ans}`")
            st.write(f"**Jawaban Benar:** `{correct_ans}`")
            st.info(f"**Penjelasan / Dasar Aturan:**\n{row['Penjelasan']}")

    if st.button("🔄 Kembali ke Menu Utama / Ulangi", type="primary"):
        st.session_state.started = False
        st.session_state.finished = False
        st.session_state.current_q = 0
        st.session_state.user_answers = {}
        st.rerun()