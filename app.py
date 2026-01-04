import streamlit as st
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt

# === Konfigurasi Halaman ===
st.set_page_config(page_title="SPK Alat Musik Tradisional", layout="wide")

# === Fungsi Load Data ===
@st.cache_data
def load_data():
    file_path = "Sistem Pendukung Keputusan dalam Menentukan Alat Musik Kesenian Tradisional Terbaik untuk Pertunjukan Seni Menggunakan Metode AHP dan TOPSIS.xlsx"
    
    # Baca semua sheet
    dataset_raw = pd.read_excel(file_path, sheet_name="Dataset", header=None)
    topsis_raw = pd.read_excel(file_path, sheet_name="Topsis")
    ahp_raw = pd.read_excel(file_path, sheet_name="AHP")
    
    # === Proses Dataset ===
    kriteria = dataset_raw.iloc[1:6, 0:2].copy()
    kriteria.columns = ["Kode", "Kriteria"]
    kriteria["Jenis"] = ["Benefit", "Benefit", "Benefit", "Benefit", "Cost"]
    kriteria = kriteria[["Kode", "Kriteria", "Jenis"]].reset_index(drop=True)
    
    # Ambil alternatif dan nilai kriteria (20 alternatif pertama yang punya data lengkap)
    alternatif = dataset_raw.iloc[1:21, 4].dropna().reset_index(drop=True)
    nilai = dataset_raw.iloc[1:21, 5:10].copy().reset_index(drop=True)
    nilai.columns = ["C1", "C2", "C3", "C4", "C5"]
    dataset = pd.concat([alternatif.to_frame(name="Alternatif"), nilai], axis=1)
    
    # === Proses AHP ===
    bobot = ahp_raw.iloc[0:5, [0, 6]].copy()
    bobot.columns = ["Kriteria", "Bobot"]
    bobot["Bobot %"] = (bobot["Bobot"] * 100).round(2)
    
    desc = ahp_raw.iloc[0:5, [10, 11]].copy()
    desc.columns = ["Kode", "Kriteria Lengkap"]
    desc = desc.dropna().reset_index(drop=True)
    
    # === Proses TOPSIS (SUDAH DIPERBAIKI SESUAI STRUKTUR EXCEL) ===
    # Hasil ranking ada di kolom "Alternatif.1" dan seterusnya
    topsis = topsis_raw[["Alternatif.1", "Rank", "Nilai Preferensi", "Keterangan"]].copy()
    
    # Ganti nama kolom agar mudah dibaca
    topsis.columns = ["Alternatif", "Rank", "Nilai Preferensi", "Keterangan"]
    
    # Hapus baris kosong (header atau NaN total)
    topsis = topsis.dropna(subset=["Alternatif"]).reset_index(drop=True)
    
    # Rank jadi integer bulat
    topsis["Rank"] = pd.to_numeric(topsis["Rank"], errors='coerce').astype("Int64")
    
    # Urutkan berdasarkan Rank
    topsis = topsis.sort_values("Rank").reset_index(drop=True)
    
    # Format angka
    topsis["Nilai Preferensi"] = pd.to_numeric(topsis["Nilai Preferensi"], errors='coerce').round(2)
    
    return kriteria, dataset, bobot, desc, topsis

# Load data
kriteria, dataset, bobot, desc, topsis = load_data()

# === Tampilan Utama ===
st.title("🎶 Sistem Pendukung Keputusan Pemilihan Alat Musik Tradisional Terbaik")
st.markdown("**Metode: Analytic Hierarchy Process (AHP) + TOPSIS**")
st.markdown("Untuk Pertunjukan Seni Kesenian Tradisional Indonesia")
st.markdown("---")

# Sidebar Navigasi
st.sidebar.header("Navigasi")
page = st.sidebar.radio("Pilih Menu:", 
                        ["Beranda", "Dataset", "Bobot Kriteria (AHP)", "Hasil Perankingan (TOPSIS)", "Visualisasi"])

if page == "Beranda":
    st.header("Selamat Datang!")
    st.write("""
    Aplikasi ini membantu menentukan alat musik tradisional terbaik untuk pertunjukan seni 
    berdasarkan 5 kriteria menggunakan metode **AHP** untuk pembobotan dan **TOPSIS** untuk perankingan.
    """)
    
    st.subheader("Kriteria Penilaian")
    st.table(kriteria)
    
    st.success("🏆 **Rekomendasi Utama:** **Kendang** (Rank 1) → Nilai Preferensi: 0.8788 (Sangat Baik)")
    st.info("Diikuti Gong Ageng (Rank 2) dan Gong Suwukan (Rank 3)")

elif page == "Dataset":
    st.header("Dataset Penilaian Alternatif")
    st.write("Skala: **1 = sangat buruk/mahal** → **5 = sangat baik/murah**")
    st.dataframe(dataset.style.background_gradient(cmap="viridis"), use_container_width=True)

elif page == "Bobot Kriteria (AHP)":
    st.header("Bobot Kriteria Hasil AHP")
    st.dataframe(bobot.style.format({"Bobot": "{:.4f}", "Bobot %": "{:.2f}%"})
                 .bar(subset=["Bobot %"], color="#FF6347"), use_container_width=True)
    
    st.subheader("Grafik Bobot Kriteria")
    fig1, ax1 = plt.subplots(figsize=(10, 6))
    ax1.barh(bobot["Kriteria"], bobot["Bobot"], color="#FF4500")
    ax1.set_xlabel("Bobot Relatif")
    ax1.set_title("Prioritas Kriteria Berdasarkan AHP")
    ax1.invert_yaxis()
    for i, v in enumerate(bobot["Bobot"]):
        ax1.text(v + 0.005, i, f"{v:.3f} ({bobot['Bobot %'].iloc[i]:.2f}%)", va='center', fontweight='bold')
    st.pyplot(fig1)

elif page == "Hasil Perankingan (TOPSIS)":
    st.header("Hasil Perankingan Metode TOPSIS")
    st.dataframe(topsis.style.background_gradient(subset=["Nilai Preferensi"], cmap="Reds")
                 .format({"Nilai Preferensi": "{:.2f}"}), 
                 use_container_width=True)
    
    st.write("**Interpretasi:**")
    st.write("- Nilai Preferensi **> 0.70** → **Sangat Baik**")
    st.write("- Nilai Preferensi **0.40 – 0.70** → **Baik**")
    st.write("- Nilai Preferensi **< 0.40** → **Terburuk**")

elif page == "Visualisasi":
    st.header("Visualisasi Ranking Alat Musik Terbaik")
    
    fig2, ax2 = plt.subplots(figsize=(10, 8))
    bars = ax2.barh(topsis["Alternatif"], topsis["Nilai Preferensi"], color="#228B22")
    ax2.set_xlabel("Nilai Preferensi (Semakin tinggi = semakin baik)")
    ax2.set_title("Ranking Alat Musik Tradisional Berdasarkan TOPSIS")
    ax2.invert_yaxis()
    ax2.grid(axis='x', alpha=0.3)
    
    # Label nilai di ujung bar
    for bar in bars:
        width = bar.get_width()
        ax2.text(width + 0.005, bar.get_y() + bar.get_height()/2, 
                 f'{width:.4f}', va='center', fontweight='bold', fontsize=11)
    
    st.pyplot(fig2)
    
    st.markdown("### Kesimpulan")
    st.success("**Kendang** adalah alat musik tradisional **terbaik secara keseluruhan** untuk pertunjukan seni berdasarkan perhitungan AHP dan TOPSIS.")

# Footer
st.markdown("---")
st.caption("Dibuat dengan ❤️ menggunakan Streamlit | Sumber data: Perhitungan AHP-TOPSIS Alat Musik Tradisional")