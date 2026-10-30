import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# Pengaturan halaman Streamlit
st.set_page_config(page_title="Dashboard Performa Manufaktur", layout="wide")

st.title("Dashboard Monitoring Performa Manufaktur")

# =====================================================
# SIMULASI DATA (Contoh, ganti dengan data Anda sendiri)
# =====================================================
@st.cache_data
def load_data():
    # Contoh data dummy, ganti dengan pd.read_excel() atau pd.read_csv() Anda
    data = {
        "UPDAT_DT": pd.date_range(start="2025-01-01", periods=100, freq="D"),
        "ANLNG_S_TP_HSD_TP_STRIP": [785 + (i % 50) - 20 for i in range(100)],
        "INLET_COIL_NO": [f"COIL-{i:03d}" for i in range(100)],
        "NO1_TRG_GRADE_CD": ["Grade A" if i % 2 == 0 else "Grade B" for i in range(100)]
    }
    df = pd.DataFrame(data)
    return df

df = load_data()

# Mendapatkan kolom numerik otomatis
numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()

# Inisialisasi Tab Dashboard
tab_trend, tab_stats, tab_heatmap = st.tabs(["Line/Scatter Trend", "Statistik & Regresi", "Korelasi Heatmap"])

# =====================================================
# TAB 1: LINE/SCATTER TREND INTERAKTIF
# =====================================================
with tab_trend:
    st.subheader("Tren Parameter Proses & Batas Kontrol")

    if not numeric_cols:
        st.warning("Tidak ada kolom numerik untuk divisualisasikan.")
    else:
        # Salin dataframe dan pastikan UPDAT_DT bertipe datetime serta di-sorting
        plot_df = df.copy()
        if "UPDAT_DT" in plot_df.columns:
            plot_df["UPDAT_DT"] = pd.to_datetime(plot_df["UPDAT_DT"], errors="coerce")
            plot_df = plot_df.sort_values(by="UPDAT_DT")

        x_options = ["UPDAT_DT", "INLET_COIL_NO"] + numeric_cols
        
        c1, c2, c3 = st.columns(3)
        x_col = c1.selectbox("Sumbu X", x_options, index=0)
        
        y_default_idx = numeric_cols.index("ANLNG_S_TP_HSD_TP_STRIP") if "ANLNG_S_TP_HSD_TP_STRIP" in numeric_cols else 0
        y_col = c2.selectbox("Parameter Y", numeric_cols, index=y_default_idx)
        
        chart_type = c3.radio("Jenis Grafik", ["Scatter", "Line", "Bar"], horizontal=True)

        x_label = "Waktu Update (UPDAT_DT)" if x_col == "UPDAT_DT" else x_col

        # Klasifikasi status untuk warna titik berdasarkan batas (820 dan 780)
        if y_col == "ANLNG_S_TP_HSD_TP_STRIP":
            plot_df["Status_Batas"] = "Normal (780 - 820)"
            plot_df.loc[plot_df[y_col] >= 820, "Status_Batas"] = "Melampaui Batas Atas (>= 820)"
            plot_df.loc[plot_df[y_col] <= 780, "Status_Batas"] = "Di Bawah Batas Bawah (<= 780)"
            
            color_arg = "Status_Batas"
            color_map = {
                "Normal (780 - 820)": "royalblue",
                "Melampaui Batas Atas (>= 820)": "red",
                "Di Bawah Batas Bawah (<= 780)": "orange"
            }
        else:
            color_arg = None
            color_map = None

        # Pembuatan Grafik menggunakan Plotly
        if chart_type == "Scatter":
            fig = px.scatter(
                plot_df, x=x_col, y=y_col,
                color=color_arg, color_discrete_map=color_map,
                title=f"Sebaran {y_col} terhadap {x_label}"
            )
            fig.update_traces(marker=dict(size=8, opacity=0.8))
        elif chart_type == "Line":
            fig = px.line(
                plot_df, x=x_col, y=y_col,
                markers=True,
                title=f"Tren {y_col} terhadap {x_label}"
            )
            if y_col == "ANLNG_S_TP_HSD_TP_STRIP":
                fig.update_traces(marker=dict(size=8, color=plot_df["Status_Batas"].map(color_map)))
        else:
            fig = px.bar(plot_df, x=x_col, y=y_col, title=f"{y_col} berdasarkan {x_label}")

        # Tambahkan Garis Batas Atas dan Bawah jika parameter yang dipilih sesuai
        if y_col == "ANLNG_S_TP_HSD_TP_STRIP":
            fig.add_hline(y=820, line_dash="dash", line_color="red", annotation_text="Batas Atas (820)")
            fig.add_hline(y=780, line_dash="dot", line_color="orange", annotation_text="Batas Bawah (780)")

        fig.update_layout(
            height=550,
            xaxis_title=x_label,
            yaxis_title=y_col,
            xaxis=dict(
                type="date" if x_col == "UPDAT_DT" else "category",
                tickformat="%Y-%m-%d %H:%M" if x_col == "UPDAT_DT" else None,
            )
        )

        st.plotly_chart(fig, use_container_width=True)

# =====================================================
# TAB 2 & 3: TEMPAT TAMBAHAN FITUR LAIN
# =====================================================
with tab_stats:
    st.subheader("Tab Statistik & Regresi Linier")
    st.write("Gunakan bagian ini untuk analisis statistik lanjutan.")

with tab_heatmap:
    st.subheader("Tab Korelasi Heatmap")
    st.write("Gunakan bagian ini untuk melihat matriks korelasi antar variabel.")