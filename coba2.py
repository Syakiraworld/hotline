import streamlit as st
import pandas as pd
import plotly.express as px

# --- CONFIG ---
st.set_page_config(page_title='QTY Dashboard with Plotly', layout='wide')

# --- LOAD DATA LANGSUNG DARI FILE ---
@st.cache_data
def load_excel_sheets(file_path):
    """Membaca semua sheet dari file Excel"""
    all_sheets = pd.read_excel(file_path, sheet_name=None)
    # Pastikan kolom Date menjadi datetime di setiap sheet
    for name, sheet in all_sheets.items():
        sheet['Date'] = pd.to_datetime(sheet['Date'], dayfirst=True, errors='coerce')
        sheet['Month'] = name  # tambahkan nama sheet sebagai kolom Bulan
    return all_sheets

@st.cache_data
def combine_sheets(all_sheets):
    """Menggabungkan semua sheet"""
    return pd.concat(all_sheets.values(), ignore_index=True)

# --- BACA DATA DARI FILE 'db.xlsx' ---
file_path = "db.xlsx"  # pastikan file ini ada di folder yang sama
all_sheets = load_excel_sheets(file_path)
sheet_names = list(all_sheets.keys())

# --- OPSI PILIHAN ---
st.sidebar.title("Opsi Tampilan Data")
view_mode = st.sidebar.radio(
    "Pilih mode tampilan:",
    ("Gabungkan semua bulan", "Lihat per bulan")
)

# --- PILIH DATA ---
if view_mode == "Lihat per bulan":
    selected_sheet = st.sidebar.selectbox("Pilih Bulan", sheet_names)
    df = all_sheets[selected_sheet]
else:
    df = combine_sheets(all_sheets)

# --- FILTER ---
st.sidebar.title('Filters')
customers = st.sidebar.multiselect(
    'Pilih Customer', 
    options=df['CUSTOMER'].unique(),
    default=df['CUSTOMER'].unique()
)

# Filter hanya berdasarkan customer
filtered_df = df[df['CUSTOMER'].isin(customers)]

# --- METRICS ---
col1, col2, col3 = st.columns(3)
col1.metric("Total QTY", f"{filtered_df['QTY'].sum():,.0f}")
col2.metric("Average QTY", f"{filtered_df['QTY'].mean():,.2f}")
col3.metric("Records", len(filtered_df))

# --- CHARTS: LINE + CUSTOMER BAR ---
col1, col2 = st.columns(2)
with col1:
    fig_line = px.line(
        filtered_df, 
        x='Date', y='QTY', color='CUSTOMER',
        title='QTY Over Time'
    )
    st.plotly_chart(fig_line, use_container_width=True)

with col2:
    customer_qty = filtered_df.groupby('CUSTOMER')['QTY'].sum().reset_index()
    fig_bar = px.bar(
        customer_qty, 
        x='CUSTOMER', y='QTY', 
        title='Total QTY by CUSTOMER'
    )
    st.plotly_chart(fig_bar, use_container_width=True)

# --- TOP 10 ITEM PALING BANYAK DIPESAN ---
st.subheader("📦 Top 10 Item Paling Banyak Dipesan")

top_items = (
    filtered_df.groupby('Item')['QTY']
    .sum()
    .reset_index()
    .sort_values(by='QTY', ascending=False)
    .head(10)
)

avg_items = (
    filtered_df.groupby('Item')['QTY']
    .mean()
    .reset_index()
    .rename(columns={'QTY': 'AvgQTY'})
)

top_items = pd.merge(top_items, avg_items, on='Item', how='left')

fig_top = px.bar(
    top_items,
    x='Item',
    y='QTY',
    text='AvgQTY',
    title='Top 10 Item Berdasarkan Jumlah Pesanan',
    labels={'QTY': 'Total Pesanan', 'Item': 'Kode Item'}
)
fig_top.update_traces(texttemplate='Avg: %{text:.2f}', textposition='outside')
st.plotly_chart(fig_top, use_container_width=True)

# --- TABLE ---
st.subheader("📋 Filtered Data")
st.dataframe(filtered_df)
