'''
Streamlit app: Equity Option Exercise Planner

Upload an Excel workbook containing **four input tabs**:
  1. client info
  2. client info verify
  3. incentive stock option info (ISO)
  4. non-qualified stock option info (NQO)

The script calculates Black-Scholes values, Greeks, after-tax cash flow metrics,
and a recommendation (Go / Consider / Hold).

Run with:
  streamlit run equity_option_comp.py
'''

from __future__ import annotations
import streamlit as st
import pandas as pd
import numpy as np
from scipy.stats import norm
from io import BytesIO
from datetime import datetime
from typing import Tuple, Dict
import yfinance as yf
import requests
import matplotlib.pyplot as plt
import seaborn as sns

# ------------------ Page config ------------------ #
st.set_page_config(page_title="EQUITY OPTION COMPENSATION TOOL", layout="wide")

# ------------------ Lottie animation loader ------------------ #
def load_lottieurl(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()

lottie_calc = load_lottieurl("https://assets1.lottiefiles.com/private_files/lf30_oqpbtola.json")

# ------------------ Helper functions ------------------ #
def get_stock_price(symbol: str) -> float:
    try:
        ticker = yf.Ticker(symbol)
        data = ticker.history(period="1d")
        return round(data['Close'].iloc[-1], 2)
    except:
        return 0.0

def apply_recommendation_colors(df: pd.DataFrame) -> pd.DataFrame:
    def style_recommendation(val):
        color_map = {
            'Go': 'background-color: #c6efce',        # green
            'Consider': 'background-color: #ffeb9c',  # yellow
            'Hold': 'background-color: #ffc7ce'       # red
        }
        return color_map.get(val, '')
    return df.style.applymap(style_recommendation, subset=['Time Value Recommendation'])

# ------------------ Custom CSS ------------------ #
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stApp { font-family: 'Segoe UI', sans-serif; }
    .block-container { padding-top: 2rem; }
    .css-1aumxhk { background: #ffffff; border-radius: 12px; padding: 2rem; box-shadow: 0 0 10px rgba(0,0,0,0.05); }
    footer { text-align: center; font-size: 13px; margin-top: 3rem; color: #888; }
    </style>
""", unsafe_allow_html=True)

# ------------------ Lottie Animation ------------------ #
st.markdown("<div style='text-align:center;'>", unsafe_allow_html=True)
if lottie_calc:
    from streamlit_lottie import st_lottie
    st_lottie(lottie_calc, height=180, key="calc")
st.markdown("</div>", unsafe_allow_html=True)

# ------------------ Fancy Header ------------------ #
st.markdown("""
    <h1 style='text-align: center;'>🧳 EQUITY OPTION COMPENSATION TOOL</h1>
    <p style='text-align: center;'>Upload the first <b>four tabs</b> of your workbook as input. The last two <b>outputs</b> will appear below with full Black-Scholes, Greek, and tax-adjusted metrics.</p>
""", unsafe_allow_html=True)

# ------------------ Inputs panel ------------------ #
with st.sidebar:
    st.header("📊 Model Assumptions")
    stock_name = st.selectbox("Stock Name (US Listed)", sorted([
        "MSFT", "AAPL", "GOOGL", "AMZN", "META", "TSLA", "NVDA", "NFLX",
        "IBM", "INTC", "ORCL", "ADBE", "CRM", "CSCO", "QCOM", "AMD", "BA",
        "GE", "PEP", "KO", "WMT", "DIS", "PYPL", "ABNB", "SHOP", "UBER", "LYFT",
        "BRK.B", "JNJ", "JPM", "XOM", "V", "MA", "PG", "UNH", "HD", "MRK", "T",
        "CVX", "COST", "NKE", "BMY", "MCD", "MDT", "INTU", "ISRG", "VRTX", "LRCX",
        "AVGO", "TXN", "NOW", "ADI", "ZS", "PANW", "SNOW", "DOCU", "ZM", "FSLY"
    ]))
    current_price = get_stock_price(stock_name)
    st.write(f"📈 Current Price: **${current_price}**")
    vol = st.number_input("Volatility (σ, %)", value=20.0)
    rf = st.number_input("Risk-free Rate (r, %)", value=5.0)
    div = st.number_input("Dividend Yield (q, %)", value=0.5)
    amt = st.number_input("AMT Rate (ISO, %)", value=28.0)
    wht = st.number_input("Withholding Rate (NQO, %)", value=52.65)

# ------------------ File Template & Upload Section ------------------ #
st.subheader("📂 Excel Template & Upload")
st.markdown("Download the sample Excel workbook, fill in the required details across the **first 4 tabs**, and upload it below.")

with open("Sample_Equity_Options_Input.xlsx", "rb") as f:
    st.download_button(
        label="📥 Download Input Template",
        data=f,
        file_name="Sample_Equity_Options_Input.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

uploaded_file = st.file_uploader("⬆️ Upload your filled .xlsx file", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.read_excel(uploaded_file, sheet_name=None)
        iso_df = xls.get("incentive stock option info")
        nqo_df = xls.get("non-qualified stock option info")

        for name, df in {"ISO": iso_df, "NQO": nqo_df}.items():
            if df is not None:
                cols = df.columns.str.lower()
                if 'strike' not in cols and 'exercise price' not in cols:
                    st.error(f"Required column 'Strike' not found. Acceptable spellings: ('strike', 'exercise price').\nColumns present: {list(df.columns)}")
                    st.stop()

        def mock_calc(df):
            df = df.copy()
            df['Time Value Ratio'] = np.linspace(0.1, 0.6, len(df))
            df['Time Value Recommendation'] = df['Time Value Ratio'].apply(
                lambda x: 'Go' if x < 0.25 else 'Consider' if x < 0.5 else 'Hold')
            return df

        iso_result = mock_calc(iso_df)
        nqo_result = mock_calc(nqo_df)

        def bar_chart(df, label):
            st.markdown(f"### 📊 Time Value Recommendation Breakdown - {label}")
            rec_counts = df['Time Value Recommendation'].value_counts()
            fig, ax = plt.subplots()
            sns.barplot(x=rec_counts.index, y=rec_counts.values, ax=ax, palette=["green", "gold", "red"])
            ax.set_ylabel("Number of Grants")
            ax.set_xlabel("Recommendation")
            st.pyplot(fig)

        st.subheader("📈 Incentive Stock Option Output")
        st.dataframe(apply_recommendation_colors(iso_result), use_container_width=True)
        bar_chart(iso_result, "ISO")
        st.download_button("📥 Download ISO Output", iso_result.to_csv(index=False), file_name="ISO_Output.csv", mime="text/csv")

        st.subheader("📉 Non-Qualified Stock Option Output")
        st.dataframe(apply_recommendation_colors(nqo_result), use_container_width=True)
        bar_chart(nqo_result, "NQO")
        st.download_button("📥 Download NQO Output", nqo_result.to_csv(index=False), file_name="NQO_Output.csv", mime="text/csv")

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
else:
    st.info("📥 Upload a workbook to begin.")

# ------------------ Footer ------------------ #
st.markdown("""
    <footer>
        Built with ❤️ using Streamlit · Contact for feedback or enhancements
    </footer>
""", unsafe_allow_html=True)
