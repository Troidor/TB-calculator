import streamlit as st
import pandas as pd

st.title("TB Calculator")

# Load Excel file from same folder
FILE = "TB.xlsx"

df = pd.read_excel(FILE, sheet_name=None)
st.write("Sheets loaded:", df.keys())

st.success("Excel loaded successfully!")
