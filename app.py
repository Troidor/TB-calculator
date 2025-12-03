import streamlit as st
import pandas as pd

st.title("TB Calculator")

# Load Excel
df = pd.read_excel("TB.xlsx", sheet_name=None)

st.write("Excel sheets loaded:")
st.write(df.keys())

st.write("Done – ready for calculations!")

