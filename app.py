import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="HEISCO HSE Dashboard", layout="wide")
st.title("HEISCO - Weekly HSE Dashboard")

base = "sample_data"

@st.cache_data
def load_data():
    names = pd.read_excel(os.path.join(base, "Names_Locations_Sample.xlsx"))
    obs = pd.read_excel(os.path.join(base, "Weekly_Observation_Sample.xlsx"))
    act = pd.read_excel(os.path.join(base, "Weekly_HSE_Activity_Sample.xlsx"))
    return names, obs, act

names_df, obs_df, act_df = load_data()

# Sidebar filters
st.sidebar.header("Filters")
areas = ["All"] + sorted(obs_df["Area"].astype(str).unique().tolist())
sel_area = st.sidebar.selectbox("Area / Location", areas)

if sel_area != "All":
    obs_f = obs_df[obs_df["Area"].astype(str) == sel_area]
else:
    obs_f = obs_df

# KPIs
c1, c2, c3, c4 = st.columns(4)
c1.metric("Total Observations", len(obs_f))
c2.metric("Major", len(obs_f[obs_f["Observation Type"]=="Major"]))
c3.metric("Minor", len(obs_f[obs_f["Observation Type"]=="Minor"]))
c4.metric("Closed", len(obs_f[obs_f["Status"]=="Closed"]))

# Charts
st.subheader("Observations by Category")
st.bar_chart(obs_f["Category"].value_counts())

st.subheader("Observations by Area")
st.bar_chart(obs_f["Area"].value_counts())

# Tables
st.subheader("Weekly Observations")
st.dataframe(obs_f, use_container_width=True)

st.subheader("Team - Locations")
st.dataframe(names_df, use_container_width=True)

st.subheader("Weekly HSE Activity / Permits")
st.dataframe(act_df, use_container_width=True)
