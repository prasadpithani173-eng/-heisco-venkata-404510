import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import date, timedelta
import random

st.set_page_config(page_title="HEISCO - Weekly HSE Dashboard", layout="wide")
st.title("HEISCO - Weekly HSE Dashboard")

# Sample data - no Excel needed
locations = ["Jubail", "Yanbu", "Khobar", "Dammam"]
names = ["Venkata", "Prasad", "Ahmed", "Sara", "John", "Fatima"]

obs_data = []
act_data = []
for i in range(60):
    d = date.today() - timedelta(days=random.randint(0,30))
    loc = random.choice(locations)
    obs_data.append({"Date": d, "Location": loc, "Observer": random.choice(names), "Observations": random.randint(1,10)})
    act_data.append({"Date": d, "Location": loc, "Owner": random.choice(names), "Open": random.randint(0,5), "Closed": random.randint(1,8)})

obs_df = pd.DataFrame(obs_data)
act_df = pd.DataFrame(act_data)

# Filters
st.sidebar.header("Filters")
sel_loc = st.sidebar.multiselect("Location", locations, default=locations)
obs_f = obs_df[obs_df["Location"].isin(sel_loc)]
act_f = act_df[act_df["Location"].isin(sel_loc)]

# KPIs
c1,c2,c3,c4 = st.columns(4)
c1.metric("Total Observations", int(obs_f["Observations"].sum()))
c2.metric("Total Open Actions", int(act_f["Open"].sum()))
c3.metric("Total Closed Actions", int(act_f["Closed"].sum()))
c4.metric("Locations", len(sel_loc))

# Charts
st.subheader("Observations by Location")
fig1 = px.bar(obs_f.groupby("Location")["Observations"].sum().reset_index(), x="Location", y="Observations")
st.plotly_chart(fig1, use_container_width=True)

st.subheader("Actions Trend")
trend = act_f.groupby("Date")[["Open","Closed"]].sum().reset_index()
fig2 = px.line(trend, x="Date", y=["Open","Closed"])
st.plotly_chart(fig2, use_container_width=True)

st.subheader("Data")
st.dataframe(obs_f.head(20))
st.dataframe(act_f.head(20))
