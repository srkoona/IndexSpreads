import pandas as pd
import plotly.express as px
import streamlit as st
import requests
from io import BytesIO

st.set_page_config(page_title="Debt Comps", 
                   page_icon=":chart_increasing:", 
                   layout="wide")

excel_file = r"OAS Spread.xlsx"

def getdata_excel():
    df = pd.read_excel(
        io = excel_file,
        engine = "openpyxl",
        sheet_name= "Datasheet",
        skiprows=1,
        usecols="E:AR",
        nrows=52
    )

    return df


df = getdata_excel()
df = df.T
df.columns = df.iloc[0]
df = df.drop(df.index[0])
df = df.reset_index(drop=True)
df = df.drop('Symbol', axis=1)
df.head()

df.columns.tolist()

#===========SIDEBAR SECTION===========

st.sidebar.header("Choose Timeframe")


offsetA, offsetB = st.sidebar.slider(
    "Offset",
    min_value = df["Offset"].min(),
    max_value = df["Offset"].max(),
    value=(0, 36),
    step = 1
)

df_selection = df.query( 
    "(Offset >= @offsetA) & (Offset <= @offsetB)"
    
)

st.dataframe(df_selection)

#===============MAINPAGE=====================

st.title("Index Spreads")
st.markdown("##")

#=============Box Plots===============

# Create figure with subplots
fig = px.box()

# Plot 1: Main Indices (columns 1-8)
main_indices = df_selection.iloc[:, 1:8]
fig1 = px.box(data_frame=main_indices, 
              title="Main Indices",
              template="plotly_white",
              points="outliers",
              labels={"value": "", "variable": ""})
fig1.update_traces(q1=main_indices.quantile(0.2),
                   q3=main_indices.quantile(0.8))
fig1.update_layout(showlegend=False,
                  height=400,
                  width=800,
                  xaxis_title=None,
                  yaxis_title=None)

# Plot 2: Ratings (columns 9-13)
ratings = df_selection.iloc[:, 8:13]
fig2 = px.box(data_frame=ratings,
              title="Ratings",
              template="plotly_white",
              points="outliers",
              labels={"value": "", "variable": ""})
fig2.update_traces(q1=ratings.quantile(0.2),
                   q3=ratings.quantile(0.8))
fig2.update_layout(showlegend=False,
                  height=400,
                  width=800,
                  xaxis_title=None,
                  yaxis_title=None)

# Plot 3: Major sub-industries (columns 14-36)
major_industries = df_selection.iloc[:, 13:36]
fig3 = px.box(data_frame=major_industries,
              title="Major sub-industries",
              template="plotly_white",
              points="outliers",
              labels={"value": "", "variable": ""})
fig3.update_traces(q1=major_industries.quantile(0.2),
                   q3=major_industries.quantile(0.8))
fig3.update_layout(showlegend=False,
                  height=400,
                  width=800,
                  xaxis_title=None,
                  yaxis_title=None)

# Plot 4: Minor sub-industries (columns 37-51)
minor_industries = df_selection.iloc[:, 36:51]
fig4 = px.box(data_frame=minor_industries,
              title="Minor sub-industries",
              template="plotly_white",
              points="outliers",
              labels={"value": "", "variable": ""})
fig4.update_traces(q1=minor_industries.quantile(0.2),
                   q3=minor_industries.quantile(0.8))
fig4.update_layout(showlegend=False,
                  height=400,
                  width=800,
                  xaxis_title=None,
                  yaxis_title=None)

# Display plots in 2x2 grid using streamlit columns
col1, col2 = st.columns(2)
with col1:
    st.plotly_chart(fig1, use_container_width=True)
    st.plotly_chart(fig3, use_container_width=True)
with col2:
    st.plotly_chart(fig2, use_container_width=True)
    st.plotly_chart(fig4, use_container_width=True)



