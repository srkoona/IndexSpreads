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

#===========SIDEBAR SECTION===========
st.sidebar.header("Index Spreads")

# Add page selection to sidebar
page = st.sidebar.selectbox(
    "Select Index",
    ["Main Indices", "Ratings", "Major Sub-Industries", "Minor Sub-Industries"]
)

offsetA, offsetB = st.sidebar.slider(
    "Choose timeframe offset",
    min_value = df["Offset"].min(),
    max_value = df["Offset"].max(),
    value=(0, 36),
    step = 1
)

df_selection = df.query( 
    "(Offset >= @offsetA) & (Offset <= @offsetB)"
)

# Get row 1 values for reference lines
row1_main = df_selection.iloc[0, 1:8]
row1_ratings = df_selection.iloc[0, 8:13]
row1_major = df_selection.iloc[0, 13:36]
row1_minor = df_selection.iloc[0, 36:51]

if page == "Main Indices":
    st.title("Main Indices")
    main_indices = df_selection.iloc[:, 1:8]
    
    # Create three columns
    col1, col2, col3 = st.columns(3)
    
    # Create a cycle of columns to iterate through
    columns = [col1, col2, col3]
    col_idx = 0
    
    # Create individual plots for each column
    for col, value in zip(main_indices.columns, row1_main):
        # Select which streamlit column to put the plot in
        with columns[col_idx]:
            fig = px.box(data_frame=main_indices, 
                        y=col,
                        title=f"Box Plot for {col}",
                        template="plotly_white",
                        points="outliers")
            
            # Add reference line for row 1 value
            fig.add_hline(y=value, line_dash="dash", line_color="red")
            
            fig.update_layout(showlegend=False,
                            height=400,
                            width=400,
                            xaxis_title=None,
                            yaxis_title=None)
            
            st.plotly_chart(fig, use_container_width=True)
        
        # Move to next column, wrap around to first column if needed
        col_idx = (col_idx + 1) % 3

elif page == "Ratings":
    st.title("Ratings")
    ratings = df_selection.iloc[:, 8:13]
    
    # Create three columns
    col1, col2, col3 = st.columns(3)
    columns = [col1, col2, col3]
    col_idx = 0
    
    for col, value in zip(ratings.columns, row1_ratings):
        with columns[col_idx]:
            fig = px.box(data_frame=ratings, 
                        y=col,
                        title=f"Box Plot for {col}",
                        template="plotly_white",
                        points="outliers")
            
            fig.add_hline(y=value, line_dash="dash", line_color="red")
            
            fig.update_layout(showlegend=False,
                            height=400,
                            width=400,
                            xaxis_title=None,
                            yaxis_title=None)
            
            st.plotly_chart(fig, use_container_width=True)
        
        col_idx = (col_idx + 1) % 3
    

elif page == "Major Sub-Industries":
    st.title("Major Sub-Industries")
    major_industries = df_selection.iloc[:, 13:36]
    
    # Create three columns
    col1, col2, col3 = st.columns(3)
    columns = [col1, col2, col3]
    col_idx = 0
    
    for col, value in zip(major_industries.columns, row1_major):
        with columns[col_idx]:
            fig = px.box(data_frame=major_industries, 
                        y=col,
                        title=f"Box Plot for {col}",
                        template="plotly_white",
                        points="outliers")
            
            fig.add_hline(y=value, line_dash="dash", line_color="red")
            
            fig.update_layout(showlegend=False,
                            height=400,
                            width=400,
                            xaxis_title=None,
                            yaxis_title=None)
            
            st.plotly_chart(fig, use_container_width=True)
        
        col_idx = (col_idx + 1) % 3
    

else:  # Minor Sub-Industries
    st.title("Minor Sub-Industries")
    minor_industries = df_selection.iloc[:, 36:51]
    
    # Create three columns
    col1, col2, col3 = st.columns(3)
    columns = [col1, col2, col3]
    col_idx = 0
    
    for col, value in zip(minor_industries.columns, row1_minor):
        with columns[col_idx]:
            fig = px.box(data_frame=minor_industries, 
                        y=col,
                        title=f"Box Plot for {col}",
                        template="plotly_white",
                        points="outliers")
            
            fig.add_hline(y=value, line_dash="dash", line_color="red")
            
            fig.update_layout(showlegend=False,
                            height=400,
                            width=400,
                            xaxis_title=None,
                            yaxis_title=None)
            
            st.plotly_chart(fig, use_container_width=True)
        
        col_idx = (col_idx + 1) % 3
    



