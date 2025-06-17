import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.express as px

# Page setup
st.set_page_config(layout="wide")
st.title("📊 Excel Data Visualizer")

# File upload
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Read Excel file
    xl = pd.ExcelFile(uploaded_file)
    sheet_names = xl.sheet_names
    
    # Sheet selection
    selected_sheet = st.selectbox("Select Sheet", sheet_names)
    df = xl.parse(selected_sheet)
    
    # Column selection
    columns = df.columns.tolist()
    selected_column = st.selectbox("Select Column", columns)
    
    # Chart type selection
    chart_types = ["Bar", "Line", "Pie", "Area", "Histogram", "Scatter"]
    selected_chart = st.selectbox("Select Chart Type", chart_types)
    
    # Data validation
    if df[selected_column].count() == 0:
        st.warning("Selected column contains no data!")
    else:
        # Generate selected chart
        st.subheader(f"{selected_chart} Chart of {selected_column}")
        
        if selected_chart == "Bar":
            st.bar_chart(df[selected_column])
        
        elif selected_chart == "Line":
            st.line_chart(df[selected_column])
        
        elif selected_chart == "Pie":
            fig = px.pie(
                names=df[selected_column].value_counts().index,
                values=df[selected_column].value_counts().values,
                title=f"Pie Chart of {selected_column}"
            )
            st.plotly_chart(fig)
        
        elif selected_chart == "Area":
            st.area_chart(df[selected_column])
        
        elif selected_chart == "Histogram":
            fig, ax = plt.subplots()
            ax.hist(df[selected_column].dropna(), bins=20, edgecolor="black")
            ax.set_xlabel(selected_column)
            ax.set_ylabel("Frequency")
            st.pyplot(fig)
        
        elif selected_chart == "Scatter":
            if len(df) > 0:
                fig = px.scatter(
                    df,
                    x=df.index,
                    y=selected_column,
                    title=f"Scatter Plot of {selected_column}"
                )
                st.plotly_chart(fig)
            else:
                st.warning("Not enough data for scatter plot")
    
    # Show raw data
    st.subheader("Raw Data Preview")
    st.dataframe(df.head(10))

else:
    st.info("Please upload an Excel file to get started")

