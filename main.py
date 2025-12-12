import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import os
from PIL import Image
from helpers import *

st.set_page_config("Data Analyst", layout= "wide")
st.title("Data Analyst App With Streamlit")

if "reports" not in st.session_state:
    st.session_state.reports = []

if "no" not in st.session_state:
    st.session_state.no = 0 
    

chart_folder_name = "charts"
chart_folder_path = f"./{chart_folder_name}"

if not os.path.exists(chart_folder_name):
    os.makedirs(chart_folder_name) 


file = st.sidebar.file_uploader("Upload File CSV", type= (["csv"]))

if file is not None:
    st.sidebar.success(f"File name {file.name} has been uploaded successfully ")
    data = pd.read_csv(file)

    # st.subheader("File đã upload thành công")
    st.subheader("Data Tables")
    st.dataframe(data.head(10))
    with st.expander("Dataset Overview"):
        st.subheader("Detailed Dataset Description")
        st.write(data.describe())
        st.subheader("Data Type")
        st.write(data.dtypes)
        st.subheader("Null Data")
        st.write(data.isnull().sum())

    st.subheader("AI-Generated Data Summary")
    if st.button("AI Overview of the Dataset"):
        with st.expander("AI Response"):
            ai_report = generate_report_from_data(data)
            st.write(ai_report)
    
    st.subheader("Data preprocessing")
    categorical_columns = data.select_dtypes(include="object").columns.tolist() #List category
    numerical_columns = data.select_dtypes(include="number").columns.tolist() #List numeric
    # Nếu DL thiếu
    if "cleaned_data" not in st.session_state:
        st.session_state.cleaned_data = None
    if st.button("Handling missing values"):
        cleaned = data.dropna(subset=categorical_columns)
        cleaned[numerical_columns] = cleaned[numerical_columns].fillna(cleaned[numerical_columns].mean())
         # Lưu vào session_state
        st.session_state.cleaned_data = cleaned
        st.success("Missing values handled!")
        #Dữ liệu sau khi làm sạch
    if st.button("The data after cleaning"):
        if st.session_state.cleaned_data is not None:
            st.dataframe(st.session_state.cleaned_data.head(10))
            with st.expander("Dataset Overview"):
                st.subheader("Detailed Dataset Description")
                st.write(st.session_state.cleaned_data.describe())
                st.subheader("Data Type")
                st.write(st.session_state.cleaned_data.dtypes)
                st.subheader("Null Data")
                st.write(st.session_state.cleaned_data.isnull().sum())
        else:
            st.warning("The data has not been processed yet! Please click 'Handling missing values' first.")
    
    st.subheader("Data Aggregation and Summarization Options") 
    category_col = st.selectbox("Select a categorical column for summary generation:", categorical_columns, key="category") 
    numeric_col = st.selectbox("Select a numerical column for analysis:", numerical_columns, key="numeric") 
    aggregation_function = st.selectbox(
        "Aggregation function:", 
        ["sum", "mean", "count", "min", "max"]
    )
    
    if aggregation_function == "sum": 
        aggregated_data = data.groupby(category_col)[numeric_col].sum().reset_index() 
    elif aggregation_function =="mean":
        aggregated_data = data.groupby(category_col)[numeric_col].mean().reset_index()
    elif aggregation_function =="count":
        aggregated_data = data.groupby(category_col)[numeric_col].count().reset_index()
    elif aggregation_function =="min":
        aggregated_data = data.groupby(category_col)[numeric_col].min().reset_index()
    elif aggregation_function =="max":
        aggregated_data = data.groupby(category_col)[numeric_col].max().reset_index()
    
    st.dataframe(aggregated_data)
    
    st.subheader("Create Visualizations for Your Data")
    
    chart_type = st.selectbox(
        "Choose chart type", 
        ["Line Chart", "Bar Chart", "Scatter Plot", "Pie Chart"] 
    )
    
    if st.button("Generate a Chart"): 
        chart_path, chart_name = plot_chart(chart_folder_name, chart_type, aggregated_data, category_col, numeric_col)
        response = generate_report_from_chart(chart_folder_path, chart_name) 
        
        report = {
            "pivot_table": aggregated_data,
            "chart_path": chart_path, 
            "sheet_name": f"Sheet {st.session_state.no}",
            "insight": response
        }
        
        st.session_state.reports.append(report) 
        st.session_state.no += 1 
        

    for i, report in enumerate(st.session_state.reports):
        filename = report["chart_path"] 
        if filename.endswith(('.png', '.jpg', '.jpeg')):
            file_path = os.path.join(chart_folder_path, filename) 
            
            image = Image.open(file_path) 
            st.image(image, caption=filename, use_column_width=True) 
    
    
    if st.button("Generate Report"):
        generate_excel_report(data, st.session_state.reports, "report")
        with open("report.xlsx", "rb") as file:
            excel_data = file.read()
        st.download_button(label="Download", data=excel_data, file_name="report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")