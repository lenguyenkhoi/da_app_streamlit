import streamlit as st
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
from PIL import Image
from dotenv import load_dotenv
import os
import xlwings as xw
import google.generativeai as genai

#Hàm sử lý dữ liệu null
def preprocess_data (data):
    # file = pd.read_csv(data)
    categorical_columns = data.select_dtypes(include= ["object"]).columns
    data = data.dropna(subset = categorical_columns)
    numerical_columns = data.select_dtypes(include = ["number"]).columns
    for numeric_col in numerical_columns:
        data[numeric_col] = data[numeric_col].fillna(data[numeric_col].mean())
    return data

#Hàm vẽ biểu đồ
def plot_chart(folder_path, chart_type, data, x_col, y_col):
    chart_name = ""
    if chart_type == "Line Chart": #Biểu đồ đường
        chart = sns.lineplot(data=data, x=x_col, y=y_col, markers='o')
        chart_fig = chart.get_figure()
        chart_name = f"{chart_type}_{x_col}_by_{y_col}.png"
        chart_fig.savefig(f"./{folder_path}/{chart_name}")
        # st.pyplot()
    elif chart_type == "Bar Chart":
        chart = sns.barplot(data=data, x=x_col, y=y_col)
        chart_fig = chart.get_figure()
        chart_name = f"{chart_type}_{x_col}_by_{y_col}.png"
        chart_fig.savefig(f"./{folder_path}/{chart_name}")
        # st.pyplot()
    elif chart_type == "Scatter Plot":
        chart = sns.scatterplot(data=data, x=x_col, y=y_col)
        chart_fig = chart.get_figure()
        chart_name = f"{chart_type}_{x_col}_by_{y_col}.png"
        chart_fig.savefig(f"./{folder_path}/{chart_name}")
        # st.pyplot()
    elif chart_type == "Pie Chart":
        pie_data = data.groupby(x_col)[y_col].sum()
        chart = pie_data.plot.pie(autopct='%1.1f%%', startangle=90)
        chart_fig = chart.get_figure()
        chart_name = f"{chart_type}_{x_col}_by_{y_col}.png"
        chart_fig.savefig(f"./{folder_path}/{chart_name}")
        # st.pyplot()

    return f"D:/cybersoft/cs_data_10/buoi16/{folder_path}/{chart_name}", chart_name

load_dotenv()
#Hàm đưa ra các insight từ chart
def generate_report_from_chart(chart_folder, chart_name):
    genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
    model = genai.GenerativeModel("gemini-2.5-flash")

    if chart_name.endswith(('.png', '.jpg', '.jpeg')):
        file_path = os.path.join(chart_folder, chart_name)
        organ = Image.open(file_path)

        response = model.generate_content(["According to the chart, generate a comprehensive report. Thus, what insight could be taken from it. Write at most 100 words", organ])
        bot_response = response.text.replace("*", "")

    return bot_response

#Hàm tóm tắt dữ liệu với AI 
def generate_report_from_data(data):
    """
    Tóm tắt dữ liệu bằng AI dựa trên DataFrame đã upload.
    - Tự động convert DataFrame thành CSV text
    - Gửi vào Gemini để phân tích
    """

    genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
    model = genai.GenerativeModel("gemini-2.5-flash")

    # Chuyển DataFrame thành dạng text ngắn gọn
    data_text = data.head(200).to_csv(index=False)  
    # (giới hạn 200 dòng để tránh input quá dài)

    prompt = f"""
    You are a data analyst AI.
    Analyze the following dataset and summarize it.
    Provide:
    - Dataset size (rows, columns)
    - Numeric vs categorical column summary
    - Missing value issues
    - Key statistical patterns
    - Notable anomalies or insights
    - 2-3 recommendations for analysis
    Respond in less than 120 words.
    
    Here is the dataset (CSV format):
    {data_text}
    """

    response = model.generate_content(prompt)
    bot_response = response.text.replace("*", "")

    return bot_response
#Hàm tạo report

def generate_excel_report(data, reports, report_name):
    """
    Generates an Excel report containing:
    - Original dataset
    - Pivot tables from aggregation
    - Charts (saved as images)
    - AI-generated insights

    Args:
    - data (pd.DataFrame): The original dataset.
    - reports (list): A list of reports with pivot tables, charts, and insights.
    - report_name (str): The output Excel file name (without extension).

    Returns:
    - str: The absolute path to the generated report.
    """

    report_filename = f"{report_name}.xlsx"
    report_path = os.path.abspath(report_filename)

    # Use Pandas' ExcelWriter with XlsxWriter as the engine
    with pd.ExcelWriter(report_path, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Save the original dataset in the first sheet
        data.to_excel(writer, sheet_name="Datasource", index=False)

        # Process each report
        for report in reports:
            sheet_name = report["sheet_name"]

            # Save the pivot table to Excel
            report["pivot_table"].to_excel(writer, sheet_name=sheet_name, index=False)

            # Get the worksheet to add images and insights
            worksheet = writer.sheets[sheet_name]

            # Add Chart Image if Exists
            chart_path = report["chart_path"]
            if os.path.exists(chart_path):
                worksheet.insert_image("F1", chart_path, {"x_scale": 0.5, "y_scale": 0.5})

            # Add AI-Generated Insight in Cell F24
            worksheet.write("F24", report["insight"])

