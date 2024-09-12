import streamlit as st
import pandas as pd
import gspread
from gspread_dataframe import get_as_dataframe
import win32com.client as win32
import os

# 공개 링크
SHEET_URL = "https://docs.google.com/spreadsheets/d/1PHYGLaxzzvhkmJo8xigYceJvtEJwAjVjQdd_JoKysEU/edit?usp=sharing"
TEMPLATE_PATH = "template.hwp"  # Hangul 템플릿 파일 경로
OUTPUT_PATH = "output.pdf"      # 생성될 PDF 파일 경로

# Google Sheets 데이터 읽기 함수
def load_sheet_data(url):
    gc = gspread.service_account()
    sheet = gc.open_by_url(url)
    worksheet = sheet.get_worksheet(0)
    data = get_as_dataframe(worksheet)
    return data

# 메일 머지 및 PDF로 저장 함수
def mail_merge_and_save_pdf(data):
    # Hangul HWP 애플리케이션 열기
    hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
    
    # 템플릿 열기
    hwp.Open(TEMPLATE_PATH)
    
    # 데이터 머지
    for index, row in data.iterrows():
        for i in range(1, 15):  # 1부터 14까지
            placeholder = f"{{{i}}}"
            if placeholder in row:
                hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                hwp.HParameterSet.HInsertText.Text = row[i]
                hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    
    # PDF로 저장
    hwp.SaveAs(OUTPUT_PATH, "PDF")
    
    # Hangul HWP 애플리케이션 종료
    hwp.Quit()

# Streamlit UI
st.title("Google Sheets to Hangul PDF")

data = load_sheet_data(SHEET_URL)
st.write(data)

if st.button("Generate PDF"):
    mail_merge_and_save_pdf(data)
    st.success("PDF 파일이 생성되었습니다.")
