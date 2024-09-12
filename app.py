import os
import json
import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import platform

# Google API 자격증명 설정
creds_json = os.getenv('GOOGLE_CREDENTIALS')
if creds_json is None:
    st.error("GOOGLE_CREDENTIALS 환경 변수가 설정되어 있지 않습니다.")
    st.stop()

# JSON 내용 파싱
try:
    creds_dict = json.loads(creds_json)
except json.JSONDecodeError:
    st.error("JSON 파싱 중 오류가 발생했습니다.")
    st.stop()

# Google Sheets API 설정
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# Google Sheets URL 입력
sheet_url = st.text_input("Google Sheets URL을 입력하세요:")

if sheet_url:
    try:
        # Google Sheets에서 데이터 가져오기
        sheet = client.open_by_url(sheet_url).sheet1
        data = sheet.get_all_records()
        df = pd.DataFrame(data)

        # 데이터 표시
        st.write("Google Sheets에서 가져온 데이터:")
        st.dataframe(df)

        # Windows에서만 HWP 및 PDF 생성
        if platform.system() == "Windows":
            import win32com.client as win32  # pywin32는 Windows에서만 사용 가능
            if st.button("HWP 파일 및 PDF 생성"):
                for index, row in df.iterrows():
                    # HWP 파일 생성 및 PDF 변환
                    hwp_template = 'template.hwp'  # 템플릿 파일 경로
                    output_hwp = f"output_{index+1}.hwp"  # 생성될 HWP 파일
                    output_pdf = f"output_{index+1}.pdf"  # PDF 출력 파일
                    
                    # HWP 메일머지 함수 호출
                    create_hwp(hwp_template, output_hwp, row)
                    
                    # HWP 파일을 PDF로 변환
                    hwp_to_pdf(output_hwp, output_pdf)
                    
                    # PDF 다운로드 제공
                    with open(output_pdf, "rb") as pdf:
                        st.download_button(f"{row['이름']}의 PDF 다운로드", pdf, file_name=output_pdf)

        else:
            st.warning("HWP 파일 및 PDF 변환은 Windows에서만 가능합니다.")

    except Exception as e:
        st.error(f"오류 발생: {e}")

# HWP 파일에 데이터를 병합하는 함수 (Windows에서만 사용 가능)
def create_hwp(template_file, output_file, data):
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.Open(template_file)

    for key, value in data.items():
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = hwp.HAction.Execute(f"{{{{{key}}}}}", value)

    hwp.SaveAs(output_file)
    hwp.Quit()

# HWP 파일을 PDF로 변환하는 함수 (Windows에서만 사용 가능)
def hwp_to_pdf(hwp_file, pdf_file):
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.Open(hwp_file)
    hwp.SaveAs(pdf_file, "PDF")
    hwp.Quit()
