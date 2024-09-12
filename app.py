import os
import json
import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from hangulmerge import HangulMerger  # 가상의 모듈, 실제로는 한글 파일을 다루는 모듈을 사용해야 합니다

# 환경 변수에서 자격증명 JSON 읽기
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

        # 한글 파일 생성 및 메일머지
        template_file = 'path/to/your/template.hwp'  # 한글 파일 경로
        output_file = 'path/to/your/output.pdf'  # PDF 출력 파일 경로
        
        # HangulMerger 사용 예시 (가상의 모듈)
        merger = HangulMerger(template_file)
        for index, row in df.iterrows():
            merger.merge(row.to_dict(), output_file)

        st.success("메일머지 및 PDF 변환이 완료되었습니다!")
    except Exception as e:
        st.error(f"오류 발생: {e}")

# requirements.txt 파일에 필요한 패키지
# streamlit
# gspread
# oauth2client
# pandas
# hangulmerge (가상의 모듈, 실제로는 한글 파일을 다루는 모듈을 사용해야 합니다)
