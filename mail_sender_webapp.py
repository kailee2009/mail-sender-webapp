import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="메일 발송 웹앱", layout="wide")
st.title("메일 발송 자동화 웹앱 (엑셀 업로드)")

uploaded_file = st.file_uploader("엑셀(.xlsx, .xlsm) 파일을 업로드하세요", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # 엑셀 파일을 메모리에서 읽기
        bytes_data = uploaded_file.read()
        wb = openpyxl.load_workbook(BytesIO(bytes_data), data_only=True)
        sheet_names = wb.sheetnames
        st.sidebar.header("시트 선택")
        selected_sheet = st.sidebar.selectbox("시트 선택", sheet_names)

        ws = wb[selected_sheet]
        data = list(ws.values)
        if not data:
            st.warning("시트에 데이터가 없습니다.")
        else:
            columns = data[0]
            if all(col is None for col in columns):
                st.warning("시트에 컬럼명이 없습니다. 첫 행에 컬럼명을 추가해 주세요.")
            else:
                df = pd.DataFrame(data[1:], columns=columns)
                st.subheader(f"시트: {selected_sheet}")
                st.dataframe(df, use_container_width=True)
                st.info(f"총 {len(df)}행, {len(df.columns)}열")
    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
else:
    st.info("엑셀 파일을 업로드하면 시트별 데이터가 표시됩니다.") 