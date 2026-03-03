import streamlit as st
import google.generativeai as genai
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.utils import get_column_letter

# --- 설정 ---
API_KEY = "AIzaSyAJvjCsfk5LuOumDZPlMwQqw6PtFG6LMW4EY"
genai.configure(api_key=API_KEY)

st.set_page_config(page_title="POS 데이터 마스터", layout="wide")
st.title("🪄 POS 데이터 자동화 (서식 보존형)")

# 파일 업로드
master_file = st.file_uploader("1. 마스터 파일 업로드", type=['xlsx'])
monthly_file = st.file_uploader("2. 월별 상세 파일 업로드", type=['xlsx'])
target_month = st.selectbox("업데이트할 월", ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])

if st.button("🚀 데이터 업데이트 시작"):
    if master_file and monthly_file:
        with st.spinner('서식을 유지하며 업데이트 중...'):
            # 1. 데이터 읽기
            df_master = pd.read_excel(master_file)
            df_monthly = pd.read_excel(monthly_file)

            # 2. SKU 매칭 및 데이터 업데이트
            # SKU ID열과 Total Items Sold열 이름이 실제 파일과 맞는지 확인 필요
            mapping = df_monthly.set_index('SKU ID')['Total Items Sold'].to_dict()
            df_master[target_month] = df_master['SKU'].map(mapping).fillna(0).astype(int)

            # 3. [개선 1] R열(Total) 합계 자동 계산 (Jan~Dec 컬럼 합산)
            month_columns = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            # 존재하는 월 컬럼만 더함
            df_master['Total'] = df_master[df_master.columns.intersection(month_columns)].sum(axis=1)

            # 4. [개선 2] 서식(너비 등) 유지를 위한 처리
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_master.to_excel(writer, index=False, sheet_name='Sheet1')
                
                workbook  = writer.book
                worksheet = writer.sheets['Sheet1']

                # 헤더 서식 지정 (색상 등)
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'fg_color': '#D7E4BC', # 연두색 계열
                    'border': 1
                })

                # 열 너비 자동 조정 및 헤더 적용
                for i, col in enumerate(df_master.columns):
                    column_len = max(df_master[col].astype(str).str.len().max(), len(col)) + 2
                    worksheet.set_column(i, i, column_len) # 열 너비 조정
                    worksheet.write(0, i, col, header_format) # 헤더 색상 적용

            st.success(f"✅ {target_month} 데이터 및 Total 합계 업데이트 완료!")
            
            st.download_button(
                label="📥 업데이트된 엑셀 다운로드",
                data=output.getvalue(),
                file_name=f"Master_Updated_{target_month}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("파일을 모두 올려주세요.")