import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta, date, time
import io
from io import BytesIO
import base64
import json
import re
import calendar
import math
import os
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import requests
from PIL import Image
from io import BytesIO
import plotly.io as pio
import numpy as np
from dateutil.relativedelta import relativedelta
import pytz
import gspread
import tempfile
from PyPDF2 import PdfMerger

# 날짜 정규화 함수
def normalize_date(date_str):
    if pd.isna(date_str) or date_str == '':
        return None
    
    # 이미 datetime 객체인 경우
    if isinstance(date_str, (datetime, pd.Timestamp)):
        return date_str
    
    # 문자열인 경우
    if isinstance(date_str, str):
        # 공백 제거
        date_str = date_str.strip()
        
        # 빈 문자열 처리
        if not date_str:
            return None
            
        # 날짜 형식 변환 시도
        try:
            # YYYY-MM-DD 형식
            if re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
                return datetime.strptime(date_str, '%Y-%m-%d')
            # YYYY.MM.DD 형식
            elif re.match(r'^\d{4}\.\d{2}\.\d{2}$', date_str):
                return datetime.strptime(date_str, '%Y.%m.%d')
            # YYYY/MM/DD 형식
            elif re.match(r'^\d{4}/\d{2}/\d{2}$', date_str):
                return datetime.strptime(date_str, '%Y/%m/%d')
            # YYYYMMDD 형식
            elif re.match(r'^\d{8}$', date_str):
                return datetime.strptime(date_str, '%Y%m%d')
        except ValueError:
            return None
    
    return None

def calculate_experience(experience_text):
    """경력기간을 계산하는 함수"""
    from datetime import datetime
    import pandas as pd
    import re
    
    # 영문 월을 숫자로 변환하는 딕셔너리
    month_dict = {
        'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06',
        'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'
    }
    
    total_months = 0
    experience_periods = []
    
    # 각 줄을 분리하여 처리
    lines = experience_text.split('\n')
    current_company = None
    
    for line in lines:
        # 공백과 탭 문자를 모두 일반 공백으로 변환하고 연속된 공백을 하나로 처리
        line = re.sub(r'[\s\t]+', ' ', line.strip())
        if not line:
            continue
            
        # 회사명 추출 (숫자나 특수문자가 없는 줄)
        if not any(c.isdigit() for c in line) and not any(c in '~-–./' for c in line):
            current_company = line
            continue
            
        # 영문 월 형식 패턴 (예: Nov 2021 – Oct 2024)
        en_pattern = r'([A-Za-z]{3})\s*(\d{4})\s*[–-]\s*([A-Za-z]{3})\s*(\d{4})'
        en_match = re.search(en_pattern, line)
        
        # 한국어 날짜 형식 패턴 (예: 2021 년 11월 – 2024 년 10월)
        kr_pattern = r'(\d{4})\s*년?\s*(\d{1,2})\s*월\s*[-–~]\s*(\d{4})\s*년?\s*(\d{1,2})\s*월'
        kr_match = re.search(kr_pattern, line)
        
        if en_match:
            start_month, start_year, end_month, end_year = en_match.groups()
            start_date = f"{start_year}-{month_dict[start_month]}-01"
            end_date = f"{end_year}-{month_dict[end_month]}-01"
            
            start = datetime.strptime(start_date, "%Y-%m-%d")
            end = datetime.strptime(end_date, "%Y-%m-%d")
            
            months = (end.year - start.year) * 12 + (end.month - start.month) + 1
            total_months += months
            
            years = months // 12
            remaining_months = months % 12
            decimal_years = round(months / 12, 1)
            
            period_str = f"{start_year}-{month_dict[start_month]}~{end_year}-{month_dict[end_month]} ({years}년 {remaining_months}개월, {decimal_years}년)"
            if current_company:
                period_str = f"{current_company}: {period_str}"
            experience_periods.append(period_str)
            continue
            
        elif kr_match:
            start_year, start_month, end_year, end_month = kr_match.groups()
            start_date = f"{start_year}-{start_month.zfill(2)}-01"
            end_date = f"{end_year}-{end_month.zfill(2)}-01"
            
            start = datetime.strptime(start_date, "%Y-%m-%d")
            end = datetime.strptime(end_date, "%Y-%m-%d")
            
            months = (end.year - start.year) * 12 + (end.month - start.month) + 1
            total_months += months
            
            years = months // 12
            remaining_months = months % 12
            decimal_years = round(months / 12, 1)
            
            period_str = f"{start_year}-{start_month.zfill(2)}~{end_year}-{end_month.zfill(2)} ({years}년 {remaining_months}개월, {decimal_years}년)"
            if current_company:
                period_str = f"{current_company}: {period_str}"
            experience_periods.append(period_str)
            continue
            
        # 날짜 패턴 처리
        # 1. 2023. 04 ~ 2024. 07 형식
        pattern1 = r'(\d{4})\.\s*(\d{1,2})\s*[~-–]\s*(\d{4})\.\s*(\d{1,2})'
        # 2. 2015.01.~2016.06 형식
        pattern2 = r'(\d{4})\.(\d{1,2})\.\s*[~-–]\s*(\d{4})\.(\d{1,2})'
        # 3. 2024.05 ~ 형식
        pattern3 = r'(\d{4})\.(\d{1,2})\s*[~-–]'
        # 4. 2024-05 ~ 형식
        pattern4 = r'(\d{4})-(\d{1,2})\s*[~-–]'
        # 5. 2024/05 ~ 형식
        pattern5 = r'(\d{4})/(\d{1,2})\s*[~-–]'
        # 6. 2024.05.01 ~ 형식 (일 부분 무시)
        pattern6 = r'(\d{4})\.(\d{1,2})\.\d{1,2}\s*[~-–]'
        # 7. 2024-05-01 ~ 형식 (일 부분 무시)
        pattern7 = r'(\d{4})-(\d{1,2})-\d{1,2}\s*[~-–]'
        # 8. 2024/05/01 ~ 형식 (일 부분 무시)
        pattern8 = r'(\d{4})/(\d{1,2})/\d{1,2}\s*[~-–]'
        # 9. 2023/05 - 2024.04 형식
        pattern9 = r'(\d{4})[/\.](\d{1,2})\s*[-]\s*(\d{4})[/\.](\d{1,2})'
        # 10. 2023-04-24 ~ 2024-05-10 형식
        pattern10 = r'(\d{4})-(\d{1,2})-(\d{1,2})\s*[~-–]\s*(\d{4})-(\d{1,2})-(\d{1,2})'
        # 11. 2021-03-2026-08 형식
        pattern11 = r'(\d{4})-(\d{1,2})-(\d{4})-(\d{1,2})'
        # 12. 2021-03~2022-08 형식
        pattern12 = r'(\d{4})-(\d{1,2})\s*[~-–]\s*(\d{4})-(\d{1,2})'
        
        # 패턴 매칭 시도
        match = None
        current_pattern = None
        
        # 먼저 패턴 10으로 시도 (2023-04-24 ~ 2024-05-10 형식)
        match = re.search(pattern10, line)
        if match:
            current_pattern = pattern10
        # 다음으로 패턴 12로 시도 (2021-03~2022-08 형식)
        elif re.search(pattern12, line):
            match = re.search(pattern12, line)
            current_pattern = pattern12
        else:
            # 다른 패턴 시도
            for pattern in [pattern1, pattern2, pattern3, pattern4, pattern5, pattern6, pattern7, pattern8, pattern9, pattern11]:
                match = re.search(pattern, line)
                if match:
                    current_pattern = pattern
                    break
                
        if match and current_pattern:
            if current_pattern in [pattern1, pattern2, pattern9]:
                start_year, start_month, end_year, end_month = match.groups()
                start_date = f"{start_year}-{start_month.zfill(2)}-01"
                end_date = f"{end_year}-{end_month.zfill(2)}-01"
                start = datetime.strptime(start_date, "%Y-%m-%d")
                end = datetime.strptime(end_date, "%Y-%m-%d")
            elif current_pattern == pattern10:
                start_year, start_month, start_day, end_year, end_month, end_day = match.groups()
                start_date = f"{start_year}-{start_month.zfill(2)}-{start_day.zfill(2)}"
                end_date = f"{end_year}-{end_month.zfill(2)}-{end_day.zfill(2)}"
                start = datetime.strptime(start_date, "%Y-%m-%d")
                end = datetime.strptime(end_date, "%Y-%m-%d")
            elif current_pattern in [pattern11, pattern12]:
                start_year, start_month, end_year, end_month = match.groups()
                start_date = f"{start_year}-{start_month.zfill(2)}-01"
                end_date = f"{end_year}-{end_month.zfill(2)}-01"
                start = datetime.strptime(start_date, "%Y-%m-%d")
                end = datetime.strptime(end_date, "%Y-%m-%d")
            else:
                start_year, start_month = match.groups()
                start_date = f"{start_year}-{start_month.zfill(2)}-01"
                start = datetime.strptime(start_date, "%Y-%m-%d")
                
                # 종료일 처리
                if '현재' in line or '재직중' in line:
                    end = datetime.now()
                else:
                    # 종료일 패턴 처리 (일 부분 무시)
                    end_pattern = r'[~-–]\s*(\d{4})[\.-/](\d{1,2})(?:[\.-/]\d{1,2})?'
                    end_match = re.search(end_pattern, line)
                    if end_match:
                        end_year, end_month = end_match.groups()
                        end_date = f"{end_year}-{end_month.zfill(2)}-01"
                        end = datetime.strptime(end_date, "%Y-%m-%d")
                    else:
                        # 종료일이 없는 경우
                        period_str = f"{start_year}-{start_month.zfill(2)}~종료일 입력 필요"
                        if current_company:
                            period_str = f"{current_company}: {period_str}"
                        experience_periods.append(period_str)
                        continue
            
            # 경력기간 계산
            if current_pattern in [pattern10, pattern11, pattern12]:
                # 패턴 10, 11, 12의 경우 정확한 일자 계산
                months = (end.year - start.year) * 12 + (end.month - start.month)
                if end.day < start.day:
                    months -= 1
                if months < 0:
                    months = 0
            else:
                # 다른 패턴의 경우 기존 로직 유지
                months = (end.year - start.year) * 12 + (end.month - start.month) + 1
            
            total_months += months
            
            years = months // 12
            remaining_months = months % 12
            decimal_years = round(months / 12, 1)
            
            # 결과 문자열 생성
            if current_pattern == pattern10:
                period_str = f"{start_year}-{start_month.zfill(2)}~{end_year}-{end_month.zfill(2)} ({years}년 {remaining_months}개월, {decimal_years}년)"
            elif current_pattern in [pattern11, pattern12]:
                period_str = f"{start_year}-{start_month.zfill(2)}~{end_year}-{end_month.zfill(2)} ({years}년 {remaining_months}개월, {decimal_years}년)"
            else:
                period_str = f"{start_year}-{start_month.zfill(2)}~{end.year}-{str(end.month).zfill(2)} ({years}년 {remaining_months}개월, {decimal_years}년)"
            
            if current_company:
                period_str = f"{current_company}: {period_str}"
            experience_periods.append(period_str)
    
    # 총 경력기간 계산
    total_years = total_months // 12
    total_remaining_months = total_months % 12
    total_decimal_years = round(total_months / 12, 1)
    
    # 결과 문자열 생성
    result = "\n".join(experience_periods)
    if result:
        result += f"\n\n총 경력기간: {total_years}년 {total_remaining_months}개월 ({total_decimal_years}년)"
    
    return result

# 페이지 설정
st.set_page_config(
    page_title="HRmate",
    page_icon="👥",
    layout="wide"
)

# CSS 스타일 추가
st.markdown("""
    <style>
    /* 비밀번호 입력 필드 스타일 */
    .password-input [data-testid="stTextInput"] {
        width: 150px !important;
        max-width: 150px !important;
        margin: 0 auto;
    }
    .password-input [data-testid="stTextInput"] input {
        width: 150px !important;
    }
    /* 검색 입력 필드 스타일 */
    .search-container [data-testid="stTextInput"] {
        width: 100px !important;
        max-width: 100px !important;
        margin: 0;
    }
    .search-container [data-testid="stTextInput"] input {
        width: 100px !important;
    }
    /* 금액 표시 스타일 */
    [data-testid="stMetricValue"] {
        font-size: 0.9rem !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.8rem !important;
    }
    .divider {
        max-width: 500px;
        margin: 1rem auto;
    }
    .header-container {
        position: relative;
        max-width: 600px;
        margin: 0 auto;
        padding: 1rem;
        text-align: center;
    }
    .logo-container {
        position: absolute;
        top: 0;
        right: 0;
        width: 130px;
    }
    .title-container {
        padding-top: 1rem;
    }
    .title-container h1 {
        margin: 0;
        color: #666;
    }
    .title-container p {
        margin: 0.5rem 0 0 0;
        color: #666;
        font-size: 0.9em;
    }
    .search-container {
        text-align: left;
        padding-left: 0;
    }
    [data-testid="stSidebar"] button {
        width: 80% !important;
        margin: 0.1rem auto !important;
        display: block !important;
        padding: 0.7rem !important;
        min-height: 0 !important;
        height: auto !important;
        line-height: 1.2 !important;
        text-align: left !important;
    }
    [data-testid="stSidebar"] [data-testid="stExpander"] {
        width: 80% !important;
        margin: 0.1rem auto !important;
        display: block !important;
    }
    [data-testid="stSidebar"] section[data-testid="stSidebarNav"] {
        padding-top: 0.5rem !important;
        padding-bottom: 0.5rem !important;
    }
    [data-testid="stSidebar"] hr {
        margin: 0.5rem 0 !important;
    }
    </style>
""", unsafe_allow_html=True)

def show_header():
    """로고와 시스템 이름을 표시하는 함수"""
    st.markdown("""
        <div class="header-container">
            <div class="logo-container">
                <img src="https://neurophethr.notion.site/image/https%3A%2F%2Fs3-us-west-2.amazonaws.com%2Fsecure.notion-static.com%2Fe3948c44-a232-43dd-9c54-c4142a1b670b%2Fneruophet_logo.png?table=block&id=893029a6-2091-4dd3-872b-4b7cd8f94384&spaceId=9453ab34-9a3e-45a8-a6b2-ec7f1cefbd7f&width=410&userId=&cache=v2" width="130">
            </div>
            <div class="title-container">
                <h1>HRmate</h1>
                <p>인원 현황 및 자동화 지원 시스템</p>
            </div>
        </div>
        <div class="divider"><hr></div>
    """, unsafe_allow_html=True)

# 비밀번호 인증
def check_password():
    """Returns `True` if the user had the correct password."""

    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state.get("password") == "0314!":
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Don't store password.
        else:
            st.session_state["password_correct"] = False

    # First run or input not cleared.
    if "password_correct" not in st.session_state:
        show_header()
        # 비밀번호 입력 필드를 중앙에 배치
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            st.markdown('<div class="password-input">', unsafe_allow_html=True)
            st.text_input(
                "비밀번호를 입력하세요", type="password", on_change=password_entered, key="password"
            )
            st.markdown('</div>', unsafe_allow_html=True)
        return False
    elif not st.session_state["password_correct"]:
        show_header()
        # Password not correct, show input + error.
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            st.markdown('<div class="password-input">', unsafe_allow_html=True)
            st.text_input(
                "비밀번호를 입력하세요", type="password", on_change=password_entered, key="password"
            )
            st.markdown('</div>', unsafe_allow_html=True)
            st.error("😕 비밀번호가 올바르지 않습니다")
        return False
    else:
        # Password correct.
        return True

# 비밀번호 확인
if not check_password():
    st.stop()  # Do not continue if check_password() returned False.

# 데이터 로드 함수
@st.cache_data(ttl=300)  # 5분마다 캐시 갱신
def load_data():
    try:
        # 엑셀 파일 경로
        file_path = "임직원 기초 데이터.xlsx"
        
        # 파일이 존재하는지 확인
        if not os.path.exists(file_path):
            st.error(f"파일을 찾을 수 없습니다: {file_path}")
            return None
             
        # 파일 수정 시간 확인
        last_modified = os.path.getmtime(file_path)
        
        # 엑셀 파일 읽기
        df = pd.read_excel(file_path)
        
        # 데이터 로드 시간 표시 (한국 시간대 적용)
        st.sidebar.markdown("<br>", unsafe_allow_html=True)
        kst_time = datetime.fromtimestamp(last_modified, pytz.timezone('Asia/Seoul'))
        st.sidebar.markdown(f"*마지막 데이터 업데이트: {kst_time.strftime('%Y년 %m월 %d일 %H:%M')}*")
        
        return df
    except Exception as e:
        st.error(f"파일을 불러오는 중 오류가 발생했습니다: {str(e)}")
        return None

# 날짜 변환 함수 캐싱
@st.cache_data(ttl=3600)  # 1시간 캐시 유지
def convert_date(date_value):
    if pd.isna(date_value):
        return pd.NaT
    try:
        # 엑셀 숫자 형식의 날짜 처리
        if isinstance(date_value, (int, float)):
            return pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(date_value))
        
        # 문자열로 변환
        date_str = str(date_value)
        
        # 여러 날짜 형식 시도
        formats = ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y%m%d']
        for fmt in formats:
            try:
                return pd.to_datetime(date_str, format=fmt)
            except:
                continue
        
        # 모든 형식이 실패하면 기본 변환 시도
        return pd.to_datetime(date_str)
    except:
        return pd.NaT

# 엑셀 다운로드 함수 캐싱
@st.cache_data(ttl=3600)  # 1시간 캐시 유지
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='임직원명부')
    processed_data = output.getvalue()
    return processed_data

# CSS 스타일 추가
st.markdown("""
    <style>
    [data-testid="stMetricValue"] {
        text-align: right;
    }
    .metric-row {
        display: flex;
        justify-content: flex-start;
        align-items: center;
        padding: 15px 30px;
        background-color: #f0f2f6;
        border-radius: 5px;
        margin-bottom: 10px;
        gap: 40px;
        max-width: 680px;
        margin-left: 0;
        margin-right: auto;
    }
    .metric-label {
        font-size: 0.9rem;
        color: #31333F;
        text-align: center;
        min-width: 60px;
        margin-bottom: 5px;
    }
    .metric-value {
        font-size: 1.6rem;
        font-weight: bold;
        color: #31333F;
        text-align: center;
        min-width: 40px;
    }
    .metric-sublabel {
        font-size: 0.8rem;
        color: #666;
        text-align: center;
        margin-top: 5px;
    }
    .total-value {
        color: #1f77b4;
    }
    [data-testid="stDataFrame"] > div {
        display: flex;
        justify-content: center;
    }
    .stRadio [role=radiogroup]{
        padding-top: 0px;
    }
     /* 사이드바 스타일 추가 */
    [data-testid="stSidebar"] {
        min-width: 200px !important;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
        font-size: 0.8rem !important;
    }
    [data-testid="stSidebar"] .stRadio [role="radiogroup"] label {
        font-size: 0.8rem !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
    }
    [data-testid="stSidebar"] a {
        font-size: 0.8rem !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
    }
    [data-testid="stSidebar"] button {
        width: 80% !important;
        margin: 0 auto !important;
        display: block !important;
    }
    </style>
""", unsafe_allow_html=True)

# 제목
st.sidebar.title("👥 HRmate")
st.sidebar.markdown("---")

# HR Data 섹션
st.sidebar.markdown("#### HR Data")
if st.sidebar.button("📊 인원현황", use_container_width=True):
    st.session_state.menu = "📊 인원현황"
if st.sidebar.button("📈 연도별 인원 통계", use_container_width=True):
    st.session_state.menu = "📈 연도별 인원 통계"
if st.sidebar.button("🚀 채용현황", use_container_width=True):
    st.session_state.menu = "🚀 채용현황"
if st.sidebar.button("🔔 인사팀 업무 공유", use_container_width=True):
    st.session_state.menu = "🔔 인사팀 업무 공유"
if st.sidebar.button("😊 임직원 명부", use_container_width=True):
    st.session_state.menu = "😊 임직원 명부"
if st.sidebar.button("🔍 연락처/생일 검색", use_container_width=True):
    st.session_state.menu = "🔍 연락처/생일 검색"


st.sidebar.markdown("#### HR Surpport")
# HR Support 섹션
if st.sidebar.button("🚀 채용 전형관리", use_container_width=True):
    st.session_state.menu = "🚀 채용 전형관리"
if st.sidebar.button("📋 채용 처우협상", use_container_width=True):
    st.session_state.menu = "📋 채용 처우협상"
if st.sidebar.button("🏦 기관제출용 인원현황", use_container_width=True):
    st.session_state.menu = "🏦 기관제출용 인원현황"
if st.sidebar.button("⏰ 초과근무 조회", use_container_width=True):
    st.session_state.menu = "⏰ 초과근무 조회"
if st.sidebar.button("📅 인사발령 내역", use_container_width=True):
    st.session_state.menu = "📅 인사발령 내역"


st.sidebar.markdown("---")
st.sidebar.markdown("<br>", unsafe_allow_html=True)
with st.sidebar.expander("💡 전사지원"):
    st.markdown('<a href="https://neuropr-lwm9mzur3rzbgoqrhzy68n.streamlit.app/" target="_blank" class="sidebar-link" style="text-decoration: none; color: #1b1b1e;">▫️PR(뉴스검색 및 기사초안)</a>', unsafe_allow_html=True)

# 기본 메뉴 설정
if 'menu' not in st.session_state:
    st.session_state.menu = "📊 인원현황"
menu = st.session_state.menu

try:
    # 데이터 로드
    df = load_data()
    
    if df is not None:
        # Excel 날짜 형식 변환 함수
        def convert_excel_date(date_value):
            try:
                if pd.isna(date_value):
                    return pd.NaT
                return pd.to_datetime('1899-12-30') + pd.Timedelta(days=int(date_value))
            except:
                return pd.to_datetime(date_value, errors='coerce')

        # 날짜 컬럼 변환
        date_columns = ['정규직전환일', '퇴사일', '생년월일', '입사일']
        for col in date_columns:
            if col in df.columns:
                df[col] = df[col].apply(convert_excel_date)
        
        # 연도 컬럼 미리 생성
        if '정규직전환일' in df.columns:
            df['정규직전환연도'] = df['정규직전환일'].dt.year
        if '퇴사일' in df.columns:
            df['퇴사연도'] = df['퇴사일'].dt.year
        
        if menu == "📊 인원현황":
            # 기본통계 분석
            st.markdown("##### 📊 인원현황")
            
            # 조회 기준일 선택
            query_date = st.date_input(
                "조회 기준일",
                value=datetime.now().date(),
                help="선택한 날짜 기준으로 인원현황을 조회합니다.",
                key="query_date_input",
                label_visibility="visible"
            )
            st.markdown(
                """
                <style>
                div[data-testid="stDateInput"] {
                    width: 200px;
                }
                </style>
                """,
                unsafe_allow_html=True
            )

            # 기준일자로 재직자 필터링
            재직자 = len(df[
                (df['입사일'].dt.date <= query_date) & 
                ((df['퇴사일'].isna()) | (df['퇴사일'].dt.date >= query_date))
            ])
            
            # 해당 연도의 입퇴사자 계산
            selected_year = query_date.year
            정규직_입사자 = len(df[(df['입사일'].dt.year == selected_year) & (df['고용구분'] == '정규직') & (df['입사일'].dt.date <= query_date)])
            정규직_퇴사자 = len(df[(df['퇴사일'].dt.year == selected_year) & (df['고용구분'] == '정규직') & (df['퇴사일'].dt.date <= query_date)])
            계약직_입사자 = len(df[(df['입사일'].dt.year == selected_year) & (df['고용구분'] == '계약직') & (df['입사일'].dt.date <= query_date)])
            계약직_퇴사자 = len(df[(df['퇴사일'].dt.year == selected_year) & (df['고용구분'] == '계약직') & (df['퇴사일'].dt.date <= query_date)])
            
            # 퇴사율 계산 (소수점 첫째자리까지)
            재직_정규직_수 = len(df[
                (df['고용구분'] == '정규직') & 
                (df['입사일'].dt.date <= query_date) & 
                ((df['퇴사일'].isna()) | (df['퇴사일'].dt.date > query_date))
            ])
            퇴사율 = round((정규직_퇴사자 / 재직_정규직_수 * 100), 1) if 재직_정규직_수 > 0 else 0
            
            # 통계 표시
            st.markdown(
                f"""
                <div class="metric-row">
                    <div>
                        <div class="metric-label">전체</div>
                        <div class="metric-value total-value">{재직자:,}</div>
                        <div class="metric-sublabel">재직자</div>
                    </div>
                    <div style="width: 2px; background-color: #ddd;"></div>
                    <div style="min-width: 100px;">
                        <div class="metric-label">정규직</div>
                        <div style="display: flex; justify-content: space-between; gap: 20px;">
                            <div>
                                <div class="metric-value">{정규직_입사자}</div>
                                <div class="metric-sublabel">입사자</div>
                            </div>
                            <div>
                                <div class="metric-value">{정규직_퇴사자}</div>
                                <div class="metric-sublabel">퇴사자</div>
                            </div>
                        </div>
                    </div>
                    <div style="width: 2px; background-color: #ddd;"></div>
                    <div style="min-width: 100px;">
                        <div class="metric-label">계약직</div>
                        <div style="display: flex; justify-content: space-between; gap: 20px;">
                            <div>
                                <div class="metric-value" style="color: #666;">{계약직_입사자}</div>
                                <div class="metric-sublabel">입사자</div>
                            </div>
                            <div>
                                <div class="metric-value" style="color: #666;">{계약직_퇴사자}</div>
                                <div class="metric-sublabel">퇴사자</div>
                            </div>
                        </div>
                    </div>
                    <div style="width: 2px; background-color: #ddd;"></div>
                    <div>
                        <div class="metric-label">퇴사율</div>
                        <div class="metric-value" style="color: #ff0000;">{퇴사율}%</div>
                        <div class="metric-sublabel">정규직 {재직_정규직_수}명</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

            st.markdown("<br>", unsafe_allow_html=True)
            
            # 3개의 컬럼 생성 (0.4:0.4:0.2 비율)
            col1, col2, col3 = st.columns([0.4, 0.3, 0.3])
            
            # 현재 재직자 필터링 (조회 기준일 기준)
            current_employees = df[
                (df['입사일'].dt.date <= query_date) & 
                ((df['퇴사일'].isna()) | (df['퇴사일'].dt.date >= query_date))
            ]
            
            with col1:
                # 본부별 인원 현황
                dept_counts = current_employees['본부'].value_counts().reset_index()
                dept_counts.columns = ['본부', '인원수']
                
                # 본부별 그래프 (수평 막대 그래프)
                fig_dept = px.bar(
                    dept_counts,
                    y='본부',
                    x='인원수',
                    title="본부별",
                    width=400,
                    height=300,
                    orientation='h'  # 수평 방향으로 변경
                )
                fig_dept.update_traces(
                    marker_color='#FF4B4B',
                    text=dept_counts['인원수'],
                    textposition='outside',
                    textfont=dict(size=14)
                )
                fig_dept.update_layout(
                    showlegend=False,
                    title_x=0.5,
                    title_y=0.95,
                    margin=dict(t=50, r=50),  # 오른쪽 여백 추가
                    xaxis=dict(
                        title="",
                        range=[0, max(dept_counts['인원수']) * 1.2]
                    ),
                    yaxis=dict(
                        title="",
                        autorange="reversed"  # 위에서 아래로 정렬
                    )
                )
                st.plotly_chart(fig_dept, use_container_width=True)
            
            with col2:
                # 직책별 인원 현황
                position_order = ['C-LEVEL', '실리드', '팀리드', '멤버', '계약직']
                position_counts = current_employees['직책'].value_counts()
                position_counts = pd.Series(position_counts.reindex(position_order).fillna(0))
                position_counts = position_counts.reset_index()
                position_counts.columns = ['직책', '인원수']
                
                # 직책별 그래프
                fig_position = px.area(
                    position_counts,
                    x='직책',
                    y='인원수',
                    title="직책별",
                    width=400,
                    height=300
                )
                fig_position.update_traces(
                    fill='tonexty',
                    line=dict(color='#666666'),
                    text=position_counts['인원수'],
                    textposition='top center'
                )
                fig_position.update_layout(
                    showlegend=False,
                    title_x=0.5,
                    title_y=0.95,
                    margin=dict(t=50),
                    yaxis=dict(range=[0, max(position_counts['인원수']) * 1.2])
                )
                st.plotly_chart(fig_position, use_container_width=True)
            
            with col3:
                # 성별 비율 계산 (조회 기준일 기준)
                gender_counts = current_employees['남/여'].value_counts()
                gender_percentages = (gender_counts / len(current_employees) * 100).round(1)
                
                # 도넛 차트 생성
                fig = go.Figure(data=[go.Pie(
                    labels=['남', '여'],
                    values=[gender_percentages['남'], gender_percentages['여']],
                    hole=0.4,
                    marker_colors=['#4A4A4A', '#FF4B4B'],
                    textinfo='label+percent',
                    textposition='inside',
                    showlegend=False,
                    textfont=dict(color='white')  # 텍스트 색상을 흰색으로 설정
                )])
                
                fig.update_layout(
                    title="성별",
                    title_x=0.4,
                    title_y=0.95,
                    width=220,
                    height=220,
                    margin=dict(t=50, b=0, l=0, r=0),  # 제목을 위한 상단 여백 추가
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)'
                )
                
                st.plotly_chart(fig)

            st.markdown("<br>", unsafe_allow_html=True)
            
            # 2025년 입퇴사자 현황
            list_col1, list_col2 = st.columns(2)
            
            with list_col1:
                st.markdown("###### 2025년 입사자")
                입사자_df = df[df['입사일'].dt.year == 2025][['성명', '팀', '직위', '입사일']]
                if not 입사자_df.empty:
                    입사자_df = 입사자_df.sort_values('입사일', ascending=False)  # 내림차순 정렬
                    입사자_df = 입사자_df.reset_index(drop=True)
                    입사자_df.index = 입사자_df.index + 1
                    입사자_df = 입사자_df.rename_axis('No.')
                    st.dataframe(입사자_df.style.format({'입사일': lambda x: x.strftime('%Y-%m-%d')}),
                               use_container_width=True)
                else:
                    st.info("2025년 입사 예정자가 없습니다.")

            with list_col2:
                st.markdown("###### 2025년 퇴사자")
                퇴사자_df = df[df['퇴사연도'] == 2025][['성명', '팀', '직위', '퇴사일']]
                if not 퇴사자_df.empty:
                    퇴사자_df = 퇴사자_df.sort_values('퇴사일', ascending=False)  # 내림차순 정렬
                    퇴사자_df = 퇴사자_df.reset_index(drop=True)
                    퇴사자_df.index = 퇴사자_df.index + 1
                    퇴사자_df = 퇴사자_df.rename_axis('No.')
                    st.dataframe(퇴사자_df.style.format({'퇴사일': lambda x: x.strftime('%Y-%m-%d')}),
                               use_container_width=True)
                else:
                    st.info("2025년 퇴사자가 없습니다.")
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            # 근속기간별 퇴사자 현황 분석
            st.markdown("##### 퇴사자 현황_정규직")
            
            # 퇴사연도 선택 드롭다운과 퇴사인원 표시를 위한 컬럼 생성
            col1, col2 = st.columns([2, 1])
            
            with col1:
                # 퇴사연도 선택 드롭다운
                available_years = sorted(df[df['재직상태'] == '퇴직']['퇴사연도'].dropna().astype(int).unique())
                default_index = list(['전체'] + list(available_years)).index(2025) if 2025 in available_years else 0
                selected_year = st.selectbox(
                    "퇴사연도 선택",
                    options=['전체'] + list(available_years),
                    index=default_index,
                    key='tenure_year_select'
                )
            
            with col2:
                # 선택된 연도의 퇴사인원 계산
                if selected_year == '전체':
                    퇴사인원 = len(df[(df['재직상태'] == '퇴직') & (df['고용구분'] == '정규직')])
                else:
                    퇴사인원 = len(df[(df['재직상태'] == '퇴직') & (df['퇴사연도'] == selected_year) & (df['고용구분'] == '정규직')])
                
                st.markdown(
                    f"""
                    <div style="padding: 0.5rem; margin-top: 1.6rem;">
                        <span style="font-size: 1rem; color: #666;">정규직 퇴사인원: </span>
                        <span style="font-size: 1.2rem; font-weight: bold; color: #FF0000;">{퇴사인원:,}명</span>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
            
            # 그래프를 위한 컬럼 생성 (60:40 비율)
            graph_col, space_col = st.columns([0.5, 0.5])
            
            with graph_col:
                def calculate_tenure_months(row):
                    if pd.isna(row['입사일']) or pd.isna(row['퇴사일']):
                        return None
                    tenure = row['퇴사일'] - row['입사일']
                    return tenure.days / 30.44  # 평균 한 달을 30.44일로 계산

                # 근속기간 계산
                df['근속월수'] = df.apply(calculate_tenure_months, axis=1)

                # 근속기간 구간 설정
                def get_tenure_category(months):
                    if pd.isna(months):
                        return None
                    elif months <= 5:
                        return "0~5개월"
                    elif months <= 11:
                        return "6~11개월"
                    elif months <= 24:
                        return "1년~2년"
                    elif months <= 36:
                        return "2년~3년"
                    else:
                        return "3년이상"

                df['근속기간_구분'] = df['근속월수'].apply(get_tenure_category)

                # 퇴직자 데이터 필터링
                퇴직자_df = df[(df['재직상태'] == '퇴직') & (df['고용구분'] == '정규직')]
                if selected_year != '전체':
                    퇴직자_df = 퇴직자_df[퇴직자_df['퇴사연도'] == selected_year]
                
                # 근속기간별 인원 집계
                tenure_counts = 퇴직자_df['근속기간_구분'].value_counts().reindex(["0~5개월", "6~11개월", "1년~2년", "2년~3년", "3년이상"], fill_value=0)

                # 그래프 생성
                fig = go.Figure()
                
                # 막대 색상 설정
                colors = ['#E0E0E0', '#E0E0E0', '#E0E0E0', '#FF0000', '#FF0000']
                
                fig.add_trace(go.Bar(
                    x=tenure_counts.index,
                    y=tenure_counts.values,
                    marker_color=colors,
                    text=tenure_counts.values,
                    textposition='outside',
                ))

                # 레이아웃 설정
                title_text = f"{'전체 기간' if selected_year == '전체' else str(selected_year) + '년'} 근속기간별 퇴사자 현황"
                fig.update_layout(
                    height=300,
                    showlegend=False,
                    plot_bgcolor='white',
                    yaxis=dict(
                        title="퇴사자 수 (명)",
                        range=[0, max(max(tenure_counts.values) * 1.2, 10)],
                        gridcolor='lightgray',
                        gridwidth=0.5,
                    ),
                    xaxis=dict(
                        showgrid=False,
                    ),
                    margin=dict(t=50, b=20)  # 하단 여백을 20으로 줄임
                )

                st.plotly_chart(fig, use_container_width=True)

            with space_col:
                st.write("")  # 빈 공간
            
            # 부서별 근속기간 분석
            본부별_근속기간 = pd.pivot_table(
                퇴직자_df,
                values='사번',
                index='본부',
                columns='근속기간_구분',
                aggfunc='count',
                fill_value=0
            ).reindex(columns=["0~5개월", "6~11개월", "1년~2년", "2년~3년", "3년이상"])

            # 재직자 수 계산
            재직자_수 = df[df['재직상태'] == '재직'].groupby('본부')['사번'].count()

            # 퇴직자 수 계산 - 선택된 연도에 따라 필터링
            if selected_year == '전체':
                퇴직자_수 = df[(df['재직상태'] == '퇴직') & (df['고용구분'] == '정규직')].groupby('본부')['사번'].count()
            else:
                퇴직자_수 = df[(df['재직상태'] == '퇴직') & (df['고용구분'] == '정규직') & (df['퇴사연도'] == selected_year)].groupby('본부')['사번'].count()

            # 퇴사율 계산
            본부별_퇴사율 = (퇴직자_수 / (재직자_수 + 퇴직자_수) * 100).round(1)

            # 조기퇴사율 계산 (1년 미만 퇴사자)
            조기퇴사자_수 = 본부별_근속기간[["0~5개월", "6~11개월"]].sum(axis=1)
            조기퇴사율 = (조기퇴사자_수 / (재직자_수 + 퇴직자_수) * 100).round(1)

            # 결과 테이블 생성
            result_df = pd.DataFrame({
                '0~5개월': 본부별_근속기간["0~5개월"],
                '6~11개월': 본부별_근속기간["6~11개월"],
                '1년~2년': 본부별_근속기간["1년~2년"],
                '2년~3년': 본부별_근속기간["2년~3년"],
                '3년이상': 본부별_근속기간["3년이상"],
                '퇴직인원': 퇴직자_수,
                '재직인원': 재직자_수,
                '퇴사율': 본부별_퇴사율.fillna(0).map('{:.1f}%'.format),
                '조기퇴사율': 조기퇴사율.fillna(0).map('{:.1f}%'.format),
                '퇴사율 비중': 본부별_퇴사율.fillna(0).map('{:.1f}%'.format)
            }).fillna(0)

            # 합계 행 추가
            total_row = pd.Series({
                '0~5개월': result_df['0~5개월'].sum(),
                '6~11개월': result_df['6~11개월'].sum(),
                '1년~2년': result_df['1년~2년'].sum(),
                '2년~3년': result_df['2년~3년'].sum(),
                '3년이상': result_df['3년이상'].sum(),
                '퇴직인원': result_df['퇴직인원'].sum(),
                '재직인원': result_df['재직인원'].sum(),
                '퇴사율': f"{(result_df['퇴직인원'].sum() / (result_df['재직인원'].sum() + result_df['퇴직인원'].sum()) * 100):.1f}%",
                '조기퇴사율': f"{(result_df['0~5개월'].sum() + result_df['6~11개월'].sum()) / (result_df['재직인원'].sum() + result_df['퇴직인원'].sum()) * 100:.1f}%",
                '퇴사율 비중': f"{(result_df['퇴직인원'].sum() / (result_df['재직인원'].sum() + result_df['퇴직인원'].sum()) * 100):.1f}%"
            }, name='총합계')

            result_df = pd.concat([result_df, pd.DataFrame(total_row).T])

            # 스타일이 적용된 테이블 표시
            st.markdown(
                """
                <style>
                .custom-table {
                    font-size: 12px;
                    width: 80%;
                    border-collapse: collapse;
                }
                .custom-table th {
                    background-color: #f0f2f6;
                    padding: 7px;
                    text-align: center;
                    border: 1px solid #ddd;
                }
                .custom-table td {
                    padding: 5px;
                    text-align: center;
                    border: 1px solid #ddd;
                }
                .custom-table tr:last-child {
                    background-color: #f0f2f6;
                    font-weight: bold;
                }
                .red-text {
                    color: red;
                }
                </style>
                """,
                unsafe_allow_html=True
            )

            # 테이블 HTML 생성
            table_html = "<table class='custom-table'><tr><th>구분</th>"
            for col in result_df.columns:
                table_html += f"<th>{col}</th>"
            table_html += "</tr>"

            for idx, row in result_df.iterrows():
                table_html += f"<tr><td>{idx}</td>"
                for col in result_df.columns:
                    value = row[col]
                    if isinstance(value, (int, float)):
                        if col in ['0~5개월', '6~11개월', '1년~2년', '2년~3년', '3년이상', '퇴직인원', '재직인원']:
                            table_html += f"<td>{int(value)}</td>"
                        else:
                            table_html += f"<td>{value}</td>"
                    else:
                        if '%' in str(value) and float(str(value).rstrip('%')) > 0:
                            table_html += f"<td class='red-text'>{value}</td>"
                        else:
                            table_html += f"<td>{value}</td>"
                table_html += "</tr>"
            table_html += "</table>"

            st.markdown(table_html, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

        elif menu == "📈 연도별 인원 통계":
            # 최근 5년간 인원 현황 분석
            st.markdown("##### 📈 연도별 인원 통계")
            
            def get_year_end_headcount(df, year):
                # 해당 연도 말일 설정
                year_end = pd.Timestamp(f"{year}-12-31")
                
                # 해당 연도 말일 기준 재직자 수 계산
                # 입사일이 연도 말일 이전이고, 퇴사일이 없거나 연도 말일과 같거나 이후인 직원
                year_end_employees = df[
                    (df['입사일'] <= year_end) & 
                    ((df['퇴사일'].isna()) | (df['퇴사일'] >= year_end))
                ]
                
                # 전체 인원
                total = len(year_end_employees)
                
                # 정규직/계약직 인원
                regular = len(year_end_employees[year_end_employees['고용구분'] == '정규직'])
                contract = len(year_end_employees[year_end_employees['고용구분'] == '계약직'])
                
                return total, regular, contract
            
            # 연도별 입/퇴사 인원 계산 함수 (get_year_end_headcount 함수 다음에 추가)
            @st.cache_data(ttl=3600)  # 1시간 캐시 유지
            def get_year_employee_stats(df, year):
                # 정규직 입사
                reg_join = len(df[(df['고용구분'] == '정규직') & 
                                  (df['입사일'].dt.year == year)])
                
                # 정규직 퇴사
                reg_leave = len(df[(df['고용구분'] == '정규직') & 
                                   (df['퇴사일'].dt.year == year)])
                
                # 계약직 입사
                contract_join = len(df[(df['고용구분'] == '계약직') & 
                                      (df['입사일'].dt.year == year)])
                
                # 계약직 퇴사
                contract_leave = len(df[(df['고용구분'] == '계약직') & 
                                       (df['퇴사일'].dt.year == year)])
                
                return reg_join, reg_leave, contract_join, contract_leave
            
            # stats_df 생성 부분을 다음과 같이 수정
            stats_df = pd.DataFrame([
                {
                    '연도': year,
                    '전체': get_year_end_headcount(df, year)[0],
                    '정규직_전체': get_year_end_headcount(df, year)[1],
                    '계약직_전체': get_year_end_headcount(df, year)[2],
                    '정규직_입사': get_year_employee_stats(df, year)[0],
                    '정규직_퇴사': get_year_employee_stats(df, year)[1],
                    '계약직_입사': get_year_employee_stats(df, year)[2],
                    '계약직_퇴사': get_year_employee_stats(df, year)[3]
                }
                for year in range(2021, 2026)  # 2021년부터 2025년까지
            ])
            
            # 그래프를 위한 컬럼 생성 (50:50 비율)
            graph_col1, space_col1,  graph_col2, space_col2 = st.columns([0.35,0.05, 0.35, 0.2])
            
            with graph_col1:
                # 전체 인원 그래프 생성
                fig = go.Figure()
                
                fig.add_trace(go.Scatter(
                    x=stats_df['연도'],
                    y=stats_df['전체'],
                    mode='lines+markers+text',
                    name='전체 인원',
                    text=stats_df['전체'],
                    textposition='top center',
                    line=dict(color='#FF4B4B', width=3),
                    marker=dict(size=10)
                ))

                fig.update_layout(
                    title="전체 인원",
                    title_x=0,
                    height=350,
                    showlegend=False,
                    plot_bgcolor='white',
                    yaxis=dict(
                        title="인원 수 (명)",
                        gridcolor='lightgray',
                        gridwidth=0.5,
                        range=[0, max(stats_df['전체']) * 1.2]
                    ),
                    xaxis=dict(
                        showgrid=False,
                        tickformat='d'  # 정수 형식으로 표시
                    ),
                    margin=dict(t=50)
                )

                st.plotly_chart(fig, use_container_width=True)

            with space_col1:
                st.write("")  # 빈 공간

            with graph_col2:
                # 정규직/계약직 막대 그래프 생성
                fig2 = go.Figure()

                # 정규직 막대
                fig2.add_trace(go.Bar(
                    x=stats_df['연도'],
                    y=stats_df['정규직_전체'],
                    name='정규직',
                    text=stats_df['정규직_전체'],
                    textposition='auto',
                    textfont=dict(color='white'),
                    marker_color='#FF4B4B'
                ))

                # 계약직 막대
                fig2.add_trace(go.Bar(
                    x=stats_df['연도'],
                    y=stats_df['계약직_전체'],
                    name='계약직',
                    text=stats_df['계약직_전체'],
                    textposition='auto',
                    marker_color='#FFB6B6'
                ))

                fig2.update_layout(
                    title="고용형태별 인원",
                    title_x=0,
                    height=350,
                    barmode='stack',
                    plot_bgcolor='white',
                    yaxis=dict(
                        gridcolor='lightgray',
                        gridwidth=0.5,
                        range=[0, max(stats_df['전체']) * 1.2]
                    ),
                    xaxis=dict(
                        showgrid=False,
                        tickformat='d'  # 정수 형식으로 표시
                    ),
                    margin=dict(t=50),
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1
                    )
                )

                st.plotly_chart(fig2, use_container_width=True)

            with space_col2:
                st.write("")  # 빈 공간
            
            # DataFrame을 직접 표시
            st.dataframe(
                stats_df.rename(columns={
                    '연도': '연도',
                    '전체': '전체 인원',
                    '정규직_전체': '정규직\n전체',
                    '계약직_전체': '계약직\n전체',
                    '정규직_입사': '정규직\n입사',
                    '정규직_퇴사': '정규직\n퇴사',
                    '계약직_입사': '계약직\n입사',
                    '계약직_퇴사': '계약직\n퇴사'
                }).style.format({
                    '연도': '{:,.0f}',
                    '전체 인원': '{:,.0f}',
                    '정규직\n전체': '{:,.0f}',
                    '계약직\n전체': '{:,.0f}',
                    '정규직\n입사': '{:,.0f}',
                    '정규직\n퇴사': '{:,.0f}',
                    '계약직\n입사': '{:,.0f}',
                    '계약직\n퇴사': '{:,.0f}'
                }).set_properties(**{
                    'text-align': 'center',
                    'vertical-align': 'middle'
                }).set_table_styles([
                    {'selector': 'th', 'props': [('text-align', 'center')]},
                    {'selector': 'td', 'props': [('text-align', 'center')]}
                ]),
                hide_index=True,
                width=800,
                use_container_width=False
            )

        elif menu == "🔍 연락처/생일 검색":
            st.markdown("##### 🔍 연연락처/생일 검색")
            
            # 검색 부분을 컬럼으로 나누기
            search_col, space_col = st.columns([0.3, 0.7])
            
            with search_col:
                st.markdown('<div class="search-container">', unsafe_allow_html=True)
                search_name = st.text_input("성명으로 검색", key="contact_search")
                st.markdown('</div>', unsafe_allow_html=True)
            
            if search_name:
                contact_df = df[df['성명'].str.contains(search_name, na=False)]
                if not contact_df.empty:
                    st.markdown("""
                        <style>
                        .dataframe {
                            text-align: left !important;
                        }
                        .dataframe td, .dataframe th {
                            text-align: left !important;
                        }
                        </style>
                    """, unsafe_allow_html=True)
                    
                    # 생년월일 컬럼을 포함하여 표시할 컬럼 선택
                    contact_info = contact_df[['성명', '생년월일', '본부', '팀', '직위', 'E-Mail', '핸드폰', '주소']].reset_index(drop=True)
                    
                    # 생년월일 형식 변환 (datetime 형식으로 변환 후 YYYY-MM-DD 형식으로 표시)
                    contact_info['생년월일'] = pd.to_datetime(contact_info['생년월일']).dt.strftime('%Y-%m-%d')
                    
                    contact_info.index = contact_info.index + 1
                    contact_info = contact_info.rename_axis('No.')
                    st.dataframe(contact_info.style.set_properties(**{'text-align': 'left'}), use_container_width=True)
                else:
                    st.info("검색 결과가 없습니다.")

            st.markdown("---")

            # 생일자 검색
            st.markdown("##### 🎂이달의 생일자")
            current_month = datetime.now(pytz.timezone('Asia/Seoul')).month
            birth_month = st.selectbox(
                "생일 월 선택",
                options=list(range(1, 13)),
                format_func=lambda x: f"{x}월",
                index=current_month - 1
            )
            
            if birth_month:
                birthday_df = df[(df['재직상태'] == '재직') & 
                               (pd.to_datetime(df['생년월일']).dt.month == birth_month)]
                if not birthday_df.empty:
                    today = pd.Timestamp.now()
                    birthday_info = birthday_df[['성명', '본부', '팀', '직위', '입사일']].copy()
                    birthday_info['근속기간'] = (today - birthday_info['입사일']).dt.days // 365
                    birthday_info['생일'] = pd.to_datetime(birthday_df['생년월일']).dt.strftime('%m-%d')
                    
                    birthday_info = birthday_info[['성명', '본부', '팀', '생일', '근속기간']]
                    birthday_info = birthday_info.sort_values('생일')
                    
                    birthday_info['근속기간'] = birthday_info['근속기간'].astype(str) + '년'
                    
                    birthday_info = birthday_info.reset_index(drop=True)
                    birthday_info.index = birthday_info.index + 1
                    birthday_info = birthday_info.rename_axis('No.')
                    
                    st.dataframe(birthday_info, use_container_width=True)
                else:
                    st.info(f"{birth_month}월 재직자 중 생일자가 없습니다.")

        elif menu == "🏦 기관제출용 인원현황":
            st.markdown("##### 🏦 기관제출용 인원현황")
            
            # 데이터 로드
            df = load_data()
            if df is not None:
                # 날짜 컬럼 변환 함수
                def convert_date(date_value):
                    if pd.isna(date_value):
                        return pd.NaT
                    try:
                        # 엑셀 숫자 형식의 날짜 처리
                        if isinstance(date_value, (int, float)):
                            return pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(date_value))
                        
                        # 문자열로 변환
                        date_str = str(date_value)
                        
                        # 여러 날짜 형식 시도
                        formats = ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y%m%d']
                        for fmt in formats:
                            try:
                                return pd.to_datetime(date_str, format=fmt)
                            except:
                                continue
                        
                        # 모든 형식이 실패하면 기본 변환 시도
                        return pd.to_datetime(date_str)
                    except:
                        return pd.NaT

                # 날짜 컬럼 변환
                df['입사일'] = df['입사일'].apply(convert_date)
                df['퇴사일'] = df['퇴사일'].apply(convert_date)
                
                
                # 조회 기준일 설정
                current_year = datetime.now().year
                current_month = datetime.now().month
                years = list(range(2016, current_year + 1))
                years.sort(reverse=True)  # 내림차순 정렬
                
                col1, col2, col3 = st.columns([0.3, 0.3, 0.4])
                with col1:
                    selected_year = st.selectbox("조회년도", years, index=0)
                with col2:
                    months = list(range(1, 13))
                    selected_month = st.selectbox("조회월", months, index=current_month-1)
                with col3:
                    st.write("")  # 공백 컬럼
                
                # 선택된 년월의 마지막 날짜 계산
                last_day = pd.Timestamp(f"{selected_year}-{selected_month:02d}-01") + pd.offsets.MonthEnd(0)
                               
                # 기준일에 재직중인 직원 필터링
                current_employees = df[
                    (df['입사일'].notna()) & 
                    (df['입사일'] <= last_day) & 
                    ((df['퇴사일'].isna()) | 
                     (df['퇴사일'] == pd.Timestamp('2050-12-31')) | 
                     (df['퇴사일'] >= last_day))
                ]
                
                st.markdown("---")
                
                if not df[df['입사일'] <= last_day].empty:
                    # 구분별 인원 현황 계산 및 표시
                    # 구분1: 주주간담회 등 IR팀 자료
                    st.markdown("1. 주주간담회 등 IR팀 자료 작성용")
                    group1_stats = current_employees['구분1'].value_counts().reset_index()
                    group1_stats.columns = ['구분', '인원수']
                    total_count1 = group1_stats['인원수'].sum()
                    
                    # '임원'이 있는 행을 찾아서 첫 번째로 이동
                    임원_row = group1_stats[group1_stats['구분'] == '임원']
                    other_rows = group1_stats[group1_stats['구분'] != '임원']
                    group1_stats = pd.concat([임원_row, other_rows]).reset_index(drop=True)
                    
                    group1_stats = group1_stats.T  # 행과 열을 바꿈
                    group1_stats.columns = group1_stats.iloc[0]  # 첫 번째 행을 컬럼명으로 설정
                    group1_stats = group1_stats.iloc[1:]  # 첫 번째 행 제외
                    group1_stats['총인원'] = total_count1  # 총인원 열 추가
                    st.dataframe(
                        group1_stats,
                        use_container_width=False,
                        width=900,
                        column_config={col: st.column_config.NumberColumn(col, width=50) for col in group1_stats.columns}
                    )
                    
                    # 구분2: 투자자 사업현황 보고1
                    st.markdown("2. 투자자 사업현황 보고")
                    group2_stats = current_employees['구분2'].value_counts().reset_index()
                    group2_stats.columns = ['구분', '인원수']
                    total_count2 = group2_stats['인원수'].sum()
                    
                    # '임원'이 있는 행을 찾아서 첫 번째로 이동
                    임원_row = group2_stats[group2_stats['구분'] == '임원']
                    other_rows = group2_stats[group2_stats['구분'] != '임원']
                    group2_stats = pd.concat([임원_row, other_rows]).reset_index(drop=True)
                    
                    group2_stats = group2_stats.T  # 행과 열을 바꿈
                    group2_stats.columns = group2_stats.iloc[0]  # 첫 번째 행을 컬럼명으로 설정
                    group2_stats = group2_stats.iloc[1:]  # 첫 번째 행 제외
                    group2_stats['총인원'] = total_count2  # 총인원 열 추가
                    st.dataframe(
                        group2_stats,
                        use_container_width=False,
                        width=600,
                        column_config={col: st.column_config.NumberColumn(col, width=50) for col in group2_stats.columns}
                    )
                    
                    # 구분3: 의료기기 생산 및 수출·수입·수리실적보고
                    st.markdown("3. 의료기기 생산 및 수출·수입·수리실적보고")
                    group3_stats = current_employees['구분3'].value_counts().reset_index()
                    group3_stats.columns = ['구분', '인원수']
                    total_count3 = group3_stats['인원수'].sum()
                    
                    # '임원'이 있는 행을 찾아서 첫 번째로 이동
                    임원_row = group3_stats[group3_stats['구분'] == '임원']
                    other_rows = group3_stats[group3_stats['구분'] != '임원']
                    group3_stats = pd.concat([임원_row, other_rows]).reset_index(drop=True)
                    
                    group3_stats = group3_stats.T  # 행과 열을 바꿈
                    group3_stats.columns = group3_stats.iloc[0]  # 첫 번째 행을 컬럼명으로 설정
                    group3_stats = group3_stats.iloc[1:]  # 첫 번째 행 제외
                    group3_stats['총인원'] = total_count3  # 총인원 열 추가
                    st.dataframe(
                        group3_stats,
                        use_container_width=False,
                        width=700,
                        column_config={col: st.column_config.NumberColumn(col, width=50) for col in group3_stats.columns}
                    )
                    
                    # 인원상세 목록
                    st.markdown("###### 🧑 인원상세")
                    detail_columns = ['성명', '본부', '실', '팀', '고용구분', '입사일', '재직상태', '남/여', '구분1', '구분2', '구분3']
                    detail_df = current_employees[detail_columns].copy()
                    detail_df['입사일'] = detail_df['입사일'].dt.strftime('%Y-%m-%d')
                    
                    # 인덱스를 1부터 시작하는 번호로 리셋
                    detail_df = detail_df.reset_index(drop=True)
                    detail_df.index = detail_df.index + 1
                    detail_df.index.name = 'No'
                    detail_df = detail_df.reset_index()
                    
                    st.dataframe(
                        detail_df,
                        hide_index=True,
                        column_config={
                            "No": st.column_config.NumberColumn("No", width=50),
                            "성명": st.column_config.TextColumn("성명", width=80),
                            "본부": st.column_config.TextColumn("본부", width=120),
                            "실": st.column_config.TextColumn("실", width=120),
                            "팀": st.column_config.TextColumn("팀", width=120),
                            "고용구분": st.column_config.TextColumn("고용구분", width=80),
                            "입사일": st.column_config.TextColumn("입사일", width=100),
                            "재직상태": st.column_config.TextColumn("재직상태", width=80),
                            "성별": st.column_config.TextColumn("남/여", width=60),
                            "구분1": st.column_config.TextColumn("구분1", width=120),
                            "구분2": st.column_config.TextColumn("구분2", width=120),
                            "구분3": st.column_config.TextColumn("구분3", width=120)
                        }
                    )
                    
                    # 엑셀 다운로드 버튼
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        detail_df.to_excel(writer, index=False)
                    excel_data = output.getvalue()
                    st.download_button(
                        label="📥 엑셀 다운로드",
                        data=excel_data,
                        file_name=f"기관제출용_인원현황_{selected_year}{selected_month:02d}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning(f"{selected_year}년 {selected_month}월 데이터가 없습니다.")
            else:
                st.error("데이터를 불러오는 중 오류가 발생했습니다.")

        elif menu == "📋 채용 처우협상":
            st.markdown("##### 🔎 처우 기본정보")
            
            # 직군 매핑 정의
            job_mapping = {
                "연구직": "직군1",
                "개발직": "직군2",
                "임상연구, QA, 인증(RA)": "직군2",
                "연구기획": "직군3",
                "디자인": "직군3",
                "SV, SCM": "직군3",
                "마케팅": "직군3",
                "기획": "직군3",
                "기술영업 / SE(5년 이상)": "직군3",
                "경영기획(전략,회계,인사,재무,법무,보안)": "직군3",
                "지원(연구, 기술, 경영 지원 등)": "직군4",
                "일반영업 /SE(5년 미만)": "직군4",
                "고객지원(CS)": "직군5",
                "레이블링": "직군5"
            }
            
            # 직군 상세 목록
            job_roles = list(job_mapping.keys())
            # 경력입력 폼 생성
            with st.form("experience_form"):
                experience_text = st.text_area("경력기간 입력 (이력서의 날짜 부분을 복사해서 붙여주세요.)", 
                                             help="# 날짜 패턴 : # 날짜 패턴 : 2023. 04, 2024.05.01, 2024.05, 2024-05, 2024-05-01, 2024/05, 2024/05/01, 2023/05, 2015.01.")
                
                # 경력기간 조회 버튼 추가
                experience_submitted = st.form_submit_button("경력기간 조회")
                
                if experience_submitted and experience_text:
                    try:
                        # 경력기간 계산
                        experience_result = calculate_experience(experience_text)
                        if experience_result:
                            # 경력기간과 총 경력기간 분리
                            experience_lines = experience_result.split('\n')
                            total_experience = experience_lines[-1]  # 마지막 줄이 총 경력기간
                            experience_periods = experience_lines[:-2]  # 마지막 두 줄(총 경력기간과 빈 줄) 제외
                            
                            # 총 경력기간을 소수점으로 변환
                            total_match = re.search(r'총 경력기간: (\d+)년 (\d+)개월', total_experience)
                            if total_match:
                                years, months = map(int, total_match.groups())
                                total_years = years + months / 12
                                total_experience = f"총 경력기간: {total_years:.1f}년"
                            
                            # 경력기간 표시
                            st.markdown(f"**{total_experience}**")
                            st.markdown("**경력기간:**")
                            for period in experience_periods:
                                st.markdown(period)
                        else:
                            st.markdown("**경력기간:** 경력 정보가 없습니다.")
                            st.session_state['years'] = 0.0
                        # 인정경력(년) 필드 업데이트
                        st.query_params["years"] = float(f"{total_years:.1f}")
                    except Exception as e:
                        st.error(f"경력기간 계산 중 오류가 발생했습니다: {str(e)}")

            # 입력 폼 생성
            with st.form("salary_form"):
                # 1줄: 포지션명, 후보자명
                col1, col2, col3 = st.columns(3)
                with col1:
                    position = st.text_input("포지션명", "")
                with col2:
                    candidate_name = st.text_input("후보자명", "")
                with col3:
                    job_role = st.selectbox("직군 선택", job_roles)
                
                # 2줄: 현재연봉, 기타 처우, 희망연봉
                col4, col5, col6, col7 = st.columns(4)
                with col4:
                    current_salary = st.number_input("현재연봉 (만원)", min_value=0, step=100)
                with col5:
                    other_salary = st.number_input("기타 보상상 (만원)", min_value=0, step=100)
                with col6:
                    desired_salary = st.number_input("희망연봉 (만원)", min_value=0, step=100)
                with col7:
                    years = st.number_input("인정경력 (년)", min_value=-4.0, value=float(st.session_state.get('years', st.query_params.get("years", 0.0))), step=0.1, format="%.1f")
                
              
                # 4줄: 특이사항
                education_notes = st.text_input("특이사항", "")
                
                # 분석하기 버튼
                submitted = st.form_submit_button("분석하기")

                if submitted:
                    try:                      
                        # salary_table.xlsx 파일 읽기
                        salary_table = pd.read_excel("salary_table.xlsx")
                        
                        # 선택된 직군상세에 해당하는 직군 가져오기
                        selected_job_category = job_mapping[job_role]
                        
                        # 해당 직군과 연차에 맞는 데이터 필터링
                        try:
                            years_int = int(float(years))  # 연차를 float로 변환 후 정수로 변환
                        except (ValueError, TypeError):
                            st.error(f"경력 기간을 정수로 변환하는 중 오류가 발생했습니다. 입력된 경력 기간: {years}")
                            st.stop()
                            
                        filtered_data = salary_table[
                            (salary_table['직군'] == selected_job_category) & 
                            (salary_table['연차'] == years_int)
                        ]
                        
                        if filtered_data.empty:
                            st.warning(f"선택하신 직군 '{job_role}' ({selected_job_category})과 연차 {years_int}년에 해당하는 데이터가 없습니다.")
                            st.stop()
                        
                        # 첫 번째 행 선택
                        filtered_data = filtered_data.iloc[0]
                        
                        # 해당 직군의 임금 데이터 가져오기
                        min_salary = filtered_data['최소연봉']
                        max_salary = filtered_data['최대연봉']
                        avg_salary = (min_salary + max_salary) / 2

                        # 분석 결과 표시
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown("#### 📊 연봉 분석 결과")
                        
                        # 직군 정보 표시
                        st.markdown(f"**선택된 직군 정보:** {selected_job_category} - {job_role}")
                        # 연봉 정보 표시
                        st.markdown(f"""
                        <div style="font-size: 1rem;">
                        <strong>현재 연봉 : {int(current_salary):,}만원 &nbsp;&nbsp;&nbsp;&nbsp; </strong>
                        <strong>최소 연봉 : {int(min_salary):,}만원 &nbsp;&nbsp;&nbsp;&nbsp;</strong>
                        <strong style="color: red;">평균 연봉 : {int(avg_salary):,}만원 &nbsp;&nbsp;&nbsp;&nbsp;</strong>
                        <strong>최대 연봉 : {int(max_salary):,}만원</strong>
                        </div>
                        """, unsafe_allow_html=True)
                        st.markdown("<br>", unsafe_allow_html=True)

                        # 컬럼으로 공간 분리
                        col1, col2 = st.columns([0.6, 0.4])
                        with col1:
                            # salary_table 관련 데이터 표시
                            related_years = [years_int-1, years_int, years_int+1]
                            related_data = salary_table[
                                (salary_table['직군'] == selected_job_category) & 
                                (salary_table['연차'].isin(related_years))
                            ].sort_values('연차')
                            
                            if not related_data.empty:
                                # 모든 연봉 컬럼을 정수로 변환
                                related_data['최소연봉'] = related_data['최소연봉'].astype(int)
                                related_data['평균연봉'] = related_data['평균연봉'].astype(int)
                                related_data['최대연봉'] = related_data['최대연봉'].astype(int)
                                
                                st.dataframe(
                                    related_data[['연차', '최소연봉', '평균연봉', '최대연봉']].rename(
                                        columns={
                                            '연차': '인정경력',
                                            '최소연봉': '최소연봉(만원)',
                                            '평균연봉': '평균연봉(만원)',
                                            '최대연봉': '최대연봉(만원)'
                                        }
                                    ),
                                    hide_index=True,
                                    column_config={
                                        '인정경력': st.column_config.Column(width=80),
                                        '최소연봉(만원)': st.column_config.Column(width=100),
                                        '평균연봉(만원)': st.column_config.Column(width=100),
                                        '최대연봉(만원)': st.column_config.Column(width=100)
                                    }
                                )
                            else:
                                st.info("해당 직군의 임금테이블 데이터가 없습니다.")
                        
                        with col2:
                            st.write("")  # 빈 공간

                        st.markdown("<br>", unsafe_allow_html=True)
                        # 2. 상세 분석 결과
                        st.markdown("##### 💡 연봉 책정 가이드")
                        
                        analysis_text = ""
                        
                        # 임금 테이블 기준 분석
                        if current_salary < min_salary:
                            analysis_text += f"⚠️ 현재 연봉(기본연봉)이 시장 최소값보다 {min_salary - current_salary:,.0f}만원 낮습니다.\n"
                            recommended_salary = min_salary
                        elif current_salary > max_salary:
                            analysis_text += f"⚠️ 현재 연봉(기본연봉)이 시장 최대값보다 {current_salary - max_salary:,.0f}만원 높습니다.\n"
                            recommended_salary = max_salary
                        else:
                            analysis_text += "✅ 현재 연봉(기본연봉)이 시장 범위 내에 있습니다.\n"
                            recommended_salary = current_salary
                                                
                        # 최종보상 계산
                        final_compensation = current_salary + other_salary
                        
                        # 제시금액 계산 로직
                        def calculate_suggested_salary(total_comp, min_salary, avg_salary, max_salary):
                            increase_10 = total_comp * 1.1
                            increase_5 = total_comp * 1.05
                            increase_2 = total_comp * 1.02
                            
                            if increase_10 <= avg_salary:
                                return int(increase_10)
                            elif increase_5 < avg_salary:
                                return int(avg_salary)
                            elif increase_5 >= avg_salary and total_comp <= avg_salary:
                                return int(increase_5)
                            elif total_comp > avg_salary and total_comp <= max_salary:
                                return int(increase_2)
                            else:
                                return "[별도 계산 필요]"

                        # 제시금액 계산
                        suggested_salary = calculate_suggested_salary(
                            final_compensation, 
                            min_salary, 
                            avg_salary, 
                            max_salary
                        )
                        # 연봉 보존율 계산
                        preservation_rate = (suggested_salary / final_compensation) * 100

                        # 현재 상황에 맞는 제시금액 계산 로직 결정
                        if final_compensation * 1.1 < avg_salary:
                            calculation_logic = "제시금액 계산 로직 : 최종보상 * 1.1 (10% 증액)으로 제안"
                        elif final_compensation * 1.05 < avg_salary:
                            calculation_logic = "제시금액 계산 로직 : 평균연봉으로 제안"
                        elif final_compensation * 1.05 >= avg_salary and final_compensation <= avg_salary:
                            calculation_logic = "제시금액 계산 로직 : 최종보상 * 1.05까지 제안 (5% 증액)"
                        elif final_compensation > avg_salary and final_compensation <= max_salary:
                            calculation_logic = "제시금액 계산 로직 : 최종보상 * 1.02까지 제안 (2% 증액)"
                        else:
                            calculation_logic = "제시금액 계산 로직 : 별도 계산 필요"

                        st.info(f"""
                        {position} 합격자 {candidate_name}님 처우 협상(안) 보고 드립니다.

                        {candidate_name}님의 경력은 {years:.1f}년으로 {selected_job_category} 임금테이블 기준으로는 
                        기준연봉 {avg_salary:,.0f}만원 ~ 상위10% {max_salary:,.0f}만원까지 고려할 수 있습니다.
                        
                        최종보상 {final_compensation:,.0f}만원, 기준(평균)연봉 {avg_salary:,.0f}만원을 고려했을 때 
                        제시금액은 {suggested_salary if isinstance(suggested_salary, str) else f'{suggested_salary:,.0f}만원'}이 어떨지 의견 드립니다.

                        [연봉산정]
                        - 인정경력: {years:.1f}년
                        - 최종연봉: 기본연봉 {current_salary:,.0f}만원 + 기타 {other_salary:,.0f}만원
                        - 희망연봉: {desired_salary:,.0f}만원
                        - 기준(임금테이블) 연봉: {avg_salary:,.0f}만원 (최소 연봉: {min_salary:,.0f}만원, 최대 연봉: {max_salary:,.0f}만원)
                        - 특이사항: {education_notes}

                        [참고]
                        - {calculation_logic}
                        - 기존 보상총액 보존율: {preservation_rate:.1f}%
                        """)
                        # 상세 분석 결과 expander
                        with st.expander("📌 분석 기준 보기"):
                            st.info(f"""
                             제시금액 계산                 
                                - 최종보상 * 1.1 < 평균연봉 : 최종보상 * 1.1 정도 제안 (10% 증액) 
                                - 최종보상 * 1.05 < 평균연봉 : 평균연봉 정도 제안 (5% 증액) 
                                - 최종보상 * 1.05 >= 평균연봉 & 최종보상 <= 평균연봉 : 최종보상 * 1.05까지 제안 (5% 증액) 
                                - 최종보상 > 평균연봉 & 최종보상 <= 최대연봉 : 최종보상 * 1.02까지 제안 (2% 증액) 
                                - 최종보상 > 최대연봉 : 별도 계산 필요
                            """)
                    except Exception as e:
                        st.error(f"임금 테이블 데이터를 불러오는 중 오류가 발생했습니다: {str(e)}")

        elif menu == "⏰ 초과근무 조회":
            st.markdown("##### ⏰ 초과근무 조회")
            
            # 엑셀 파일 업로드
            uploaded_file = st.file_uploader("초과근무 엑셀 파일을 업로드하세요", type=['xlsx'])
            
            if uploaded_file is not None:
                try:
                    # 엑셀 파일 읽기
                    overtime_df = pd.read_excel(uploaded_file)
                    
                    # 연월 구분 드롭다운 생성
                    if '연월구분' in overtime_df.columns:
                        months = overtime_df['연월구분'].unique()
                        selected_month = st.selectbox('조회 기준을 선택하세요.', sorted(months, reverse=True))
                        
                        # 선택된 연월에 해당하는 데이터 필터링
                        filtered_df = overtime_df[overtime_df['연월구분'] == selected_month]
                        
                        # 필터링된 데이터가 있을 때만 표시
                        if not filtered_df.empty:
                            # 월별 본부별 초과근무 합계 표시                                                      
                            # 시간을 숫자로 변환
                            filtered_df['초과시간'] = filtered_df['초과시간'].apply(lambda x: float(x.hour) + float(x.minute)/60 if hasattr(x, 'hour') and hasattr(x, 'minute') else float(x))
                            
                            # 피벗 테이블 생성
                            pivot_df = pd.pivot_table(
                                filtered_df,
                                values='초과시간',
                                index='연월구분',
                                columns='본부',
                                aggfunc='sum',
                                fill_value=0
                            )
                            
                            # 전체 합계 열 추가
                            pivot_df['전체 합계'] = pivot_df.sum(axis=1)
                            
                            # 본부별 인원수 계산
                            employee_count = filtered_df.groupby('본부')['이름'].nunique()
                            employee_count['전체 합계'] = employee_count.sum()
                            
                            # 인원수 행 추가
                            pivot_df.loc['인원수'] = employee_count
                            
                            # 시간을 소수점 한 자리로 변환 (인원수 행 제외)
                            for col in pivot_df.columns:
                                pivot_df.loc[pivot_df.index != '인원수', col] = pivot_df.loc[pivot_df.index != '인원수', col].apply(lambda x: f"{float(x):.1f}시간")
                            
                            # 피벗 테이블이 비어있지 않을 때만 표시
                            if not pivot_df.empty:
                                st.dataframe(
                                    pivot_df,
                                    use_container_width=True,
                                )
                            # 이름과 이메일로 그룹화하여 초과근무 내역과 시간 합계 계산
                            # 시간을 숫자로 변환하여 합산
                            filtered_df['초과시간'] = filtered_df['초과시간'].apply(lambda x: float(x.hour) + float(x.minute)/60 if hasattr(x, 'hour') and hasattr(x, 'minute') else float(x))
                            
                            # 초과근무 내용 컬럼명 확인
                            content_column = '초과근무 내용' if '초과근무 내용' in filtered_df.columns else '초과근무내용'
                            
                            result_df = filtered_df.groupby(['이름', '이메일']).agg({
                                content_column: lambda x: '\n'.join(x),  # 일반 줄바꿈 문자 사용
                                '초과시간': 'sum'
                            }).reset_index()
                            
                            # 시간을 시:분 형식으로 변환
                            result_df['초과근무시간 합'] = result_df['초과시간'].apply(lambda x: f"{int(x)}시간 {int((x % 1) * 60)}분")
                            
                            # 컬럼명 변경
                            result_df = result_df.rename(columns={content_column: '초과근무 내역'})
                            result_df = result_df[['이름', '초과근무시간 합',  '초과근무 내역', '이메일']]
                            
                            # 인덱스를 1부터 시작하도록 설정
                            result_df.index = range(1, len(result_df) + 1)
                            
                            # 테이블 표시
                            st.markdown("""
                                <style>
                                [data-testid="stDataFrame"] {
                                    width: 80%;
                                }
                                [data-testid="stDataFrame"] td {
                                    white-space: pre-wrap !important;
                                    min-height: fit-content !important;
                                    height: auto !important;
                                    line-height: 1.5 !important;
                                    padding: 8px !important;
                                    vertical-align: top !important;
                                }
                                [data-testid="stDataFrame"] div[data-testid="StyledDataFrameDataCell"] {
                                    min-height: fit-content !important;
                                    height: auto !important;
                                    white-space: pre-wrap !important;
                                    overflow: visible !important;
                                }
                                [data-testid="stDataFrame"] div[data-testid="StyledDataFrameDataCell"] > div {
                                    min-height: fit-content !important;
                                    height: auto !important;
                                    white-space: pre-wrap !important;
                                    overflow: visible !important;
                                }
                                [data-testid="stDataFrame"] div[role="cell"] {
                                    min-height: fit-content !important;
                                    height: auto !important;
                                    white-space: pre-wrap !important;
                                    overflow: visible !important;
                                }
                                [data-testid="stDataFrame"] div[role="row"] {
                                    min-height: fit-content !important;
                                    height: auto !important;
                                }
                                [data-testid="stDataFrame"] div[data-testid="StyledDataFrameRowMain"] {
                                    min-height: fit-content !important;
                                    height: auto !important;
                                }
                                </style>
                            """, unsafe_allow_html=True)
                            
                            st.dataframe(
                                result_df,
                                column_config={
                                    "이름": st.column_config.TextColumn("이름", width=50),
                                    "초과근무시간 합": st.column_config.TextColumn("초과근무시간 합", width=70),
                                    "초과근무 내역": st.column_config.TextColumn("초과근무 내역", width=400),
                                    "이메일": st.column_config.TextColumn("이메일", width=100)
                                },
                                hide_index=False,
                                use_container_width=True,
                                height=400
                            )
                            # 엑셀 다운로드 버튼
                            output = BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                result_df.to_excel(writer, sheet_name='초과근무내역', index=True, index_label='No')
                                # 열 너비 자동 조정
                                worksheet = writer.sheets['초과근무내역']
                                worksheet.column_dimensions['B'].width = 10 # 이름
                                worksheet.column_dimensions['C'].width = 15  # 초과근무시간 합
                                worksheet.column_dimensions['D'].width = 70  # 초과근무 내역
                                worksheet.column_dimensions['E'].width = 25  # 이메일
                            excel_data = output.getvalue()
                                    
                            st.download_button(
                                        label="📥 엑셀 파일 다운로드",
                                data=excel_data,
                                file_name=f"초과근무내역_{selected_month}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            


                        else:
                            st.error("엑셀 파일에 '연월구분' 컬럼이 없습니다.")
                    
                except Exception as e:
                    st.error(f"파일을 읽는 중 오류가 발생했습니다: {str(e)}")
            else:
                st.info("초과근무 엑셀 파일을 업로드하세요.")

        elif menu == "😊 임직원 명부":
            st.markdown("##### 😊 임직원 명부")
            # 조회 조건
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                query_date = st.date_input("조회일자", datetime.now())
            
            with col2:
                name = st.text_input("성명")
            
            with col3:
                employment_type = st.selectbox(
                    "고용구분",
                    ["전체", "정규직", "계약직"]
                )
            
            with col4:
                employment_status = st.selectbox(
                    "재직상태",
                    ["전체", "재직", "퇴직"]
                )
            
            with col5:
                show_department_history = st.checkbox("해당 시점부서 추가")
            
            # 데이터 로드
            @st.cache_data
            def load_employee_data():
                try:
                    # 파일 경로를 절대 경로로 변경
                    import os
                    current_dir = os.path.dirname(os.path.abspath(__file__))
                    file_path = os.path.join(current_dir, "임직원 기초 데이터.xlsx")
                    
                    # 파일이 존재하는지 확인
                    if not os.path.exists(file_path):
                        st.error(f"파일을 찾을 수 없습니다: {file_path}")
                        return None, None
                    
                    # 파일 읽기
                    df = pd.read_excel(file_path, sheet_name=0)  # 첫 번째 시트 사용
                    df_history = pd.read_excel(file_path, sheet_name=1)  # 두 번째 시트 사용
                    
                    # 컬럼 이름 재정의
                    df.columns = df.columns.str.strip()  # 컬럼 이름의 공백 제거
                    df_history.columns = df_history.columns.str.strip()  # 컬럼 이름의 공백 제거
                    
                    # 날짜 컬럼 형식 통일
                    date_columns = ['입사일', '퇴사일', '발령일']
                    for col in date_columns:
                        if col in df.columns:
                            df[col] = pd.to_datetime(df[col], errors='coerce')
                        if col in df_history.columns:
                            df_history[col] = pd.to_datetime(df_history[col], errors='coerce')
                    
                    # None 값 처리
                    df = df.fillna('')
                    df_history = df_history.fillna('')
                    
                    return df, df_history
                except Exception as e:
                    st.error(f"파일을 불러오는 중 오류가 발생했습니다: {str(e)}")
                    return None, None
            
            df, df_history = load_employee_data()
            
            # 조회일자 기준으로 재직중인 직원 필터링
            df = df[
                (df['입사일'] <= pd.Timestamp(query_date)) &  # 입사일이 조회일자 이전
                (
                    (df['퇴사일'].isna()) |  # 퇴사일이 없는 경우
                    (df['퇴사일'] >= pd.Timestamp(query_date))  # 퇴사일이 조회일자 이후
                )
            ]
            
            # 조회일자 기준으로 인사발령 데이터 필터링
            df_history_filtered = df_history[df_history['발령일'] <= pd.Timestamp(query_date)]
            
            # 각 직원별 가장 최근 발령 데이터만 선택
            df_history_filtered = df_history_filtered.sort_values('발령일').groupby('성명').last().reset_index()
            
            # 기본 컬럼 설정
            se_columns = [
                "사번", "성명", "본부", "팀", "직무", "직위", "직책", "입사일", 
                "재직기간", "정규직전환일", "고용구분", "재직상태", "생년월일", 
                "남/여", "만나이", "퇴사일", "학력", "최종학교", "전공", 
                "경력사항", "휴직상태"
            ]
            
            history_columns = [
                "발령일", "구분", "성명", "변경후_본부",  "변경후_팀", "변경후_직책"
            ]
            
            # 재직기간 계산 함수
            def calculate_employment_period(row):
                if pd.isna(row['입사일']):
                    return None
                
                start_date = pd.to_datetime(row['입사일'])
                
                # 재직상태가 '퇴직'인 경우 퇴사일을 기준으로 계산
                if row['재직상태'] == '퇴직' and pd.notna(row['퇴사일']):
                    end_date = pd.to_datetime(row['퇴사일'])
                else:
                    # 그 외의 경우 조회일자를 기준으로 계산
                    end_date = pd.Timestamp(query_date)
                
                years = (end_date - start_date).days // 365
                months = ((end_date - start_date).days % 365) // 30
                
                return f"{years}년 {months}개월"
            
            # 데이터 필터링
            if name:
                df = df[df['성명'].str.contains(name, na=False)]
            
            if employment_type != "전체":
                df = df[df['고용구분'] == employment_type]
            
            if employment_status != "전체":
                df = df[df['재직상태'] == employment_status]
            
            # 재직기간 계산
            df['재직기간'] = df.apply(calculate_employment_period, axis=1)
            
            # 부서 이력 데이터 처리
            if show_department_history:
                # 인사발령 데이터와 조인
                df_merged = pd.merge(
                    df, 
                    df_history_filtered, 
                    left_on='성명', 
                    right_on='성명', 
                    how='left',
                    suffixes=('', '_history')  # 중복 컬럼에 접미사 추가
                )
                
                # 발령이 없는 경우 기본값 설정
                df_merged['변경후_본부'] = df_merged['변경후_본부'].fillna(df_merged['본부'])
                df_merged['변경후_팀'] = df_merged['변경후_팀'].fillna(df_merged['팀'])
                df_merged['변경후_직책'] = df_merged['변경후_직책'].fillna(df_merged['직책'])
                
                # 컬럼 순서 조정
                display_columns = se_columns + [col for col in history_columns if col not in se_columns]
                df_display = df_merged[display_columns]
            else:
                df_display = df[se_columns]
            
            # 데이터 표시
            df_display = df_display.reset_index(drop=True)
            df_display.index = df_display.index + 1
            df_display = df_display.reset_index()
            df_display = df_display.rename(columns={'index': 'No'})
            
            # 날짜 컬럼의 시간 제거
            date_columns = ['정규직전환일', '입사일', '퇴사일', '생년월일', '발령일']
            for col in date_columns:
                if col in df_display.columns:
                    df_display[col] = pd.to_datetime(df_display[col]).dt.date
            
            # 데이터 수에 따라 높이 동적 조정 (행당 35픽셀)
            row_height = 35  # 각 행의 예상 높이
            dynamic_height = min(len(df_display) * row_height + 40, 600)  # 헤더 높이 추가, 최대 600픽셀로 제한
            
            st.dataframe(
                df_display,
                use_container_width=True,
                hide_index=True,
                height=dynamic_height,
                column_config={
                   "직무": st.column_config.Column(width=70),
                   "최종학교": st.column_config.Column(width=70),
                   "전공": st.column_config.Column(width=70),
                   "경력사항": st.column_config.Column(width=70)
                }
            )
            
            # 엑셀 다운로드 버튼
            @st.cache_data
            def convert_df_to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='임직원명부')
                processed_data = output.getvalue()
                return processed_data
            
            excel_data = convert_df_to_excel(df_display)
            st.download_button(
                label="📥 엑셀 다운로드",
                data=excel_data,
                file_name=f"임직원명부_{query_date.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        elif menu == "📅 인사발령 내역":
            st.markdown("##### 📅 인사발령 내역")
            
            # 데이터 로드
            @st.cache_data(ttl=300)  # 5분마다 캐시 갱신
            def load_promotion_data():
                try:
                    # 파일 경로를 절대 경로로 변경
                    import os
                    current_dir = os.path.dirname(os.path.abspath(__file__))
                    file_path = os.path.join(current_dir, "임직원 기초 데이터.xlsx")
                    
                    # 파일이 존재하는지 확인
                    if not os.path.exists(file_path):
                        st.error(f"파일을 찾을 수 없습니다: {file_path}")
                        return None
                    
                    # 파일 읽기 (sheet2)
                    df_promotion = pd.read_excel(file_path, sheet_name=1)
                    
                    # 컬럼 이름 재정의
                    df_promotion.columns = df_promotion.columns.str.strip()
                    
                    # 날짜 컬럼 형식 통일
                    df_promotion['발령일'] = pd.to_datetime(df_promotion['발령일'], errors='coerce')
                    
                    # None 값 처리
                    df_promotion = df_promotion.fillna('')
                    
                    # 발령일이 유효한 날짜인 행만 필터링
                    df_promotion = df_promotion[pd.notna(df_promotion['발령일'])]
                    
                    # 발령년도 추출 (NA 값 처리)
                    df_promotion['발령년도'] = df_promotion['발령일'].dt.year
                    df_promotion['발령년도'] = df_promotion['발령년도'].fillna(0).astype(int)
                    
                    return df_promotion
                except Exception as e:
                    st.error(f"파일을 불러오는 중 오류가 발생했습니다: {str(e)}")
                    return None
            
            df_promotion = load_promotion_data()
            
            if df_promotion is not None:
                # 조회 조건
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    current_year = datetime.now().year
                    years = sorted(df_promotion['발령일'].dt.year.unique(), reverse=True)
                    selected_year = st.selectbox("발령 연도", ["전체"] + years, index=0)
                
                with col2:
                    name = st.text_input("성명")
                
                with col3:
                    promotion_types = sorted(df_promotion['구분'].unique())
                    selected_types = st.multiselect("발령구분", promotion_types)
                
                # 데이터 필터링
                filtered_df = df_promotion.copy()
                
                if selected_year != "전체":
                    filtered_df = filtered_df[filtered_df['발령일'].dt.year == selected_year]
                
                if name:
                    filtered_df = filtered_df[filtered_df['성명'].str.contains(name, na=False)]
                
                if selected_types:
                    filtered_df = filtered_df[filtered_df['구분'].isin(selected_types)]
                
                # 표시할 컬럼 설정
                display_columns = [
                    "발령일", "구분", "성명", 
                    "변경전_본부", "변경전_실", "변경전_팀", "변경전_직책",
                    "변경후_본부", "변경후_실", "변경후_팀", "변경후_직책", "비고"
                ]
                
                # 데이터 표시
                df_display = filtered_df[display_columns].copy()
                df_display = df_display.sort_values('발령일', ascending=False)
                df_display = df_display.reset_index(drop=True)
                df_display.index = df_display.index + 1
                df_display = df_display.reset_index()
                df_display = df_display.rename(columns={'index': 'No'})
                
                # 날짜 컬럼의 시간 제거
                df_display['발령일'] = pd.to_datetime(df_display['발령일']).dt.date
                
                # 데이터프레임 표시
                if not filtered_df.empty:
                    # 데이터 정렬 및 인덱스 설정
                    display_df = filtered_df[display_columns].sort_values('발령일', ascending=False).reset_index(drop=True)
                    display_df.index = display_df.index + 1  # 인덱스를 1부터 시작하도록 설정
                    
                    # 발령일 컬럼의 시간 제거
                    display_df['발령일'] = pd.to_datetime(display_df['발령일']).dt.strftime('%Y-%m-%d')
                    
                    # 데이터 수에 따라 높이 동적 조정 (행당 35픽셀)
                    row_height = 35  # 각 행의 예상 높이
                    dynamic_height = min(len(display_df) * row_height + 40, 600)  # 헤더 높이 추가, 최대 600픽셀로 제한
                    
                    st.dataframe(
                        display_df,
                        use_container_width=True,
                        height=dynamic_height
                    )
                else:
                    st.warning("조회된 데이터가 없습니다.")
                
                # 엑셀 다운로드 버튼
                @st.cache_data
                def convert_df_to_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='인사발령내역')
                    processed_data = output.getvalue()
                    return processed_data
                
                excel_data = convert_df_to_excel(df_display)
                st.download_button(
                    label="📥 엑셀 다운로드",
                    data=excel_data,
                    file_name=f"인사발령내역_{selected_year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        elif menu == "🔔 인사팀 업무 공유":
            st.markdown("##### 🔔 인사팀 업무 공유")
            # 업무보고 데이터 가져오기
            @st.cache_data(ttl=60)  # 1분마다 캐시 갱신
            def get_work_report_data():
                try:
                    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
                    credentials_dict = {
                        "type": st.secrets["google_credentials"]["type"],
                        "project_id": st.secrets["google_credentials"]["project_id"],
                        "private_key_id": st.secrets["google_credentials"]["private_key_id"],
                        "private_key": st.secrets["google_credentials"]["private_key"],
                        "client_email": st.secrets["google_credentials"]["client_email"],
                        "client_id": st.secrets["google_credentials"]["client_id"],
                        "auth_uri": st.secrets["google_credentials"]["auth_uri"],
                        "token_uri": st.secrets["google_credentials"]["token_uri"],
                        "auth_provider_x509_cert_url": st.secrets["google_credentials"]["auth_provider_x509_cert_url"],
                        "client_x509_cert_url": st.secrets["google_credentials"]["client_x509_cert_url"]
                    }
                    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
                    gc = gspread.authorize(credentials)
                    
                    try:
                        # 업무보고 시트 ID
                        sheet_id = st.secrets["google_sheets"]["work_report_id"]
                        worksheet = gc.open_by_key(sheet_id).worksheet('시트1')  # '업무보고' 시트 선택
                    except Exception as e:
                        st.error(f"시트 접근 중 오류 발생: {str(e)}")
                        return pd.DataFrame()
                    
                    try:
                        # 데이터 가져오기
                        data = worksheet.get_all_records()
                        
                        # 데이터프레임으로 변환
                        df = pd.DataFrame(data)
                        
                        # 보고일 컬럼을 datetime으로 변환
                        if '보고일' in df.columns:
                            df['보고일'] = pd.to_datetime(df['보고일'])
                        
                        return df
                    except Exception as e:
                        st.error(f"데이터 처리 중 오류 발생: {str(e)}")
                        return pd.DataFrame()
                        
                except Exception as e:
                    st.error(f"인증 중 오류 발생: {str(e)}")
                    return pd.DataFrame()

            # 업무보고 데이터 로드
            report_df = get_work_report_data()
            st.markdown("<br>", unsafe_allow_html=True)
            if not report_df.empty:
                st.markdown("###### 업무 공유/보고")
                
# 조회 조건 컬럼 생성
                col1, col2, col3 = st.columns([0.15, 0.3, 0.55]) 
                
                with col1:
                    # 보고상태 선택
                    status_options = ['보고예정', '보고완료']
                    selected_status = st.selectbox('보고상태', status_options)

                    # 선택된 보고상태에 해당하는 데이터만 필터링
                    status_filtered_df = report_df[report_df['보고상태'] == selected_status]
                    
                with col2:
                    # 타입과 보고일을 합친 옵션 생성
                    type_date_options = ['전체']
                    for type_val in status_filtered_df['타입'].unique():
                        dates = status_filtered_df[status_filtered_df['타입'] == type_val]['보고일'].dt.strftime('%Y-%m-%d').unique()
                        for date in dates:
                            type_date_options.append(f"{type_val} - {date}")
                    
                    selected_type_date = st.selectbox('타입 - 보고일자', type_date_options)

                with col3:
                    st.write("")

                # 데이터 필터링
                filtered_df = report_df[report_df['보고상태'] == selected_status]
                
                if selected_type_date != '전체':
                    type_val, date_val = selected_type_date.split(' - ')
                    filtered_df = filtered_df[
                        (filtered_df['타입'] == type_val) & 
                        (filtered_df['보고일'].dt.strftime('%Y-%m-%d') == date_val)
                    ]

                # 데이터프레임 정렬
                filtered_df = filtered_df.sort_values('보고일', ascending=False)

                if not filtered_df.empty:
                    html_output = []
                    html_output.append('<table style="width: 70%;">')
                    
                    for _, row in filtered_df.iterrows():
                        html_output.append("<tr>")
                        # 업무구분 
                        html_output.append(f'<td style="width: 20%; text-align: left; background-color: #f0f2f6; font-size: 13px;""> {row["업무구분"]}</td>')
                        # 업무내용
                        업무내용 = row["업무내용"]
                        if not 업무내용.startswith("<"):
                            # 여러 줄 지원 및 URL 자동 링크 변환
                            업무내용 = 업무내용.replace("\n", "<br>")
                            # URL 패턴 찾기
                            url_pattern = r'https?://[^\s<>"]+|www\.[^\s<>"]+'
                            # "링크" 텍스트 찾기
                            link_pattern = r'링크' 
                            
                            # URL이 있는지 확인
                            urls = re.findall(url_pattern, 업무내용)
                            if urls:
                                # 각 URL에 대해
                                for url in urls:
                                    # "링크" 텍스트가 있으면 해당 텍스트를 URL로 대체
                                    if re.search(link_pattern, 업무내용):
                                        업무내용 = re.sub(link_pattern, f'<a href="{url}" target="_blank">링크</a>', 업무내용, count=1)
                                    else:
                                        # "링크" 텍스트가 없으면 URL 자체를 링크로 변환
                                        업무내용 = 업무내용.replace(url, f'<a href="{url}" target="_blank">링크</a>')
                        
                        html_output.append(f'<td style="width: 85%; text-align: left; padding-left: 15px; font-size: 13px;">{업무내용}</td>')
                        html_output.append("</tr>")
                    
                    html_output.append("</table>")
                    
                    # HTML 출력
                    final_html = "\n".join(html_output)
                    st.markdown(final_html, unsafe_allow_html=True)
                else:
                    st.info("조회된 데이터가 없습니다.")
            
            
            try:
                # 구글 시트 인증
                scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
                credentials_dict = {
                    "type": st.secrets["google_credentials"]["type"],
                    "project_id": st.secrets["google_credentials"]["project_id"],
                    "private_key_id": st.secrets["google_credentials"]["private_key_id"],
                    "private_key": st.secrets["google_credentials"]["private_key"],
                    "client_email": st.secrets["google_credentials"]["client_email"],
                    "client_id": st.secrets["google_credentials"]["client_id"],
                    "auth_uri": st.secrets["google_credentials"]["auth_uri"],
                    "token_uri": st.secrets["google_credentials"]["token_uri"],
                    "auth_provider_x509_cert_url": st.secrets["google_credentials"]["auth_provider_x509_cert_url"],
                    "client_x509_cert_url": st.secrets["google_credentials"]["client_x509_cert_url"]
                }
                credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
                gc = gspread.authorize(credentials)

                try:
                    # 업무보고 시트 ID
                    sheet_id = st.secrets["google_sheets"]["work_report_id"]
                    worksheet = gc.open_by_key(sheet_id).worksheet('주요일정')  # '업무보고' 시트 선택
                    schedule_data = worksheet.get_all_values()
                except Exception as e:
                    st.error(f"시트 접근 중 오류 발생: {str(e)}")
                    schedule_data = []
                
                # 데이터가 있는 경우에만 DataFrame 생성
                if schedule_data:
                    # 데이터프레임으로 변환
                    schedule_df = pd.DataFrame(schedule_data[1:], columns=schedule_data[0])
                    
                    # NaN 값을 빈 문자열로 변환
                    schedule_df = schedule_df.fillna("")
                    
                    # 모든 열을 문자열로 변환하고 앞뒤 공백 제거
                    for col in schedule_df.columns:
                        schedule_df[col] = schedule_df[col].astype(str).str.strip()

                    # 스타일이 적용된 테이블 표시
                    st.markdown("""
                    <style>
                    .schedule-table {
                        width: 90%;
                        border-collapse: collapse;
                        margin: 0px 0;
                        font-size: 13px; 
                    }
                    .schedule-table th, .schedule-table td {
                        border: 1px solid #ddd;
                        padding: 6px;
                        text-align: center;
                        min-width: 50px;
                        color: #A6A6A6;
                    }
                    .schedule-table th {
                        background-color: #F2F2F2;
                        position: sticky;
                        top: 0;
                        z-index: 1;
                        white-space: nowrap;
                        color: #000000;
                    }
                    .schedule-table td {
                        background-color: white;
                    }
                    .schedule-table tr:nth-child(even) td {
                        background-color: #ffffff; 
                    }
                    .schedule-table td:first-child {
                        background-color: #F2F2F2;
                        position: sticky;
                        left: 0;
                        z-index: 1;
                    }
                    .schedule-container {
                        overflow-x: auto;
                        margin-top: 0px;
                        max-height: 800px;
                        overflow-y: auto;
                    }
                    </style>
                    """, unsafe_allow_html=True)

                    # HTML 테이블 생성
                    table_html = '<div class="schedule-container">'
                    table_html += '<div style="margin-bottom: 10px; font-weight: bold;">연간 주요일정</div>'
                    table_html += '<table class="schedule-table">'
                    
                    # 헤더 행 추가
                    table_html += '<tr><th style="color: #000000; background-color: #f0f2f6; font-weight: normal;">구분</th>'
                    for col in schedule_df.columns[1:]:
                        table_html += f'<th style="color: #000000; background-color: #f0f2f6; font-weight: normal;">{col}</th>'
                    table_html += '</tr>'
                    
                    # 데이터 행 추가
                    for _, row in schedule_df.iterrows():
                        table_html += '<tr>'
                        current_month = int(datetime.now().month)  # 현재 월을 정수형으로 가져오기
                        for idx, col in enumerate(schedule_df.columns):
                            cell_value = row[col]
                            if idx == 0:  # 첫 번째 열(구분)
                                table_html += f'<td style="background-color: #f0f2f6; text-align: center; color: #000000;">{cell_value}</td>'
                            else:
                                # 현재 월에 해당하는 열인지 확인 (1월은 첫 번째 열이므로 idx가 1)
                                is_current_month = (idx == current_month)
                                
                                if is_current_month and cell_value and cell_value != "":
                                    # 현재 월이고 내용이 있는 경우 빨간 배경과 흰색 글씨
                                    table_html += f'<td style="background-color: #ff3333; text-align: center; color: #FFFFFF;">{cell_value}</td>'
                                elif "진행" in str(cell_value).lower():
                                    table_html += f'<td style="background-color: #FFE5E5; text-align: center; color: #EE6C6C;">{cell_value}</td>'
                                elif "계획" in str(cell_value).lower():
                                    table_html += f'<td style="background-color: #F2F2F2; text-align: center; color: #A6A6A6;">{cell_value}</td>'
                                elif cell_value and cell_value != "":  # 그 외 텍스트가 있는 경우
                                    table_html += f'<td style="background-color: #FFE5E6; text-align: center; color: #EE6C6C;">{cell_value}</td>'
                                else:
                                    table_html += f'<td style="text-align: center; color: #A6A6A6;">{cell_value}</td>'
                        table_html += '</tr>'
                    
                    table_html += '</table></div>'
                    
                    # 테이블 표시
                    st.markdown(table_html, unsafe_allow_html=True)
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown("###### 수시/상시 일정")
                    
                    st.markdown("""
                    <div style="font-size: 13px;">
                    ㆍ채용 진행 : 정시(연간 인원계획)/수시/결원에 대한 채용 진행<br>                
                    ㆍ온보딩/수습평가 운영 : 온보딩 프로그램 / CEO 환영 미팅 / 3개월 후 수습평가 실시<br>                
                    ㆍ인력운영/관리 : 근태(휴가/초과근무/출퇴근) 관리, 조직개편 및 인사발령, 입퇴사 4대보험 처리<br>                
                    ㆍ복지제도 운영 : 경조비/경조휴가, 근속 포상(휴가, 상품) 지급<br>                
                    ㆍ사내 시스템 운영 : 뉴로웍스, 뉴로핏 커리어 콘텐츠 업데이트, MS/비즈박스 라이선스 관리 등<br>                
                    ㆍ교육 운영 : 직무 전문 교육, 특강 등 교육 지원, 각종 이러닝 콘텐츠 공유<br>                
                    ㆍ노무 이슈 가이드/조치 : 고충처리(동료간 어려움, 컴플레인 등) 상담, 규정/제도 가이드<br>                
                    ㆍ각종 대관 업무 : 노동부(실사/ 인원통계 /출산 및 육아 휴직), 병무청, 산학협력 등<br>
                    </div>
                    """, unsafe_allow_html=True)

            except Exception as e:
                st.error(f"연간일정을 불러오는 중 오류가 발생했습니다: {str(e)}")
            st.markdown("<br>", unsafe_allow_html=True)   
            st.markdown("<br>", unsafe_allow_html=True)              
            st.markdown('''
            <a href="https://docs.google.com/spreadsheets/d/1KjlfACJIzNLerJQ38ti4VlPbJh3t5gDobpi-wr28zf8/edit?gid=0#gid=0" 
            target="_blank" 
            style="
                text-decoration: none; 
                color: #1b1b1e;
                background-color: #f0f2f6;
                padding: 5px 10px;
                border-radius: 5px;
                font-size: 12px;
                display: inline-block;
                ">
                🔗 업무보고 및 주요일정 DB
            </a>
            ''', unsafe_allow_html=True)

        # 지원서 관리 메뉴
        elif menu == "🚀 채용 전형관리":
            st.markdown("##### 🚀 채용 전형관리")
            st.markdown("<br>", unsafe_allow_html=True)
            # CSS 스타일 정의
            st.markdown("""
                <style>
                a {
                    text-decoration: none !important;
                }
                .link-hover {
                    color: #1b1b1e;
                    font-size: 13px;
                    transition: color 0.3s;
                    display: block;
                    margin: 0;
                    padding: 0;
                    line-height: 1;
                }
                .link-hover:hover {
                    color: #0066ff !important;
                    text-decoration: none !important;
                }
                .category-title {
                    color: #1b1b1e;
                    font-size: 14px;
                    font-weight: 600;
                    margin-top: 5px;
                    margin-bottom: 2px;
                    line-height: 1;
                }
                .link-container {
                    margin-left: 10px;
                    line-height: 1;
                }
                </style>
            """, unsafe_allow_html=True)
            st.markdown("###### 📝 채용 관리 시스템")
            
            with st.expander("👇 링크 바로가기 "):
                # 1. 지원자 접수
                st.markdown('<div class="category-title">1️⃣ 채용공고 관리</div>', unsafe_allow_html=True)
                st.markdown('<div class="link-container">', unsafe_allow_html=True)
                st.markdown('<a href="https://www.notion.so/neurophethr/Career_ADMIN-74f617b482894f5ba7196833eeaed2ef" target="_blank" class="link-hover">▫️뉴로핏 커리어 공고 업데이트</a>', unsafe_allow_html=True)
                st.markdown('<a href="https://app.oopy.io/home?utm_source=oopy&utm_medium=homepage" target="_blank" class="link-hover">▫️뉴로핏 커리어 웹호스팅(우피)</a>', unsafe_allow_html=True)
                st.markdown('<a href="https://career.neurophet.com/" target="_blank" class="link-hover">▫️뉴로핏 커리어 </a>', unsafe_allow_html=True)
                st.markdown('<a href="https://docs.google.com/spreadsheets/d/1SfVtvaHgXesDFtdFozt9CJD8aQpPBrK76AxNj-OOfFE/edit?gid=0#gid=0" target="_blank" class="link-hover">▫️평가기준 및 채용공고 DB</a>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

                # 1. 지원자 접수
                st.markdown('<div class="category-title">2️⃣ 지원자 접수</div>', unsafe_allow_html=True)
                st.markdown('<div class="link-container">', unsafe_allow_html=True)
                st.markdown('<a href="https://docs.google.com/spreadsheets/d/1o5tLJr-6NbYZiImU7IKBUTtjVaeU-HI_pNxNvvF2f5c/edit?gid=126612072#gid=126612072" target="_blank" class="link-hover">▫️구글 지원자 DB</a>', unsafe_allow_html=True)
                st.markdown('<a href="https://neurophet.sharepoint.com/sites/HR2/SitePages/%EC%B1%84%EC%9A%A9-%EC%A0%84%ED%98%95%EA%B4%80%EB%A6%AC.aspx" target="_blank" class="link-hover">▫️지원자 정보 업데이트</a>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

                # 2. 면접 전형
                st.markdown('<div class="category-title">3️⃣ 면접 전형</div>', unsafe_allow_html=True)
                st.markdown('<div class="link-container">', unsafe_allow_html=True)
                st.markdown('<a href="https://hr-resume-uzu5bngyefgcv5ykngnhcd.streamlit.app" target="_blank" class="link-hover">▫️채용 가이드 및 AI분석</a>', unsafe_allow_html=True)
                st.markdown('<a href="https://hr-resume-uzu5bngyefgcv5ykngnhcd.streamlit.app/~/+/?page=admin" target="_blank" class="link-hover">▫️면접평가서 조회 및 PDF 다운로드</a>', unsafe_allow_html=True)
                st.markdown('<a href="https://docs.google.com/spreadsheets/d/1zwYJ2hwneCeSgd6p4s9ngll8PDmhLhq9qOTRo5SLCz8/edit?gid=0#gid=0" target="_blank" class="link-hover">▫️면접평가서 DB</a>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            
            # PDF 병합 기능
            st.markdown("###### 📑이력서 PDF 병합")
            
            
            
            tab1, tab2 = st.tabs(["구글 드라이브 링크로 병합", "파일 업로드로 병합"])
            
            with tab1:
                # 1. 파일 ID 추출 함수
                def extract_file_id(link):
                    try:
                        return link.split("/d/")[1].split("/")[0]
                    except:
                        return None

                # 2. 다운로드 함수
                def download_pdf_from_drive(file_id, save_path):
                    try:
                        url = f"https://drive.google.com/uc?export=download&id={file_id}"
                        response = requests.get(url, allow_redirects=True)
                        
                        # PDF 여부 확인
                        if response.status_code == 200 and b"%PDF" in response.content[:1024]:
                            with open(save_path, "wb") as f:
                                f.write(response.content)
                            return True
                        else:
                            st.error(f"PDF 파일이 아니거나 다운로드 실패: {url}")
                            return False
                    except Exception as e:
                        st.error(f"파일 다운로드 중 오류 발생: {str(e)}")
                        return False

                # 3. PDF 병합 UI
                links = st.text_area("Google Drive PDF 링크를 '링크가있는 모든 사용자'로 공유하고, 한 줄에 하나씩 입력해주세요.", height=100)

                if st.button("구글 드라이브 PDF 병합"):
                    link_list = [l.strip() for l in links.splitlines() if l.strip()]
                    if not link_list:
                        st.warning("PDF 링크를 입력해주세요.")
                    else:
                        with st.spinner("PDF 병합 중..."):
                            # Windows 환경에서 임시 디렉토리 생성
                            temp_dir = os.path.join(tempfile.gettempdir(), 'pdf_merge_temp')
                            os.makedirs(temp_dir, exist_ok=True)
                            
                            try:
                                merger = PdfMerger()
                                download_success = False
                                
                                for i, link in enumerate(link_list):
                                    file_id = extract_file_id(link)
                                    if not file_id:
                                        st.error(f"링크 오류: {link}")
                                        continue
                                    
                                    # Windows 경로 형식으로 PDF 파일 경로 생성
                                    pdf_path = os.path.join(temp_dir, f'file_{i}.pdf')
                                    
                                    # 다운로드 시도
                                    if download_pdf_from_drive(file_id, pdf_path):
                                        merger.append(pdf_path)
                                        download_success = True
                                    else:
                                        st.error(f"{link} 다운로드에 실패했습니다.")
                                
                                if download_success:
                                    try:
                                        # 병합된 PDF 저장
                                        output_path = os.path.join(temp_dir, 'merged_result.pdf')
                                        merger.write(output_path)
                                        merger.close()

                                        # 파일이 실제로 생성되었는지 확인
                                        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                                            with open(output_path, "rb") as f:
                                                st.download_button(
                                                    label="📥 병합된 PDF 다운로드",
                                                    data=f,
                                                    file_name="merged_result.pdf",
                                                    mime="application/pdf"
                                                )
                                        else:
                                            st.error("PDF 병합 파일 생성에 실패했습니다.")
                                    except Exception as e:
                                        st.error(f"PDF 병합 중 오류 발생: {str(e)}")
                                else:
                                    st.error("다운로드에 성공한 PDF 파일이 없습니다.")
                            finally:
                                # 임시 파일들 정리
                                try:
                                    import shutil
                                    if os.path.exists(temp_dir):
                                        shutil.rmtree(temp_dir)
                                except Exception as e:
                                    st.warning(f"임시 파일 정리 중 오류 발생: {str(e)}")
            
            with tab2:
                uploaded_files = st.file_uploader("PDF 파일들을 선택하세요", type=['pdf'], accept_multiple_files=True)
                
                if st.button("업로드한 PDF 병합") and uploaded_files:
                    if len(uploaded_files) < 2:
                        st.warning("최소 2개 이상의 PDF 파일을 선택해주세요.")
                    else:
                        with st.spinner("PDF 병합 중..."):
                            try:
                                merger = PdfMerger() 
                                
                                # 업로드된 파일들을 병합
                                for uploaded_file in uploaded_files:
                                    merger.append(uploaded_file)
                                
                                # 병합된 PDF를 메모리에 저장
                                merged_pdf = BytesIO()
                                merger.write(merged_pdf)
                                merger.close() 
                                
                                # 다운로드 버튼 생성
                                st.download_button(
                                    label="📥 병합된 PDF 다운로드",
                                    data=merged_pdf.getvalue(),
                                    file_name="merged_result.pdf",
                                    mime="application/pdf"
                                )
                                
                            except Exception as e:
                                st.error(f"PDF 병합 중 오류가 발생했습니다: {str(e)}")
                elif not uploaded_files:
                    st.info("PDF 파일을 선택해주세요.")

        # 채용현황 메뉴
        elif menu == "🚀 채용현황":
            st.markdown("##### 🚀 채용현황")
            
            # 채용현황 데이터 로드
            @st.cache_data(ttl=300)  # 5분마다 캐시 갱신
            def load_recruitment_data():
                try:
                    # 현재 디렉토리에서 엑셀 파일 경로 설정
                    current_dir = os.path.dirname(os.path.abspath(__file__))
                    file_path = os.path.join(current_dir, "임직원 기초 데이터.xlsx")
                    
                    # 엑셀 파일에서 "채용-공고현황" 시트 읽기 
                    df = pd.read_excel(file_path, sheet_name="채용-공고현황")
                    
                    # 채용진행년도를 문자열로 변환
                    if '채용진행년도' in df.columns:
                        df['채용진행년도'] = df['채용진행년도'].astype(str)
                        # 빈 문자열이나 'nan'은 제외
                        df = df[df['채용진행년도'].str.strip() != '']
                        df = df[df['채용진행년도'] != 'nan']
                    
                    # TO와 확정 컬럼을 숫자로 변환
                    if 'TO' in df.columns:
                        df['TO'] = pd.to_numeric(df['TO'], errors='coerce').fillna(0).astype(int)
                    if '확정' in df.columns:
                        df['확정'] = pd.to_numeric(df['확정'], errors='coerce').fillna(0).astype(int)
                    
                    # 날짜 컬럼 변환 시도
                    if '공고게시일자' in df.columns:
                        # 원본 값 보존
                        df['공고게시일자_원본'] = df['공고게시일자'].astype(str)
                        # 날짜 변환 시도
                        df['공고게시일자'] = pd.to_datetime(df['공고게시일자'], errors='coerce')
                        # 변환 실패한 경우 원본 값으로 복원
                        df.loc[df['공고게시일자'].isna(), '공고게시일자'] = df.loc[df['공고게시일자'].isna(), '공고게시일자_원본']
                        # 임시 컬럼 삭제
                        df = df.drop('공고게시일자_원본', axis=1)
                    
                    return df
                except Exception as e:
                    st.error(f"채용현황 데이터를 불러오는 중 오류가 발생했습니다: {str(e)}")
                    return None

            # 데이터 로드
            recruitment_df = load_recruitment_data()     
            if recruitment_df is not None:
                # 조회 조건 설정
                col1, col2, col3 = st.columns([0.2, 0.2, 0.6])
                
                with col1:
                    # 채용진행년도 선택 (문자열 처리, '0' 제외)
                    years = sorted([year for year in recruitment_df['채용진행년도'].unique() 
                                  if year not in ['0', 'nan', ''] and year.strip()], reverse=True)
                    if not years:
                        st.error("유효한 채용진행년도 데이터가 없습니다.")
                    selected_year = st.selectbox("채용진행년도", years if years else [str(datetime.now().year)])
                
                with col2:
                    # 채용상태 선택 ('0' 제외)
                    statuses = ['전체'] + sorted([str(status) for status in recruitment_df['채용상태'].unique() 
                                                if pd.notna(status) and str(status) not in ['0', 'nan', ''] and str(status).strip()])
                    selected_status = st.selectbox("채용상태", statuses)
                
                with col3:
                    # 여백 컬럼
                    st.empty()

                # 데이터 필터링 (문자열 비교)
                filtered_df = recruitment_df[recruitment_df['채용진행년도'] == selected_year]
                if selected_status != '전체':
                    filtered_df = filtered_df[filtered_df['채용상태'].astype(str) == selected_status]

                # 통계 계산
                stats_df = filtered_df.groupby('본부').agg({
                    'TO': 'sum',
                    '확정': 'sum',
                    '채용상태': lambda x: ', '.join(sorted(set(x)))  # 중복 제거하고 정렬하여 표시
                }).reset_index()

                # 합계 행 추가
                total_row = pd.DataFrame({
                    '본부': ['합계'],
                    'TO': [stats_df['TO'].sum()],
                    '확정': [stats_df['확정'].sum()],
                    '채용상태': ['']  # 합계 행의 채용상태는 빈 값으로
                })
                stats_df = pd.concat([stats_df, total_row])

                # 통계 표시
                col_stats1, col_stats2 = st.columns([0.5, 0.5])
                
                with col_stats1:
                    st.dataframe(
                        stats_df,
                        column_config={
                            "본부": st.column_config.TextColumn("본부", width=150),
                            "TO": st.column_config.NumberColumn("TO", width=80),
                            "확정": st.column_config.NumberColumn("확정", width=80),
                            "채용상태": st.column_config.TextColumn("채용상태", width=200)
                        },
                        hide_index=True
                    )
                
                with col_stats2:
                    # 여백 컬럼
                    st.empty()

                # 상세 리스트 표시
                st.markdown("###### 📋 채용 포지션 리스트")
                
                # 데이터프레임 인덱스 재설정 (1부터 시작)
                filtered_df = filtered_df.reset_index(drop=True)
                filtered_df.index = filtered_df.index + 1
                
                # 표시할 컬럼 선택 및 정렬
                display_df = filtered_df[['본부', '부서', '포지션명', 'TO', '확정', '채용상태', '공고게시일자', '채용진행년도']]
                
                st.dataframe(
                    display_df,
                    column_config={
                        "본부": st.column_config.TextColumn("본부", width=120),
                        "부서": st.column_config.TextColumn("부서", width=120),
                        "포지션명": st.column_config.TextColumn("포지션명", width=150),
                        "TO": st.column_config.NumberColumn("TO", width=80),
                        "확정": st.column_config.NumberColumn("확정", width=80),
                        "채용상태": st.column_config.TextColumn("채용상태", width=100),
                        "공고게시일자": st.column_config.TextColumn("공고게시일자", width=120),
                        "채용진행년도": st.column_config.NumberColumn("채용진행년도", width=100)
                    },
                    hide_index=False
                )
            else:
                st.warning("채용현황 데이터를 불러올 수 없습니다.")
                
            st.markdown("---")
            st.markdown("##### 👥 면접자 현황")
            
            # 면접 현황 데이터 로드
            @st.cache_data(ttl=300)  # 5분마다 캐시 갱신
            def load_interview_data():
                try:
                    # 현재 디렉토리에서 엑셀 파일 경로 설정
                    current_dir = os.path.dirname(os.path.abspath(__file__))
                    file_path = os.path.join(current_dir, "임직원 기초 데이터.xlsx")
                    
                    # 엑셀 파일에서 "채용-면접" 시트 읽기
                    df = pd.read_excel(file_path, sheet_name="채용-면접")
                    
                    # 면접일자가 비어있는 행 제거
                    df = df.dropna(subset=['면접일자'])
                    
                    # 성명이 0인 행 제거
                    df = df[df['성명'] != 0]
                    df = df[df['성명'] != '0']
                    
                    # 면접일자를 datetime으로 변환
                    def convert_to_datetime(x):
                        try:
                            if pd.isna(x):
                                return None
                            elif isinstance(x, (datetime, pd.Timestamp)):
                                return x
                            elif isinstance(x, date):
                                return datetime.combine(x, time())
                            elif isinstance(x, time):
                                return datetime.combine(datetime.now().date(), x)
                            elif isinstance(x, str):
                                return pd.to_datetime(x)
                            elif isinstance(x, (int, float)):
                                # 엑셀 날짜 숫자 처리
                                return pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(x))
                            else:
                                return None
                        except:
                            return None

                    df['면접일자'] = df['면접일자'].apply(convert_to_datetime)
                    
                    # 변환 실패한 데이터 제거
                    df = df.dropna(subset=['면접일자'])
                    
                    return df
                except Exception as e:
                    st.error(f"면접 현황 데이터를 불러오는 중 오류가 발생했습니다: {str(e)}")
                    return None

            # 데이터 로드
            interview_df = load_interview_data()
            
            if interview_df is not None and len(interview_df) > 0:
                # 조회 조건 설정
                col1, col2, col3, col4 = st.columns([0.2, 0.2, 0.2, 0.4 ])
                
                with col1:
                    # 시작일 선택 (오늘 - 15일)
                    start_date = st.date_input(
                        "시작일",
                        value=datetime.now().date() - timedelta(days=15),
                        help="면접일정 조회 시작일을 선택하세요."
                    )
                
                with col2:
                    # 종료일 선택 (오늘 + 30일)
                    end_date = st.date_input(
                        "종료일",
                        value=datetime.now().date() + timedelta(days=30),
                        help="면접일정 조회 종료일을 선택하세요."
                    )
                
                with col3:
                    # 전형구분 선택 (None 값과 0 값 처리)
                    interview_types = ['전체'] + sorted([
                        str(t) for t in interview_df['전형구분'].unique() 
                        if pd.notna(t) and str(t) != '0' and str(t) != '0.0' and t != 0
                    ])
                    selected_type = st.selectbox("전형구분", interview_types)
                
                with col4:
                    # 여백 컬럼
                    st.empty()
                # 데이터 필터링
                filtered_df = interview_df[
                    (interview_df['면접일자'].dt.date >= start_date) &
                    (interview_df['면접일자'].dt.date <= end_date)
                ]
                
                if selected_type != '전체':
                    filtered_df = filtered_df[filtered_df['전형구분'].astype(str) == selected_type]

                if len(filtered_df) > 0:
                    # 표시할 컬럼 선택
                    display_columns = ['채용분야', '성명', '전형구분', '면접일자', '면접일시', '특이사항']
                    display_df = filtered_df[display_columns].copy()
                    
                    # 면접일자 기준으로 내림차순 정렬
                    display_df = display_df.sort_values('면접일자', ascending=False)
                    
                    # 면접일자 포맷 변경
                    display_df['면접일자'] = display_df['면접일자'].dt.strftime('%Y-%m-%d')
                    
                    # 인덱스 1부터 시작하도록 설정
                    display_df = display_df.reset_index(drop=True)
                    display_df.index = display_df.index + 1
                    
                    # 데이터프레임 표시
                    st.dataframe(
                        display_df,
                        column_config={
                            "채용분야": st.column_config.TextColumn("채용분야", width=150),
                            "성명": st.column_config.TextColumn("성명", width=100),
                            "전형구분": st.column_config.TextColumn("전형구분", width=100),
                            "면접일자": st.column_config.TextColumn("면접일자", width=100),
                            "면접일시": st.column_config.TextColumn("면접일시", width=200),
                            "특이사항": st.column_config.TextColumn("특이사항", width=300)
                        },
                        hide_index=False
                    )
                else:
                    st.info("선택한 기간에 해당하는 면접 데이터가 없습니다.")
            else:
                st.warning("면접 현황 데이터를 불러올 수 없습니다.")

            st.markdown("---")
            st.markdown("##### 💡 지원자 접수 통계")
            
            # 지원자 통계 데이터 로드
            @st.cache_data(ttl=300)  # 5분마다 캐시 갱신
            def load_applicant_stats():
                try:
                    # 현재 디렉토리에서 엑셀 파일 경로 설정
                    current_dir = os.path.dirname(os.path.abspath(__file__))
                    file_path = os.path.join(current_dir, "임직원 기초 데이터.xlsx")
                    
                    # 엑셀 파일에서 "채용-지원자" 시트 읽기
                    df = pd.read_excel(file_path, sheet_name="채용-지원자")
                    
                    # 성명이 0인 행 제거
                    df = df[df['성명'] != 0]
                    df = df[df['성명'] != '0']
                    
                    # 등록날짜에서 연도 추출
                    df['지원연도'] = pd.to_datetime(df['등록날짜']).dt.year
                    
                    return df
                except Exception as e:
                    st.error(f"지원자 통계 데이터를 불러오는 중 오류가 발생했습니다: {str(e)}")
                    return None

            # 데이터 로드
            applicant_df = load_applicant_stats()
            
            if applicant_df is not None and len(applicant_df) > 0:
                # 연도 선택
                years = sorted(applicant_df['지원연도'].unique(), reverse=True)
                selected_year = st.selectbox("조회연도", years, key="applicant_year")
                
                # 선택된 연도의 데이터만 필터링
                year_df = applicant_df[applicant_df['지원연도'] == selected_year]
                
                # 접수방법 통계
                col1, col2 = st.columns(2)
                
                with col1:
                    
                    # 접수방법 순서 정의
                    channel_order = ['뉴로핏커리어', '사내추천', '원티드', '헤드헌팅', '점핏', '인재서치', '기타']
                    
                    # 접수방법별 카운트
                    channel_stats = year_df['접수방법'].value_counts().reindex(channel_order).fillna(0)
                    total_channel = channel_stats.sum()
                    # 차트 생성
                    fig_channel = px.bar(
                        x=channel_stats.index,
                        y=channel_stats.values,
                        labels={'x': '', 'y': '지원자 수'},
                        title=f"{selected_year}년 접수방법별 지원자 현황 (총 {int(total_channel):,}명)"
                    )
                    
                    # 차트 스타일 설정
                    colors = ['#FF4B4B' if x == '뉴로핏커리어' else '#FFB6B6' for x in channel_stats.index]
                    fig_channel.update_traces(marker_color=colors)
                    # 막대 위에 값 표시 추가
                    fig_channel.update_traces(
                        text=channel_stats.values.astype(int),
                        textposition='outside'
                    )
                    fig_channel.update_layout(
                        showlegend=False,
                        height=450,
                        title_x=0,
                        title_y=0.95,
                        margin=dict(t=70)  # 상단 여백을 더 크게 증가
                    )
                    
                    # 차트 표시
                    st.plotly_chart(fig_channel, use_container_width=True)
                with col2:
                    # 여백 컬럼
                    st.empty()

                col1, col2 = st.columns([0.9, 0.1])
                with col1:
                    
                    # 전형결과 순서 정의
                    result_order = [
                        '[1]서류검토', '[2]서류합격', '[3]1차면접합격', '[4]2차면접합격', '[5]최종합격','입사포기',
                        '서류불합격', '1차면접불합격', '2차면접불합격', '면접불참',  '보류', '연락안됨'
                    ]
                    
                    # 전형결과별 카운트
                    result_stats = year_df['전형 결과'].value_counts().reindex(result_order).fillna(0)
                    total = result_stats.sum()
                    
                    # '합계' 항목 제외
                    result_stats = result_stats[result_stats.index != '합계']
                    
                    # 차트 생성
                    fig_result = px.bar(
                        x=result_stats.index,
                        y=result_stats.values,
                        labels={'x': '전형 결과', 'y': '지원자 수'},
                        title=f"{selected_year}년 전형결과별 현황 (총 {int(total):,}명)"
                    )
                    
                    # 차트 스타일 설정
                    colors = ['#FF4B4B' if x in ['[5]최종합격', '입사포기'] else '#FFB6B6' for x in result_stats.index]
                    fig_result.update_traces(marker_color=colors) 
                    # 막대 위에 값 표시 추가
                    fig_result.update_traces(
                        text=result_stats.values.astype(int),
                        textposition='outside'
                    )
                    fig_result.update_layout(
                        showlegend=False,
                        height=450,
                        title_x=0,
                        title_y=0.95,
                        margin=dict(t=70),  # 상단 여백 설정
                        xaxis_tickangle=45
                    )
                    
                    # 차트 표시
                    st.plotly_chart(fig_result, use_container_width=True)
                
            else:
                st.warning("지원자 통계 데이터를 불러올 수 없습니다.")

except Exception as e:
    st.error(f"데이터를 불러오는 중 오류가 발생했습니다: {str(e)}")   