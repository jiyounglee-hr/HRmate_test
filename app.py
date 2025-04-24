import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import io
import requests
from PIL import Image
from io import BytesIO
import re
import plotly.io as pio
import numpy as np
from dateutil.relativedelta import relativedelta
from docx import Document
import tempfile
import PyPDF2
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer
import subprocess
import platform

def docx_to_pdf(docx_path, pdf_path):
    try:
        # Word 문서 읽기
        doc = Document(docx_path)
        
        # PDF 생성
        c = canvas.Canvas(pdf_path, pagesize=letter)
        width, height = letter
        
        # 한글 폰트 등록
        try:
            pdfmetrics.registerFont(TTFont('HanSans', 'NanumGothic.ttf'))
            c.setFont('HanSans', 12)
        except:
            c.setFont('Helvetica', 12)
        
        y = height - 50
        for para in doc.paragraphs:
            if y < 50:
                c.showPage()
                y = height - 50
                try:
                    c.setFont('HanSans', 12)
                except:
                    c.setFont('Helvetica', 12)
            
            text = para.text
            c.drawString(50, y, text)
            y -= 20
        
        c.save()
        return True
    except Exception as e:
        st.error(f"변환 중 오류 발생: {str(e)}")
        return False

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
        
        # 데이터 로드 시간 표시
        st.sidebar.markdown(f"*마지막 데이터 업데이트: {datetime.fromtimestamp(last_modified).strftime('%Y-%m-%d %H:%M:%S')}*")
        
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
        max-width: 1000px;
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
st.sidebar.markdown("### HR Data")
if st.sidebar.button("📊 현재 인원현황", use_container_width=True):
    st.session_state.menu = "📊 현재 인원현황"
if st.sidebar.button("📈 연도별 인원 통계", use_container_width=True):
    st.session_state.menu = "📈 연도별 인원 통계"
if st.sidebar.button("🔍 임직원 검색", use_container_width=True):
    st.session_state.menu = "🔍 임직원 검색"
if st.sidebar.button("😊 임직원 명부", use_container_width=True):
    st.session_state.menu = "😊 임직원 명부"

st.sidebar.markdown("---")

# HR Support 섹션
st.sidebar.markdown("### HR Support")
if st.sidebar.button("🏦 기관제출용 인원현황", use_container_width=True):
    st.session_state.menu = "🏦 기관제출용 인원현황"
if st.sidebar.button("📋 채용_처우협상", use_container_width=True):
    st.session_state.menu = "📋 채용_처우협상"
if st.sidebar.button("⏰ 초과근무 조회", use_container_width=True):
    st.session_state.menu = "⏰ 초과근무 조회"
if st.sidebar.button("📅 인사발령 내역", use_container_width=True):
    st.session_state.menu = "📅 인사발령 내역"
if st.sidebar.button("🔗 채용_이력서 pdf변환", use_container_width=True):
    st.session_state.menu = "🔗 채용_이력서 pdf변환"

# 채용서포트 링크 추가
st.sidebar.markdown("---")
st.sidebar.markdown("##### 참고 사이트")
st.sidebar.markdown('<a href="https://hr-resume-uzu5bngyefgcv5ykngnhcd.streamlit.app/" target="_blank" class="sidebar-link" style="text-decoration: none;">📋 채용(이력서 분석)</a>', unsafe_allow_html=True)
st.sidebar.markdown('<a href="https://neuropr-lwm9mzur3rzbgoqrhzy68n.streamlit.app/" target="_blank" class="sidebar-link" style="text-decoration: none;">📰 PR(뉴스검색 및 기사초안)</a>', unsafe_allow_html=True)

# 기본 메뉴 설정
if 'menu' not in st.session_state:
    st.session_state.menu = "📊 현재 인원현황"
menu = st.session_state.menu

def convert_to_pdf(uploaded_files):
    pdf_paths = []
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            for uploaded_file in uploaded_files:
                # 임시 파일 경로 생성
                input_path = os.path.join(temp_dir, uploaded_file.name)
                output_path = os.path.join(temp_dir, f"{os.path.splitext(uploaded_file.name)[0]}.pdf")
                
                # 파일 저장
                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # LibreOffice로 변환
                if platform.system() == "Windows":
                    subprocess.run([
                        "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
                        "--headless",
                        "--convert-to", "pdf",
                        "--outdir", temp_dir,
                        input_path
                    ], check=True)
                else:
                    subprocess.run([
                        "libreoffice",
                        "--headless",
                        "--convert-to", "pdf",
                        "--outdir", temp_dir,
                        input_path
                    ], check=True)
                
                # 변환된 PDF 파일 경로 저장
                pdf_paths.append(output_path)
        return pdf_paths
    except Exception as e:
        st.error(f"PDF 변환 중 오류가 발생했습니다: {str(e)}")
        return None

# 데이터 로드
try:
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
        
        # 메뉴 선택에 따른 처리
        if menu == "📊 대시보드":
            # ... existing code ...
            pass
        elif menu == "📝 이력서 변환":
            st.title("📝 이력서 변환")
            st.write("여러 형식의 이력서를 PDF로 변환하고 병합합니다.")
            
            # 파일 업로드
            uploaded_files = st.file_uploader(
                "이력서 파일을 업로드하세요 (PDF, DOCX, PPTX)",
                type=["pdf", "docx", "pptx"],
                accept_multiple_files=True
            )
            
            if uploaded_files:
                pdf_paths = convert_to_pdf(uploaded_files)
                if pdf_paths:
                    try:
                        # PDF 병합
                        merger = PyPDF2.PdfMerger()
                        for pdf_path in pdf_paths:
                            merger.append(pdf_path)
                        
                        # 병합된 PDF를 메모리에 저장
                        output = io.BytesIO()
                        merger.write(output)
                        merger.close()
                        output.seek(0)
                        
                        # 다운로드 버튼
                        st.download_button(
                            label="📥 병합된 PDF 다운로드",
                            data=output,
                            file_name="merged_resume.pdf",
                            mime="application/pdf"
                        )
                    except Exception as e:
                        st.error(f"PDF 병합 중 오류가 발생했습니다: {str(e)}")
except Exception as e:
    st.error(f"데이터를 불러오는 중 오류가 발생했습니다: {str(e)}")

# ... existing code ...