import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import io
import requests
from PIL import Image
from io import BytesIO

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
@st.cache_data
def load_data():
    try:
        # 현재 스크립트의 디렉토리 경로 가져오기
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # 자동으로 엑셀 파일 찾기
        excel_files = [f for f in os.listdir(current_dir) if f.endswith(('.xlsx', '.xls'))]
        if excel_files:
            # 가장 최근 수정된 엑셀 파일 선택
            latest_file = max(excel_files, key=lambda x: os.path.getmtime(os.path.join(current_dir, x)))
            file_path = os.path.join(current_dir, latest_file)
            df = pd.read_excel(file_path)
            return df
        else:
            st.warning("Excel 파일을 찾을 수 없습니다.")
            return None
    except Exception as e:
        st.error(f"파일을 불러오는 중 오류가 발생했습니다: {str(e)}")
        return None

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
    </style>
""", unsafe_allow_html=True)

# 제목
st.sidebar.title("👥 HRmate")
st.sidebar.markdown("---")

# 네비게이션 메뉴
menu = st.sidebar.radio(
    " ",
    ["현재 인원현황", "연도별 인원 통계", "🔍 임직원 검색", "🏦 기관제출용 인원현황", "📋 채용_처우협상", "⏰ 초과근무 조회"],
    index=0,
    format_func=lambda x: f"📊 {x}" if x == "현재 인원현황" else (f"📈 {x}" if x == "연도별 인원 통계" else f"{x}")
)
# 채용서포트 링크 추가
st.sidebar.markdown("---")
st.sidebar.markdown("##### 참고 사이트")
st.sidebar.markdown('<a href="https://hr-resume-uzu5bngyefgcv5ykngnhcd.streamlit.app/" target="_blank" class="sidebar-link" style="text-decoration: none;">📋 채용(이력서 분석)</a>', unsafe_allow_html=True)
st.sidebar.markdown('<a href="https://neuropr-lwm9mzur3rzbgoqrhzy68n.streamlit.app/" target="_blank" class="sidebar-link" style="text-decoration: none;">📰 PR(뉴스검색 및 기사초안)</a>', unsafe_allow_html=True)
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
        
        if menu == "현재 인원현황":
            # 기본 통계
            if '재직상태' in df.columns and '정규직전환일' in df.columns:
                재직자 = len(df[df['재직상태'] == '재직'])
                
                # 정규직/계약직 입퇴사자 계산
                정규직_입사자 = len(df[(df['정규직전환연도'] == 2025) & (df['고용구분'] == '정규직')])
                정규직_퇴사자 = len(df[(df['퇴사연도'] == 2025) & (df['고용구분'] == '정규직')])
                계약직_입사자 = len(df[(df['정규직전환연도'] == 2025) & (df['고용구분'] == '계약직')])
                계약직_퇴사자 = len(df[(df['퇴사연도'] == 2025) & (df['고용구분'] == '계약직')])
                
                # 퇴사율 계산 (소수점 첫째자리까지)
                재직_정규직_수 = len(df[(df['고용구분'] == '정규직') & (df['재직상태'] == '재직')])
                퇴사율 = round((정규직_퇴사자 / 재직_정규직_수 * 100), 1) if 재직_정규직_수 > 0 else 0
                
                 # 기본통계 분석
                st.markdown("##### ㆍ현재 인원 현황")
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
                col1, col2, col3 = st.columns([0.4, 0.4, 0.2])
                
                # 현재 재직자 필터링
                current_employees = df[df['재직상태'] == '재직']
                
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
                        height=400,
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
                    st.plotly_chart(fig_dept, use_container_width=True, key="dept_chart")
                
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
                        height=400
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
                    st.plotly_chart(fig_position, use_container_width=True, key="position_chart")
                
                with col3:
                    st.write("")  # 빈 공간

                st.markdown("<br>", unsafe_allow_html=True)
                
                # 2025년 입퇴사자 현황
                list_col1, list_col2 = st.columns(2)
                
                with list_col1:
                    st.markdown("##### ㆍ2025년 입사자")
                    입사자_df = df[df['정규직전환연도'] == 2025][['성명', '팀', '직위', '정규직전환일']]
                    if not 입사자_df.empty:
                        입사자_df = 입사자_df.sort_values('정규직전환일')
                        입사자_df = 입사자_df.reset_index(drop=True)
                        입사자_df.index = 입사자_df.index + 1
                        입사자_df = 입사자_df.rename_axis('No.')
                        st.dataframe(입사자_df.style.format({'정규직전환일': lambda x: x.strftime('%Y-%m-%d')}),
                                   use_container_width=True)
                    else:
                        st.info("2025년 입사 예정자가 없습니다.")

                with list_col2:
                    st.markdown("##### ㆍ2025년 퇴사자")
                    퇴사자_df = df[df['퇴사연도'] == 2025][['성명', '팀', '직위', '퇴사일']]
                    if not 퇴사자_df.empty:
                        퇴사자_df = 퇴사자_df.sort_values('퇴사일')
                        퇴사자_df = 퇴사자_df.reset_index(drop=True)
                        퇴사자_df.index = 퇴사자_df.index + 1
                        퇴사자_df = 퇴사자_df.rename_axis('No.')
                        st.dataframe(퇴사자_df.style.format({'퇴사일': lambda x: x.strftime('%Y-%m-%d')}),
                                   use_container_width=True)
                    else:
                        st.info("2025년 퇴사자가 없습니다.")
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # 근속기간별 퇴사자 현황 분석
                st.markdown("##### ㆍ정규직 퇴사자 현황")
                
                # 퇴사연도 선택 드롭다운과 퇴사인원 표시를 위한 컬럼 생성
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    # 퇴사연도 선택 드롭다운
                    available_years = sorted(df[df['재직상태'] == '퇴직']['퇴사연도'].dropna().unique())
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
                        title=title_text,
                        height=400,
                        showlegend=False,
                        plot_bgcolor='white',
                        yaxis=dict(
                            title="퇴사자 수 (명)",
                            range=[0, max(max(tenure_counts.values) * 1.2, 10)],
                            gridcolor='lightgray',
                            gridwidth=0.5,
                        ),
                        xaxis=dict(
                            title="근속기간",
                            showgrid=False,
                        ),
                        margin=dict(t=50)
                    )

                    st.plotly_chart(fig, use_container_width=True)

                with space_col:
                    st.write("")  # 빈 공간

                # 본부별 근속기간 분석을 위한 테이블 생성
                st.markdown("<br>", unsafe_allow_html=True)
                
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
                        font-size: 14px;
                        width: 100%;
                        border-collapse: collapse;
                    }
                    .custom-table th {
                        background-color: #f0f2f6;
                        padding: 8px;
                        text-align: center;
                        border: 1px solid #ddd;
                    }
                    .custom-table td {
                        padding: 8px;
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

        elif menu == "연도별 인원 통계":
            # 최근 5년간 인원 현황 분석
            st.markdown("##### ㆍ최근 5년간 입퇴사 현황")
            
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
            
            # 하드코딩된 데이터로 DataFrame 생성
            stats_df = pd.DataFrame([
                {'연도': 2021, '전체': get_year_end_headcount(df, 2021)[0], '정규직_전체': get_year_end_headcount(df, 2021)[1], '계약직_전체': get_year_end_headcount(df, 2021)[2], '정규직_입사': 40, '정규직_퇴사': 24, '계약직_입사': 6, '계약직_퇴사': 6},
                {'연도': 2022, '전체': get_year_end_headcount(df, 2022)[0], '정규직_전체': get_year_end_headcount(df, 2022)[1], '계약직_전체': get_year_end_headcount(df, 2022)[2], '정규직_입사': 46, '정규직_퇴사': 16, '계약직_입사': 12, '계약직_퇴사': 11},
                {'연도': 2023, '전체': get_year_end_headcount(df, 2023)[0], '정규직_전체': get_year_end_headcount(df, 2023)[1], '계약직_전체': get_year_end_headcount(df, 2023)[2], '정규직_입사': 30, '정규직_퇴사': 14, '계약직_입사': 21, '계약직_퇴사': 19},
                {'연도': 2024, '전체': get_year_end_headcount(df, 2024)[0], '정규직_전체': get_year_end_headcount(df, 2024)[1], '계약직_전체': get_year_end_headcount(df, 2024)[2], '정규직_입사': 55, '정규직_퇴사': 23, '계약직_입사': 6, '계약직_퇴사': 10},
                {'연도': 2025, '전체': get_year_end_headcount(df, 2025)[0], '정규직_전체': get_year_end_headcount(df, 2025)[1], '계약직_전체': get_year_end_headcount(df, 2025)[2], '정규직_입사': 7, '정규직_퇴사': 3, '계약직_입사': 1, '계약직_퇴사': 1}
            ])
            
            # 그래프를 위한 컬럼 생성 (50:50 비율)
            graph_col, space_col = st.columns([0.5, 0.5])
            
            with graph_col:
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
                    title="연도별 전체 인원 현황",
                    title_x=0.5,
                    height=400,
                    showlegend=False,
                    plot_bgcolor='white',
                    yaxis=dict(
                        title="인원 수 (명)",
                        gridcolor='lightgray',
                        gridwidth=0.5,
                        range=[0, max(stats_df['전체']) * 1.2]
                    ),
                    xaxis=dict(
                        title="연도",
                        showgrid=False,
                        tickformat='d'  # 정수 형식으로 표시
                    ),
                    margin=dict(t=50)
                )

                st.plotly_chart(fig, use_container_width=True)

            with space_col:
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
                }),
                hide_index=True,
                width=800,
                use_container_width=False
            )

        elif menu == "🔍 임직원 검색":
            st.markdown("##### 🔍 연락처 검색")
            
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
            current_month = datetime.now().month
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
                    detail_columns = ['성명', '본부', '실', '팀', '고용구분', '입사일', '재직상태', '구분1', '구분2', '구분3']
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

        elif menu == "📋 채용_처우협상":
            st.markdown("##### 🔎 처우 기본정보")
            
            # 직군 매핑 정의
            job_mapping = {
                "연구직": "직군1",
                "개발직": "직군2",
                "임상연구, QA": "직군2",
                "연구기획": "직군3",
                "디자인": "직군3",
                "인증(RA), SV, SCM": "직군3",
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
            
            # 입력 폼 생성
            with st.form("salary_form"):
                # 1줄: 포지션명, 후보자명
                col1, col2 = st.columns(2)
                with col1:
                    position = st.text_input("포지션명", "")
                with col2:
                    candidate_name = st.text_input("후보자명", "")
                
                # 2줄: 직군선택
                job_role = st.selectbox("직군 선택", job_roles)
                
                # 3줄: 현재연봉, 기타 처우, 희망연봉
                col3, col4, col5 = st.columns(3)
                with col3:
                    current_salary = st.number_input("현재연봉 (만원)", min_value=0, step=100)
                with col4:
                    other_salary = st.number_input("기타 보상상 (만원)", min_value=0, step=100)
                with col5:
                    desired_salary = st.number_input("희망연봉 (만원)", min_value=0, step=100)
                
                # 4줄: 인정경력 연차, 학력특이사항
                col6, col7 = st.columns(2)
                with col6:
                    years = st.number_input("인정경력 (년)", min_value=-4.0, value=0.0, step=0.1, format="%.1f")
                with col7:
                    education_notes = st.text_input("특이사항", "")
                
                # 전체 경력을 년 단위로 변환 (분석용) - 반올림 적용
                years_exp = round(years)
                
                
                # 분석하기 버튼
                submitted = st.form_submit_button("분석하기")

                if submitted:
                    try:
                        # salary_table.xlsx 파일 읽기
                        salary_table = pd.read_excel("salary_table.xlsx")
                        
                        # 선택된 직군상세에 해당하는 직군 가져오기
                        selected_job_category = job_mapping[job_role]
                        
                        # 해당 직군과 연차에 맞는 데이터 필터링
                        filtered_data = salary_table[
                            (salary_table['직군'] == selected_job_category) & 
                            (salary_table['연차'] == years_exp)
                        ]
                        
                        if filtered_data.empty:
                            st.warning(f"선택하신 직군 '{job_role}' ({selected_job_category})과 연차 {years_exp}년에 해당하는 데이터가 없습니다.")
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
                            related_years = [years_exp-1, years_exp, years_exp+1]
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
                        - 인정경력: {years:.1f}년 (인정경력 기준: {years_exp}년)
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
                        selected_month = st.selectbox('조회 기준 연월을 선택하세요', sorted(months, reverse=True))
                        
                        # 선택된 연월에 해당하는 데이터 필터링
                        filtered_df = overtime_df[overtime_df['연월구분'] == selected_month]
                        
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
                        result_df = result_df[['이름', '이메일', '초과근무 내역', '초과근무시간 합']]
                        
                        # 테이블 표시
                        st.markdown("""
                            <style>
                            [data-testid="stDataFrame"] td {
                                white-space: pre-wrap;
                                min-height: fit-content;
                                height: auto !important;
                                line-height: 1.5;
                                padding: 8px !important;
                            }
                            [data-testid="stDataFrame"] div[data-testid="StyledDataFrameDataCell"] {
                                min-height: fit-content;
                                height: auto !important;
                            }
                            </style>
                        """, unsafe_allow_html=True)
                        
                        st.dataframe(
                            result_df,
                            column_config={
                                "이름": st.column_config.TextColumn("이름", width=100),
                                "초과근무시간 합": st.column_config.TextColumn("초과근무시간 합", width=100),
                                "초과근무 내역": st.column_config.TextColumn("초과근무 내역", width=300),
                                "이메일": st.column_config.TextColumn("이메일", width=100)
                            },
                            hide_index=True,
                            use_container_width=True,
                            height=400
                        )
                    else:
                        st.error("엑셀 파일에 '연월구분' 컬럼이 없습니다.")
                    
                except Exception as e:
                    st.error(f"파일을 읽는 중 오류가 발생했습니다: {str(e)}")
            else:
                st.info("초과근무 엑셀 파일을 업로드하세요.")

except Exception as e:
    st.error(f"데이터를 불러오는 중 오류가 발생했습니다: {str(e)}") 