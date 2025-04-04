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

# 페이지 설정
st.set_page_config(
    page_title="인사팀 대시보드",
    page_icon="👥",
    layout="wide"
)

# 데이터 로드 함수
@st.cache_data
def load_data(uploaded_file=None):
    try:
        if uploaded_file is not None:
            df = pd.read_excel(uploaded_file)
            return df
        else:
            st.warning("Excel 파일을 업로드해주세요.")
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
st.sidebar.title("👥 임직원 현황")
st.sidebar.markdown("---")

# 네비게이션 메뉴
menu = st.sidebar.radio(
    "메뉴",
    ["현재 인원현황", "연도별 인원 통계", "🔍 임직원 검색"],
    index=0,
    format_func=lambda x: f"📊 {x}" if x == "현재 인원현황" else (f"📈 {x}" if x == "연도별 인원 통계" else f"{x}")
)

# 사이드바 설정
with st.sidebar:
    
    # 하단에 파일 업로드 추가
    st.sidebar.markdown("---")  # 구분선 추가
    st.sidebar.markdown("### 데이터 업데이트")
    uploaded_file = st.sidebar.file_uploader(
        "Excel 파일 업로드",
        type=["xlsx", "xls"],
        help="Excel 파일을 업로드해주세요."
    )

try:
    # 데이터 로드
    df = load_data(uploaded_file)
    
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
            
            # 그래프를 위한 컬럼 생성 (60:40 비율)
            graph_col, space_col = st.columns([0.6, 0.4])
            
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

        else:  # 임직원 검색            # 연락처 검색
            st.markdown("#### 🔍 연락처 검색")
            search_name = st.text_input("성명으로 검색", key="contact_search")
            
            if search_name:
                contact_df = df[df['성명'].str.contains(search_name, na=False)]
                if not contact_df.empty:
                    contact_info = contact_df[['본부', '팀', 'E-Mail', '핸드폰', '주소']].reset_index(drop=True)
                    contact_info.index = contact_info.index + 1
                    contact_info = contact_info.rename_axis('No.')
                    st.dataframe(contact_info, use_container_width=True)
                else:
                    st.info("검색 결과가 없습니다.")

            st.markdown("---")

            # 생일자 검색
            st.markdown("#### 🎂생일자 검색")
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

except Exception as e:
    st.error(f"데이터를 불러오는 중 오류가 발생했습니다: {str(e)}") 