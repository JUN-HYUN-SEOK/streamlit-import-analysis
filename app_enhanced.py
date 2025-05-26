import streamlit as st
import pandas as pd
import io
import datetime
import traceback
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import PieChart, Reference, LineChart

# 페이지 설정
st.set_page_config(
    page_title="수입신고 분석 도구 (완전판)", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# 제목과 설명
st.title("🚢 수입신고 분석 도구 (완전판)")
st.markdown("""
이 도구는 `app-new202505.py`의 모든 분석 기능을 웹으로 구현한 완전판입니다.  
**8% 환급검토, 0% Risk, 세율 Risk, 단가 Risk 분석**을 모두 제공합니다.
""")

# 사이드바 메뉴
with st.sidebar:
    st.header("📋 분석 메뉴")
    analysis_type = st.selectbox(
        "원하는 분석을 선택하세요:",
        ["전체 분석", "8% 환급 검토", "0% Risk 분석", "세율 Risk 분석", "단가 Risk 분석"]
    )
    
    st.markdown("---")
    st.subheader("📁 파일 업로드")

# 파일 업로드
uploaded_file = st.file_uploader(
    "분석할 엑셀 파일을 업로드하세요", 
    type=["xlsx", "xls", "csv"],
    help="원본 app-new202505.py와 동일한 형식의 파일을 업로드하세요"
)

def read_and_process_excel(file):
    """엑셀 파일 읽기 및 기본 전처리 (원본과 동일)"""
    try:
        st.info("📂 엑셀 파일 읽기 시작...")
        
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        
        st.write(f"원본 데이터 크기: {df.shape}")
        
        # 컬럼명 정리 (원본과 동일)
        df.columns = df.columns.str.strip()
        
        # 중복 컬럼명 처리 - 더 안전하게
        st.info("중복 컬럼명 처리 중...")
        original_columns = df.columns.tolist()
        st.write(f"원본 컬럼 목록 (처음 10개): {original_columns[:10]}")
        
        # 중복 컬럼명 찾기 및 처리
        seen = {}
        new_columns = []
        for col in df.columns:
            if col in seen:
                seen[col] += 1
                new_columns.append(f"{col}_{seen[col]}")
            else:
                seen[col] = 0
                new_columns.append(col)
        
        df.columns = new_columns
        st.write(f"중복 처리 후 컬럼 수: {len(df.columns)}")
        
        # 강제 컬럼 인식 전에 기존 컬럼 확인
        target_columns = ['세율구분', '관세실행세율']
        existing_target_cols = [col for col in target_columns if col in df.columns]
        
        if existing_target_cols:
            st.info(f"이미 존재하는 대상 컬럼: {existing_target_cols}")
        
        # 강제 컬럼 인식 (원본과 동일) - 더 안전하게 처리
        if len(df.columns) > 71:
            try:
                # 70번째, 71번째 컬럼 확인
                col_70 = df.columns[70]
                col_71 = df.columns[71]
                
                st.write(f"70번째 컬럼: '{col_70}'")
                st.write(f"71번째 컬럼: '{col_71}'")
                
                # 기존에 세율구분, 관세실행세율이 없는 경우만 매핑
                rename_dict = {}
                
                if '세율구분' not in df.columns:
                    rename_dict[col_70] = '세율구분'
                    st.info(f"'{col_70}' → '세율구분' 매핑")
                else:
                    st.info("'세율구분' 컬럼이 이미 존재합니다.")
                    
                if '관세실행세율' not in df.columns:
                    rename_dict[col_71] = '관세실행세율'
                    st.info(f"'{col_71}' → '관세실행세율' 매핑")
                else:
                    st.info("'관세실행세율' 컬럼이 이미 존재합니다.")
                
                # 실제 매핑 실행
                if rename_dict:
                    df.rename(columns=rename_dict, inplace=True)
                    st.success(f"컬럼 매핑 완료: {rename_dict}")
                else:
                    st.info("매핑할 컬럼이 없습니다.")
                
            except Exception as e:
                st.error(f"컬럼 매핑 중 오류: {e}")
                # 기본값 설정
                if '세율구분' not in df.columns:
                    df['세율구분'] = 'Unknown'
                if '관세실행세율' not in df.columns:
                    df['관세실행세율'] = 0
        else:
            st.warning(f"컬럼 수가 부족합니다. 현재: {len(df.columns)}개, 필요: 72개 이상")
            # 기본값 설정
            if '세율구분' not in df.columns:
                df['세율구분'] = 'Unknown'
            if '관세실행세율' not in df.columns:
                df['관세실행세율'] = 0
        
        # 관세실행세율 안전한 숫자 변환
        try:
            if '관세실행세율' in df.columns:
                st.info("관세실행세율 변환 시작...")
                
                # 안전한 컬럼 접근
                try:
                    # 컬럼이 중복되어 있는지 확인
                    rate_columns = [col for col in df.columns if '관세실행세율' in col]
                    st.write(f"관세실행세율 관련 컬럼들: {rate_columns}")
                    
                    # 첫 번째 관세실행세율 컬럼 사용
                    if rate_columns:
                        target_col = rate_columns[0]
                        st.write(f"사용할 컬럼: '{target_col}'")
                        
                        # 안전한 데이터 접근
                        rate_data = df[target_col]
                        st.write(f"컬럼 타입: {type(rate_data)}")
                        
                        if hasattr(rate_data, 'dtype'):
                            st.write(f"데이터 타입: {rate_data.dtype}")
                            st.write(f"샘플 데이터: {rate_data.head().tolist()}")
                            
                            # Series인지 확인하고 변환
                            if isinstance(rate_data, pd.Series):
                                # 안전한 변환
                                rate_data_clean = rate_data.fillna(0)
                                df['관세실행세율'] = pd.to_numeric(rate_data_clean, errors='coerce').fillna(0)
                                st.success("관세실행세율 변환 완료")
                            else:
                                st.warning("관세실행세율이 Series가 아닙니다. 기본값으로 설정합니다.")
                                df['관세실행세율'] = 0
                        else:
                            st.warning("관세실행세율 컬럼에 dtype 속성이 없습니다. 기본값으로 설정합니다.")
                            df['관세실행세율'] = 0
                    else:
                        st.warning("관세실행세율 컬럼을 찾을 수 없습니다.")
                        df['관세실행세율'] = 0
                        
                except Exception as inner_e:
                    st.error(f"컬럼 접근 중 오류: {inner_e}")
                    df['관세실행세율'] = 0
                    
            else:
                st.warning("관세실행세율 컬럼이 없습니다. 새로 생성합니다.")
                df['관세실행세율'] = 0
                
        except Exception as e:
            st.error(f"관세실행세율 변환 중 오류: {e}")
            df['관세실행세율'] = 0
        
        # 세율구분 안전 처리
        try:
            if '세율구분' in df.columns:
                rate_type_data = df['세율구분']
                if isinstance(rate_type_data, pd.Series):
                    df['세율구분'] = rate_type_data.astype(str).fillna('Unknown')
                else:
                    df['세율구분'] = 'Unknown'
            else:
                df['세율구분'] = 'Unknown'
        except Exception as e:
            st.error(f"세율구분 처리 중 오류: {e}")
            df['세율구분'] = 'Unknown'
        
        # 최종 중복 컬럼 확인 및 정리
        final_columns = df.columns.tolist()
        duplicate_cols = [col for col in final_columns if final_columns.count(col) > 1]
        
        if duplicate_cols:
            st.warning(f"여전히 중복된 컬럼: {list(set(duplicate_cols))}")
            # 중복 컬럼 제거 (첫 번째만 유지)
            df = df.loc[:, ~df.columns.duplicated()]
            st.info("중복 컬럼 제거 완료")
        
        st.success(f"✅ 파일 읽기 완료: {df.shape[0]:,}행 {df.shape[1]}열")
        
        # 최종 확인
        if '관세실행세율' in df.columns and '세율구분' in df.columns:
            st.write("최종 관세실행세율 샘플:", df['관세실행세율'].head().tolist())
            st.write("최종 세율구분 샘플:", df['세율구분'].head().tolist())
        else:
            st.error("필수 컬럼이 생성되지 않았습니다!")
        
        return df
        
    except Exception as e:
        st.error(f"❌ 파일 읽기 실패: {str(e)}")
        st.code(traceback.format_exc())
        return None

def create_eight_percent_refund_analysis(df):
    """8% 환급 검토 분석 (원본과 동일)"""
    try:
        st.subheader("🎯 8% 환급 검토 분석")
        
        # 필요한 컬럼 선택 (원본과 동일)
        selected_columns = [
            '수입신고번호', 'B/L번호', '세번부호', '세율구분', '세율설명',
            '관세실행세율', '규격1', '규격2', '규격3', '성분1', '성분2', '성분3',
            '실제관세액', '결제방법', '결제통화단위', '거래품명', '란번호', '행번호',
            '수량_1', '수량단위_1', '단가', '금액', '수리일자'
        ]
        
        # 존재하는 컬럼만 선택
        available_columns = [col for col in selected_columns if col in df.columns]
        df_work = df[available_columns].copy()
        
        # 데이터 전처리 (원본과 동일)
        df_work['세율구분'] = df_work['세율구분'].astype(str).str.strip()
        df_work['관세실행세율'] = pd.to_numeric(
            df_work['관세실행세율'].fillna(0), errors='coerce'
        ).fillna(0)
        
        if '실제관세액' in df_work.columns:
            df_work['실제관세액'] = pd.to_numeric(
                df_work['실제관세액'].fillna(0), errors='coerce'
            ).fillna(0)
        
        df_work.fillna(0, inplace=True)
        df_work = df_work.infer_objects(copy=False)
        
        # 필터링 조건 적용 (원본과 동일)
        df_filtered = df_work[
            (df_work['세율구분'] == 'A') & 
            (df_work['관세실행세율'] >= 8)
        ]
        
        # 결과 표시
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("🔍 필터링 조건", "세율구분='A' AND 관세실행세율≥8%")
        with col2:
            st.metric("📊 총 대상 건수", f"{len(df_filtered):,}건")
        with col3:
            if len(df_work) > 0:
                ratio = (len(df_filtered) / len(df_work)) * 100
                st.metric("📈 비율", f"{ratio:.1f}%")
        
        if len(df_filtered) > 0:
            # 관세실행세율 분포 차트
            fig = px.histogram(
                df_filtered, 
                x='관세실행세율',
                title="8% 환급 검토 대상 - 관세실행세율 분포",
                nbins=20
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # 데이터 테이블
            st.dataframe(df_filtered, use_container_width=True)
            
            # 엑셀 다운로드
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_filtered.to_excel(writer, sheet_name='8% 환급 검토', index=False)
            
            st.download_button(
                label="📥 8% 환급검토 엑셀 다운로드",
                data=excel_buffer.getvalue(),
                file_name=f"8퍼센트_환급검토_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ 8% 환급 검토 조건에 맞는 데이터가 없습니다.")
            
        return df_filtered
        
    except Exception as e:
        st.error(f"❌ 8% 환급 검토 분석 중 오류: {str(e)}")
        return pd.DataFrame()

def create_zero_percent_risk_analysis(df):
    """0% Risk 분석 (원본과 동일)"""
    try:
        st.subheader("⚠️ 0% Risk 분석")
        
        # 필요한 컬럼 선택 (원본과 동일)
        selected_columns = [
            '수입신고번호', 'B/L번호', '세번부호', '세율구분', '관세실행세율',
            '규격1', '규격2', '성분1', '실제관세액', '거래품명', '란번호', '행번호',
            '수량_1', '수량단위_1', '단가', '금액', '수리일자'
        ]
        
        # 존재하는 컬럼만 선택
        available_columns = [col for col in selected_columns if col in df.columns]
        
        # 0% Risk 조건 적용 (원본과 동일)
        df_zero_risk = df[
            (df['관세실행세율'] < 8) & 
            (~df['세율구분'].astype(str).str.match(r'^F.{3}$'))
        ]
        
        df_zero_risk.fillna(0, inplace=True)
        df_zero_risk = df_zero_risk.infer_objects(copy=False)
        df_zero_risk = df_zero_risk[available_columns]
        
        # 결과 표시
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("🔍 필터링 조건", "관세실행세율<8% AND 세율구분≠F***")
        with col2:
            st.metric("📊 총 Risk 건수", f"{len(df_zero_risk):,}건")
        with col3:
            if len(df) > 0:
                ratio = (len(df_zero_risk) / len(df)) * 100
                st.metric("📈 비율", f"{ratio:.1f}%")
        
        if len(df_zero_risk) > 0:
            # 세율구분별 분포 차트
            fig = px.pie(
                df_zero_risk['세율구분'].value_counts().reset_index(),
                values='count',
                names='세율구분',
                title="0% Risk - 세율구분별 분포"
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # 데이터 테이블
            st.dataframe(df_zero_risk, use_container_width=True)
            
            # 엑셀 다운로드
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_zero_risk.to_excel(writer, sheet_name='0% Risk', index=False)
            
            st.download_button(
                label="📥 0% Risk 엑셀 다운로드",
                data=excel_buffer.getvalue(),
                file_name=f"0퍼센트_리스크_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ 0% Risk 조건에 맞는 데이터가 없습니다.")
            
        return df_zero_risk
        
    except Exception as e:
        st.error(f"❌ 0% Risk 분석 중 오류: {str(e)}")
        return pd.DataFrame()

def create_tariff_risk_analysis(df):
    """세율 Risk 분석 (원본과 동일)"""
    try:
        st.subheader("📊 세율 Risk 분석")
        st.info("동일한 규격1에 대해 서로 다른 세번부호가 적용된 경우를 찾습니다.")
        
        # 필요한 컬럼 체크
        required_columns = [
            '수입신고번호', 'B/L번호', '수리일자', '규격1', '규격2', '규격3',
            '성분1', '성분2', '성분3', '세번부호', '세율구분', '세율설명',
            '과세가격달러', '실제관세액', '결제방법'
        ]
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.warning(f"⚠️ 누락된 컬럼: {missing_columns}")
            # 존재하는 컬럼만 사용
            available_columns = [col for col in required_columns if col in df.columns]
            if '규격1' not in available_columns or '세번부호' not in available_columns:
                st.error("❌ 필수 컬럼(규격1, 세번부호)이 없어 분석할 수 없습니다.")
                return pd.DataFrame()
        else:
            available_columns = required_columns
        
        # 규격1별 세번부호 분석 (원본과 동일)
        risk_specs = df.groupby('규격1')['세번부호'].nunique()
        risk_specs = risk_specs[risk_specs > 1]
        
        if len(risk_specs) == 0:
            st.success("✅ 세율 Risk가 발견되지 않았습니다.")
            return pd.DataFrame()
        
        # Risk 데이터 추출
        risk_data = df[df['규격1'].isin(risk_specs.index)][available_columns].copy()
        risk_data = risk_data.sort_values('규격1').fillna('')
        
        # 결과 표시
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("🔍 위험 규격1 수", f"{len(risk_specs)}개")
        with col2:
            st.metric("📊 총 Risk 건수", f"{len(risk_data):,}건")
        with col3:
            if len(df) > 0:
                ratio = (len(risk_data) / len(df)) * 100
                st.metric("📈 비율", f"{ratio:.1f}%")
        
        # 위험 규격1별 세번부호 수 차트
        fig = px.bar(
            x=risk_specs.index[:20],  # 상위 20개만
            y=risk_specs.values[:20],
            title="위험 규격1별 세번부호 수 (상위 20개)",
            labels={'x': '규격1', 'y': '세번부호 수'}
        )
        st.plotly_chart(fig, use_container_width=True)
        
        # 상세 분석 테이블
        st.subheader("🔍 세율 Risk 상세 분석")
        
        # 규격1별 그룹화하여 표시
        for spec1 in risk_specs.index[:10]:  # 상위 10개만
            with st.expander(f"📋 규격1: {spec1} (세번부호 {risk_specs[spec1]}개)"):
                spec_data = risk_data[risk_data['규격1'] == spec1]
                
                # 해당 규격1의 세번부호별 분포
                tariff_counts = spec_data['세번부호'].value_counts()
                fig_spec = px.pie(
                    values=tariff_counts.values,
                    names=tariff_counts.index,
                    title=f"규격1 '{spec1}' - 세번부호 분포"
                )
                st.plotly_chart(fig_spec, use_container_width=True)
                
                # 데이터 테이블
                st.dataframe(spec_data, use_container_width=True)
        
        # 전체 데이터 테이블
        st.subheader("📋 전체 세율 Risk 데이터")
        st.dataframe(risk_data, use_container_width=True)
        
        # 엑셀 다운로드
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            risk_data.to_excel(writer, sheet_name='세율 Risk', index=False)
            
            # 요약 시트 추가
            summary_df = pd.DataFrame({
                '규격1': risk_specs.index,
                '세번부호_수': risk_specs.values
            })
            summary_df.to_excel(writer, sheet_name='세율Risk_요약', index=False)
        
        st.download_button(
            label="📥 세율 Risk 엑셀 다운로드",
            data=excel_buffer.getvalue(),
            file_name=f"세율리스크_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        return risk_data
        
    except Exception as e:
        st.error(f"❌ 세율 Risk 분석 중 오류: {str(e)}")
        st.code(traceback.format_exc())
        return pd.DataFrame()

def create_price_risk_analysis(df):
    """단가 Risk 분석 (원본과 동일)"""
    try:
        st.subheader("💰 단가 Risk 분석")
        st.info("동일 조건에서 단가 편차가 큰 경우를 찾습니다.")
        
        # 필요한 컬럼 체크
        required_columns = ['수입신고번호', '규격1', '세번부호', '거래구분', '결제방법', '수리일자', 
                          '단가', '결제통화단위', '거래품명', 
                          '란번호', '행번호', '수량_1', '수량단위_1', '금액']
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.warning(f"⚠️ 누락된 컬럼: {missing_columns}")
            available_columns = [col for col in required_columns if col in df.columns]
            if '단가' not in available_columns:
                st.error("❌ 필수 컬럼(단가)이 없어 분석할 수 없습니다.")
                return pd.DataFrame()
        else:
            available_columns = required_columns
        
        # 단가를 숫자형으로 변환
        df_work = df.copy()
        df_work['단가'] = pd.to_numeric(df_work['단가'].fillna(0), errors='coerce').fillna(0)
        
        # 단가가 0보다 큰 데이터만 분석
        df_work = df_work[df_work['단가'] > 0]
        
        if len(df_work) == 0:
            st.warning("⚠️ 단가 데이터가 없어 분석할 수 없습니다.")
            return pd.DataFrame()
        
        # 그룹화 기준 선택
        group_columns = st.multiselect(
            "그룹화 기준을 선택하세요:",
            ['규격1', '세번부호', '거래구분', '결제방법', '결제통화단위'],
            default=['규격1', '세번부호', '결제통화단위']
        )
        
        if not group_columns:
            st.warning("⚠️ 최소 하나의 그룹화 기준을 선택하세요.")
            return pd.DataFrame()
        
        # 데이터 그룹화 및 분석 (원본과 동일 로직)
        st.info("📊 데이터 그룹화 중...")
        
        agg_dict = {
            '단가': ['mean', 'max', 'min', 'std', 'count'],
            '수입신고번호': ['min', 'max'],
            '결제통화단위': 'first',
            'B/L번호': 'first',
            '수리일자': 'first',
            '거래품명': 'first',
            '란번호': 'first',
            '행번호': 'first',
            '수량_1': 'first',
            '수량단위_1': 'first',
            '금액': 'sum'
        }
        
        grouped = df_work.groupby(group_columns).agg(agg_dict).reset_index()
        
        # 집계 후 실제 컬럼명에 맞춰 new_columns를 동적으로 생성
        grouped_columns = list(grouped.columns)
        new_columns = []
        for col in grouped_columns:
            # 다중 컬럼(튜플) 처리: ('단가', 'mean') 등
            if isinstance(col, tuple):
                if col[0] == '단가' and col[1] == 'mean':
                    new_columns.append('평균단가')
                elif col[0] == '단가' and col[1] == 'max':
                    new_columns.append('최고단가')
                elif col[0] == '단가' and col[1] == 'min':
                    new_columns.append('최저단가')
                elif col[0] == '단가' and col[1] == 'std':
                    new_columns.append('단가표준편차')
                elif col[0] == '단가' and col[1] == 'count':
                    new_columns.append('데이터수')
                elif col[0] == '수입신고번호' and col[1] == 'min':
                    new_columns.append('Min신고번호')
                elif col[0] == '수입신고번호' and col[1] == 'max':
                    new_columns.append('Max신고번호')
                elif col[0] == '결제통화단위' and col[1] == 'first':
                    new_columns.append('결제통화단위')
                elif col[0] == 'B/L번호' and col[1] == 'first':
                    new_columns.append('B/L번호')
                elif col[0] == '수리일자' and col[1] == 'first':
                    new_columns.append('수리일자')
                else:
                    new_columns.append(f'{col[0]}_{col[1]}')
            else:
                new_columns.append(col)
        grouped.columns = new_columns
        
        # 위험도 계산 (원본과 동일)
        grouped['단가편차율'] = np.where(
            grouped['평균단가'] > 0,
            (grouped['최고단가'] - grouped['최저단가']) / grouped['평균단가'],
            0
        )
        
        # 위험도 분류
        def classify_risk(row):
            if row['평균단가'] == 0:
                return '확인필요'
            elif row['단가편차율'] > 0.5:  # 50% 이상 편차
                return '매우높음'
            elif row['단가편차율'] > 0.3:  # 30% 이상 편차
                return '높음'
            elif row['단가편차율'] > 0.1:  # 10% 이상 편차
                return '보통'
            else:
                return '낮음'
        
        grouped['위험도'] = grouped.apply(classify_risk, axis=1)
        
        # 비고 생성
        grouped['비고'] = grouped.apply(lambda row: 
            f'평균단가 확인 필요' if row['평균단가'] == 0 
            else f'단가편차: {row["단가편차율"]*100:.1f}%', axis=1
        )
        
        # 결과 표시
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📊 총 그룹 수", f"{len(grouped):,}개")
        with col2:
            high_risk = len(grouped[grouped['위험도'].isin(['높음', '매우높음'])])
            st.metric("⚠️ 고위험 그룹", f"{high_risk:,}개")
        with col3:
            avg_deviation = grouped['단가편차율'].mean() * 100
            st.metric("📈 평균 편차율", f"{avg_deviation:.1f}%")
        with col4:
            zero_price = len(grouped[grouped['평균단가'] == 0])
            st.metric("🔍 확인필요", f"{zero_price:,}개")
        
        # 위험도별 분포 차트
        risk_counts = grouped['위험도'].value_counts()
        fig_risk = px.pie(
            values=risk_counts.values,
            names=risk_counts.index,
            title="단가 Risk 위험도 분포",
            color_discrete_map={
                '매우높음': '#FF0000',
                '높음': '#FF8C00', 
                '보통': '#FFD700',
                '낮음': '#32CD32',
                '확인필요': '#808080'
            }
        )
        st.plotly_chart(fig_risk, use_container_width=True)
        
        # 단가 편차율 히스토그램
        fig_hist = px.histogram(
            grouped[grouped['단가편차율'] <= 2],  # 200% 이하만 표시
            x='단가편차율',
            title="단가 편차율 분포",
            nbins=30
        )
        fig_hist.update_layout(xaxis_title="단가 편차율", yaxis_title="빈도")
        st.plotly_chart(fig_hist, use_container_width=True)
        
        # 위험도별 필터링
        st.subheader("🔍 위험도별 상세 분석")
        
        selected_risk = st.selectbox(
            "조회할 위험도를 선택하세요:",
            ['전체'] + list(risk_counts.index)
        )
        
        if selected_risk == '전체':
            display_data = grouped
        else:
            display_data = grouped[grouped['위험도'] == selected_risk]
        
        # 정렬 옵션
        sort_column = st.selectbox(
            "정렬 기준:",
            ['단가편차율', '평균단가', '데이터수'],
            index=0
        )
        
        display_data = display_data.sort_values(sort_column, ascending=False)
        
        st.write(f"**{selected_risk} 위험도 데이터: {len(display_data):,}건**")
        st.dataframe(display_data, use_container_width=True)
        
        # '수입신고번호', '수리일자' 컬럼을 가장 왼쪽에 오도록 재정렬 (엑셀 저장용)
        left_cols = []
        if '수입신고번호' in grouped.columns:
            left_cols.append('수입신고번호')
        if '수리일자' in grouped.columns:
            left_cols.append('수리일자')
        other_cols = [col for col in grouped.columns if col not in left_cols]
        grouped_for_excel = grouped[left_cols + other_cols]
        
        return grouped_for_excel
        
    except Exception as e:
        st.error(f"❌ 단가 Risk 분석 중 오류: {str(e)}")
        st.code(traceback.format_exc())
        return pd.DataFrame()

def create_summary_analysis(df):
    """종합 요약 분석 (원본과 동일)"""
    try:
        st.subheader("📈 종합 분석 요약")
        
        # 기본 통계
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            total_count = len(df)
            st.metric("📊 전체 데이터", f"{total_count:,}건")
        
        with col2:
            unique_declarations = df['수입신고번호'].nunique() if '수입신고번호' in df.columns else 0
            st.metric("📋 고유 신고번호", f"{unique_declarations:,}개")
        
        with col3:
            if '관세실행세율' in df.columns:
                avg_rate = df['관세실행세율'].mean()
                st.metric("📈 평균 관세율", f"{avg_rate:.2f}%")
        
        with col4:
            if '실제관세액' in df.columns:
                total_tax = df['실제관세액'].sum()
                st.metric("💰 총 관세액", f"{total_tax:,.0f}")
        
        # 세율구분별 분석
        if '세율구분' in df.columns:
            st.subheader("📊 세율구분별 분포")
            
            rate_analysis = df.groupby('세율구분').agg({
                '수입신고번호': 'nunique',
                '관세실행세율': ['mean', 'min', 'max'] if '관세실행세율' in df.columns else 'count',
                '실제관세액': 'sum' if '실제관세액' in df.columns else 'count'
            }).round(2)
            
            # 차트로 표시
            rate_counts = df['세율구분'].value_counts()
            
            fig_rate = make_subplots(
                rows=1, cols=2,
                subplot_titles=('세율구분별 건수', '세율구분별 비율'),
                specs=[[{"type": "bar"}, {"type": "pie"}]]
            )
            
            # 막대 차트
            fig_rate.add_trace(
                go.Bar(x=rate_counts.index, y=rate_counts.values, name="건수"),
                row=1, col=1
            )
            
            # 파이 차트
            fig_rate.add_trace(
                go.Pie(labels=rate_counts.index, values=rate_counts.values, name="비율"),
                row=1, col=2
            )
            
            fig_rate.update_layout(height=400, showlegend=False)
            st.plotly_chart(fig_rate, use_container_width=True)
            
            # 테이블로도 표시
            st.dataframe(rate_analysis, use_container_width=True)
        
        # 거래구분별 분석
        if '거래구분' in df.columns:
            st.subheader("🚢 거래구분별 분포")
            
            trade_counts = df['거래구분'].value_counts()
            fig_trade = px.bar(
                x=trade_counts.index,
                y=trade_counts.values,
                title="거래구분별 신고 건수"
            )
            st.plotly_chart(fig_trade, use_container_width=True)
        
        # 시계열 분석
        if '수리일자' in df.columns:
            st.subheader("📅 시계열 분석")
            
            try:
                df['수리일자_converted'] = pd.to_datetime(df['수리일자'], errors='coerce')
                if df['수리일자_converted'].notna().sum() > 0:
                    daily_counts = df.groupby(df['수리일자_converted'].dt.date).size()
                    
                    fig_time = px.line(
                        x=daily_counts.index,
                        y=daily_counts.values,
                        title="일별 수입신고 건수"
                    )
                    st.plotly_chart(fig_time, use_container_width=True)
                else:
                    st.info("수리일자 데이터를 파싱할 수 없습니다.")
            except Exception as e:
                st.warning(f"시계열 분석 중 오류: {e}")
        
        return True
        
    except Exception as e:
        st.error(f"❌ 종합 분석 중 오류: {str(e)}")
        return False

def create_comprehensive_excel_report(df, eight_percent_df, zero_risk_df, tariff_risk_df, price_risk_df):
    """종합 엑셀 리포트 생성 (보고서 스타일 + 차트 포함)"""
    try:
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # 1. Summary 시트 - DataFrame으로 먼저 작성
            summary_data = [
                ['수입신고 분석 보고서', ''],
                ['', ''],
                ['1. 기본 분석 정보', ''],
                ['전체 데이터 건수', len(df)],
                ['고유 신고번호 수', df['수입신고번호'].nunique() if '수입신고번호' in df.columns else 0],
                ['8% 환급검토 대상', len(eight_percent_df)],
                ['0% Risk 대상', len(zero_risk_df)],
                ['세율 Risk 대상', len(tariff_risk_df)],
                ['단가 Risk 그룹', len(price_risk_df)],
                ['', '']
            ]
            summary_df = pd.DataFrame(summary_data, columns=['구분', '건수'])
            summary_df.to_excel(writer, sheet_name='Summary', index=False, startrow=0)
            start_row = len(summary_data) + 2
            
            # 세율구분 분석
            rate_df = None
            if '세율구분' in df.columns:
                rate_counts = df['세율구분'].value_counts()
                rate_df = pd.DataFrame({
                    '세율구분': rate_counts.index,
                    '건수': rate_counts.values,
                    '비율(%)': (rate_counts.values / len(df) * 100).round(1)
                })
                rate_df.to_excel(writer, sheet_name='Summary', index=False, startrow=start_row)
                rate_start = start_row
                start_row += len(rate_df) + 3
            
            # 거래구분 분석
            trade_df = None
            if '거래구분' in df.columns:
                trade_counts = df['거래구분'].value_counts()
                trade_df = pd.DataFrame({
                    '거래구분': trade_counts.index,
                    '건수': trade_counts.values,
                    '비율(%)': (trade_counts.values / len(df) * 100).round(1)
                })
                trade_df.to_excel(writer, sheet_name='Summary', index=False, startrow=start_row)
                trade_start = start_row
                start_row += len(trade_df) + 3
            
            # 시계열 분석
            time_df = None
            if '수리일자' in df.columns:
                try:
                    df['수리일자_converted'] = pd.to_datetime(df['수리일자'], errors='coerce')
                    if df['수리일자_converted'].notna().sum() > 0:
                        daily_counts = df.groupby(df['수리일자_converted'].dt.date).size()
                        time_df = pd.DataFrame({
                            '날짜': daily_counts.index,
                            '건수': daily_counts.values
                        })
                        time_df.to_excel(writer, sheet_name='Summary', index=False, startrow=start_row)
                        time_start = start_row
                except:
                    pass
            
            # openpyxl로 스타일 및 차트 적용
            workbook = writer.book
            summary_sheet = workbook['Summary']
            # 스타일
            title_font = Font(name='맑은 고딕', size=14, bold=True)
            header_font = Font(name='맑은 고딕', size=11, bold=True)
            normal_font = Font(name='맑은 고딕', size=10)
            center = Alignment(horizontal='center', vertical='center')
            bold = Font(bold=True)
            fill = PatternFill(start_color='EAF1FB', end_color='EAF1FB', fill_type='solid')
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            # 제목
            summary_sheet['A1'].font = title_font
            summary_sheet['A1'].alignment = center
            # 섹션 헤더
            summary_sheet['A3'].font = header_font
            # 표 헤더들
            for row in summary_sheet.iter_rows(min_row=4, max_row=4, min_col=1, max_col=2):
                for cell in row:
                    cell.font = header_font
                    cell.alignment = center
                    cell.fill = fill
                    cell.border = border
            # 데이터
            for row in summary_sheet.iter_rows(min_row=5, min_col=1, max_col=2):
                for cell in row:
                    if cell.font != header_font:
                        cell.font = normal_font
                        cell.border = border
            # 열 너비 자동 조정
            for column_cells in summary_sheet.columns:
                max_length = 0
                column = column_cells[0].column_letter
                for cell in column_cells:
                    try:
                        if cell.value and len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                summary_sheet.column_dimensions[column].width = adjusted_width
            # 차트 추가
            from openpyxl.utils import get_column_letter
            chart_row_offset = 2
            # 세율구분 파이차트
            if rate_df is not None:
                pie = PieChart()
                pie.title = "세율구분별 비율"
                data_ref = Reference(summary_sheet, min_col=2, min_row=rate_start+2, max_row=rate_start+1+len(rate_df))
                labels_ref = Reference(summary_sheet, min_col=1, min_row=rate_start+2, max_row=rate_start+1+len(rate_df))
                pie.add_data(data_ref, titles_from_data=False)
                pie.set_categories(labels_ref)
                pie.height = 7
                pie.width = 7
                summary_sheet.add_chart(pie, f"E{rate_start+2}")
            # 거래구분 파이차트
            if trade_df is not None:
                pie2 = PieChart()
                pie2.title = "거래구분별 비율"
                data_ref = Reference(summary_sheet, min_col=2, min_row=trade_start+2, max_row=trade_start+1+len(trade_df))
                labels_ref = Reference(summary_sheet, min_col=1, min_row=trade_start+2, max_row=trade_start+1+len(trade_df))
                pie2.add_data(data_ref, titles_from_data=False)
                pie2.set_categories(labels_ref)
                pie2.height = 7
                pie2.width = 7
                summary_sheet.add_chart(pie2, f"E{trade_start+2}")
            # 시계열 꺾은선그래프
            if time_df is not None:
                line = LineChart()
                line.title = "일별 수입신고 건수"
                data_ref = Reference(summary_sheet, min_col=2, min_row=time_start+2, max_row=time_start+1+len(time_df))
                cats_ref = Reference(summary_sheet, min_col=1, min_row=time_start+2, max_row=time_start+1+len(time_df))
                line.add_data(data_ref, titles_from_data=False)
                line.set_categories(cats_ref)
                line.height = 7
                line.width = 14
                summary_sheet.add_chart(line, f"E{time_start+2}")
            # 이하 기존 시트 저장 로직 동일
            if len(eight_percent_df) > 0:
                eight_percent_df.to_excel(writer, sheet_name='8% 환급 검토', index=False)
            if len(zero_risk_df) > 0:
                zero_risk_df.to_excel(writer, sheet_name='0% Risk', index=False)
            if len(tariff_risk_df) > 0:
                tariff_risk_df.to_excel(writer, sheet_name='세율 Risk', index=False)
            if len(price_risk_df) > 0:
                cols = price_risk_df.columns.tolist()
                if '수입신고번호' in cols:
                    cols.remove('수입신고번호')
                    cols.insert(0, '수입신고번호')
                    price_risk_df = price_risk_df[cols]
                price_risk_df.to_excel(writer, sheet_name='단가 Risk', index=False)
            df.to_excel(writer, sheet_name='전체데이터', index=False)
            if '세율구분' in df.columns:
                rate_stats = df.groupby('세율구분').agg({
                    '수입신고번호': 'nunique',
                    '관세실행세율': ['mean', 'min', 'max'] if '관세실행세율' in df.columns else 'count'
                }).round(2)
                rate_stats.to_excel(writer, sheet_name='세율구분별_통계')
        return output.getvalue()
    except Exception as e:
        st.error(f"❌ 종합 엑셀 리포트 생성 오류: {e}")
        return None

# 메인 실행 부분
if uploaded_file is not None:
    # 파일 읽기
    df = read_and_process_excel(uploaded_file)
    
    if df is not None:
        st.success("✅ 파일 로딩 완료!")
        
        # 데이터 기본 정보
        with st.expander("📊 데이터 기본 정보", expanded=True):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("총 행 수", f"{len(df):,}")
            with col2:
                st.metric("총 열 수", len(df.columns))
            with col3:
                memory_mb = df.memory_usage(deep=True).sum() / 1024**2
                st.metric("메모리 사용량", f"{memory_mb:.1f} MB")
            
            # 데이터 미리보기
            st.dataframe(df.head(10), use_container_width=True)
        
        # 분석 실행
        if analysis_type == "전체 분석":
            st.header("🎯 전체 분석 실행")
            
            with st.spinner("분석 중..."):
                # 모든 분석 실행
                eight_percent_df = create_eight_percent_refund_analysis(df)
                zero_risk_df = create_zero_percent_risk_analysis(df)
                tariff_risk_df = create_tariff_risk_analysis(df)
                price_risk_df = create_price_risk_analysis(df)
                create_summary_analysis(df)
                
                # 종합 엑셀 리포트 다운로드
                st.subheader("📥 종합 리포트 다운로드")
                excel_data = create_comprehensive_excel_report(
                    df, eight_percent_df, zero_risk_df, tariff_risk_df, price_risk_df
                )
                
                if excel_data:
                    st.download_button(
                        label="📊 종합 분석 엑셀 다운로드",
                        data=excel_data,
                        file_name=f"수입신고_종합분석_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        
        elif analysis_type == "8% 환급 검토":
            create_eight_percent_refund_analysis(df)
            
        elif analysis_type == "0% Risk 분석":
            create_zero_percent_risk_analysis(df)
            
        elif analysis_type == "세율 Risk 분석":
            create_tariff_risk_analysis(df)
            
        elif analysis_type == "단가 Risk 분석":
            create_price_risk_analysis(df)

# 사용 안내
with st.expander("📋 사용 방법 및 기능 설명"):
    st.markdown("""
    ### 🎯 주요 기능 (app-new202505.py 완전 구현)
    
    1. **8% 환급 검토**: 세율구분='A' AND 관세실행세율≥8% 조건의 데이터 분석
    2. **0% Risk 분석**: 관세실행세율<8% AND 세율구분≠F*** 조건의 리스크 데이터
    3. **세율 Risk 분석**: 동일 규격1에 대해 서로 다른 세번부호가 적용된 경우
    4. **단가 Risk 분석**: 동일 조건에서 단가 편차가 큰 경우 분석
    5. **종합 요약**: 모든 분석 결과의 통계 및 시각화
    
    ### 📊 분석 조건 (원본과 동일)
    - **컬럼 매핑**: 70번째 컬럼 → 세율구분, 71번째 컬럼 → 관세실행세율
    - **8% 환급**: 세율구분 = 'A' AND 관세실행세율 >= 8%
    - **0% 리스크**: 관세실행세율 < 8% AND 세율구분이 F로 시작하는 4자리가 아님
    - **세율 리스크**: 규격1당 세번부호 수 > 1
    - **단가 리스크**: 그룹별 단가 편차율 > 30%
    
    ### 💡 사용 팁
    - 대용량 파일은 처리 시간이 다소 소요될 수 있습니다
    - 분석 유형을 개별 선택하여 빠른 확인이 가능합니다
    - 모든 결과는 엑셀 형태로 다운로드 가능합니다
    - 차트와 그래프로 직관적인 분석 결과를 제공합니다
    """)

# 푸터
st.markdown("---")
st.markdown("🚀 **완전판**: 원본 `app-new202505.py`의 모든 분석 기능을 웹으로 구현한 완전판입니다.") 
