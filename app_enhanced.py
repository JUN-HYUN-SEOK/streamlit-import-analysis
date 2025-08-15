import streamlit as st
import pandas as pd
import numpy as np
from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml
import datetime
import io
import traceback
import openpyxl
from docx.enum.text import WD_COLOR_INDEX
import xlsxwriter

# 페이지 설정
st.set_page_config(
    page_title="[관세법인우신] 수입신고 Risk Management System v2",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 제목과 설명
st.title("🚢 수입신고 분석 도구 v2")
st.markdown("""
이 도구는 수입신고 데이터를 분석하여 다음과 같은 리포트를 생성합니다:
- **8% 환급 검토**: 환급 대상 분석
- **0% Risk 분석**: 저위험 항목 검토
- **세율 Risk 분석**: 세율 불일치 위험
- **단가 Risk 분석**: 단가 변동성 분석
- **Summary**: 전체 분석 요약
""")

# 사이드바 메뉴
with st.sidebar:
    st.header("📋 분석 옵션")
    analysis_options = st.multiselect(
        "포함할 분석을 선택하세요:",
        ["8% 환급 검토", "0% Risk 분석", "세율 Risk 분석", "단가 Risk 분석", "Summary"],
        default=["8% 환급 검토", "0% Risk 분석", "Summary"]
    )
    
    st.markdown("---")
    st.subheader("📁 파일 업로드")

# 파일 업로드
uploaded_file = st.file_uploader(
    "분석할 엑셀 파일을 업로드하세요", 
    type=["xlsx", "xls"],
    help="수입신고 데이터가 포함된 엑셀 파일을 업로드하세요"
)

def read_excel_file(file):
    """엑셀 파일 읽기"""
    try:
        # Streamlit 업로드 파일 처리
        df = pd.read_excel(file)
        
        # 컬럼 이름 정리
        df.columns = [str(col).strip() for col in df.columns]
        
        st.info(f"데이터 로드 완료: {df.shape[0]}행, {df.shape[1]}열")
        
        # 컬럼 매핑 - 기존 컬럼명이 있는지 먼저 확인
        if '세율구분' not in df.columns and '관세실행세율' not in df.columns:
            # 컬럼 인덱스 기반 매핑 시도
            if len(df.columns) > 71:
                try:
                    df = df.rename(columns={
                        df.columns[70]: '세율구분',
                        df.columns[71]: '관세실행세율'
                    })
                    st.info("컬럼 매핑 완료: 인덱스 기반")
                except Exception as e:
                    st.warning(f"컬럼 매핑 실패: {str(e)}")
            else:
                st.warning(f"컬럼 수 부족. 현재: {len(df.columns)}개")
        
        # 필수 컬럼이 없으면 기본값으로 생성
        if '관세실행세율' not in df.columns:
            st.warning("관세실행세율 컬럼이 없어 기본값(0)으로 생성합니다.")
            df['관세실행세율'] = 0
        else:
            # 숫자형으로 변환
            df['관세실행세율'] = pd.to_numeric(df['관세실행세율'], errors='coerce').fillna(0)
        
        if '세율구분' not in df.columns:
            st.warning("세율구분 컬럼이 없어 기본값('A')으로 생성합니다.")
            df['세율구분'] = 'A'
        
        # 기타 필요한 컬럼들도 기본값으로 생성
        required_columns = ['수입신고번호', '규격1', '세번부호', '거래구분']
        for col in required_columns:
            if col not in df.columns:
                df[col] = f'기본값_{col}'
                st.info(f"'{col}' 컬럼이 없어 기본값으로 생성했습니다.")
        
        return df
        
    except Exception as e:
        st.error(f"엑셀 파일 읽기 실패: {str(e)}")
        st.error(f"오류 상세: {traceback.format_exc()}")
        return None

def process_data(df):
    """데이터 전처리"""
    try:
        # 컬럼 이름의 공백 제거
        df.columns = df.columns.str.strip()
        
        # 필요한 컬럼이 있는지 확인
        required_columns = ['관세실행세율', '세율구분']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"누락된 컬럼: {missing_columns}")
            return None

        # 0% Risk 조건에 맞는 데이터 필터링
        df_zero_risk = df[
            (df['관세실행세율'] < 8) & 
            (~df['세율구분'].astype(str).str.match(r'^F.{3}$'))
        ]
        
        # '세율구분'이 4자리인 행 제외
        df_filtered = df_zero_risk[df_zero_risk['세율구분'].apply(lambda x: len(str(x)) != 4)]

        return df_filtered
        
    except Exception as e:
        st.error(f"데이터 전처리 중 오류 발생: {e}")
        return None

def create_eight_percent_refund_analysis(df):
    """8% 환급 검토 분석"""
    try:
        # 필터링 조건 적용
        df_filtered = df[
            (df['세율구분'] == 'A') & 
            (df['관세실행세율'] >= 8)
        ]
        
        if len(df_filtered) == 0:
            return None, "8% 환급 검토 대상이 없습니다."
        
        # 행별관세 계산 (필요한 컬럼이 있는 경우)
        if all(col in df_filtered.columns for col in ['실제관세액', '금액', '란결제금액']):
            df_filtered = df_filtered.copy()
            df_filtered['실제관세액'] = pd.to_numeric(df_filtered['실제관세액'].fillna(0), errors='coerce').fillna(0)
            df_filtered['금액'] = pd.to_numeric(df_filtered['금액'].fillna(0), errors='coerce').fillna(0)
            df_filtered['란결제금액'] = pd.to_numeric(df_filtered['란결제금액'].fillna(0), errors='coerce').fillna(0)
            
            df_filtered['행별관세'] = np.where(
                df_filtered['란결제금액'] != 0,
                (df_filtered['실제관세액'] * df_filtered['금액']) / df_filtered['란결제금액'],
                0
            )
        
        return df_filtered, f"총 {len(df_filtered)}건의 8% 환급 검토 대상이 발견되었습니다."
        
    except Exception as e:
        return None, f"8% 환급 검토 분석 중 오류 발생: {str(e)}"

def create_zero_percent_risk_analysis(df):
    """0% Risk 분석"""
    try:
        # 0% Risk 조건에 맞는 데이터 필터링
        df_zero_risk = df[
            (df['관세실행세율'] < 8) & 
            (~df['세율구분'].astype(str).str.match(r'^F.{3}$'))
        ]
        
        if len(df_zero_risk) == 0:
            return None, "0% Risk 대상이 없습니다."
        
        return df_zero_risk, f"총 {len(df_zero_risk)}건의 0% Risk가 발견되었습니다."
        
    except Exception as e:
        return None, f"0% Risk 분석 중 오류 발생: {str(e)}"

def create_tariff_risk_analysis(df):
    """세율 Risk 분석"""
    try:
        if '규격1' not in df.columns or '세번부호' not in df.columns:
            return None, "세율 Risk 분석에 필요한 컬럼이 누락되었습니다."
        
        # 규격1별로 세번부호의 고유값 개수를 계산
        risk_specs = df.groupby('규격1')['세번부호'].nunique()
        risk_specs = risk_specs[risk_specs > 1]
        
        if len(risk_specs) == 0:
            return None, "세율 Risk가 발견되지 않았습니다."
        
        # 위험한 규격1에 해당하는 데이터 추출
        risk_data = df[df['규격1'].isin(risk_specs.index)].copy()
        risk_data = risk_data.sort_values(['규격1', '세번부호']).fillna('')
        
        return risk_data, f"총 {len(risk_data)}건의 세율 Risk가 발견되었습니다."
        
    except Exception as e:
        return None, f"세율 Risk 분석 중 오류 발생: {str(e)}"

def create_summary_analysis(df):
    """Summary 분석"""
    try:
        summary_data = {}
        
        # 전체 신고 건수
        if '수입신고번호' in df.columns:
            total_declarations = df['수입신고번호'].nunique()
        else:
            total_declarations = len(df)
        summary_data['전체_신고건수'] = total_declarations
        
        # 거래구분별 분석
        if '거래구분' in df.columns:
            trade_analysis = df['거래구분'].value_counts()
            summary_data['거래구분별'] = trade_analysis
        
        # 세율구분별 분석
        if '세율구분' in df.columns:
            tariff_analysis = df['세율구분'].value_counts()
            summary_data['세율구분별'] = tariff_analysis
        
        # Risk 분석
        if all(col in df.columns for col in ['관세실행세율', '세율구분']):
            zero_risk_count = len(df[
                (df['관세실행세율'] < 8) & 
                (~df['세율구분'].astype(str).str.match(r'^F.{3}$'))
            ])
            
            eight_percent_count = len(df[
                (df['세율구분'] == 'A') & 
                (df['관세실행세율'] >= 8)
            ])
            
            summary_data['Risk분석'] = {
                '0% Risk': zero_risk_count,
                '8% 환급검토': eight_percent_count
            }
        
        return summary_data, "Summary 분석이 완료되었습니다."
        
    except Exception as e:
        return None, f"Summary 분석 중 오류 발생: {str(e)}"

def create_excel_report(df, analysis_results):
    """엑셀 보고서 생성"""
    try:
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # 헤더 포맷 정의
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#D9E1F2',
                'border': 1,
                'align': 'center'
            })
            
            # 분석 결과 시트들 생성
            sheet_count = 0
            
            # 8% 환급 검토 시트
            if ("8% 환급 검토" in analysis_results and 
                analysis_results["8% 환급 검토"][0] is not None and
                not analysis_results["8% 환급 검토"][0].empty):
                df_eight = analysis_results["8% 환급 검토"][0]
                df_eight.to_excel(writer, sheet_name='8% 환급 검토', index=False)
                sheet_count += 1
            
            # 0% Risk 시트
            if ("0% Risk 분석" in analysis_results and 
                analysis_results["0% Risk 분석"][0] is not None and
                not analysis_results["0% Risk 분석"][0].empty):
                df_zero = analysis_results["0% Risk 분석"][0]
                df_zero.to_excel(writer, sheet_name='0% Risk', index=False)
                sheet_count += 1
            
            # 세율 Risk 시트
            if ("세율 Risk 분석" in analysis_results and 
                analysis_results["세율 Risk 분석"][0] is not None and
                not analysis_results["세율 Risk 분석"][0].empty):
                df_tariff = analysis_results["세율 Risk 분석"][0]
                df_tariff.to_excel(writer, sheet_name='세율 Risk', index=False)
                sheet_count += 1
            
            # 원본 데이터 시트 (항상 포함)
            df_sample = df.head(1000)  # 처음 1000행만
            df_sample.to_excel(writer, sheet_name='원본데이터', index=False)
            sheet_count += 1
            
            st.info(f"총 {sheet_count}개 시트가 생성되었습니다.")
        
        output.seek(0)
        return output.getvalue()
        
    except Exception as e:
        st.error(f"엑셀 보고서 생성 중 오류 발생: {str(e)}")
        st.error(f"상세 오류: {traceback.format_exc()}")
        return None

def create_word_report(analysis_results):
    """워드 보고서 생성"""
    try:
        doc = Document()
        
        # 제목 추가
        doc.add_heading('수입신고 분석 보고서', 0)
        doc.add_paragraph(datetime.datetime.now().strftime("%Y년 %m월 %d일"))
        
        # 분석 결과 요약
        doc.add_heading('분석 결과 요약', level=1)
        summary_para = doc.add_paragraph()
        
        # 각 분석 결과 추가
        for analysis_name, (data, message) in analysis_results.items():
            doc.add_heading(analysis_name, level=2)
            doc.add_paragraph(message)
            
            if analysis_name == "Summary":
                # Summary는 특별 처리
                if isinstance(data, dict):
                    for key, value in data.items():
                        doc.add_paragraph(f"{key}: {value}", style='List Bullet')
            elif data is not None and hasattr(data, 'head') and len(data) > 0:
                # 데이터프레임인 경우 상위 5개만 테이블로 표시
                table_data = data.head(5)
                if len(table_data) > 0 and len(table_data.columns) <= 10:  # 컬럼이 너무 많으면 제외
                    try:
                        table = doc.add_table(rows=len(table_data)+1, cols=len(table_data.columns))
                        table.style = 'Table Grid'
                        
                        # 헤더 추가
                        for j, column in enumerate(table_data.columns):
                            table.cell(0, j).text = str(column)[:20]  # 컬럼명 길이 제한
                        
                        # 데이터 추가
                        for i, row in enumerate(table_data.values):
                            for j, value in enumerate(row):
                                cell_value = str(value)[:30] if value is not None else ""  # 셀 값 길이 제한
                                table.cell(i+1, j).text = cell_value
                        
                        doc.add_paragraph("※ 상위 5건만 표시됨")
                    except Exception as table_error:
                        doc.add_paragraph(f"테이블 생성 실패: {str(table_error)}")
                        doc.add_paragraph(f"데이터 건수: {len(data)}건")
                else:
                    doc.add_paragraph(f"데이터 건수: {len(data)}건 (테이블 표시 생략)")
        
        # 메모리에 저장
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        return output.getvalue()
        
    except Exception as e:
        st.error(f"워드 보고서 생성 중 오류 발생: {str(e)}")
        st.error(f"상세 오류: {traceback.format_exc()}")
        return None

# 메인 로직
if uploaded_file is not None:
    st.info(f"업로드된 파일: {uploaded_file.name}")
    st.info(f"파일 크기: {uploaded_file.size} bytes")
    
    # 데이터 로드
    with st.spinner("파일을 읽는 중..."):
        try:
            # 파일 포인터를 처음으로 이동
            uploaded_file.seek(0)
            df = read_excel_file(uploaded_file)
        except Exception as e:
            st.error(f"파일 읽기 중 오류 발생: {str(e)}")
            st.error(f"오류 상세: {traceback.format_exc()}")
            df = None
    
    if df is not None:
        st.success(f"✅ 파일 로드 완료: {len(df):,}행, {len(df.columns)}열")
        
        # 데이터 미리보기
        with st.expander("📊 데이터 미리보기"):
            st.dataframe(df.head(10))
            st.info(f"컬럼 목록 (처음 20개): {', '.join(df.columns.tolist()[:20])}")
            if len(df.columns) > 20:
                st.info(f"... 총 {len(df.columns)}개 컬럼")
        
        # 분석 실행
        if st.button("🔍 분석 시작", type="primary"):
            analysis_results = {}
            
            with st.spinner("분석을 수행하는 중..."):
                # 선택된 분석 수행
                if "8% 환급 검토" in analysis_options:
                    st.info("8% 환급 검토 분석 중...")
                    result = create_eight_percent_refund_analysis(df)
                    analysis_results["8% 환급 검토"] = result
                
                if "0% Risk 분석" in analysis_options:
                    st.info("0% Risk 분석 중...")
                    result = create_zero_percent_risk_analysis(df)
                    analysis_results["0% Risk 분석"] = result
                
                if "세율 Risk 분석" in analysis_options:
                    st.info("세율 Risk 분석 중...")
                    result = create_tariff_risk_analysis(df)
                    analysis_results["세율 Risk 분석"] = result
                
                if "Summary" in analysis_options:
                    st.info("Summary 분석 중...")
                    result = create_summary_analysis(df)
                    analysis_results["Summary"] = result
            
            # 결과 표시
            st.header("📈 분석 결과")
            
            for analysis_name, (data, message) in analysis_results.items():
                with st.expander(f"{analysis_name} 결과"):
                    st.info(message)
                    
                    if data is not None:
                        if analysis_name == "Summary":
                            # Summary 특별 처리
                            if isinstance(data, dict):
                                for key, value in data.items():
                                    st.subheader(key)
                                    if isinstance(value, pd.Series):
                                        st.bar_chart(value)
                                    elif isinstance(value, dict):
                                        st.json(value)
                                    else:
                                        st.write(value)
                        else:
                            # 데이터프레임 표시
                            st.dataframe(data)
                            
                            # 간단한 통계
                            if len(data) > 0:
                                st.metric("총 건수", len(data))
            
            # 보고서 다운로드
            st.header("📥 보고서 다운로드")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("📊 엑셀 보고서")
                with st.spinner("엑셀 보고서 생성 중..."):
                    excel_data = create_excel_report(df, analysis_results)
                
                if excel_data:
                    st.download_button(
                        label="📊 엑셀 파일 다운로드",
                        data=excel_data,
                        file_name=f"수입신고분석_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="excel_download"
                    )
                    st.success("엑셀 보고서가 준비되었습니다!")
                else:
                    st.error("엑셀 보고서 생성에 실패했습니다.")
            
            with col2:
                st.subheader("📝 워드 보고서")
                with st.spinner("워드 보고서 생성 중..."):
                    word_data = create_word_report(analysis_results)
                
                if word_data:
                    st.download_button(
                        label="📝 워드 파일 다운로드",
                        data=word_data,
                        file_name=f"수입신고분석_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="word_download"
                    )
                    st.success("워드 보고서가 준비되었습니다!")
                else:
                    st.error("워드 보고서 생성에 실패했습니다.")

else:
    st.info("👆 분석할 엑셀 파일을 업로드해주세요.")
    
    # 사용법 안내
    with st.expander("📖 사용법 안내"):
        st.markdown("""
        ### 📋 지원하는 분석 종류
        
        1. **8% 환급 검토**: 세율구분이 'A'이고 관세실행세율이 8% 이상인 항목
        2. **0% Risk 분석**: 관세실행세율이 8% 미만이고 특별세율이 아닌 항목
        3. **세율 Risk 분석**: 동일 규격에 다른 세번부호가 적용된 항목
        4. **단가 Risk 분석**: 단가 변동성이 큰 항목
        5. **Summary**: 전체 분석 요약
        
        ### 📂 파일 형식
        - 지원 형식: `.xlsx`, `.xls`
        - 필요 컬럼: `세율구분`, `관세실행세율`, `규격1`, `세번부호` 등
        
        ### 📥 결과물
        - 엑셀 보고서: 각 분석별 시트가 포함된 통합 파일
        - 워드 보고서: 분석 결과 요약 문서
        """) 
