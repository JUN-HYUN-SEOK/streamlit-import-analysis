import pandas as pd
import numpy as np
from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml
import datetime
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import traceback
import openpyxl
from docx.enum.text import WD_COLOR_INDEX
import win32com.client

def init_gui():
    """GUI 초기화"""
    root = tk.Tk()
    root.withdraw()
    return root

def select_input_file():
    """입력 파일 선택"""
    try:
        file_path = filedialog.askopenfilename(
            title="분석할 엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not file_path:
            print("파일 선택이 취소되었습니다.")
            return None
        if not os.path.exists(file_path):
            messagebox.showerror("오류", "선택한 파일이 존재하지 않습니다.")
            return None
        return file_path
    except Exception as e:
        messagebox.showerror("오류", f"파일 선택 중 오류 발생: {str(e)}")
        return None

def read_excel_file(file_path):
    """엑셀 파일 읽기"""
    try:
        print(f"\n엑셀 파일 읽기 시작: {file_path}")
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()  # 컬럼 이름의 공백 제거
        print(f"- 데이터 크기: {df.shape}")
        print("- 컬럼 목록:", df.columns.tolist())  # 컬럼 목록 출력
        
        # 1. 컬럼 존재 여부 확인
        if len(df.columns) > 71:
            # 안전한 컬럼 매핑
            df.rename(columns={
                df.columns[70]: '세율구분',
                df.columns[71]: '관세실행세율'
            }, inplace=True)
        else:
            # 기본값 설정
            df.rename(columns={
                df.columns[70]: '세율구분',
                df.columns[71]: '관세실행세율'
            }, inplace=True)
        
        # 2. Series 타입 확인
        rate_column = df['관세실행세율']
        if isinstance(rate_column, pd.Series):
            # 관세실행세율 컬럼을 숫자형으로 변환
            df['관세실행세율'] = pd.to_numeric(
                df['관세실행세율'].fillna(0), errors='coerce'
            ).fillna(0)
        else:
            # 기본값으로 대체
            df['관세실행세율'] = 0
        
        return df
    except Exception as e:
        messagebox.showerror("오류", f"엑셀 파일 읽기 실패: {str(e)}")
        return None

def process_data(df):
    """데이터 전처리"""
    try:
        print("\n데이터 전처리 시작...")
        
        # 컬럼 이름의 공백 제거
        df.columns = df.columns.str.strip()
        
        # 필요한 컬럼이 있는지 확인
        required_columns = ['관세실행세율', '세율구분']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print("\n누락된 컬럼:")
            for col in missing_columns:
                print(f"- {col}")
            return None

        # 0% Risk 조건에 맞는 데이터 필터링
        print("\n데이터 필터링 디버깅:")
        print(f"- 전체 데이터 수: {len(df)}")
        print(f"- 관세실행세율이 0%인 데이터 수: {len(df[df['관세실행세율'] == 0])}")
        print(f"- 세율구분이 'CIT'인 데이터 수: {len(df[df['세율구분'] == 'CIT'])}")
        print(f"- 세율구분이 'C'인 데이터 수: {len(df[df['세율구분'] == 'C'])}")
        print(f"- 두 조건 모두 만족하는 데이터 수: "
              f"{len(df[(df['관세실행세율'] == 0) & (df['세율구분'] == 'CIT')])}")

        df_zero_risk = df[
            (df['관세실행세율'] < 8) & 
            (~df['세율구분'].astype(str).str.match(r'^F.{3}$'))  # F로 시작하는 4자리 코드 제외
        ]
        print(f"- 최종 필터링된 데이터 수: {len(df_zero_risk)}")

        # '세율구분'이 4자리인 행 제외
        df_filtered = df_zero_risk[df_zero_risk['세율구분'].apply(lambda x: len(str(x)) != 4)]

        print(f"필터링된 데이터프레임: {df_filtered.shape}")
        return df_filtered
        
    except Exception as e:
        print(f"데이터 전처리 중 오류 발생: {e}")
        traceback.print_exc()
        return None

def save_files_dialog():
    """파일 저장 경로 선택"""
    try:
        root = tk.Tk()
        root.withdraw()
        
        # 저장 경로의 기본 디렉토리 설정
        default_dir = os.path.expanduser("~\\Documents")  # 사용자의 문서 폴더를 기본값으로
        
        # 엑셀 파일 저장 경로 선택
        excel_path = filedialog.asksaveasfilename(
            title="엑셀 파일 저장 위치 선택",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="수입신고분석.xlsx",
            initialdir=default_dir
        )
        
        if not excel_path:
            print("엑셀 파일 저장이 취소되었습니다.")
            return None, None
        
        # 엑셀 파일 경로가 유효한지 확인
        try:
            excel_dir = os.path.dirname(excel_path)
            if not os.path.exists(excel_dir):
                os.makedirs(excel_dir)
        except Exception as e:
            print(f"엑셀 파일 경로 생성 중 오류: {str(e)}")
            return None, None
        
        # 워드 파일 저장 경로 선택
        word_path = filedialog.asksaveasfilename(
            title="워드 파일 저장 위치 선택",
            defaultextension=".docx",
            filetypes=[("Word files", "*.docx")],
            initialfile="수입신고분석.docx",
            initialdir=os.path.dirname(excel_path)  # 엑셀 파일과 같은 디렉토리를 기본값으로
        )
        
        if not word_path:
            print("워드 파일 저장이 취소되었습니다.")
            return None, None
        
        # 워드 파일 경로가 유효한지 확인
        try:
            word_dir = os.path.dirname(word_path)
            if not os.path.exists(word_dir):
                os.makedirs(word_dir)
        except Exception as e:
            print(f"워드 파일 경로 생성 중 오류: {str(e)}")
            return None, None
        
        # 파일이 이미 존재하는지 확인
        if os.path.exists(excel_path):
            try:
                os.remove(excel_path)
            except Exception as e:
                print(f"기존 엑셀 파일 삭제 중 오류: {str(e)}")
                return None, None
        
        if os.path.exists(word_path):
            try:
                os.remove(word_path)
            except Exception as e:
                print(f"기존 워드 파일 삭제 중 오류: {str(e)}")
                return None, None
        
        print(f"선택된 엑셀 경로: {excel_path}")
        print(f"선택된 워드 경로: {word_path}")
        
        return excel_path, word_path
        
    except Exception as e:
        print(f"파일 저장 경로 선택 중 오류: {str(e)}")
        return None, None

def create_eight_percent_refund_sheet(df, writer, document):
    """8% 환급 검토 시트 생성"""
    try:
        print("\n- 8% 환급 검토 시트 생성 중...")
        
        # 필요한 컬럼만 선택 - 거래품명 위치 조정하여 란번호 앞에 배치
        selected_columns = [
            '수입신고번호',
            '수리일자',    # 수입신고번호 옆에 수리일자 추가
            'B/L번호',  # B/L번호 추가
            '세번부호', 
            '세율구분',
            '세율설명',
            '관세실행세율',
            '적출국코드',  # 1. 적출국코드 추가
            '원산지코드',  # 2. 원산지코드 추가
            'FTA사후환급 검토',  # 5. FTA사후환급 검토 컬럼 추가
            '규격1',
            '규격2',
            '규격3',
            '성분1',
            '성분2',
            '성분3',
            '실제관세액',
            '결제방법',
            '결제통화단위',
            '무역거래처상호',      # 3. 무역거래처상호 추가
            '무역거래처국가코드',  # 4. 무역거래처국가코드 추가
            '거래품명',    # 란번호 앞에 거래품명 추가
            '란번호',     # 추가된 컬럼
            '행번호',     # 추가된 컬럼
            '수량_1',     # 추가된 컬럼
            '수량단위_1', # 추가된 컬럼
            '단가',       # 추가된 컬럼
            '금액',       # 추가된 컬럼
            '란결제금액',  # 행별관세 계산용 추가
            '행별관세'     # 행별관세 컬럼 추가
        ]
        
        # 수리일자 컬럼 매핑 (다양한 이름 지원)
        if '수리일자' not in df.columns:
            possible_date_columns = ['수리일자_converted', '수리일자_변환', '수리일자_날짜', '수리일자_변환됨']
            for col in possible_date_columns:
                if col in df.columns:
                    df = df.rename(columns={col: '수리일자'})
                    print(f"수리일자 컬럼을 '{col}'에서 매핑했습니다.")
                    break
        
        # 행별관세 계산에 필요한 컬럼들을 제외하고 존재하는 컬럼만 선택
        base_columns = [col for col in selected_columns 
                       if col not in ['행별관세'] and col in df.columns]
        
        if len(base_columns) < len(selected_columns) - 1:
            missing_columns = [col for col in selected_columns 
                             if col not in df.columns and col != '행별관세']
            print(f"경고: 다음 컬럼들이 누락되었습니다: {missing_columns}")
        
        # 원본 데이터를 복사하여 사용
        df_work = df[base_columns].copy()
        
        # 데이터 전처리
        df_work['세율구분'] = df_work['세율구분'].astype(str).str.strip()
        df_work['관세실행세율'] = pd.to_numeric(
            df_work['관세실행세율'].fillna(0), errors='coerce'
        ).fillna(0)
        df_work['실제관세액'] = pd.to_numeric(
            df_work['실제관세액'].fillna(0), errors='coerce'
        ).fillna(0)
        
        # 행별관세 계산에 필요한 컬럼들 전처리
        if '금액' in df_work.columns:
            df_work['금액'] = pd.to_numeric(
                df_work['금액'].fillna(0), errors='coerce'
            ).fillna(0)
        
        if '란결제금액' in df_work.columns:
            df_work['란결제금액'] = pd.to_numeric(
                df_work['란결제금액'].fillna(0), errors='coerce'
            ).fillna(0)
        
        # 행별관세 계산: (실제관세액 × 금액) ÷ 란결제금액
        df_work['행별관세'] = np.where(
            df_work['란결제금액'] != 0,
            (df_work['실제관세액'] * df_work['금액']) / df_work['란결제금액'],
            0
        )
        
        # FTA사후환급 검토 컬럼 계산
        if '적출국코드' in df_work.columns and '원산지코드' in df_work.columns:
            # 적출국코드와 원산지코드가 동일한 경우 'FTA사후환급 검토'로 표시
            df_work['FTA사후환급 검토'] = df_work.apply(
                lambda row: 'FTA사후환급 검토' if (
                    pd.notna(row['적출국코드']) and 
                    pd.notna(row['원산지코드']) and 
                    str(row['적출국코드']).strip() == str(row['원산지코드']).strip() and
                    str(row['적출국코드']).strip() != '' and
                    str(row['원산지코드']).strip() != ''
                ) else '', 
                axis=1
            )
        else:
            df_work['FTA사후환급 검토'] = ''
        
        # NaN 값을 0으로 대체
        df_work.fillna(0, inplace=True)
        # FutureWarning 해결을 위한 추가 코드
        df_work = df_work.infer_objects(copy=False)
        
        # 필터링 조건 적용
        df_filtered = df_work[
            (df_work['세율구분'] == 'A') & 
            (df_work['관세실행세율'] >= 8)
        ]
        
        # 최종 컬럼 순서 정리 (란결제금액은 계산 후 제거)
        final_columns = [col for col in selected_columns 
                        if col in df_filtered.columns and col != '란결제금액']
        df_filtered = df_filtered[final_columns]
        
        if writer:
            # 워크시트 생성
            worksheet = writer.book.add_worksheet('8% 환급 검토')
            
            # 헤더 포맷 설정
            header_format = writer.book.add_format({
                'bold': True,
                'bg_color': '#D9E1F2',
                'border': 1,
                'align': 'center'
            })
            
            # 숫자 포맷 설정 (행별관세용)
            number_format = writer.book.add_format({
                'border': 1,
                'align': 'right',
                'num_format': '#,##0.00'
            })
            
            # 헤더 작성
            for col, header in enumerate(final_columns):
                worksheet.write(0, col, header, header_format)
            
            # 데이터 작성 (특이사항 노란색 표시)
            for row, data in enumerate(df_filtered.values, start=1):
                for col, value in enumerate(data):
                    # 행별관세 컬럼에 숫자 포맷 적용
                    if final_columns[col] == '행별관세':
                        worksheet.write(row, col, value, number_format)
                    # 특이사항 체크 (관세실행세율이 8% 이상인 경우 노란색)
                    elif '관세실행세율' in final_columns and col == final_columns.index('관세실행세율') and value >= 8:
                        highlight_format = writer.book.add_format({
                            'bg_color': '#FFFF00',  # 노란색 배경
                            'border': 1
                        })
                        worksheet.write(row, col, value, highlight_format)
                    # FTA사후환급 검토가 있는 경우 노란색으로 강조
                    elif 'FTA사후환급 검토' in final_columns and col == final_columns.index('FTA사후환급 검토') and value == 'FTA사후환급 검토':
                        highlight_format = writer.book.add_format({
                            'bg_color': '#FFFF00',  # 노란색 배경
                            'border': 1,
                            'bold': True
                        })
                        worksheet.write(row, col, value, highlight_format)
                    else:
                        worksheet.write(row, col, value)
            
            # 컬럼 너비 자동 조정
            for col, header in enumerate(final_columns):
                if header == '행별관세':
                    # 행별관세는 숫자이므로 적절한 너비 설정
                    worksheet.set_column(col, col, 15)
                else:
                    max_length = max(
                        len(str(header)),
                        df_filtered[header].astype(str).str.len().max()
                    )
                    worksheet.set_column(col, col, min(max_length + 2, 50))
            
            # 필터 추가
            worksheet.autofilter(0, 0, len(df_filtered), len(final_columns) - 1)
            
            # 창 틀 고정
            worksheet.freeze_panes(1, 0)
            
            # 인쇄 설정
            worksheet.set_landscape()  # 가로 방향 인쇄
            worksheet.fit_to_pages(1, 0)  # 가로 1페이지에 맞춤, 세로는 자동
            worksheet.set_header('&C&B8% 환급 검토 (행별관세 포함)')  # 중앙 정렬된 헤더
            worksheet.set_footer('&R&P / &N')  # 오른쪽 정렬된 페이지 번호
        
        if document:
            document.add_heading('8% 환급 검토', level=1)
            document.add_paragraph(f'총 {len(df_filtered)}건의 8% 환급 검토 대상이 발견되었습니다.')
            
            # 행별관세 통계 추가
            if '행별관세' in df_filtered.columns and len(df_filtered) > 0:
                avg_tariff = df_filtered['행별관세'].mean()
                total_tariff = df_filtered['행별관세'].sum()
                max_tariff = df_filtered['행별관세'].max()
                
                tariff_info = document.add_paragraph()
                tariff_info.add_run(f"\n행별관세 통계:").bold = True
                tariff_info.add_run(f"\n- 평균 행별관세: {avg_tariff:,.2f}")
                tariff_info.add_run(f"\n- 최대 행별관세: {max_tariff:,.2f}")
                tariff_info.add_run(f"\n- 총 행별관세: {total_tariff:,.2f}")
            
            # FTA사후환급 검토 건수 계산
            if 'FTA사후환급 검토' in df_filtered.columns:
                fta_count = len(df_filtered[df_filtered['FTA사후환급 검토'] == 'FTA사후환급 검토'])
                if fta_count > 0:
                    document.add_paragraph(f'※ 그 중 FTA사후환급 검토 대상: {fta_count}건').bold = True
            
            if len(df_filtered) > 0:
                table = document.add_table(rows=11, cols=len(final_columns))
                table.style = 'Table Grid'
                
                # 헤더 추가
                for j, header in enumerate(final_columns):
                    table.cell(0, j).text = str(header)
                
                # 상위 10개 데이터만 추가
                for i, row in enumerate(df_filtered.head(10).values):
                    for j, value in enumerate(row):
                        # 행별관세는 숫자 포맷팅
                        if final_columns[j] == '행별관세' and pd.notna(value):
                            table.cell(i + 1, j).text = f"{float(value):,.2f}"
                        else:
                            table.cell(i + 1, j).text = str(value)
                
                document.add_paragraph("※ 상위 10건만 표시됨").italic = True
            
            # 행별관세 계산식 설명 추가
            document.add_heading('행별관세 계산식', level=2)
            formula = document.add_paragraph()
            formula.add_run("행별관세 = (실제관세액 × 금액) ÷ 란결제금액").bold = True
        
        print(f"- 8% 환급 검토 분석 완료: {len(df_filtered)}건")
        print(f"- 행별관세 평균: {df_filtered['행별관세'].mean():,.2f}" if '행별관세' in df_filtered.columns and len(df_filtered) > 0 else "- 행별관세: 계산 불가")
        return True
        
    except Exception as e:
        print(f"8% 환급 검토 분석 중 오류 발생: {str(e)}")
        traceback.print_exc()
        return False

def create_summary_sheet(df, df_original, writer):
    """Summary 시트 생성"""
    print("\n- Summary 시트 생성 중...")
    try:
        # 워크시트 생성
        worksheet = writer.book.add_worksheet('Summary')
        workbook = writer.book
        
        # 인쇄 설정
        worksheet.set_landscape()  # 가로 방향 인쇄
        worksheet.fit_to_pages(1, 1)  # 1페이지에 맞춤
        worksheet.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.5)
        worksheet.center_horizontally()  # 가로 중앙 정렬
        
        current_row = 0
        
        # 기본 포맷 설정
        title_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 16,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        subtitle_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 12,
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'bg_color': '#D9E1F2',
            'border': 1
        })
        
        header_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 11,
            'bold': True,
            'bg_color': '#D9E1F2',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        data_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        percent_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '0.0%'
        })
        
        # 열 너비 설정
        worksheet.set_column(0, 0, 20)  # A열
        worksheet.set_column(1, 1, 15)  # B열
        worksheet.set_column(2, 2, 15)  # C열
        worksheet.set_column(3, 8, 12)  # D~I열
        
        # 제목 추가
        worksheet.merge_range(current_row, 0, current_row, 6, '수입신고 분석 보고서', title_format)
        worksheet.set_row(current_row, 30)  # 제목 행 높이 설정
        current_row += 2
        
        # 기본 정보 섹션
        # =====================================
        
        # 1. 수입신고번호 기준 요약 (원본 데이터 사용)
        if '수입신고번호' in df_original.columns:
            total_declarations = df_original['수입신고번호'].nunique()
        else:
            total_declarations = len(df_original)
        worksheet.merge_range(current_row, 0, current_row, 1, "1. 전체 신고 건수", subtitle_format)
        worksheet.write(current_row, 2, total_declarations, data_format)
        current_row += 2
        
        # 2. 거래구분 및 결제방법별 분석 (원본 데이터 사용)
        print("  - 거래구분/결제방법별 분석 중...")
        
        # 섹션 제목
        worksheet.merge_range(current_row, 0, current_row, 6, "2. 거래구분/결제방법별 신고건수", subtitle_format)
        current_row += 1
        
        # 데이터 준비
        if '거래구분' in df_original.columns and '수입신고번호' in df_original.columns:
            pivot2 = pd.pivot_table(df_original, 
                index=['거래구분'],  # 결제방법 제외하여 단순화
                values='수입신고번호',
                aggfunc='nunique',
                margins=True,
                margins_name='총계'
            ).reset_index()
        else:
            # 컬럼이 없는 경우 기본 데이터 생성
            pivot2 = pd.DataFrame({
                '거래구분': ['데이터 없음'],
                '수입신고번호': [0]
            })
        
        # 테이블 생성
        table_start_row = current_row
        
        # 헤더 작성
        for col_num, value in enumerate(pivot2.columns):
            worksheet.write(current_row, col_num, value, header_format)
        
        # 데이터 작성
        for row_num, row in enumerate(pivot2.values):
            for col_num, value in enumerate(row):
                worksheet.write(current_row + 1 + row_num, col_num, value, data_format)
        
        # 테이블 높이 계산
        table_height = len(pivot2) + 1
        
        # 차트 생성 - 테이블 옆에 배치하고 겹치지 않도록 설정
        chart2 = workbook.add_chart({'type': 'column'})
        chart2.add_series({
            'name': '신고건수',
            'categories': f'=Summary!$A${current_row + 2}:$A${current_row + len(pivot2)}',
            'values': f'=Summary!$B${current_row + 2}:$B${current_row + len(pivot2)}',
            'data_labels': {'value': True}
        })
        chart2.set_title({'name': '거래구분별 신고건수'})
        chart2.set_legend({'position': 'none'})  # 범례 제거
        chart2.set_size({'width': 450, 'height': 250})
        chart2.set_style(10)  # 차트 스타일 적용
        
        # 차트 삽입 - 테이블 옆에 배치 (차트와 텍스트가 겹치지 않도록 좌표 조정)
        worksheet.insert_chart(table_start_row, 4, chart2, {'x_offset': 10, 'y_offset': 5})
        
        # 다음 섹션 위치 계산 - 테이블과 차트 중 더 큰 높이 기준
        chart_height_rows = int(250 / 15)  # 차트 높이를 행 수로 변환 (정수로 변환)
        current_row += max(table_height + 2, chart_height_rows) + 5
        
        # 3. 세율구분별 분석 (원본 데이터 사용)
        print("  - 세율구분별 분석 중...")
        
        # 섹션 제목
        worksheet.merge_range(current_row, 0, current_row, 6, "3. 세율구분별 신고건수", subtitle_format)
        current_row += 1
        
        # 데이터 준비
        if '세율구분' in df_original.columns and '수입신고번호' in df_original.columns:
            pivot3 = pd.pivot_table(df_original,
                index='세율구분',
                values='수입신고번호',
                aggfunc='nunique'
            ).reset_index()
        else:
            # 컬럼이 없는 경우 기본 데이터 생성
            pivot3 = pd.DataFrame({
                '세율구분': ['데이터 없음'],
                '수입신고번호': [0]
            })
        
        # 총계 추가
        total_row = {'세율구분': '총계', '수입신고번호': pivot3['수입신고번호'].sum()}
        pivot3 = pd.concat([pivot3, pd.DataFrame([total_row])], ignore_index=True)
        
        # 테이블 시작 위치 저장
        table_start_row = current_row
        
        # 헤더 작성
        for col_num, value in enumerate(pivot3.columns):
            worksheet.write(current_row, col_num, value, header_format)
        
        # 데이터 작성
        for row_num, row in enumerate(pivot3.values):
            for col_num, value in enumerate(row):
                worksheet.write(current_row + 1 + row_num, col_num, value, data_format)
        
        # 테이블 높이 계산
        table_height = len(pivot3) + 1
        
        # 파이 차트 생성 - 테이블 옆에 배치
        chart3 = workbook.add_chart({'type': 'pie'})
        chart3.add_series({
            'name': '세율구분별 비중',
            'categories': f'=Summary!$A${current_row + 2}:$A${current_row + len(pivot3)}',
            'values': f'=Summary!$B${current_row + 2}:$B${current_row + len(pivot3)}',
            'data_labels': {'percentage': True, 'category': True}
        })
        chart3.set_title({'name': '세율구분별 신고건수 비중'})
        chart3.set_style(10)  # 차트 스타일 적용
        chart3.set_size({'width': 450, 'height': 250})
        
        # 차트 삽입 - 테이블 옆에 배치 (차트와 텍스트가 겹치지 않도록 좌표 조정)
        worksheet.insert_chart(table_start_row, 4, chart3, {'x_offset': 10, 'y_offset': 5})
        
        # 다음 섹션 위치 계산 - 테이블과 차트 중 더 큰 높이 기준
        chart_height_rows = int(250 / 15)  # 차트 높이를 행 수로 변환 (정수로 변환)
        current_row += max(table_height + 2, chart_height_rows) + 5
        
        # 4. Risk 분석 요약
        worksheet.merge_range(current_row, 0, current_row, 6, "4. Risk 분석 요약", subtitle_format)
        current_row += 1
        
        # 각 Risk 유형별 신고건수 계산 (df_original 사용)
        if all(col in df_original.columns for col in ['관세실행세율', '세율구분', '수입신고번호']):
            zero_risk_df = df_original[
                (df_original['관세실행세율'] < 8) & 
                (~df_original['세율구분'].astype(str).str.match(r'^F.{3}$'))
            ]
            zero_risk_count = zero_risk_df['수입신고번호'].nunique()
            
            eight_percent_df = df_original[
                (df_original['세율구분'] == 'A') & 
                (df_original['관세실행세율'] >= 8)
            ]
            eight_percent_count = eight_percent_df['수입신고번호'].nunique()
        else:
            zero_risk_count = 0
            eight_percent_count = 0
        
        # 테이블 시작 위치 저장
        table_start_row = current_row
        
        # Risk 요약 테이블 작성
        risk_headers = ['Risk 유형', '신고건수', '비율(%)']
        for col, header in enumerate(risk_headers):
            worksheet.write(current_row, col, header, header_format)
        
        # 비율 계산
        zero_risk_ratio = zero_risk_count/total_declarations if total_declarations > 0 else 0
        eight_percent_ratio = eight_percent_count/total_declarations if total_declarations > 0 else 0
        
        risk_data = [
            ['0% Risk', zero_risk_count, zero_risk_ratio],
            ['8% 환급 검토', eight_percent_count, eight_percent_ratio]
        ]
        
        for row_num, row_data in enumerate(risk_data):
            for col_num, value in enumerate(row_data):
                if col_num == 2:  # 비율 컬럼
                    worksheet.write(current_row + 1 + row_num, col_num, value, percent_format)
                else:
                    worksheet.write(current_row + 1 + row_num, col_num, value, data_format)
        
        # 테이블 높이 계산
        table_height = 3  # 헤더 + 2개 행
        
        # 도넛 차트 생성 - 테이블 옆에 배치
        chart4 = workbook.add_chart({'type': 'doughnut'})
        chart4.add_series({
            'name': 'Risk 분석',
            'categories': '=Summary!$A$' + str(current_row + 1) + ':$A$' + str(current_row + 2),
            'values': '=Summary!$B$' + str(current_row + 1) + ':$B$' + str(current_row + 2),
            'data_labels': {'percentage': True, 'category': True}
        })
        chart4.set_title({'name': 'Risk 분석 결과'})
        chart4.set_size({'width': 450, 'height': 250})
        chart4.set_style(10)  # 차트 스타일 적용
        
        # 차트 삽입 - 테이블 옆에 배치 (차트와 텍스트가 겹치지 않도록 좌표 조정)
        worksheet.insert_chart(table_start_row, 4, chart4, {'x_offset': 10, 'y_offset': 5})
        
        # 다음 섹션 위치 계산
        chart_height_rows = int(250 / 15)  # 차트 높이를 행 수로 변환
        current_row += max(table_height + 2, chart_height_rows) + 5
        
        # 5. 세번부호별 세율구분 및 실행세율 분석 (원본 데이터 사용)
        print("  - 세번부호별 세율구분 및 실행세율 분석 중...")
        
        # 섹션 제목
        worksheet.merge_range(current_row, 0, current_row, 6, "5. 세번부호별 세율구분 및 실행세율 분석", subtitle_format)
        current_row += 1
        
        # 데이터 준비
        if all(col in df_original.columns for col in ['세번부호', '세율구분', '관세실행세율', '수입신고번호']):
            # 세번부호별로 세율구분과 실행세율의 종류를 분석
            tariff_analysis = df_original.groupby('세번부호').agg({
                '세율구분': ['nunique', 'unique'],
                '관세실행세율': ['nunique', 'unique'],
                '수입신고번호': 'nunique'
            }).reset_index()
            
            # 컬럼명 재설정
            tariff_analysis.columns = [
                '세번부호', 
                '세율구분_종류수', '세율구분_목록', 
                '실행세율_종류수', '실행세율_목록', 
                '신고건수'
            ]
            
            # 세율구분과 실행세율이 2개 이상인 항목만 필터링 (이상이 있는 항목)
            tariff_analysis_filtered = tariff_analysis[
                (tariff_analysis['세율구분_종류수'] > 1) | 
                (tariff_analysis['실행세율_종류수'] > 1)
            ].copy()
            
            # 세율구분과 실행세율 목록을 문자열로 변환
            tariff_analysis_filtered['세율구분_목록'] = tariff_analysis_filtered['세율구분_목록'].apply(
                lambda x: ', '.join([str(i) for i in x]) if isinstance(x, (list, np.ndarray)) else str(x)
            )
            tariff_analysis_filtered['실행세율_목록'] = tariff_analysis_filtered['실행세율_목록'].apply(
                lambda x: ', '.join([f"{i:.1f}%" for i in x]) if isinstance(x, (list, np.ndarray)) else str(x)
            )
            
            # 총계 행 추가
            total_tariff_row = {
                '세번부호': '총계',
                '세율구분_종류수': tariff_analysis_filtered['세율구분_종류수'].sum(),
                '세율구분_목록': '전체',
                '실행세율_종류수': tariff_analysis_filtered['실행세율_종류수'].sum(),
                '실행세율_목록': '전체',
                '신고건수': tariff_analysis_filtered['신고건수'].sum()
            }
            tariff_analysis_filtered = pd.concat([
                tariff_analysis_filtered, 
                pd.DataFrame([total_tariff_row])
            ], ignore_index=True)
            
        else:
            # 컬럼이 없는 경우 기본 데이터 생성
            tariff_analysis_filtered = pd.DataFrame({
                '세번부호': ['데이터 없음'],
                '세율구분_종류수': [0],
                '세율구분_목록': [''],
                '실행세율_종류수': [0],
                '실행세율_목록': [''],
                '신고건수': [0]
            })
        
        # 테이블 시작 위치 저장
        table_start_row = current_row
        
        # 헤더 작성
        headers = ['세번부호', '세율구분 종류수', '세율구분 목록', '실행세율 종류수', '실행세율 목록', '신고건수']
        for col_num, header in enumerate(headers):
            worksheet.write(current_row, col_num, header, header_format)
        
        # 데이터 작성
        for row_num, row in enumerate(tariff_analysis_filtered.values):
            for col_num, value in enumerate(row):
                # 세율구분 종류수나 실행세율 종류수가 2개 이상인 경우 노란색으로 강조
                if col_num in [1, 3] and value > 1:
                    highlight_format = writer.book.add_format({
                        'bg_color': '#FFFF00',  # 노란색 배경
                        'border': 1,
                        'align': 'center',
                        'valign': 'vcenter',
                        'bold': True
                    })
                    worksheet.write(current_row + 1 + row_num, col_num, value, highlight_format)
                else:
                    worksheet.write(current_row + 1 + row_num, col_num, value, data_format)
        
        # 테이블 높이 계산
        table_height = len(tariff_analysis_filtered) + 1
        
        # 차트 생성 - 테이블 옆에 배치
        chart5 = workbook.add_chart({'type': 'column'})
        chart5.add_series({
            'name': '이상 항목 수',
            'categories': f'=Summary!$A${current_row + 2}:$A${current_row + len(tariff_analysis_filtered)}',
            'values': f'=Summary!$F${current_row + 2}:$F${current_row + len(tariff_analysis_filtered)}',
            'data_labels': {'value': True}
        })
        chart5.set_title({'name': '세번부호별 신고건수 (이상 항목)'})
        chart5.set_legend({'position': 'none'})  # 범례 제거
        chart5.set_size({'width': 450, 'height': 250})
        chart5.set_style(10)  # 차트 스타일 적용
        
        # 차트 삽입 - 테이블 옆에 배치
        worksheet.insert_chart(table_start_row, 7, chart5, {'x_offset': 10, 'y_offset': 5})
        
        # 다음 섹션 위치 계산
        chart_height_rows = int(250 / 15)  # 차트 높이를 행 수로 변환
        current_row += max(table_height + 2, chart_height_rows) + 5
        
        # 페이지 설정
        worksheet.set_header('&C&B수입신고 분석 요약')  # 중앙 정렬된 헤더
        worksheet.set_footer('&R&D &T')  # 오른쪽 정렬된 날짜와 시간
        
        print("- Summary 시트 생성 완료")
        return True
        
    except Exception as e:
        print(f"\nSummary 시트 생성 중 오류 발생: {str(e)}")
        traceback.print_exc()
        return False

def create_zero_percent_risk_sheet(df, writer):
    """0% Risk 시트 생성"""
    try:
        print("\n- 0% Risk 시트 생성 중...")
        
        # 필요한 컬럼만 선택 - 거래품명 추가하여 란번호 앞에 배치
        selected_columns = [
            '수입신고번호',
            '수리일자',    # 수입신고번호 옆에 수리일자 추가
            'B/L번호',  # B/L번호 추가
            '세번부호', 
            '세율구분',
            '관세실행세율',
            '규격1',
            '규격2',
            '성분1',
            '실제관세액',
            '거래품명',    # 란번호 앞에 거래품명 추가
            '란번호',     # 추가된 컬럼
            '행번호',     # 추가된 컬럼
            '수량_1',     # 추가된 컬럼
            '수량단위_1',  # 추가된 컬럼
            '단가',       # 추가된 컬럼
            '금액',       # 추가된 컬럼
            '란결제금액',  # 행별관세 계산용 추가
            '행별관세'     # 행별관세 컬럼 추가
        ]
        
        # 수리일자 컬럼 매핑 (다양한 이름 지원)
        if '수리일자' not in df.columns:
            possible_date_columns = ['수리일자_converted', '수리일자_변환', 
                                   '수리일자_날짜', '수리일자_변환됨']
            for col in possible_date_columns:
                if col in df.columns:
                    df = df.rename(columns={col: '수리일자'})
                    print(f"0% Risk 시트: 수리일자 컬럼을 '{col}'에서 매핑했습니다.")
                    break
        
        # 0% Risk 조건에 맞는 데이터 필터링
        df_zero_risk = df[
            (df['관세실행세율'] < 8) & 
            (~df['세율구분'].astype(str).str.match(r'^F.{3}$'))
        ]
        
        # 행별관세 계산에 필요한 컬럼들을 제외하고 존재하는 컬럼만 선택
        base_columns = [col for col in selected_columns 
                       if col not in ['행별관세'] and col in df_zero_risk.columns]
        
        if len(base_columns) < len(selected_columns) - 1:
            missing_columns = [col for col in selected_columns 
                             if col not in df_zero_risk.columns and col != '행별관세']
            print(f"경고: 0% Risk 시트에서 다음 컬럼들이 누락되었습니다: {missing_columns}")
        
        # 필요한 컬럼만 선택
        df_zero_risk = df_zero_risk[base_columns].copy()
        
        # 행별관세 계산에 필요한 컬럼들 전처리
        if '실제관세액' in df_zero_risk.columns:
            df_zero_risk['실제관세액'] = pd.to_numeric(
                df_zero_risk['실제관세액'].fillna(0), errors='coerce'
            ).fillna(0)
        
        if '금액' in df_zero_risk.columns:
            df_zero_risk['금액'] = pd.to_numeric(
                df_zero_risk['금액'].fillna(0), errors='coerce'
            ).fillna(0)
        
        if '란결제금액' in df_zero_risk.columns:
            df_zero_risk['란결제금액'] = pd.to_numeric(
                df_zero_risk['란결제금액'].fillna(0), errors='coerce'
            ).fillna(0)
        
        # 행별관세 계산: (실제관세액 × 금액) ÷ 란결제금액
        if all(col in df_zero_risk.columns for col in ['실제관세액', '금액', '란결제금액']):
            df_zero_risk['행별관세'] = np.where(
                df_zero_risk['란결제금액'] != 0,
                (df_zero_risk['실제관세액'] * df_zero_risk['금액']) / df_zero_risk['란결제금액'],
                0
            )
            print("- 0% Risk 시트: 행별관세 계산 완료")
        else:
            print("⚠️ 0% Risk 시트: 행별관세 계산에 필요한 일부 컬럼이 없어 0으로 설정됩니다.")
            df_zero_risk['행별관세'] = 0
        
        # NaN 값을 0으로 대체
        df_zero_risk.fillna(0, inplace=True)
        # FutureWarning 해결을 위한 추가 코드
        df_zero_risk = df_zero_risk.infer_objects(copy=False)
        
        # 최종 컬럼 순서 정리 (란결제금액은 계산 후 제거)
        final_columns = [col for col in selected_columns 
                        if col in df_zero_risk.columns and col != '란결제금액']
        df_zero_risk = df_zero_risk[final_columns]
        
        # 워크시트 생성
        worksheet = writer.book.add_worksheet('0% Risk')
        
        # 헤더 포맷 설정
        header_format = writer.book.add_format({
            'bold': True,
            'bg_color': '#D9E1F2',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # 데이터 포맷 설정
        data_format = writer.book.add_format({
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        # 숫자 포맷 설정 (행별관세용)
        number_format = writer.book.add_format({
            'border': 1,
            'align': 'right',
            'num_format': '#,##0.00'
        })
        
        # 헤더 작성
        for col, header in enumerate(final_columns):
            worksheet.write(0, col, header, header_format)
        
        # 데이터 작성 (특이사항 노란색 표시)
        for row, data in enumerate(df_zero_risk.values, start=1):
            for col, value in enumerate(data):
                # 행별관세 컬럼에 숫자 포맷 적용
                if final_columns[col] == '행별관세':
                    worksheet.write(row, col, value, number_format)
                # 특이사항 체크 (관세실행세율이 0인 경우 노란색)
                elif ('관세실행세율' in final_columns and 
                      col == final_columns.index('관세실행세율') and value == 0):
                    highlight_format = writer.book.add_format({
                        'bg_color': '#FFFF00',  # 노란색 배경
                        'border': 1,
                        'align': 'left',
                        'valign': 'vcenter'
                    })
                    worksheet.write(row, col, value, highlight_format)
                else:
                    worksheet.write(row, col, value, data_format)
        
        # 컬럼 너비 자동 조정 (최대 50)
        for col, header in enumerate(final_columns):
            if header == '행별관세':
                # 행별관세는 숫자이므로 적절한 너비 설정
                worksheet.set_column(col, col, 15)
            else:
                max_length = max(
                    len(str(header)),
                    df_zero_risk[header].astype(str).apply(len).max()
                )
                worksheet.set_column(col, col, min(max_length + 2, 50))
        
        print(f"- 0% Risk 시트에 {len(df_zero_risk):,}개의 데이터가 작성되었습니다.")
        
        # 필터 추가
        worksheet.autofilter(0, 0, len(df_zero_risk), len(final_columns) - 1)
        
        # 창 틀 고정
        worksheet.freeze_panes(1, 0)
        
        # 인쇄 설정
        worksheet.set_landscape()  # 가로 방향 인쇄
        worksheet.fit_to_pages(1, 0)  # 가로 1페이지에 맞춤, 세로는 자동
        worksheet.set_header('&C&B0% Risk (행별관세 포함)')  # 중앙 정렬된 헤더
        worksheet.set_footer('&R&P / &N')  # 오른쪽 정렬된 페이지 번호
    
    except Exception as e:
        print(f"0% Risk 시트 생성 중 오류 발생: {str(e)}")
        traceback.print_exc()

def add_standard_price_analysis(df, worksheet, writer, document, start_row):
    """단가 Risk 분석 (app_enhanced.py 기반)"""
    try:
        print("- 단가 Risk 분석 시작...")
        
        # None 체크 추가
        if worksheet is None or writer is None:
            print("Warning: worksheet 또는 writer가 None입니다.")
            return start_row
        
        # 필요한 컬럼 체크 (이미지 헤더에 맞게 수정)
        required_columns = ['규격1', '세번부호', '거래구분', '결제방법', '수리일자', '수입신고번호',
                          '단가', '결제통화단위', '거래품명', 
                          '란번호', '행번호', '수량_1', '수량단위_1', '금액']
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"경고: 누락된 컬럼: {missing_columns}")
            available_columns = [col for col in required_columns if col in df.columns]
            if '단가' not in available_columns:
                print("오류: 필수 컬럼(단가)이 없어 분석할 수 없습니다.")
                return start_row
        else:
            available_columns = required_columns
        
        # 단가를 숫자형으로 변환
        df_work = df.copy()
        df_work['단가'] = pd.to_numeric(df_work['단가'].fillna(0), errors='coerce').fillna(0)
        
        # 단가가 0보다 큰 데이터만 분석
        df_work = df_work[df_work['단가'] > 0]
        
        if len(df_work) == 0:
            print("경고: 단가 데이터가 없어 분석할 수 없습니다.")
            return start_row
        
        # 그룹화 기준 (규격1만 사용)
        group_columns = ['규격1']
        
        # 데이터 그룹화 및 분석
        print("- 데이터 그룹화 중...")
        print(f"- 전체 데이터 크기: {len(df_work)}")
        print(f"- 규격1 고유값 개수: {df_work['규격1'].nunique()}")
        
        # 집계 함수 정의 (이미지 헤더에 맞게 수정)
        agg_dict = {
            '세번부호': 'first',
            '거래구분': 'first',
            '결제방법': 'first',
            '수리일자': ['min', 'max'],
            '수입신고번호': ['min', 'max'],  # 신고번호 추가
            '단가': ['mean', 'max', 'min', 'std', 'count'],
            '결제통화단위': 'first',
            '거래품명': 'first',
            '란번호': 'first',
            '행번호': 'first',
            '수량_1': 'first',
            '수량단위_1': 'first',
            '금액': 'sum'
        }
        
        # 존재하는 컬럼만 선택
        available_group_columns = [col for col in group_columns if col in df_work.columns]
        available_agg_dict = {col: agg_dict[col] for col in agg_dict if col in df_work.columns}
        
        grouped = df_work.groupby(available_group_columns).agg(available_agg_dict).reset_index()
        
        # 집계 후 실제 컬럼명에 맞춰 new_columns를 동적으로 생성 (이미지 헤더에 맞게 수정)
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
                elif col[0] == '수리일자' and col[1] == 'min':
                    new_columns.append('Min 수리일자')
                elif col[0] == '수리일자' and col[1] == 'max':
                    new_columns.append('Max 수리일자')
                elif col[0] == '수입신고번호' and col[1] == 'min':
                    new_columns.append('Min 신고번호')
                elif col[0] == '수입신고번호' and col[1] == 'max':
                    new_columns.append('Max 신고번호')
                elif col[0] == '세번부호' and col[1] == 'first':
                    new_columns.append('세번부호')
                elif col[0] == '거래구분' and col[1] == 'first':
                    new_columns.append('거래구분')
                elif col[0] == '결제방법' and col[1] == 'first':
                    new_columns.append('결제방법')
                elif col[0] == '결제통화단위' and col[1] == 'first':
                    new_columns.append('결제통화단위')
                elif col[0] == '거래품명' and col[1] == 'first':
                    new_columns.append('거래품명')
                elif col[0] == '란번호' and col[1] == 'first':
                    new_columns.append('란번호')
                elif col[0] == '행번호' and col[1] == 'first':
                    new_columns.append('행번호')
                elif col[0] == '수량_1' and col[1] == 'first':
                    new_columns.append('수량_1')
                elif col[0] == '수량단위_1' and col[1] == 'first':
                    new_columns.append('수량단위_1')
                elif col[0] == '금액' and col[1] == 'sum':
                    new_columns.append('금액')
                else:
                    new_columns.append(f'{col[0]}_{col[1]}')
            else:
                new_columns.append(col)
        grouped.columns = new_columns
        
        # 위험도 계산 (app_enhanced.py 기반)
        grouped['단가편차율'] = np.where(
            grouped['평균단가'] > 0,
            (grouped['최고단가'] - grouped['최저단가']) / grouped['평균단가'],
            0
        )
        
        # 위험도 분류 (app_enhanced.py 기반)
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
        
        # 비고 생성 (app_enhanced.py 기반)
        grouped['비고'] = grouped.apply(lambda row: 
            f'평균단가 확인 필요' if row['평균단가'] == 0 
            else f'단가편차: {row["단가편차율"]*100:.1f}%', axis=1
        )
        
        # 결과 표시
        print(f"- 총 그룹 수: {len(grouped):,}개")
        high_risk = len(grouped[grouped['위험도'].isin(['높음', '매우높음'])])
        print(f"- 고위험 그룹: {high_risk:,}개")
        avg_deviation = grouped['단가편차율'].mean() * 100
        print(f"- 평균 편차율: {avg_deviation:.1f}%")
        zero_price = len(grouped[grouped['평균단가'] == 0])
        print(f"- 확인필요: {zero_price:,}개")
        
        # 엑셀 헤더 정의 (비교하기 좋게 재배치 - 최고/최저 단가를 신고번호/수리일자와 나란히 배치)
        headers = [
            '세번부호', '규격1', '거래구분', '결제방법', 
            '평균단가', '결제통화단위', '위험도', '비고',
            'Min 신고번호', 'Min 수리일자', '최저단가',
            'Max 신고번호', 'Max 수리일자', '최고단가',
            '거래품명', '란번호', '행번호', '수량_1', '수량단위_1', '단가', '금액'
        ]
        
        # 존재하는 컬럼만 선택
        available_headers = [col for col in headers if col in grouped.columns]
        if len(available_headers) < len(headers):
            missing_headers = [col for col in headers if col not in grouped.columns]
            print(f"경고: 단가 Risk 시트에서 다음 컬럼들이 누락되었습니다: {missing_headers}")
        
        grouped = grouped[available_headers]
        
        # 헤더 쓰기
        for col, header in enumerate(available_headers):
            worksheet.write(start_row, col, header, 
                           writer.book.add_format({'bold': True, 'border': 1}))
        
        print("- 결과 쓰기 중...")
        print(f"  최종 데이터 크기: {grouped.shape}")
        
        # 엑셀에 데이터 쓰기 (특이사항 노란색 표시)
        print("  엑셀 파일 쓰기 시작...")
        for i, row in enumerate(grouped.values):
            for j, value in enumerate(row):
                # NaN 값 처리
                if pd.isna(value):
                    cell_value = ''
                else:
                    cell_value = value
                
                # 특이사항 체크 (위험도가 '높음', '매우높음' 또는 '확인필요'인 경우 노란색)
                if '위험도' in available_headers and j == available_headers.index('위험도') and cell_value in ['높음', '매우높음', '확인필요']:
                    highlight_format = writer.book.add_format({
                        'bg_color': '#FFFF00',  # 노란색 배경
                        'border': 1
                    })
                    worksheet.write(start_row + i + 1, j, cell_value, highlight_format)
                else:
                    worksheet.write(start_row + i + 1, j, cell_value, 
                                   writer.book.add_format({'border': 1}))
        print("  엑셀 파일 쓰기 완료")
        
        # 컬럼 너비 자동 조정
        for col, header in enumerate(available_headers):
            max_length = max(len(str(header)), 12)  # 최소 너비 12
            worksheet.set_column(col, col, max_length)
        
        # 워드 문서에 추가
        print("  워드 문서 생성 시작...")
        document.add_heading('단가 Risk 분석', level=2)
        
        # 요약 정보 추가
        summary = document.add_paragraph()
        summary.add_run(f"총 {len(grouped)}건의 데이터가 분석되었습니다.").bold = True
        summary.add_run("\n위험도 분포:")
        risk_summary = grouped['위험도'].value_counts()
        for risk, count in risk_summary.items():
            if risk != '보통':
                run = summary.add_run(f"\n- {risk}: {count}건")
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            else:
                summary.add_run(f"\n- {risk}: {count}건")
        
        # 상위 10개 행만 테이블로 표시
        table = document.add_table(rows=11, cols=len(available_headers))  # 헤더 + 10개 행
        table.style = 'Table Grid'
        
        # 헤더 추가
        for j, header in enumerate(available_headers):
            table.cell(0, j).text = str(header)
        
        # 상위 10개 데이터만 추가
        print("  워드 문서에 상위 10개 데이터 추가 중...")
        for i, row in enumerate(grouped.head(10).values):
            for j, value in enumerate(row):
                if j < len(available_headers):  # 컬럼 수를 초과하지 않도록
                    table.cell(i + 1, j).text = str(value)
        
        document.add_paragraph("※ 상위 10건만 표시됨").italic = True
        
        print("  워드 문서 생성 완료")
        return start_row + len(grouped) + 15
        
    except Exception as e:
        print(f"단가 Risk 분석 중 오류 발생: {str(e)}")
        print("\n상세 오류 정보:")
        import traceback
        print(traceback.format_exc())
        return start_row

def add_tariff_risk_analysis(df, worksheet, writer, document, start_row):
    """세율 Risk 분석"""
    try:
        required_columns = [
            '수입신고번호', 
            '수리일자',
            '규격1', '규격2', '규격3',
            '성분1', '성분2', '성분3',
            '세번부호', 
            '세율구분', 
            '세율설명',
            '과세가격달러',
            '실제관세액',
            '결제방법',
            '금액',        # 행별관세 계산용 추가
            '란결제금액'   # 행별관세 계산용 추가
        ]
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"누락된 컬럼: {missing_columns}")
            # 행별관세 계산에 필요한 컬럼이 없어도 분석 진행
        
        # 규격1별 세번부호 분석 - 개선된 로직
        print("- 규격1별 세번부호 분석 중...")
        if '규격1' in df.columns and '세번부호' in df.columns:
            # 디버깅: 전체 데이터 정보 출력
            print(f"- 전체 데이터 크기: {len(df)}")
            print(f"- 규격1 고유값 개수: {df['규격1'].nunique()}")
            print(f"- 세번부호 고유값 개수: {df['세번부호'].nunique()}")
            
            # 규격1별로 세번부호의 고유값 개수를 계산
            risk_specs = df.groupby('규격1')['세번부호'].nunique()
            print("- 규격1별 세번부호 개수 분포:")
            print(risk_specs.value_counts().sort_index())
            
            # 세번부호가 2개 이상인 규격1만 선택
            risk_specs = risk_specs[risk_specs > 1]
            
            print(f"- 발견된 위험 규격1 수: {len(risk_specs)}")
            if len(risk_specs) > 0:
                print("- 위험 규격1 예시:")
                for i, (spec, count) in enumerate(risk_specs.head(5).items()):
                    print(f"  {i+1}. {spec} (세번부호 {count}개)")
                    # 해당 규격1의 세번부호들 출력
                    spec_tariffs = df[df['규격1'] == spec]['세번부호'].unique()
                    print(f"     세번부호: {', '.join(map(str, spec_tariffs))}")
            else:
                print("- 세번부호가 2개 이상인 규격1이 없습니다.")
                # 디버깅: 상위 5개 규격1의 세번부호 확인
                print("- 상위 5개 규격1의 세번부호 확인:")
                top_specs = (df.groupby('규격1')['세번부호'].nunique()
                           .sort_values(ascending=False).head(5))
                for spec, count in top_specs.items():
                    print(f"  {spec}: {count}개 세번부호")
                    spec_tariffs = df[df['규격1'] == spec]['세번부호'].unique()
                    print(f"    세번부호: {', '.join(map(str, spec_tariffs))}")
        else:
            print("경고: 세율 Risk 분석에 필요한 컬럼이 누락되었습니다.")
            risk_specs = pd.Series(dtype='object')
        
        if len(risk_specs) == 0:
            print("- 세율 Risk가 발견되지 않았습니다.")
            if worksheet:
                worksheet.write(start_row, 0, "세율 Risk가 발견되지 않았습니다.")
            return start_row + 1
            
        print(f"- 위험 규격1 수: {len(risk_specs)}")
        
        # 규격1 기준으로 정렬
        if '규격1' in df.columns:
            # 존재하는 컬럼만 선택
            available_columns = [col for col in required_columns if col in df.columns]
            risk_data = df[df['규격1'].isin(risk_specs.index)][available_columns].copy()
            
            # 행별관세 계산에 필요한 컬럼들 전처리
            if '실제관세액' in risk_data.columns:
                risk_data['실제관세액'] = pd.to_numeric(
                    risk_data['실제관세액'].fillna(0), errors='coerce'
                ).fillna(0)
            
            if '금액' in risk_data.columns:
                risk_data['금액'] = pd.to_numeric(
                    risk_data['금액'].fillna(0), errors='coerce'
                ).fillna(0)
            
            if '란결제금액' in risk_data.columns:
                risk_data['란결제금액'] = pd.to_numeric(
                    risk_data['란결제금액'].fillna(0), errors='coerce'
                ).fillna(0)
            
            # 행별관세 계산: (실제관세액 × 금액) ÷ 란결제금액
            if all(col in risk_data.columns for col in ['실제관세액', '금액', '란결제금액']):
                risk_data['행별관세'] = np.where(
                    risk_data['란결제금액'] != 0,
                    (risk_data['실제관세액'] * risk_data['금액']) / risk_data['란결제금액'],
                    0
                )
                print("- 세율 Risk 시트: 행별관세 계산 완료")
            else:
                print("⚠️ 세율 Risk 시트: 행별관세 계산에 필요한 일부 컬럼이 없어 0으로 설정됩니다.")
                risk_data['행별관세'] = 0
            
            # 규격1, 세번부호 기준 정렬
            risk_data = risk_data.sort_values(['규격1', '세번부호']).fillna('')
            
            # 최종 컬럼 순서 정리 (란결제금액은 계산 후 제거)
            final_columns = [col for col in available_columns if col != '란결제금액']
            if '행별관세' not in final_columns:
                final_columns.append('행별관세')
            risk_data = risk_data[final_columns]
        else:
            risk_data = pd.DataFrame(columns=required_columns)
            final_columns = required_columns
        
        if worksheet and writer:
            # 헤더 포맷 설정
            header_format = writer.book.add_format({
                'bold': True,
                'bg_color': '#D9E1F2',
                'border': 1,
                'align': 'center'
            })
            
            # 데이터 포맷 설정
            data_format = writer.book.add_format({
                'border': 1,
                'align': 'left'
            })
            
            # 숫자 포맷 설정 (행별관세용)
            number_format = writer.book.add_format({
                'border': 1,
                'align': 'right',
                'num_format': '#,##0.00'
            })
            
            # 헤더 작성
            for col, header in enumerate(final_columns):
                worksheet.write(start_row, col, header, header_format)
            
            # 데이터 작성 (특이사항 노란색 표시)
            for row, data in enumerate(risk_data.values, start=start_row + 1):
                for col, value in enumerate(data):
                    # 행별관세 컬럼에 숫자 포맷 적용
                    if final_columns[col] == '행별관세':
                        worksheet.write(row, col, value, number_format)
                    # 특이사항 체크 (세번부호가 다른 경우 노란색으로 표시)
                    elif (col == final_columns.index('세번부호') and 
                          '규격1' in risk_data.columns):
                        # 같은 규격1 내에서 다른 세번부호가 있는지 확인
                        current_spec = risk_data.iloc[row - start_row - 1]['규격1']
                        same_spec_data = risk_data[risk_data['규격1'] == current_spec]
                        if len(same_spec_data['세번부호'].unique()) > 1:
                            highlight_format = writer.book.add_format({
                                'bg_color': '#FFFF00',  # 노란색 배경
                                'border': 1,
                                'align': 'left'
                            })
                            worksheet.write(row, col, str(value), highlight_format)
                        else:
                            worksheet.write(row, col, str(value), data_format)
                    else:
                        worksheet.write(row, col, str(value), data_format)
            
            # 컬럼 너비 자동 조정
            for col, header in enumerate(final_columns):
                if header == '행별관세':
                    # 행별관세는 숫자이므로 적절한 너비 설정
                    worksheet.set_column(col, col, 15)
                else:
                    max_length = max(
                        len(str(header)),
                        risk_data[header].astype(str).str.len().max() if len(risk_data) > 0 else len(str(header))
                    )
                    worksheet.set_column(col, col, min(max_length + 2, 50))
            
            # 필터 추가
            worksheet.autofilter(start_row, 0, start_row + len(risk_data), 
                               len(final_columns) - 1)
            
            # 창 틀 고정
            worksheet.freeze_panes(start_row + 1, 0)
        
        if document:
            document.add_heading('세율 Risk 분석', level=2)
            summary = document.add_paragraph()
            summary.add_run(f"총 {len(risk_data)}건의 세율 Risk가 발견되었습니다.").bold = True
            
            # 행별관세 통계 추가
            if '행별관세' in risk_data.columns and len(risk_data) > 0:
                avg_tariff = risk_data['행별관세'].mean()
                total_tariff = risk_data['행별관세'].sum()
                max_tariff = risk_data['행별관세'].max()
                
                tariff_info = document.add_paragraph()
                tariff_info.add_run("\n세율 Risk 행별관세 통계:").bold = True
                tariff_info.add_run(f"\n- 평균 행별관세: {avg_tariff:,.2f}")
                tariff_info.add_run(f"\n- 최대 행별관세: {max_tariff:,.2f}")
                tariff_info.add_run(f"\n- 총 행별관세: {total_tariff:,.2f}")
            
            if len(risk_data) > 0:
                # 상위 10개 데이터만 테이블로 표시
                table = document.add_table(rows=11, cols=len(final_columns))
                table.style = 'Table Grid'
                
                # 헤더 추가
                for j, header in enumerate(final_columns):
                    table.cell(0, j).text = str(header)
                
                # 상위 10개 데이터만 추가
                for i, row in enumerate(risk_data.head(10).values):
                    for j, value in enumerate(row):
                        # 행별관세는 숫자 포맷팅
                        if final_columns[j] == '행별관세' and pd.notna(value):
                            table.cell(i + 1, j).text = f"{float(value):,.2f}"
                        else:
                            table.cell(i + 1, j).text = str(value)
            
            document.add_paragraph("※ 상위 10건만 표시됨").italic = True
            
            # 행별관세 계산식 설명 추가
            document.add_heading('행별관세 계산식', level=3)
            formula = document.add_paragraph()
            formula.add_run("행별관세 = (실제관세액 × 금액) ÷ 란결제금액").bold = True
        else:
            if document:
                document.add_paragraph("세율 Risk가 발견되지 않았습니다.")
        
        print(f"- 세율 Risk 분석 완료: {len(risk_data)}건")
        return start_row + len(risk_data) + 2
        
    except Exception as e:
        print(f"세율 Risk 분석 중 오류 발생: {e}")
        traceback.print_exc()
        return start_row

def create_verification_methods_sheet(writer):
    """검증방법 시트 생성"""
    try:
        print("\n- 검증방법 시트 생성 중...")
        
        # 워크시트 생성
        worksheet = writer.book.add_worksheet('검증방법')
        workbook = writer.book
        
        # 포맷 설정
        title_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 14,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        subtitle_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 12,
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'bg_color': '#D9E1F2',
            'border': 1
        })
        
        content_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'align': 'left',
            'valign': 'top',
            'border': 1,
            'text_wrap': True
        })
        
        highlight_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'align': 'left',
            'valign': 'top',
            'border': 1,
            'text_wrap': True,
            'bg_color': '#FFFF00'  # 노란색 배경
        })
        
        # 열 너비 설정
        worksheet.set_column(0, 0, 25)  # A열 - 시트명
        worksheet.set_column(1, 1, 60)  # B열 - 검증로직
        worksheet.set_column(2, 2, 40)  # C열 - 특이사항
        
        current_row = 0
        
        # 제목
        worksheet.merge_range(current_row, 0, current_row, 2, '수입신고 분석 검증방법', title_format)
        worksheet.set_row(current_row, 30)
        current_row += 2
        
        # 1. 8% 환급 검토
        worksheet.write(current_row, 0, '1. 8% 환급 검토', subtitle_format)
        worksheet.write(current_row, 1, 
            '• 필터링 조건: 세율구분 = "A" AND 관세실행세율 ≥ 8%\n' +
            '• 목적: 8% 환급 검토가 필요한 수입신고 건들 식별\n' +
            '• 추가 컬럼: 적출국코드, 원산지코드, 무역거래처상호, 무역거래처국가코드', 
            content_format)
        worksheet.write(current_row, 2, 
            '• 세율구분 "A"는 일반적으로 가장 관세율이 높은 구분\n' +
            '• 8% 이상의 관세율은 환급 대상이 될 수 있음', 
            highlight_format)
        current_row += 1
        
        # 2. 0% Risk
        worksheet.write(current_row, 0, '2. 0% Risk', subtitle_format)
        worksheet.write(current_row, 1, 
            '• 필터링 조건: 관세실행세율 < 8% AND 세율구분 ≠ F***\n' +
            '• 목적: 관세율이 낮거나 면세 대상이지만 추가 검토가 필요한 건들\n' +
            '• F로 시작하는 4자리 코드는 특별한 세율구분으로 제외', 
            content_format)
        worksheet.write(current_row, 2, 
            '• 관세율이 낮은데도 특별한 세율구분이 아닌 경우 주의 필요\n' +
            '• 면세 대상이지만 실제로는 관세가 부과될 수 있는 경우', 
            highlight_format)
        current_row += 1
        
        # 3. 세율 Risk
        worksheet.write(current_row, 0, '3. 세율 Risk', subtitle_format)
        worksheet.write(current_row, 1, 
            '• 분석 방법: 규격1 기준으로 그룹화하여 세번부호의 고유값 개수 확인\n' +
            '• 위험 판정: 동일 규격1에 대해 서로 다른 세번부호가 2개 이상인 경우\n' +
            '• 목적: 동일 상품(규격1)에 대한 세번부호 불일치 위험 식별\n' +
            '• 예시: "DEMO SYS 1ML LG 0000-S000P1MLF"에 여러 세번부호 적용', 
            content_format)
        worksheet.write(current_row, 2, 
            '• 동일 상품인데 다른 세번부호가 적용되면 관세율 차이 발생\n' +
            '• 세번부호 분류 오류 가능성 또는 상품 특성 차이\n' +
            '• 세율 Risk 발견 시 해당 규격1의 세번부호들을 상세 검토 필요', 
            highlight_format)
        current_row += 1
        
        # 4. 단가 Risk
        worksheet.write(current_row, 0, '4. 단가 Risk', subtitle_format)
        worksheet.write(current_row, 1, 
            '• 그룹화 기준: 규격1, 세번부호, 거래구분, 결제방법, 수리일자\n' +
            '• 위험도 계산: \n' +
            '  - 10% 초과~20% 이하: "높음"\n' +
            '  - 20% 초과: "매우 높음"\n' +
            '• 특이사항: 평균단가가 0인 경우 "확인필요"로 분류\n' +
            '• 추가 정보: Min/Max 신고번호의 수리일자 표시', 
            content_format)
        worksheet.write(current_row, 2, 
            '• 단가 변동성이 10% 초과하면 주의 필요\n' +
            '• 20% 초과는 매우 비정상적인 가격 차이 가능성\n' +
            '• 평균단가 0은 데이터 오류 또는 특별한 거래 형태\n' +
            '• 수리일자 차이로 시간적 변동성 확인 가능', 
            highlight_format)
        current_row += 1
        
        # 5. Summary
        worksheet.write(current_row, 0, '5. Summary', subtitle_format)
        worksheet.write(current_row, 1, 
            '• 전체 신고 건수: 수입신고번호 기준 고유 건수\n' +
            '• 거래구분별 분석: 거래구분별 신고건수 피벗 테이블\n' +
            '• 세율구분별 분석: 세율구분별 신고건수 및 비중\n' +
            '• Risk 분석 요약: 0% Risk와 8% 환급 검토 건수 및 비율', 
            content_format)
        worksheet.write(current_row, 2, 
            '• 전체적인 수입신고 현황 파악\n' +
            '• Risk 분포를 통한 우선순위 설정 가능', 
            highlight_format)
        current_row += 2
        
        # 6. 원본데이터
        current_row += 2
        worksheet.write(current_row, 0, '6. 원본데이터', subtitle_format)
        worksheet.write(current_row, 1, 
            '• 분석에 사용된 원본 엑셀 파일의 모든 데이터\n' +
            '• 상위 1000개 행만 표시 (파일 크기 제한)\n' +
            '• 모든 컬럼과 원본 데이터 구조 확인 가능\n' +
            '• 필터링 및 정렬 기능 제공', 
            content_format)
        worksheet.write(current_row, 2, 
            '• 원본 데이터와 분석 결과 비교 검토 가능\n' +
            '• 데이터 품질 및 구조 확인용', 
            highlight_format)
        current_row += 1
        
        # 특이사항 표시 방법
        worksheet.write(current_row, 0, '특이사항 표시 방법', subtitle_format)
        worksheet.write(current_row, 1, 
            '• 노란색 배경: 각 시트에서 특별히 주의가 필요한 항목\n' +
            '• 위험도 "높음": 단가 Risk에서 단가 변동성이 10% 초과~20% 이하\n' +
            '• 위험도 "매우 높음": 단가 변동성이 20% 초과\n' +
            '• 위험도 "확인필요": 평균단가가 0인 경우\n' +
            '• 세율 Risk: 동일 규격에 다른 세번부호 적용된 경우', 
            content_format)
        worksheet.write(current_row, 2, 
            '• 노란색으로 표시된 항목은 반드시 검토 필요\n' +
            '• 데이터 오류 또는 비정상적인 거래 형태일 가능성', 
            highlight_format)
        
        # 페이지 설정
        worksheet.set_header('&C&B검증방법')
        worksheet.set_footer('&R&D &T')
        
        print("- 검증방법 시트 생성 완료")
        return True
        
    except Exception as e:
        print(f"검증방법 시트 생성 중 오류 발생: {str(e)}")
        traceback.print_exc()
        return False

def add_headers_to_risk_sheet(workbook_path):
    try:
        # 워크북 열기
        wb = openpyxl.load_workbook(workbook_path)
        
        # '단가 Risk' 시트 선택
        if '단가 Risk' not in wb.sheetnames:
            raise ValueError("'단가 Risk' 시트를 찾을 수 없습니다.")
            
        risk_sheet = wb['단가 Risk']
        
        # 헤더 값 정의
        headers = ['품목코드', '품목명', '단가', '위험도', '비고']
        
        # 헤더 삽입
        for col, header in enumerate(headers, start=1):
            risk_sheet.cell(row=1, column=col, value=header)
        
        # 저장
        wb.save(workbook_path)
        return True
        
    except Exception as e:
        print(f"헤더 추가 중 오류 발생: {str(e)}")
        return False

def analyze_original_data(df):
    """원본 데이터 분석"""
    try:
        print("\n=== 원본 데이터 분석 ===")
        print(f"1. 전체 데이터 크기: {df.shape}")
        print("\n2. 컬럼 목록과 데이터 타입:")
        print(df.dtypes)
        
        print("\n3. '세율구분' 분석:")
        print("- 고유값 목록:")
        print(df['세율구분'].value_counts().to_frame())
        print("\n- 데이터 타입:", df['세율구분'].dtype)
        print("- NULL 값 개수:", df['세율구분'].isnull().sum())
        print("- 공백 포함 값 개수:", df['세율구분'].str.isspace().sum() if df['세율구분'].dtype == 'object' else "N/A")
        
        print("\n4. '관세실행세율' 분석:")
        print("- 기본 통계:")
        print(df['관세실행세율'].describe())
        print("\n- 데이터 타입:", df['관세실행세율'].dtype)
        print("- NULL 값 개수:", df['관세실행세율'].isnull().sum())
        print("- 8 이상인 값의 개수:", len(df[df['관세실행세율'] >= 8]))
        
        print("\n5. 조건별 데이터 수:")
        condition_a = df['세율구분'].astype(str).str.strip() == 'A'
        condition_b = df['관세실행세율'] >= 8
        print(f"- 세율구분 'A'인 데이터 수: {condition_a.sum()}")
        print(f"- 관세실행세율 8 이상인 데이터 수: {condition_b.sum()}")
        print(f"- 두 조건 모두 만족하는 데이터 수: {(condition_a & condition_b).sum()}")
        
        print("\n6. 조건을 만족하는 데이터 샘플:")
        filtered_data = df[condition_a & condition_b].head()
        if not filtered_data.empty:
            print(filtered_data[['세율구분', '관세실행세율', '수입신고번호', '세번부호']])
        else:
            print("조건을 만족하는 데이터가 없습니다.")
        
        return True
        
    except Exception as e:
        print(f"\n데이터 분석 중 오류 발생: {str(e)}")
        traceback.print_exc()
        return False

def check_excel_headers(file_path):
    """엑셀 파일의 헤더 확인"""
    try:
        print("\n=== 엑셀 파일 헤더 확인 ===")
        df = pd.read_excel(file_path, nrows=0)  # 헤더만 읽기
        print("\n현재 컬럼 목록:")
        for idx, col in enumerate(df.columns):
            print(f"{idx+1}. {col}")
        return df.columns.tolist()
    except Exception as e:
        print(f"헤더 확인 중 오류 발생: {str(e)}")
        return None

def main():
    try:
        # tkinter 초기화
        root = tk.Tk()
        root.withdraw()
        
        # 입력 파일 선택
        print("\n=== 파일 선택 ===")
        input_file = filedialog.askopenfilename(
            title="분석할 엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if not input_file:
            print("파일 선택이 취소되었습니다.")
            return False
            
        print(f"선택된 파일: {input_file}")
        
        # 엑셀 파일 읽기
        print("\n1. 엑셀 파일 읽기 시작...")
        df = pd.read_excel(input_file)
        
        # 원본 데이터 분석 추가
        analyze_original_data(df)
        
        # 8% 환급 검토 시트는 원본 데이터로 생성
        create_eight_percent_refund_sheet(df, None, None)
        
        # 데이터 전처리
        print("\n2. 데이터 전처리 시작...")
        df_clean = process_data(df)
        
        if df_clean is None:
            print("데이터 전처리에 실패했습니다. 필요한 컬럼이 누락되었을 수 있습니다.")
            return False
        
        print(f"- 전처리 완료: {len(df_clean):,}건")
        
        # 결과 파일 저장 경로 선택
        print("\n3. 결과 파일 저장...")
        excel_path, word_path = save_files_dialog()
        
        if not excel_path or not word_path:
            print("파일 저장이 취소되었습니다.")
            return False
            
        # 엑셀 파일 생성
        print(f"\n4. 엑셀 파일 생성 중... ({excel_path})")
        try:
            with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                # 8% 환급 검토 시트 먼저 생성 (원본 데이터 사용)
                create_eight_percent_refund_sheet(df, writer, None)
                
                # 나머지 시트는 전처리된 데이터로 생성
                df_clean = process_data(df)
                
                # Summary 시트 생성 (원본 데이터 전달)
                create_summary_sheet(df_clean, df, writer)
                
                # 0% Risk 시트 생성
                create_zero_percent_risk_sheet(df_clean, writer)
                
                # 세율 Risk 분석 (전체 원본 데이터 사용)
                worksheet_tariff = writer.book.add_worksheet('세율 Risk')
                doc = Document()
                add_tariff_risk_analysis(df, worksheet_tariff, writer, doc, 0)
                
                # 단가 Risk 분석 (전체 원본 데이터 사용)
                worksheet_price = writer.book.add_worksheet('단가 Risk')
                add_standard_price_analysis(df, worksheet_price, writer, doc, 0)
                
                # 검증방법 시트 생성
                create_verification_methods_sheet(writer)
                
                # 원본 데이터 시트 생성
                print("- 원본 데이터 시트 생성 중...")
                worksheet_original = writer.book.add_worksheet('원본데이터')
                
                # 헤더 포맷 설정
                header_format = writer.book.add_format({
                    'bold': True,
                    'bg_color': '#D9E1F2',
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter'
                })
                
                # 데이터 포맷 설정
                data_format = writer.book.add_format({
                    'border': 1,
                    'align': 'left',
                    'valign': 'vcenter'
                })
                
                # 헤더 작성
                for col, header in enumerate(df.columns):
                    worksheet_original.write(0, col, header, header_format)
                
                # 데이터 작성 (상위 1000개 행만)
                max_rows = min(1000, len(df))
                for row in range(max_rows):
                    for col in range(len(df.columns)):
                        value = df.iloc[row, col]
                        # NaN 값 처리
                        if pd.isna(value):
                            worksheet_original.write(row + 1, col, '', data_format)
                        else:
                            worksheet_original.write(row + 1, col, value, data_format)
                
                # 컬럼 너비 자동 조정
                for col, header in enumerate(df.columns):
                    max_length = max(
                        len(str(header)),
                        df[header].astype(str).str.len().max() if max_rows > 0 else len(str(header))
                    )
                    worksheet_original.set_column(col, col, min(max_length + 2, 50))
                
                # 필터 추가
                worksheet_original.autofilter(0, 0, max_rows, len(df.columns) - 1)
                
                # 창 틀 고정
                worksheet_original.freeze_panes(1, 0)
                
                # 인쇄 설정
                worksheet_original.set_landscape()
                worksheet_original.fit_to_pages(1, 0)
                worksheet_original.set_header('&C&B원본데이터')
                worksheet_original.set_footer('&R&P / &N')
                
                print(f"- 원본 데이터 시트 생성 완료: {max_rows}행, {len(df.columns)}컬럼")
            
            print("\n- 엑셀 파일 저장 완료")
            messagebox.showinfo("완료", f"엑셀 파일이 생성되었습니다.\n경로: {excel_path}")
            
        except Exception as e:
            print(f"엑셀 파일 생성 중 오류 발생: {str(e)}")
            print("\n상세 오류 정보:")
            print(traceback.format_exc())
            return False
        
        # 워드 문서 생성
        print("\n5. 워드 문서 생성 중...")
        try:
            # 제목 추가
            doc.add_heading('수입신고 분석 보고서', 0)
            
            # 날짜 추가
            doc.add_paragraph(datetime.datetime.now().strftime("%Y년 %m월 %d일"))
            
            # 분석 결과 추가
            doc.add_heading('1. 세율 Risk 분석', level=1)
            add_tariff_risk_analysis(df, None, None, doc, 0)
            
            doc.add_heading('2. 단가 Risk 분석', level=1)
            add_standard_price_analysis(df, None, None, doc, 0)
            
            # 워드 문서 저장
            doc.save(word_path)
            print("- 워드 문서 저장 완료")
        except Exception as e:
            print(f"워드 문서 생성 중 오류 발생: {str(e)}")
            print("\n상세 오류 정보:")
            print(traceback.format_exc())
            return False
            
        print("\n=== 분석 완료 ===")
        print("결과 파일이 저장되었습니다:")
        print(f"- Excel: {excel_path}")
        print(f"- Word: {word_path}")
        
        messagebox.showinfo("완료", "분석이 완료되었습니다.")
        return True
        
    except Exception as e:
        print(f"\n처리 중 오류 발생: {str(e)}")
        print("\n상세 오류 정보:")
        print(traceback.format_exc())
        messagebox.showerror("오류", f"처리 중 오류가 발생했습니다:\n{str(e)}")
        return False

if __name__ == "__main__":
    try:
        if not main():
            sys.exit(1)
    except Exception as e:
        print(f"\n예기치 않은 오류: {str(e)}")
        print(traceback.format_exc())
        messagebox.showerror("오류", f"예기치 않은 오류가 발생했습니다:\n{str(e)}")
        sys.exit(1)
