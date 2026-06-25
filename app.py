import streamlit as st
import pandas as pd
import pdfplumber
import re

# 화면 넓게 쓰기
st.set_page_config(layout="wide")

st.title("📋 견적의뢰서 vs 견적서 3중 검증 시스템 (Excel & PDF 지원)")
st.write("의뢰하신 내용과 받은 견적서(Excel/PDF)가 일치하는지 세 번씩 철저하게 검토합니다.")
st.divider()

# 1. 파일 업로드 상자 (이제 PDF도 올릴 수 있게 확장했어요!)
col1, col2 = st.columns(2)
with col1:
    st.subheader("1️⃣ 내가 보낸 견적의뢰서 (RFQ)")
    rfq_file = st.file_uploader("의뢰서 엑셀 파일을 올려주세요.", type=["xlsx"], key="rfq")

with col2:
    st.subheader("2️⃣ 업체에서 받은 견적서 (Quotation)")
    quotation_file = st.file_uploader("견적서 파일(Excel 또는 PDF)을 올려주세요.", type=["xlsx", "pdf"], key="quote")

st.divider()

# PDF에서 텍스트를 파싱하여 데이터프레임으로 만드는 마법의 함수
def parse_pdf_quotation(file):
    data = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            # 페이지에서 표(Table) 구조가 있다면 먼저 추출 시도
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table:
                        # 빈 줄이나 제목 줄 제외 가공
                        filtered_row = [cell.strip() if cell else "" for cell in row]
                        if filtered_row and any(filtered_row):
                            data.append(filtered_row)
            else:
                # 표 형태가 흐릿할 경우 글자를 통째로 읽어서 줄바꿈 단위로 분석
                text = page.extract_text()
                if text:
                    for line in text.split("\n"):
                        # 공백으로 글자들을 쪼개기
                        row = line.split()
                        if len(row) >= 3:
                            data.append(row)
    
    # 임시 표 만들기
    df_raw = pd.DataFrame(data)
    
    # 💡 인공지능 규칙: 행에서 '품명', '수량', '단가'와 매칭되는 열 찾기
    # PDF 마다 양식이 다르므로 컴퓨터가 글자를 보고 똑똑하게 유추합니다.
    refined_data = []
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(c) for c in row])
        
        # 숫자가 포함된 행만 품목으로 인정 (수량, 단가가 있어야 하므로)
        numbers = re.findall(r'\d[\d,]*', row_str)
        if len(numbers) >= 2:
            # 가장 그럴싸한 규칙으로 품명, 수량, 단가 추출 시도
            # 보통 줄에서 글자가 먼저 나오고 숫자가 뒤에 나옵니다.
            words = [str(c) for c in row if c and not re.match(r'^\d[\d,]*$', str(c))]
            p_name = words[0] if words else "미확인 품목"
            
            # 숫자 중 100미만 혹은 작은 숫자를 수량으로, 큰 숫자를 단가로 유추하는 똑똑한 로직
            clean_nums = [int(re.sub(r'[^0-9]', '', n)) for n in numbers if re.sub(r'[^0-9]', '', n)]
            if len(clean_nums) >= 2:
                # 정렬해서 작은 건 수량, 큰 건 단가로 매칭 (일반적인 견적서 특징 활용)
                qty = min(clean_nums[0], clean_nums[1])
                price = max(clean_nums[0], clean_nums[1])
                refined_data.append({'품명': p_name, '수량': qty, '단가': price})
                
    return pd.DataFrame(refined_data)

# 두 파일이 모두 올라왔을 때 작동
if rfq_file is not None and quotation_file is not None:
    if st.button("🔍 3중 매칭 및 금액 검증 시작", type="primary"):
        with st.spinner("컴퓨터가 PDF와 엑셀을 동시에 분석하고 있습니다..."):
            try:
                # 1. 의뢰서(엑셀) 읽기
                df_rfq = pd.read_excel(rfq_file)
                df_rfq.columns = df_rfq.columns.str.strip()
                
                # 2. 받은 견적서(엑셀 또는 PDF) 읽기
                if quotation_file.name.endswith('.pdf'):
                    df_quote = parse_pdf_quotation(quotation_file)
                else:
                    df_quote = pd.read_excel(quotation_file)
                    df_quote.columns = df_quote.columns.str.strip()
                
                # 필수 열 검사 (의뢰서 기준)
                if '품명' not in df_rfq.columns or '수량' not in df_rfq.columns or '단가' not in df_rfq.columns:
                    st.error("❌ 견적의뢰서(Excel) 첫 줄에 '품명', '수량', '단가'라는 열 이름이 정확히 있어야 합니다.")
                elif df_quote.empty or '품명' not in df_quote.columns:
                    st.error("❌ PDF 견적서에서 '품명', '수량', '단가' 데이터를 자동으로 추출하지 못했습니다. PDF 양식이 너무 독특하거나 이미지 스캔본일 수 있습니다.")
                else:
                    # 데이터 정제
                    df_rfq_clean = df_rfq[['품명', '수량', '단가']].copy()
                    df_quote_clean = df_quote[['품명', '수량', '단가']].copy()
                    
                    # 짝꿍 매칭 (Outer Join)
                    merged = pd.merge(df_rfq_clean, df_quote_clean, on='품명', how='outer', suffixes=('_의뢰', '_견적'))
                    
                    results = []
                    for idx, row in merged.iterrows():
                        status = "✅ 일치"
                        reason = "의뢰 내용과 정확히 일치합니다."
                        
                        if pd.isna(row['수량_의뢰']):
                            status = "❌ 의뢰외 품목"
                            reason = "의뢰서에는 없는 품목이 견적서에 추가되었습니다."
                        elif pd.isna(row['수량_견적']):
                            status = "❌ 품목 누락"
                            reason = "의뢰서에 있는 품목이 견적서에서 빠졌습니다."
                        else:
                            qty_match = float(row['수량_의뢰']) == float(row['수량_견적'])
                            price_match = float(row['단가_의뢰']) == float(row['단가_견적'])
                            
                            if not qty_match and not price_match:
                                status = "❌ 수량/단가 불일치"
                                reason = f"수량 불일치({row['수량_의뢰']}≠{row['수량_견적']}) 및 단가 불일치"
                            elif not qty_match:
                                status = "⚠️ 수량 불일치"
                                reason = f"의뢰 {row['수량_의뢰']}개 ➡️ 견적 {row['수량_견적']}개로 다릅니다."
                            elif not price_match:
                                status = "⚠️ 단가 불일치"
                                reason = f"의뢰 단가 {row['단가_의뢰']:,}원 ➡️ 견적 단가 {row['단가_견적']:,}원"
                        
                        results.append({
                            "품명": row['품명'],
                            "의뢰수량": row['수량_의뢰'] if not pd.isna(row['수량_의뢰']) else 0,
                            "견적수량": row['수량_견적'] if not pd.isna(row['수량_견적']) else 0,
                            "의뢰단가": row['단가_의뢰'] if not pd.isna(row['단가_의뢰']) else 0,
                            "견적단가": row['단가_견적'] if not pd.isna(row['단가_견적']) else 0,
                            "검증결과": status,
                            "상세이유": reason
                        })
                    
                    df_result = pd.DataFrame(results)
                    
                    # 틀린 부분에 색 칠하기 규칙
                    def color_rows(val):
                        if "❌" in val:
                            return 'background-color: #ffcccc; color: #cc0000; font-weight: bold;'
                        elif "⚠️" in val:
                            return 'background-color: #fff3cd; color: #856404;'
                        return 'background-color: #d4edda; color: #155724;'
                    
                    styled_df = df_result.style.applymap(color_rows, subset=['검증결과'])
                    
                    # 대시보드 상단 요약
                    err_count = sum(df_result['검증결과'].str.contains('❌|⚠️'))
                    if err_count == 0:
                        st.balloons()
                        st.success("🎉 완벽합니다! PDF 견적서의 모든 숫자가 의뢰 내용과 100% 일치합니다.")
                    else:
                        st.error(f"🚨 총 {err_count}개의 품목에서 불일치(오류/누락/추가)가 발견되었습니다. 아래 빨간 형광펜을 확인하세요.")
                    
                    # 결과 화면 출력
                    st.dataframe(styled_df, use_container_width=True)
                    
                    # CSV 다운로드
                    @st.cache_data
                    def convert_df(df):
                        return df.to_csv(index=False).encode('utf-8-sig')
                    csv = convert_df(df_result)
                    st.download_button(
                        label="📥 검증 리포트 다운로드 (CSV)",
                        data=csv,
                        file_name='견적검증결과_리포트.csv',
                        mime='text/csv',
                    )
                    
            except Exception as e:
                st.error(f"파일을 읽는 도중 오류가 발생했습니다: {e}")
else:
    st.info("💡 왼쪽에는 내가 보낸 의뢰서(Excel)를, 오른쪽에는 업체 견적서(Excel 또는 PDF)를 올려주세요.")