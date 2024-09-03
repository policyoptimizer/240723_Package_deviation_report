# [기존 ppt]
# 부적합 추가

        if '일탈번호' in df.columns:
            columns_to_keep = ['일탈번호', '일탈등급', '제목', 'QA 검토자', '작성일', '일탈기준', '일탈내용', '작업자오류내용']
            iltal_dfs.append(df[columns_to_keep])
        elif '고객불만번호' in df.columns:
            columns_to_keep = ['고객불만번호', '제목', '불만발생일', '조사담당자의견', '고객요구사항']
            gongmun_dfs.append(df[columns_to_keep])

'부적합번호', '제목', 'QA 검토자', '발생부서', '제품', '근본 원인', '부적합품 후속 처리 계획', '조사 세부사항', '사유', '완료 요약', '후속 조치 완료 여부'

# [엑셀 파일]
# 각 유형별로 개별 시트 출력됨
# 3가지 폴더 추가함
# 나머지는 칼럼만 추가
