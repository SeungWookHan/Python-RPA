from openpyxl import load_workbook
wb = load_workbook("sample.xlsx")
ws = wb.active

# # ws.delete_rows(8)  # 8번째 줄 즉 7번 학생 삭제
# ws.delete_rows(8, 3)  # 8번째 줄 즉 7번 학생부터 3명 7, 8, 9 삭제

# wb.save("sample_delete.xlsx")

# ws.delete_cols(2)  # 2번째 열 즉 B 전체 삭제
ws.delete_cols(2, 2)  # 2번째 열 즉 B부터 총 2개 B C 열 전부 삭제
wb.save("sample_delete_cols.xlsx")
