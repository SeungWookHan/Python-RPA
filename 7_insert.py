from openpyxl import load_workbook
wb = load_workbook("sample.xlsx")
ws = wb.active

# ws.insert_rows(8) #8번째 줄이 비워짐 = 엑셀에서 우클릭후 row 삽입
# ws.insert_rows(8, 5)  # 8번째 줄 위치에 5줄을 삽입
wb.save("samle_insert_rows.xlsx")

# ws.insert_cols(2)  # B 번째 열이 추가됨
ws.insert_cols(2, 3)  # B 번째 열로 부터 3칸, B, C, D
wb.save("samle_insert_cols.xlsx")
