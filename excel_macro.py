import openpyxl

# 엑셀파일 열기 #파일명 변경해줘
wb = openpyxl.load_workbook('sample.xlsx')
ws = wb['Sheet1']

#Name BC-  넘버
start_num = 48
#J1에다가 결과값 붙여 넣기
start_value = 1
#마지막 번호
end_value = 185
#Cover 시작 번호
start_cover = 1

for i in range(1, end_value*3 ,3):
    ws['B' + str(i)] = '** Name: BC-' + str(start_num) + ' Type: Temperature'
    ws['B' + str(i + 1)] = '*Boundary'
    ws['B' + str(i + 2)] =  'cover'+ str(start_cover) + ', 11, 11, ' + str(ws['J' + str(start_value)].value)
    print(ws['B' + str(i)].value)
    print(ws['B' + str(i+1)].value)
    print(ws['B' + str(i + 2)].value)
    start_num+=1
    start_value+=1
    start_cover +=1

# 엑셀 파일 저장
wb.save("result.xlsx")
wb.close()