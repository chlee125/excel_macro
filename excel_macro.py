import openpyxl

# 엑셀파일 열기 #파일명 변경해줘
wb = openpyxl.load_workbook('sample.xlsx')
ws = wb['Sheet1']

#Name BC-  넘버
start_num = int(input("BC넘버 입력해라 : "))
#타입입력
input_type = str(input("Type 입력 해라 : "))
#J1에다가 결과값 붙여 넣기
start_value = 1
#마지막 번호
end_value = int(input("마지막 번호 입력 해라 : "))
#Cover 시작 번호
start_cover = 1

for i in range(1, end_value*3 ,3):
    ws['A' + str(i)] = '** Name: BC-' + str(start_num) + ' Type: ' + str(input_type)
    ws['A' + str(i + 1)] = '*Boundary'
    ws['A' + str(i + 2)] = 'cover'+ str(start_cover) + ', 11, 11, ' + str(ws['J' + str(start_value)].value) + '.'
    print(ws['A' + str(i)].value)
    print(ws['A' + str(i+1)].value)
    print(ws['A' + str(i+2)].value)
    start_num += 1
    start_value += 1
    start_cover += 1

# 엑셀 파일 저장
save_name = input("저장할 파일명 입력 해라 : ")
wb.save(save_name+".xlsx")
wb.close()
a = input("파일 확인해 보거라\n아무키나 누르면 종료가 된다.")
exit()