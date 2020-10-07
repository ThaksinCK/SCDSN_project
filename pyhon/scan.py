from openpyxl import Workbook
import openpyxl as op

def enter_score():
    file = ('C:\\Users\\SCDSN0\\PycharmProjects\\SCDSNproject\\Excle\\คะแนนความสะอาด\\ภาคเรียนที่1\\Term1.xlsx')
    wb = op.load_workbook(file)
    ws = wb.active

    for row in ws.iter_rows(min_row=2,min_col=1,max_col=1,max_row=15):

        #เรียกใช้Function iter โดยกำหนดจุดเริ่มต้นที่ (2,1)(minrow2,mincol1)
        #กำหนดจุดเสิ้นสุพที่่ (15,1)(maxrow15,maxcol1)
        #มาเก็บไว้ในตัวแปล row

        for cell in row: #นำค่า row มาใส่ใน cell เพื่อใช้ในการึำนวณ
            print(cell.value) #ดึงค่าทีี่ในcellออกมา

    return cell.value #คืนค่ากลับไปเพื่อใช้คำนวณที่อื่น

print('Done!')
