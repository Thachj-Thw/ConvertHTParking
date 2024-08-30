from openpyxl import load_workbook

START_USER = 5
MAX_ROW = 10000
ROOM = "C"
CARD_NO = "W"

def read_user(sheet):
    result = {}
    for row in range(START_USER, MAX_ROW):
        str_row = str(row)
        card_no = sheet[CARD_NO + str_row].value
        room = sheet[ROOM + str_row].value
        if not card_no and not room:
            break
        result[card_no] = room
    return result

wb = load_workbook("Output/Dang ky khach hang.xlsx")
sheet = wb.active

data = read_user(sheet)

with open("query.sql", "w", encoding="utf-8") as f:
    query = ""
    for card_no in data.keys():
        query += "UPDATE card2 SET note = '%s' WHERE code = '%s';\n" % (data[card_no], card_no)
    f.write(query)
