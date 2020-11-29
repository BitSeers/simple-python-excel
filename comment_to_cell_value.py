# consult https://openpyxl.readthedocs.io/en/stable/ for more info
from openpyxl import load_workbook

# open excel file
wb = load_workbook(filename='Petroleum_Upstream_v721_vs_Petroleum_Upstream_v610.xlsx')

# list of worksheets to process
ws_list = [f'{i}.0' for i in range(1, 14)]

# get a worksheet
ws = wb[ws_list[0]]

comment_column_letter = 'C'
desc_column_letter = 'D'
comment_counter = 0

for ws_name in ws_list:
    ws = wb[ws_name]
    for t in range(1, 1000):
        read_cell_name = comment_column_letter + str(t)
        # print(ws[ref].value)
        comment = ws[read_cell_name].comment
        if comment:
            comment_counter += 1
            # print(f"{read_cell_name}: {comment.text}")
            write_cell_name = desc_column_letter + str(t)
            ws[write_cell_name].value = comment.text

wb.save(filename='ExcelData01_Test.xlsx')
wb.close()
print(f"Successfully processed {comment_counter} comments.")
