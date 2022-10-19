import xlwings as xw
from model_bar_info import BarInfoBlock
from pathlib import Path


EXCEL_COLUMNS = {'BAR_MARK':'B', 'BAR_DIA':'C', 'BAR_COUNT':'E', 'A':'K',
'B':'L', 'C':'M','D':'N','E':'O','F':'P','G':'Q','H':'R'}


BAR_INFO_DIMENSIONS = ['A','B','C','D','E','F','G','H']

def main():
    delete_all_images()


def test1():
    sheet = xw.sheets.active
    sheet['A1'].value = 'Hello'

    sheet = None


def delete_all_images():
    sheet = xw.sheets.active
    for pic in sheet.pictures:
        pic.delete()
    sheet = None

def add_bar(bar_info:BarInfoBlock,total_count, row = 17, images_folder = r'C:\BBS_SOURCE\IMG'):
    sheet = xw.sheets.active

    sheet[EXCEL_COLUMNS['BAR_MARK']+str(row)].value = bar_info.attributes['BAR_MARK']
    sheet[EXCEL_COLUMNS['BAR_DIA']+str(row)].value = bar_info.get_bar_dia()
    sheet[EXCEL_COLUMNS['BAR_COUNT']+str(row)].value = total_count

    notes = bar_info.attributes['BAR_LOCATION'] + '\n'
    var = bar_info.attributes['VARIABLES']
    if len(var) > 0 and var != '0':
        notes = notes + var + '\n'
    for i in range(4):
        y = bar_info.attributes['Y' + str(i+1)]
        if len(y) > 0 and y != '0':
            notes = notes + f'Y{i+1}={round(float(y)/1000,2)}, '

    for i in range(3):
        r = bar_info.attributes['R' + str(i+1)]
        if len(r) > 0 and r != '0':
            notes = notes + f'R{i+1}={round(float(r)/1000,2)}, '

    sheet[EXCEL_COLUMNS['BAR_MARK']+str(row+1)].value = notes

    for k,v in bar_info.attributes.items():
        if k in BAR_INFO_DIMENSIONS:
            sheet[EXCEL_COLUMNS[k]+str(row)].value = round(float(v)/1000,2)

    image_file = Path( images_folder ) / bar_info.get_image_filename()
    sheet.pictures.add(image_file,left=sheet[EXCEL_COLUMNS['A']+str(row+1)].left, top=sheet[EXCEL_COLUMNS['A']+str(row+1)].top,height=99)

    sheet = None

if __name__ == '__main__':
    main()