
from calendar import day_abbr
from sqlite3 import Row
from tkinter.tix import ROW
from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
doc = Document('/Users/rustammirzoev/Desktop/Rustam_super.docx')


list_ = [
    {
        "type": "filled",
        "data": {

            'name': "Наименование работ",
            'units': 'Единица измерения',
            "count": "100",
            "price": "100",
            "total": "1000",
        }
    }
]


table1 = doc.tables[0]

name = doc.tables[0].cell(2, 1).text = list_[0]['data'].get('name')
units = doc.tables[0].cell(2, 2).text = list_[0]['data'].get('units')
count = doc.tables[0].cell(2, 3).text = list_[0]['data'].get('count')
price = doc.tables[0].cell(2, 4).text = list_[0]['data'].get('price')
total = doc.tables[0].cell(2, 5).text = list_[0]['data'].get('total')
    

for paper in list_[0]['data'].values():
    print(paper)
    cells = table1.add_row().cells
    for qwas in ROW:
        print(qwas)
        

    

doc.save('demo2.docx')














list2=[

    {
        "type": "subtitle",
        "data": {
            "title": "Подзаголовок",
        }
    },

]

list3=[
    {
        "type": "total",
        "data": {
            "value": "2324 р."
        }
    }
]