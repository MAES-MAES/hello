
from calendar import day_abbr
from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
doc = Document('/Users/rustammirzoev/Desktop/Rustam_super.docx')


list_ = [
    {
        "type": "filled",
        "data": {

            'name': "Наименование работ",'units': 'Единица измерения',"count": "100","price": "100","total": "1000",
        }
    },
]


table1 = doc.tables[0]

for datas in list_[0]['data']:
    
    cells = table1.add_row().cells
   
    print(datas)
    
    cells[1].text = datas
   







    

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