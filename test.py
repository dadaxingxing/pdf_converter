from docx import Document
import comtypes.client
import os


doc = Document()

doc.add_heading('fdasjiofdsaoifjdsaiofjdiosafads')

doc.save('test2.doc')


# def convert(path):
#     path = os.path.abspath(path)

#     docx_path = os.path.splitext(path)[0] + '.docx'

#     word = comtypes.client.CreateObject('word.application')
#     word.Visible = False 
#     doc = word.Documents.Open(path)
#     doc.saveas(docx_path, FileFormat=16)
#     doc.Close()
#     word.Quit()
#     os.remove(path)

# convert('test2.doc')