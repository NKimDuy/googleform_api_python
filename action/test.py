from apiclient import discovery
from httplib2 import Http
from oauth2client import client, file, tools
import os
from openpyxl import load_workbook
import time


# Xác định phạm vi của form
SCOPES = "https://www.googleapis.com/auth/forms.body"
DISCOVERY_DOC = "https://forms.googleapis.com/$discovery/rest?version=v1"


# đường dẫn đến các folder
root_folder = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
authentication_dir = os.path.join(root_folder, "authentication")
export_file_dir = os.path.join(root_folder, "export_file")
files_dir = os.path.join(root_folder, "files")
files_subject = os.path.join(files_dir, "241_test.xlsx")


# Đọc file excel
wb = load_workbook(files_subject)
ws = wb.active


for i in range(2, ws.max_row + 1):
    id_semester = "241"
    id_subject = ws.cell(i, 1).value
    name_subject = ws.cell(i, 2).value
    id_class = ws.cell(i, 3).value
    id_group = ws.cell(i, 4).value
    id_link_unit = ws.cell(i, 7).value
    name_link_unit = ws.cell(i, 8).value
    id_techer = ws.cell(i, 5).value
    name_teacher = ws.cell(i, 6).value
    name_manager = ws.cell(i, 9).value
    form = ws.cell(i, 10).value

    if not os.path.exists(os.path.join(export_file_dir, name_manager)):
        os.makedirs(os.path.join(export_file_dir, name_manager))
    link_manager = os.path.join(export_file_dir, name_manager)
    if not os.path.exists(os.path.join(link_manager, " - ".join([id_link_unit, name_link_unit]))):
        os.makedirs(os.path.join(link_manager, " - ".join([id_link_unit, name_link_unit])))
    link_unit = os.path.join(link_manager, " - ".join([id_link_unit, name_link_unit]))
    if not os.path.exists(os.path.join(link_unit, id_class)):
        os.makedirs(os.path.join(link_unit, id_class))
    class_dir = os.path.join(link_unit, id_class)

    name_file = "".join([id_semester, "-", 
                         id_subject, "-", 
                         id_group, "-", 
                         name_subject, "-", 
                         name_teacher,
                         ".txt"])
    with open(os.path.join(class_dir, name_file), "w", encoding="utf-8") as file:
      file.write(form)
    
    print(f" -------------------- Đang tạo form có id là: {i - 1} -----------------------")
    if i == ws.max_row:
       print(f"Đã thêm thành công {i} dòng")
    


    