# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from rpw.ui.forms import TextInput, Alert # import giao diện

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

try:
	active_view = doc.ActiveView 
	# selection = uidoc.Selection.GetElementIds() # chọn đối tượng 
	# for i in selection:  
	ele = active_view
	tableData = ele.GetTableData()
	sectionData = tableData.GetSectionData(SectionType.Body)
	numberOfRows = sectionData.NumberOfRows
	numberOfColumns = sectionData.NumberOfColumns
	from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox,
									Separator, Button, CheckBox)
	components = [Label('ĐÁNH SỐ THỨ TỰ:'),
					Label('Nhập Tên Cột Cần Đánh Số Thứ Tự:'),
					ComboBox('combobox1', [ele.GetCellText(SectionType.Body, 0, i) for i in range(numberOfColumns)]),
					Label('Nhập Tên Cột Sheet number:'),
					ComboBox('combobox2', [ele.GetCellText(SectionType.Body, 0, i) for i in range(numberOfColumns)]),
					Label('Bắt Đầu Từ Số:'),
					TextBox('textbox1', Text="1"),
					Separator(),
					Button('Run')]
	form = FlexForm('LNQ tools', components)
	form.show()
	form.values
	cot_can_nhap_so= str(form.values["combobox1"])
	cot_tham_chieu= str(form.values["combobox2"])
	so_bat_dau = int(form.values["textbox1"])
	# print(cot_can_nhap_so)
	# print(cot_tham_chieu)
	# print(so_bat_dau)
	sheet_stt_dict = {} # tạo từ điển
	count = so_bat_dau  # Biến đếm bắt đầu từ số....
	for i in range(numberOfColumns):
		columnValue = ele.GetCellText(SectionType.Body, 0, i)
		# print(columnValue)

		if columnValue == cot_tham_chieu:
			sheet = FilteredElementCollector(doc).OfClass(ViewSheet).WhereElementIsNotElementType().ToElements() # tìm tất cả sheet trong dự án kể cả file link
			for j in sheet:
				sheet_number_param = j.LookupParameter("Sheet Number")   # lọc ra tên sheet Number
				stt_param = j.LookupParameter(cot_can_nhap_so) 
				sheet_number = sheet_number_param.AsString()
				stt = stt_param.AsInteger()
				sheet_stt_dict[sheet_number] = stt_param  # Lưu trữ tham chiếu đến tham số STT của tờ trong từ điển
				# print(sheet_number)
				# print(stt_param)
			t = Transaction(doc, 'Đánh số thứ tự')
			t.Start()
			for rowIndex in range(1, numberOfRows):
				rowValue = ele.GetCellText(SectionType.Body, rowIndex, i)
				if rowValue != "" :#and rowValue in sheet_stt_dict: # nếu sheet number trong bảng thống kê có trong thư viện của dự án và không đếm dòng trống
					print("Sheet Number:", rowValue, "STT:", count)
					# print(count)
					# print(rowValue)
					if rowValue in sheet_stt_dict:
						sheet_stt_dict[rowValue].Set(count)
					count += 1  # Tăng giá trị của biến đếm sau mỗi lần in ra giá trị
			t.Commit()
except:
	pass

# LẤY TÊN TIÊU ĐỀ CỦA SHEET LIST
	# tableName = ele.GetCellText(SectionType.Header, 0, 0)
	# print(tableName)



# THỂ HIỆN KẾT QUẢ NẾU KHÔNG CÓ LINK FILE
# for rowIndex in range(1, numberOfRows):
# 	rowValue = ele.GetCellText(SectionType.Body, rowIndex, i)
# 	if rowValue != "" and rowValue in sheet_stt_dict: # nếu sheet number trong bảng thống kê có trong thư viện của dự án và không đếm dòng trống
# 		sheet_stt_dict[rowValue].Set(count)
# 		count += 1  # Tăng giá trị của biến đếm sau mỗi lần in ra giá trị


# import clr
# clr.AddReference('RevitAPI')
# from Autodesk.Revit.DB import *
# from rpw.ui.forms import TextInput, Alert # import giao diện

# uidoc = __revit__.ActiveUIDocument
# doc = uidoc.Document

# t = Transaction(doc, 'Đánh số thứ tự')
# t.Start()
# selection = uidoc.Selection.GetElementIds() # chọn đối tượng 


# cot_can_nhap_so = TextInput("Nhập Tên Cột Cần Đánh Số Thứ Tự:", default = "STT") # khai báo nhập số STT và sheet number
# cot_tham_chieu = TextInput("Nhập Tên Cột Sheet number:", default = "KÝ HIỆU BẢN VẼ")
# so_bat_dau_text = TextInput("Bắt Đầu Từ Số:", default = "1")
# so_bat_dau = int(so_bat_dau_text)
# sheet_stt_dict = {} # tạo từ điển


# for i in selection:  
# 	ele = doc.GetElement(i)
# 	tableData = ele.GetTableData()
# 	sectionData = tableData.GetSectionData(SectionType.Body)
# 	numberOfRows = sectionData.NumberOfRows
# 	numberOfColumns = sectionData.NumberOfColumns


# 	sheet_test = {}
# 	count = so_bat_dau  # Biến đếm bắt đầu từ số....
# 	for i in range(numberOfColumns):
# 		columnValue = ele.GetCellText(SectionType.Body, 0, i)
# 		# print(columnValue)

# 		if columnValue == cot_tham_chieu:
# 			sheet = FilteredElementCollector(doc).OfClass(ViewSheet).WhereElementIsNotElementType().ToElements() # tìm tất cả sheet trong dự án
# 			for j in sheet:
# 				sheet_number_param = j.LookupParameter("Sheet Number")   # lọc ra tên sheet Number
# 				stt_param = j.LookupParameter(cot_can_nhap_so)
# 				sheet_number = sheet_number_param.AsString()
# 				stt = stt_param.AsInteger()
# 				sheet_stt_dict[sheet_number] = stt_param  # Lưu trữ tham chiếu đến tham số STT của tờ trong từ điển
# 				# print(sheet_number)
# 				# print(stt_param)

# 			for rowIndex in range(1, numberOfRows):
# 				rowValue = ele.GetCellText(SectionType.Body, rowIndex, i)
# 				if rowValue != "" :#and rowValue in sheet_stt_dict: # nếu sheet number trong bảng thống kê có trong thư viện của dự án và không đếm dòng trống
# 					print("Sheet Number:", rowValue, "STT:", count)
# 					# print(count)
# 					# print(rowValue)
# 					if rowValue in sheet_stt_dict:
# 						sheet_stt_dict[rowValue].Set(count)
# 					count += 1  # Tăng giá trị của biến đếm sau mỗi lần in ra giá trị
# t.Commit()

