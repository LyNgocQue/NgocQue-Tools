
# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from rpw.ui.forms import TextInput, Alert # import giao diện

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
t = Transaction(doc, 'Đánh sheet number')
t.Start()
selection = uidoc.Selection.GetElementIds() # chọn đối tượng 
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox,
							Separator, Button, CheckBox)
components = [Label('ĐÁNH SHEET NUMBER:'),
				Label('Tiền tố:'),
				TextBox('textbox1', Text="KT-"),
				Label('Bắt đầu từ số:'),
				TextBox('textbox2', Text="1"),
				Label('kiểu seri:'),
				ComboBox('combobox1', [1,10,100, 1000,10000]),
				Label('Hậu tố:'),
				TextBox('textbox3', ),
				Separator(),
				Button('Run')]
form = FlexForm('LNQ tools', components)
form.show()
form.values
ky_tu_ban_ve = str(form.values["textbox1"])
so_bat_dau = int(form.values["textbox2"])
Hau_to = str(form.values["textbox3"])
kieu_seri= int(form.values["combobox1"])
list_sheet = []
for i in selection:
	ele = doc.GetElement(i)
	if isinstance(ele, ViewSheet):
		sheet_name = ele.LookupParameter("Sheet Name").AsString()
		sheet_number = ele.LookupParameter("Sheet Number").AsString()
		sheet_number_para = ele.LookupParameter("Sheet Number")
		# print(sheet_number)
		list_sheet.append(sheet_number)

if kieu_seri == 1:
	count = so_bat_dau
	sorted_list_sheet = sorted(list_sheet)
	for sheet_number in sorted_list_sheet:
		sheet_number_formatted = str(count).zfill(1)  # Định dạng số đếm thành chuỗi có độ dài 1 ký tự, bằng cách thêm số 0 vào đầu nếu cần
		print(sheet_number_formatted, "-", sheet_number)
		sheet_number_new = str(ky_tu_ban_ve + str(count).zfill(1) + Hau_to)
		for i in selection:
			ele = doc.GetElement(i)
			if isinstance(ele, ViewSheet) and ele.LookupParameter("Sheet Number").AsString() == sheet_number:
				ele.LookupParameter("Sheet Number").Set(sheet_number_new)
		count += 1
elif kieu_seri == 10:
	count = so_bat_dau
	sorted_list_sheet = sorted(list_sheet)
	for sheet_number in sorted_list_sheet:
		sheet_number_formatted = str(count).zfill(2)  # Định dạng số đếm thành chuỗi có độ dài 2 ký tự, bằng cách thêm số 0 vào đầu nếu cần
		print(sheet_number_formatted, "-", sheet_number)
		sheet_number_new = str(ky_tu_ban_ve + str(count).zfill(2) + Hau_to)
		for i in selection:
			ele = doc.GetElement(i)
			if isinstance(ele, ViewSheet) and ele.LookupParameter("Sheet Number").AsString() == sheet_number:
				ele.LookupParameter("Sheet Number").Set(sheet_number_new)
		count += 1
elif kieu_seri == 100:
	count = so_bat_dau
	sorted_list_sheet = sorted(list_sheet)
	for sheet_number in sorted_list_sheet:
		sheet_number_formatted = str(count).zfill(3)  # Định dạng số đếm thành chuỗi có độ dài 3 ký tự, bằng cách thêm số 0 vào đầu nếu cần
		print(sheet_number_formatted, "-", sheet_number)
		sheet_number_new = str(ky_tu_ban_ve + str(count).zfill(3) + Hau_to)
		for i in selection:
			ele = doc.GetElement(i)
			if isinstance(ele, ViewSheet) and ele.LookupParameter("Sheet Number").AsString() == sheet_number:
				ele.LookupParameter("Sheet Number").Set(sheet_number_new)
		count += 1
elif kieu_seri == 1000:
	count = so_bat_dau
	sorted_list_sheet = sorted(list_sheet)
	for sheet_number in sorted_list_sheet:
		sheet_number_formatted = str(count).zfill(4)  # Định dạng số đếm thành chuỗi có độ dài 4 ký tự, bằng cách thêm số 0 vào đầu nếu cần
		print(sheet_number_formatted, "-", sheet_number)
		sheet_number_new = str(ky_tu_ban_ve + str(count).zfill(4) + Hau_to)
		for i in selection:
			ele = doc.GetElement(i)
			if isinstance(ele, ViewSheet) and ele.LookupParameter("Sheet Number").AsString() == sheet_number:
				ele.LookupParameter("Sheet Number").Set(sheet_number_new)
		count += 1
else:
	if kieu_seri == 10000:
		count = so_bat_dau
		sorted_list_sheet = sorted(list_sheet)
		for sheet_number in sorted_list_sheet:
			sheet_number_formatted = str(count).zfill(5)  # Định dạng số đếm thành chuỗi có độ dài 3 ký tự, bằng cách thêm số 0 vào đầu nếu cần
			print(sheet_number_formatted, "-", sheet_number)
			sheet_number_new = str(ky_tu_ban_ve + str(count).zfill(5) + Hau_to)
			for i in selection:
				ele = doc.GetElement(i)
				if isinstance(ele, ViewSheet) and ele.LookupParameter("Sheet Number").AsString() == sheet_number:
					ele.LookupParameter("Sheet Number").Set(sheet_number_new)
			count += 1
t.Commit()


