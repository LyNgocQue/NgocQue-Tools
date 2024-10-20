# -*- coding: utf-8 -*-
import sys
from Autodesk.Revit.DB import *
from pyrevit.forms import select_views, select_sheets
from pyrevit import forms
from pyrevit import revit, DB, UI
from System.Collections.Generic import List
from System.Windows.Forms import MessageBox

__doc__ = 'LNQ'
__title__ = 'Select 2D On Sheets'

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = doc.ActiveView

elementToCopy = uidoc.Selection.GetElementIds()
if not elementToCopy:
	MessageBox.Show("Chọn đối tượng trước rồi mới click vào Tool", "LNQ-Tools")
	sys.exit()
selected_sheets = select_sheets(__title__, multiple=True, button_name='Select Sheets')
if not selected_sheets:
	sys.exit()
elements_to_select = []  # Danh sách để lưu các phần tử sẽ được chọn
for j in elementToCopy:
	element_copy = doc.GetElement(j)
	type_element_copy = element_copy.GetTypeId().IntegerValue
	category = element_copy.Category.Name
	if category == "Viewports":
		viewid = element_copy.ViewId
		for sht in selected_sheets:
			existing_vps = FilteredElementCollector(doc) \
				.OfCategory(BuiltInCategory.OST_Viewports) \
				.OwnedByView(sht.Id) \
				.ToElements()

			for k in existing_vps:
				viewid_on_sheet = k.ViewId
				if viewid_on_sheet == viewid :
					elements_to_select.append(k.Id)  # Thêm ID của các viewport vào danh sách
	elif category == "Title Blocks":
		viewid = element_copy.GetTypeId()
		for sht in selected_sheets:
			existing_vps = FilteredElementCollector(doc) \
				.OfCategory(BuiltInCategory.OST_TitleBlocks) \
				.OwnedByView(sht.Id) \
				.ToElements()

			for k in existing_vps:
				viewid_on_sheet = k.GetTypeId()
				if viewid_on_sheet == viewid :
					elements_to_select.append(k.Id)  # Thêm ID của các viewport vào danh sách
	elif category == "Generic Annotations":
		viewid = element_copy.GetTypeId()
		for sht in selected_sheets:
			existing_vps = FilteredElementCollector(doc) \
				.OfCategory(BuiltInCategory.OST_GenericAnnotation) \
				.OwnedByView(sht.Id) \
				.ToElements()

			for k in existing_vps:
				viewid_on_sheet = k.GetTypeId()
				if viewid_on_sheet == viewid :
					elements_to_select.append(k.Id)  # Thêm ID của các viewport vào danh sách

# Chọn tất cả các phần tử đã thu thập
if elements_to_select:
	# Chuyển đổi danh sách thành ICollection[ElementId]
	element_ids = List[ElementId](elements_to_select)
	uidoc.Selection.SetElementIds(element_ids)
else:
	forms.alert("Không tìm thấy phần tử nào trên các sheet đã chọn.")