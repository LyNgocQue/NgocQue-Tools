# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from System.Collections.Generic import *
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
active_view = uidoc.ActiveView

selection = uidoc.Selection.GetElementIds()

# Lưu tên gia đình đã chọn
selected_family = None
if selection:
	ele = doc.GetElement(selection[0])  # Lấy phần tử đầu tiên trong selection
	family_param = ele.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)
	if family_param:
		selected_family = family_param.AsValueString()  # Lưu tên gia đình đã chọn
		# print(f"Gia đình đã chọn: {selected_family}")

# Tạo set để lưu tên gia đình và danh sách các phần tử cùng family
family_names = set()
elements_to_select = []  # Danh sách để lưu ID của các phần tử sẽ chọn

# Lặp qua tất cả các phần tử trong view hiện tại
active_view = uidoc.ActiveView
collector = FilteredElementCollector(doc, active_view.Id)
elements = collector.WhereElementIsNotElementType().ToElements()

for ele in elements:
	# Lấy thông tin gia đình
	family_param = ele.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)
	if family_param:
		family_name = family_param.AsValueString()
		if family_name:
			family_names.add(family_name)  # Thêm tên gia đình vào set
			# Kiểm tra xem tên gia đình có giống với tên đã chọn không
			if family_name == selected_family:
				elements_to_select.append(ele.Id)  # Lưu ID của phần tử cùng family

# Chọn tất cả các phần tử cùng family
if elements_to_select:
	# Chuyển danh sách thành ICollection[ElementId]
	uidoc.Selection.SetElementIds(List[ElementId](elements_to_select))
