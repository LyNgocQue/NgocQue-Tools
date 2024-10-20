# -*- coding: utf-8 -*-
import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox

# Lấy view hiện tại
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
active_view = uidoc.ActiveView
try:
	# Chọn phần tử nguồn
	select1 = uidoc.Selection.PickObject(ObjectType.Element, "Chọn phần tử nguồn")
	element1_id = select1.ElementId

	# Lấy phần tử nguồn
	element1 = doc.GetElement(element1_id)
	MessageBox.Show('Chọn đối tượng để gán thuộc tính...','LNQ-TOOL')
	# Chọn nhiều phần tử đích
	select2 = uidoc.Selection.PickObjects(ObjectType.Element, "Chọn các phần tử đích")
		
	# Lấy OverrideGraphicSettings từ phần tử nguồn trong view hiện tại
	override_settings = active_view.GetElementOverrides(element1_id)

	# Bắt đầu giao dịch để cập nhật thuộc tính
	t = Transaction(doc, "Copy Override Graphics")
	t.Start()

	# Lấy view chứa phần tử đích
	target_view = uidoc.ActiveView  # Mặc định là view hiện tại

	# Áp dụng OverrideGraphicSettings cho tất cả các phần tử đích đã chọn
	if target_view:
		for target in select2:
			element2_id = target.ElementId
			# Áp dụng cài đặt cho từng phần tử đích
			target_view.SetElementOverrides(element2_id, override_settings)
	t.Commit()
except:
	pass
