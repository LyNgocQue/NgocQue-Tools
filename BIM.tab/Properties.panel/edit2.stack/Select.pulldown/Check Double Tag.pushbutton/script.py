# -*- coding: utf-8 -*-
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from System.Collections.Generic import List  
import collections
from rpw.ui.forms import Alert

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

def get_selected_elements():
	selection = uidoc.Selection
	selection_ids = selection.GetElementIds()
	elements = [doc.GetElement(id) for id in selection_ids]
	return elements

Ele = get_selected_elements()
ListDuplicate = []
select = uidoc.Selection
NewList = []

t = Transaction(doc, "Check double tag")
t.Start()

try:
	for i in Ele:
		if i.Category.Name != "Room Tags":
			for host in i.GetTaggedLocalElementIds():
				NewList.append(host)
		else:
			host = i.TaggedLocalRoomId
			# Chỉ thêm vào danh sách nếu ID hợp lệ
			if host != ElementId.InvalidElementId:
				NewList.append(host)

	# Tìm các ID trùng lặp
	ListDuplicate = [item for item, count in collections.Counter(NewList).items() if count > 1]
	
	if ListDuplicate:  # Nếu có ID trùng lặp
		Icollection = List[ElementId](ListDuplicate)
		select.SetElementIds(Icollection)  # Chọn các phần tử trùng lặp
	else:
		Alert("No duplicate tags found.")
except Exception as e:
	Alert(str(e))  # Hiển thị thông báo lỗi nếu có
finally:
	t.Commit()  # Đảm bảo commit transaction

