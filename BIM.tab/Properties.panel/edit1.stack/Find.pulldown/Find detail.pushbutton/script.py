# -*- coding: utf-8 -*-
import clr
from Autodesk.Revit.DB import *
from pyrevit import revit, DB, script, forms
output = script.get_output()
output.close_others()
__doc__ = 'LNQ'
__title__ = 'Find detail'

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

selection = uidoc.Selection.GetElementIds()
for i in selection:
	ele = doc.GetElement(i)
	id_type =  (ele.GetTypeId()).IntegerValue
	# category = (ele.Category).Name
	# print(category)
	if id_type:
		groups = list(FilteredElementCollector(doc).OfClass(Group))
		for group in groups:
			# print(group)
			results = []
			eleid = (group.GetTypeId()).IntegerValue
			if eleid == id_type:
				id_group = group.Id
				ten = group.Name
				category = (group.Category).Name
				owner_Viewid = (group.OwnerViewId)
				sheet_name_para = doc.GetElement(owner_Viewid).get_Parameter (BuiltInParameter.VIEWPORT_SHEET_NAME).AsString()
				sheet_name = sheet_name_para if sheet_name_para != None else "View chưa được đặt vào sheet"
				sheet_number = doc.GetElement(owner_Viewid).get_Parameter (BuiltInParameter.VIEWER_SHEET_NUMBER).AsValueString()
				if sheet_number:
					sheet = FilteredElementCollector(doc).OfClass(ViewSheet).WhereElementIsNotElementType().ToElements() # tìm tất cả sheet trong dự án 
					for j in sheet:
						sheet_number_param = (j.LookupParameter("Sheet Number")).AsString()  # lọc ra tên sheet Number
						if sheet_number_param == sheet_number:
							# print(sheet_number_param)
							sheet_id = j.Id
							# print(sheet_id)
							results.append((output.linkify(id_group), ten, output.linkify(sheet_id), sheet_number, sheet_name))
	# print(eleid)
			if len(results) != 0:
				results = sorted(results, key=lambda x: (x[1], x[3]))
			# output.print_md("## Legends on Sheets")
				headers = ["Detail selector", "Detail Name",
						"Sheet selector", "Sheet Number", "Sheet Name"]
				output.print_table(results, headers)