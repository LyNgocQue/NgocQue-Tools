# -*- coding: utf-8 -*-
import clr
from Autodesk.Revit.DB import *
from pyrevit import revit, DB, script, forms
output = script.get_output()
output.close_others()
__doc__ = 'LNQ'
__title__ = 'Find Legends, Title Blocks'

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

selection = uidoc.Selection.GetElementIds()
for i in selection:
	ele = doc.GetElement(i)
	category = (ele.Category).Name
	if category == "Title Blocks":
		id_type =  (ele.GetTypeId()).IntegerValue
		if id_type:
			list_titleblock = []
			title_blocks = FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()
			for j in title_blocks:
				id_type_all =  (j.GetTypeId()).IntegerValue
				if id_type_all == id_type:
					title_block_name = j.Name
					owner_Viewid = j.OwnerViewId
					sheet_name = (doc.GetElement(owner_Viewid)).Name
					sheet_number = (doc.GetElement(owner_Viewid)).SheetNumber
					list_titleblock.append((output.linkify(j.Id), title_block_name, output.linkify(owner_Viewid), sheet_number, sheet_name))
		if len(list_titleblock) != 0:
			list_titleblock = sorted(list_titleblock, key=lambda x: (x[1], x[3]))
			output.print_md("## Title Blocks on Sheets")
			headers = ["Title Blocks selector", "Title Blocks Name",
					"Sheet selector", "Sheet Number", "Sheet Name"]
			output.print_table(list_titleblock, headers)
	if category == "Viewports":
		viewid = ele.ViewId
		sheets =DB. FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()
		results = []
		list_legend =[]
		for sheet in sheets:
			vps = sheet.GetAllPlacedViews()
			for vp in vps:
				# print(vp)
				if viewid == vp:
					list_legend.append(vp.IntegerValue)
					results.append((output.linkify(doc.GetElement(vp).Id), doc.GetElement(
						vp).Name, output.linkify(sheet.Id), sheet.SheetNumber, sheet.Name))
		if len(results) != 0:
			results = sorted(results, key=lambda x: (x[1], x[3]))
			output.print_md("## Legends on Sheets")
			headers = ["Legend selector", "Legend Name",
					"Sheet selector", "Sheet Number", "Sheet Name"]
			output.print_table(results, headers)
		else:
			forms.alert("No legends found on sheets.")