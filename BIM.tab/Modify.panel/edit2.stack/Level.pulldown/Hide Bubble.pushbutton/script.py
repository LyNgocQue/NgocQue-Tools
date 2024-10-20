# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from pyrevit import forms
__doc__ = 'LNQ'
__title__ = 'Hide Bubble'


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = doc.ActiveView

t = Transaction(doc, __title__)
t.Start()

selected = forms.SelectFromList.show(
		{'All': 'Trái Phải'.split(),},
		title='Chọn Hướng Để Tắt',
		multiselect=True
		
	)
try:
	if not selected:
		for v in FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements():
			res_end0 = v.ShowBubbleInView(DatumEnds.End0, view)
			res_end1 = v.ShowBubbleInView(DatumEnds.End1, view)
		# forms.alert('Chưa chọn Hiển thị đầu grid. Vui Lòng Thử Lại.', exitscript=True)
	all_levels = FilteredElementCollector(doc, view.Id) \
				.OfCategory(BuiltInCategory.OST_Levels) \
				.WhereElementIsNotElementType() \
				.ToElements()

	for j in selected:
		if j == "Trái":
			for i in all_levels:
				i.HideBubbleInView(DatumEnds.End0, view)
		elif j == "Phải":
			for i in all_levels:
				i.HideBubbleInView(DatumEnds.End1, view)
except Exception as e:
	pass
t.Commit()