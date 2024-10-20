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
		{'All': 'Dưới Trên Trái Phải'.split(),},
		title='Chọn Hướng Để Tắt',
		multiselect=True
		
	)
try:
	if not selected:
		for v in FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements():
			res_end0 = v.ShowBubbleInView(DatumEnds.End0, view)
			res_end1 = v.ShowBubbleInView(DatumEnds.End1, view)
		# forms.alert('Chưa chọn Hiển thị đầu grid. Vui Lòng Thử Lại.', exitscript=True)
	all_grids = FilteredElementCollector(doc, view.Id) \
				.OfCategory(BuiltInCategory.OST_Grids) \
				.WhereElementIsNotElementType() \
				.ToElements()
	grid_y = []
	grid_x = []
	for i in all_grids:
		vector = i.Curve.Direction
		if abs(vector.Y) == 1 and abs(vector.X) < 0.1:
			grid_y.append(i)
		else:
			if abs(vector.X) == 1 and abs(vector.Y) < 0.1:
				grid_x.append(i)
	for j in selected:
		if j == "Dưới":
			for k in grid_y:
				k.HideBubbleInView(DatumEnds.End0, view)
		elif j == "Trên":
			for k in grid_y:
				k.HideBubbleInView(DatumEnds.End1, view)
		elif j == "Trái":
			for k in grid_x:
				k.HideBubbleInView(DatumEnds.End0, view)
		else:
			for k in grid_x:
				k.HideBubbleInView(DatumEnds.End1, view)
except Exception as e:
	pass
t.Commit()