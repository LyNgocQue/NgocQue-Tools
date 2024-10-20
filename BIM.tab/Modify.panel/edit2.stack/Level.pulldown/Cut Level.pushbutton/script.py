# -*- coding: utf-8 -*-
import Autodesk
from Autodesk.Revit.DB import *
from rpw.ui.forms import * 
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI import TaskDialog
__doc__ = 'LNQ'
__title__ = 'Cut level'


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = doc.ActiveView

try:
	components = [
					Label('Offset Head With Crop Region:'),
					TextBox('textbox1', Text="8"),
					Label('Offset Tail With Crop Region:'),
					TextBox('textbox2', Text="0"),
					Separator(),
					Button('OK')]
	form = FlexForm('LNQ tools', components)
	form.show()
	form.values
	offset_head = float(form.values["textbox1"])
	offset_tail = float(form.values["textbox2"])


	t = Transaction(doc, __title__)
	t.Start()
	selection = uidoc.Selection.GetElementIds() #Trả về 1 list ElementId
	for i in selection:
		ele = doc.GetElement(i)
		category = ele.Category
		category_name = category.Name
		boundingbox = ele.get_BoundingBox(view)
		min_x = boundingbox.Min.X 
		max_x = boundingbox.Max.X 
		min_y = boundingbox.Min.Y 
		max_y = boundingbox.Max.Y


	all_levels = FilteredElementCollector(doc, view.Id) \
					.OfCategory(BuiltInCategory.OST_Levels) \
					.WhereElementIsNotElementType() \
					.ToElements()
	for ele in all_levels:
		levelCurves = ele.GetCurvesInView(  DatumExtentType.ViewSpecific, view )
		for c in levelCurves:
			start = c.GetEndPoint( 0 )
			end = c.GetEndPoint( 1 )
			if start.X < end.X :
				newStart =  XYZ(min_x - offset_tail, start.Y, start.Z)
				newEnd  =  XYZ(max_x + offset_head, end.Y, end.Z)
			elif start.X > end.X :
				newStart =  XYZ(max_x + offset_tail, start.Y, start.Z)
				newEnd  =  XYZ(min_x - offset_head, end.Y, end.Z)
			elif start.Y > end.Y :
				newStart =  XYZ(start.X , max_y + offset_tail , start.Z)
				newEnd  =  XYZ(end.X, min_y - offset_head, end.Z)
			elif start.Y < end.Y :
				newStart =  XYZ(start.X , min_y - offset_tail, start.Z)
				newEnd  =  XYZ(end.X , max_y + offset_head, end.Z)
			else:
				continue
			newLine = Line.CreateBound( newStart, newEnd )
			ele.SetCurveInView( DatumExtentType.ViewSpecific, view, newLine )		
	t.Commit()

except:
	pass