# -*- coding: utf-8 -*-
import Autodesk
from Autodesk.Revit.DB import *
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox,
								Separator, Button, CheckBox)
from Autodesk.Revit.UI.Selection import ObjectType 
__doc__ = 'LNQ'
__title__ = 'Cut Grid'


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = doc.ActiveView


# fecGrids = FilteredElementCollector(doc, view.Id).OfClass(DatumPlane).ToElements()
# for i in fecGrids:
# 	curve =  i.Curve
# 	startpointXYZ = curve.GetEndPoint(0)
# 	endpointXYZ = curve.GetEndPoint(1)
	# print(startpointXYZ)
	# print(endpointXYZ)
# all_grids = FilteredElementCollector(doc, view.Id) \
# 			.OfCategory(BuiltInCategory.OST_Grids) \
# 			.WhereElementIsNotElementType() \
# 			.ToElements()
# for j in all_grids:
# 	startpointXYZ = j.GetEndPoint(0)
# 	print(startpointXYZ)
# cropBox = view.CropBox

# cutOffset =  fecGrids[0].Curve.Origin.Z
# fecGrids = [x for x in fecGrids if isinstance(x, Autodesk.Revit.DB.Grid)]

# shpMannager = view.GetCropRegionShapeManager()
# boundLines = shpMannager.GetCropShape()[0]
# print(shpMannager)
# print(boundLines)
# try:
components = [
				Label('Offset With Crop Region:'),
				TextBox('textbox1', Text="0"),
				Separator(),
				Button('OK')]
form = FlexForm('LNQ tools', components)
form.show()
form.values
offset = float(form.values["textbox1"])

t = Transaction(doc, __title__)
t.Start()
selection = uidoc.Selection.GetElementIds() #Trả về 1 list ElementId
for i in selection:
	ele = doc.GetElement(i)
	category = ele.Category
	# category_name = category.Name
	# print(category.Name)
	if category:
		if category.Name == "Scope Boxes":
			boundingbox = ele.get_BoundingBox(view)
			min_x = boundingbox.Min.X - offset
			min_y = boundingbox.Min.Y - offset
			max_x = boundingbox.Max.X + offset
			max_y = boundingbox.Max.Y + offset

			all_grids = FilteredElementCollector(doc, view.Id) \
							.OfCategory(BuiltInCategory.OST_Grids) \
							.WhereElementIsNotElementType() \
							.ToElements()
			for ele in all_grids:
				vector = ele.Curve.Direction
				if abs(vector.Y) == 1 and abs(vector.X) < 0.1:
					gridCurves = ele.GetCurvesInView(  DatumExtentType.ViewSpecific, view )
					for c in gridCurves:
						start = c.GetEndPoint( 0 )
						end = c.GetEndPoint( 1 )
						newStart =  XYZ(start.X, min_y, start.Z)
						newEnd  =  XYZ(end.X, max_y, end.Z)
						newLine = Line.CreateBound( newStart, newEnd )
						ele.SetCurveInView( DatumExtentType.ViewSpecific, view, newLine )
				elif abs(vector.X) == 1 and abs(vector.Y) < 0.1:
					gridCurves_x = ele.GetCurvesInView(  DatumExtentType.ViewSpecific, view )
					for c in gridCurves_x:
						start = c.GetEndPoint( 0 )
						end = c.GetEndPoint( 1 )
						newStart =  XYZ(min_x, start.Y, start.Z)
						newEnd  =  XYZ(max_x, end.Y, end.Z)
						newLine = Line.CreateBound( newStart, newEnd )
						ele.SetCurveInView( DatumExtentType.ViewSpecific, view, newLine )
		else:
			boundingbox = ele.get_BoundingBox(view)
			min_x = boundingbox.Min.X 
			min_y = boundingbox.Min.Y 
			min_z = boundingbox.Min.Z - offset
			max_x = boundingbox.Max.X 
			max_y = boundingbox.Max.Y
			max_z = boundingbox.Max.Z + offset

			all_grids = FilteredElementCollector(doc, view.Id) \
							.OfCategory(BuiltInCategory.OST_Grids) \
							.WhereElementIsNotElementType() \
							.ToElements()
			for ele in all_grids:
				vector = ele.Curve.Direction
				# if abs(vector.Y) == 1 and abs(vector.X) < 0.1:
				gridCurves = ele.GetCurvesInView(  DatumExtentType.ViewSpecific, view )
				for c in gridCurves:
					start = c.GetEndPoint( 0 )
					end = c.GetEndPoint( 1 )
					newStart =  XYZ(start.X, start.Y, min_z)
					newEnd  =  XYZ(end.X, end.Y, max_z)
					newLine = Line.CreateBound( newStart, newEnd )
					ele.SetCurveInView( DatumExtentType.ViewSpecific, view, newLine )
	else:
		boundingbox = ele.get_BoundingBox(view)
		min_x = boundingbox.Min.X - offset
		min_y = boundingbox.Min.Y - offset
		max_x = boundingbox.Max.X + offset
		max_y = boundingbox.Max.Y + offset


		all_grids = FilteredElementCollector(doc, view.Id) \
						.OfCategory(BuiltInCategory.OST_Grids) \
						.WhereElementIsNotElementType() \
						.ToElements()
		for ele in all_grids:
			vector = ele.Curve.Direction
			if abs(vector.Y) == 1 and abs(vector.X) < 0.1:
				gridCurves = ele.GetCurvesInView(  DatumExtentType.ViewSpecific, view )
				for c in gridCurves:
					start = c.GetEndPoint( 0 )
					end = c.GetEndPoint( 1 )
					newStart =  XYZ(start.X, min_y, start.Z)
					newEnd  =  XYZ(end.X, max_y, end.Z)
					newLine = Line.CreateBound( newStart, newEnd )
					ele.SetCurveInView( DatumExtentType.ViewSpecific, view, newLine )
			elif abs(vector.X) == 1 and abs(vector.Y) < 0.1:
				gridCurves_x = ele.GetCurvesInView(  DatumExtentType.ViewSpecific, view )
				for c in gridCurves_x:
					start = c.GetEndPoint( 0 )
					end = c.GetEndPoint( 1 )
					newStart =  XYZ(min_x, start.Y, start.Z)
					newEnd  =  XYZ(max_x, end.Y, end.Z)
					newLine = Line.CreateBound( newStart, newEnd )
					ele.SetCurveInView( DatumExtentType.ViewSpecific, view, newLine )

t.Commit()
# except:
# 	pass



































# def createDatumline(boundline, grid):
# 	gridline = None
# 	curveG = grid.Curve
# 	vectGrid = curveG.Direction
# 	lstPtToLine=[]
# 	for lineBound in boundLines:
# 		rayc = DB.Line.CreateUnbound (XYZ (curveG.Origin.X,
# 		curveG.Origin.Y, lineBound.GetEndPoint(0).Z), vectGrid) 
# 		outInterR = clr.Reference[IntersectionResultArray] ()
# 		result = rayc. Intersect (lineBound, outInterR) 
# 		print(result)
# 		if result == SetComparisonResult.Overlap: 
# 			interResult = outInterR.Value
# 			lstPtToLine.append(interResult[0].XYZPoint)
# 	if len(lstPtToLine) == 2:
# 		P1 = lstPtToLine[0].ToPoint()
# 		P2 = lstPtToLine[1].ToPoint()
# 		TransXYZ1 = Geometry. Translate (P1,0,0,0)
# 		TransXYZ2 = Geometry. Translate(P2,0,0,0) 
# 		TransPoint1 = TransXYZ1.ToXyz()
# 		TransPoint2 = TransXYZ2.ToXyz()
# 		gridline = Autodesk.Revit.DB.Line.CreateBound(TransPoint1,TransPoint2)
# 	return gridline
# if view.CropBoxVisible == True:
# 	# doc.Regenerate()
# 	cropBox = view.CropBox
# 	fecGrids = FilteredElementCollector(doc, view.Id).OfClass(DatumPlane).ToElements()
# 	cutOffset =  fecGrids[0].Curve.Origin.Z
# 	fecGrids = [x for x in fecGrids if isinstance(x, Autodesk.Revit.DB.Grid)]

# 	outLst=[]
# 	newGLineList = []
# 	shpMannager = view.GetCropRegionShapeManager()
# 	boundLines = shpMannager.GetCropShape()[0]
# 	startpointList = []
# 	endpointList = []
# 	rotationAngleList = []
# 	if view.ViewDirection.IsAlmostEqualTo(XYZ(0,0,1)):
# 		currentZ = list(boundLines)[0].GetEndPoint(0).Z
# 		tf = Transform.CreateTranslation(XYZ(0,0,cutOffset-currentZ))
# 		boundLines = CurveLoop.CreateViaTransform(boundLines,tf)
# 		for grid in fecGrids:
# 			outLst.append(grid.Curve.ToPrototype())
# 			newGLine = createDatumline(boundline,grid)
# 			newGLineList.append(newGLine)
# 			if newGLine is not None:
# 				grid.SetCurveInView(DatumExtentType.ViewSpecific,doc.ActiveView, newGLine)
# 	else:
# 		for grid in fecGrids:
# 			startpointXYZ = newGLine.GetEndPoint(0)
# 			endpointXYZ = newGLine.GetEndPoint(1)
# 			startpoint = endpointXYZ.ToPoint()
# 			endpoint = endpointXYZ.ToPoint()
# 			startpointList.append(startpoint)
# 			endpointList.append(endpoint)
# 			vector = newGLine.Direction
# 			rotationAngle=abs(math.degrees(vector.AngleOnPlaneTo(view.RightDirection.XYZ.BasicZ)))













		# if abs(start.X - end.X) < 1e-9:
		# 	new_start= XYZ(start.X, max_y, start.Z)
		# 	new_end = XYZ(end.X, min_y, end.Z)
		# 	revitCurves = []
		# 	new_curve = Line.CreateBound(new_start, new_end)
		# 	revitCurves.append(new_curve)
		# 	i.SetCurveInView(DatumExtentType.ViewSpecific,doc.ActiveView, new_curve)
				











