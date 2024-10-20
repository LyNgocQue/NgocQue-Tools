# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from pyrevit import forms
__doc__ = 'LNQ'
__title__ = 'Convert Grid & Level 2D'


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = doc.ActiveView

try:
	t = Transaction(doc, __title__)
	t.Start()

	all_grids = FilteredElementCollector(doc, view.Id) \
				.OfCategory(BuiltInCategory.OST_Grids) \
				.WhereElementIsNotElementType() \
				.ToElements()
	for i in all_grids:
		i.SetDatumExtentType(DatumEnds.End0, view, DatumExtentType.ViewSpecific)
		i.SetDatumExtentType(DatumEnds.End1, view, DatumExtentType.ViewSpecific)
	all_levels = FilteredElementCollector(doc, view.Id) \
				.OfCategory(BuiltInCategory.OST_Levels) \
				.WhereElementIsNotElementType() \
				.ToElements()
	for i in all_levels:
		i.SetDatumExtentType(DatumEnds.End0, view, DatumExtentType.ViewSpecific)
		i.SetDatumExtentType(DatumEnds.End1, view, DatumExtentType.ViewSpecific)
	t.Commit()
except:
	pass








































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
				











