# -*- coding: utf-8 -*-
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import*
from Autodesk.Revit.UI.Selection import ISelectionFilter
__doc__ = 'LNQ'
__title__ = 'Cut Grid theo Line'

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = doc.ActiveView
# try:
class GridSelectionFilter(ISelectionFilter):
	def AllowElement(self, element):
		return isinstance(element, Grid)

	def AllowReference(self, reference, position):
		return False

def select_grids(sender, args):
	# Lấy các đối tượng Grid được chọn
	selection = uidoc.Selection
	selected_grids = selection.GetElementIds()
	
	# In ra thông tin về các đối tượng Grid được chọn
	for grid_id in selected_grids:
		grid = doc.GetElement(grid_id)
	
	return selected_grids
# Đăng ký sự kiện 'selectionChanged'
uidoc.Application.ViewActivated += select_grids

try:
	# Cho phép người dùng chọn các đối tượng
	selection = uidoc.Selection
	filter = GridSelectionFilter()
	selected_grids = selection.PickObjects(ObjectType.Element, filter)
finally:
	# Hủy đăng ký sự kiện 'selectionChanged'
	uidoc.Application.ViewActivated -= select_grids

	def get_intersection(c1, c2):
		p1 = c1.GetEndPoint(0)
		q1 = c1.GetEndPoint(1)
		p2 = c2.GetEndPoint(0)
		q2 = c2.GetEndPoint(1)
		v1 = q1 - p1
		v2 = q2 - p2
		w = p2 - p1
		p5 = None
		denominator = v2.X * v1.Y - v2.Y * v1.X
		if not abs(denominator) < 1e-9:
			c = (v2.X * w.Y - v2.Y * w.X) / denominator
			x = p1.X + c * v1.X
			y = p1.Y + c * v1.Y
			p5 = XYZ(x, y, 0)
		return p5
# def get_intersection(c1, c2):
#     p1 = c1.GetEndPoint(0)
#     q1 = c1.GetEndPoint(1)
#     p2 = c2.GetEndPoint(0)
#     q2 = c2.GetEndPoint(1)
#     p1_x, p1_y, p1_z = p1.X, p1.Y, p1.Z
#     q1_x, q1_y, q1_z = q1.X, q1.Y, q1.Z
#     p2_x, p2_y, p2_z = p2.X, p2.Y, p2.Z
#     q2_x, q2_y, q2_z = q2.X, q2.Y, q2.Z

#     # Kiểm tra trường hợp 2 đường thẳng trùng nhau
#     if (p2_x - p1_x) / (q2_x - q1_x) == (p2_y - p1_y) / (q2_y - q1_y):
#         return None

#     m1 = (q1_y - p1_y) / (q1_x - p1_x)
#     b1 = p1_y - m1 * p1_x
#     m2 = (q2_y - p2_y) / (q2_x - p2_x)
#     b2 = p2_y - m2 * p2_x

#     # Kiểm tra trường hợp 2 đường thẳng song song
#     if m1 == m2:
#         return None

#     x = (b2 - b1) / (m1 - m2)
#     y = m1 * x + b1
#     p5 = XYZ(x, y, 0)
#     return p5


pick_line = uidoc.Selection.PickObject(ObjectType.Element)
eleid_line = pick_line.ElementId
ele_line = doc.GetElement(pick_line)

t = Transaction(doc, __title__)
t.Start()
intersection_points = []
for pick_grid in selected_grids:
	eleid_grid = pick_grid.ElementId
	ele_grid = doc.GetElement(pick_grid)
	point = get_intersection(ele_grid.Curve, ele_line.GeometryCurve)
	# print(point)
	if point:
		grid_Curves = ele_grid.GetCurvesInView(DatumExtentType.ViewSpecific, view)
		intersection_points.append(point)
		for curve in grid_Curves:
			start = curve.GetEndPoint(0)
			end = curve.GetEndPoint(1)
			distance_to_start = point.DistanceTo(start)
			distance_to_end = point.DistanceTo(end)
			if distance_to_start > distance_to_end:
				new_start = XYZ(start.X, start.Y, start.Z)
				new_end = XYZ(point.X, point.Y, end.Z)
				new_curve = Line.CreateBound(new_start, new_end)
				ele_grid.SetCurveInView(DatumExtentType.ViewSpecific, view, new_curve)
			elif distance_to_start < distance_to_end:
				new_start = XYZ(point.X, point.Y, start.Z)
				new_end = XYZ(end.X, end.Y, end.Z)
				new_curve = Line.CreateBound(new_start, new_end)
				ele_grid.SetCurveInView(DatumExtentType.ViewSpecific, view, new_curve)
			else:
				pass
t.Commit()
# except:
# 	pass



# def get_intersection(curve, line):
# 	results = clr.Reference[IntersectionResultArray]()
	
# 	# Kiểm tra xem line có phải là DetailLine không
# 	if isinstance(line, DetailLine):
# 		# Kiểm tra xem line có phải là đường cong không
# 		if isinstance(line.GeometryCurve, Curve):
# 			line = line.GeometryCurve
# 		else:
# 			line = line.Curve
	
# 	result = curve.Intersect(line, results)

# 	if result != SetComparisonResult.Overlap:
# 		print('No Intersection')
# 		return None

# 	intersection = results.Item[0]
# 	return intersection.XYZPoint

# selected_grids = uidoc.Selection.PickObjects(ObjectType.Element) #Trả về 1 ref

# pick_line = uidoc.Selection.PickObject(ObjectType.Element) #Trả về 1 ref
# eleid_line = pick_line.ElementId
# ele_line = doc.GetElement(pick_line)   #Tu pick ra duoc Element

# t = Transaction(doc, __title__)
# t.Start()
# intersection_points = []
# for pick_grid in selected_grids:
# 	eleid_grid = pick_grid.ElementId
# 	ele_grid = doc.GetElement(pick_grid)
# 	point = get_intersection(ele_grid.Curve, ele_line)
# 	print(point)
# 	if point:
# 		grid_Curves = ele_grid.GetCurvesInView(  DatumExtentType.ViewSpecific, view )
# 		intersection_points.append(point)
# 		for curve in grid_Curves:
# 			# Lấy start và end của đường cong
# 			start = curve.GetEndPoint(0)
# 			end = curve.GetEndPoint(1)
			
# 			# Cập nhật lại start và end của đường cong dựa trên điểm giao
# 			new_start = XYZ(start.X, start.Y, start.Z)
# 			new_end = XYZ(point.X, point.Y, end.Z)
			
# 			# Tạo đường cong mới với giá trị mới
# 			new_curve = Line.CreateBound(new_start, new_end)
			
# 			# Cập nhật đường cong của grid
# 			ele_grid.SetCurveInView(DatumExtentType.ViewSpecific, view, new_curve)

# t.Commit()