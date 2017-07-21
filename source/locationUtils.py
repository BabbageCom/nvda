#locationUtils.py
#A part of NonVisual Desktop Access (NVDA)
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.
#Copyright (C) 2017 NV Access Limited, Babbage B.V.

"""Utilities and classes for working with rectangles and coordinates."""

from collections import namedtuple
import windowUtils
import winUser
from ctypes.wintypes import RECT, SMALL_RECT, POINT
import textInfos

class Point(namedtuple("Point",("x","y"))):
	"""Represents a point on the screen."""

	def __add__(self,other):
		if not isinstance(other,_POINT_CLASSES):
			return super(Point,self).__add__(other)
		return Point((self.x+other.x),(self.y+other.y))

	def __radd__(self,other):
		if not isinstance(other,_POINT_CLASSES):
			return super(Point,self).__radd__(other)
		return Point((self.x+other.x),(self.y+other.y))

	def __sub__(self,other):
		if not isinstance(other,_POINT_CLASSES):
			return super(Point,self).__sub__(other)
		return Point((self.x-other.x),(self.y-other.y))

	def __rsub__(self,other):
		if not isinstance(other,_POINT_CLASSES):
			return super(Point,self).__rsub__(other)
		return Point((self.x-other.x),(self.y-other.y))

class _RectMixin:
	"""Mix-in class for properties shared between Location and Rect classes"""

	def toCRect(self):
		return RECT(self.left,self.top,self.right,self.bottom)

	@property
	def topLeft(self):
		return Point(self.left,self.top)

	@property
	def bottomRight(self):
		return Point(self.right,self.bottom)

	@property
	def center(self):
		return Point((self.left+self.right)/2, (self.top+self.bottom)/2)

	def __contains__(self,other):
		if isinstance(other,_POINT_CLASSES):
			return other.x >= self.left < self.right and other.y >= self.top < self.bottom
		if isinstance(other,_RECT_CLASSES):
			return self<other
		return super(type(self),self).__contains__(other)

	def __lt__(self, other):
		if isinstance(other,_RECT_CLASSES):
			return self.left > other.left and self.top > other.top and self.right < other.right and self.bottom < other.bottom
		return super(type(self),self).__lt__(other)

	def __le__(self, other):
		if isinstance(other,_RECT_CLASSES):
			return self.left >= other.left and self.top >= other.top and self.right <= other.right and self.bottom <= other.bottom
		return super(type(self),self).__le__(other)

	def __gt__(self, other):
		if isinstance(other,_RECT_CLASSES):
			return other.left > self.left and other.top > self.top and other.right < self.right and other.bottom < self.bottom
		return super(type(self),self).__gt__(other)

	def __ge__(self, other):
		if isinstance(other,_RECT_CLASSES):
			return other.left >= self.left and other.top >= self.top and other.right <= self.right and other.bottom <= self.bottom
		return super(type(self),self).__ge__(other)

	def __eq__(self, other):
		if isinstance(other,_RECT_CLASSES):
			return other.left == self.left and other.top == self.top and other.right == self.right and other.bottom == self.bottom
		return super(type(self),self).__eq__(other)

	def __neq__(self, other):
		if isinstance(other,_RECT_CLASSES):
			return not (other.left == self.left and other.top == self.top and other.right == self.right and other.bottom == self.bottom)
		return super(type(self),self).__neq__(other)

class Location(_RectMixin, namedtuple("Location",("left","top","width","height"))):
	"""Represents a rectangle on the screen, based on left and top coordinates, width and height."""

	@property
	def right(self):
		return self.left+self.width

	@property
	def bottom(self):
		return self.top+self.height

	def toRect(self):
		return Rect(self.left,self.top,self.right,self.bottom)

	def toLogical(self, hwnd):
		left,top=windowUtils.physicalToLogicalPoint(hwnd, self.left, self.top)
		right,bottom=windowUtils.physicalToLogicalPoint(hwnd, self.right, self.bottom)
		return Location(left,top,(right-left),(bottom-top))

	def toPhysical(self, hwnd):
		left,top=windowUtils.logicalToPhysicalPoint(hwnd, self.left, self.top)
		right,bottom=windowUtils.logicalToPhysicalPoint(hwnd, self.right, self.bottom)
		return Location(left,top,(right-left),(bottom-top))

	def toClient(self, hwnd):
		left,top=winUser.ScreenToClient(hwnd, self.left, self.top)
		right,bottom=winUser.ScreenToClient(hwnd, self.right, self.bottom)
		return Location(left,top,(right-left),(bottom-top))

	def toScreen(self, hwnd):
		left,top=winUser.ClientToScreen(hwnd, self.left, self.top)
		right,bottom=winUser.ClientToScreen(hwnd, self.right, self.bottom)
		return Location(left,top,(right-left),(bottom-top))

class Rect(_RectMixin, namedtuple("Rect",("left","top","right","bottom"))):
	"""Represents a rectangle on the screen.
	By convention, the right and bottom edges of the rectangle are normally considered exclusive.
	"""

	@property
	def width(self):
		return self.right-self.left

	@property
	def height(self):
		return self.bottom-self.top

	def toLocation(self):
		return Location(self.left,self.top,self.width,self.height)

	def toLogical(self, hwnd):
		left,top=windowUtils.physicalToLogicalPoint(hwnd, self.left, self.top)
		right,bottom=windowUtils.physicalToLogicalPoint(hwnd, self.right, self.bottom)
		return Rect(left,top,right,bottom)

	def toPhysical(self, hwnd):
		left,top=windowUtils.logicalToPhysicalPoint(hwnd, self.left, self.top)
		right,bottom=windowUtils.logicalToPhysicalPoint(hwnd, self.right, self.bottom)
		return Rect(left,top,right,bottom)

	def toClient(self, hwnd):
		left,top=winUser.ScreenToClient(hwnd, self.left, self.top)
		right,bottom=winUser.ScreenToClient(hwnd, self.right, self.bottom)
		return Rect(left,top,right,bottom)

	def toScreen(self, hwnd):
		left,top=winUser.ClientToScreen(hwnd, self.left, self.top)
		right,bottom=winUser.ClientToScreen(hwnd, self.right, self.bottom)
		return Rect(left,top,right,bottom)

def toRect(*params):
	if len(params)==0:
		raise TypeError("This function takes at least 1 argument (0 given)")
	if len(params)==1:
		param=params[0]
		if isinstance(param,Rect):
			return param
		if isinstance(param,_RECT_CLASSES):
			return Rect(param.left,param.top,param.right,param.bottom)
		if isinstance(param,_POINT_CLASSES):
			# Right and bottom edges of the resulting rectangle are considered exclusive
			x,y=point.x,point.y
			return Rect(x,y,x+1,y+1)
		if isinstance(param,(tuple,list)):
			# One indexable in another indexable doesn't make sence, so treat the inner indexable as outer indexable
			params=param
	if all(isinstance(param,int) for param in params):
		if len(params)==4:
			# Assume that we are converting from a tuple rectangle.
			# To convert from a tuple location, use L{toLocation} instead.
			# To convert from a tuple rectangle to L{Location}, use this function and execute L{toLocation} on the resulting object.
			return Rect(*params)
		elif len(param)==2:
			x,y=params
			return Rect(x,y,x+1,y+1)
	xs=[]
	ys=[]
	for param in params:
		if isinstance(param,_RECT_CLASSES):
			xs.extend((param.left,param.right))
			ys.extend((param.top,param.bottom))
		elif isinstance(param,_POINT_CLASSES):
			xs.append(param.x)
			ys.append(param.y)
		else:
			raise ValueError("Unexpected parameter %s"%param)
	left=min(xs)
	top=min(ys)
	right=max(xs)
	bottom=max(ys)
	return Rect(left,top,right,bottom)

def toLocation(*params):
	if len(params)==0:
		raise TypeError("This function takes at least 1 argument (0 given)")
	if len(params)==1:
		param=params[0]
		if isinstance(param,Location):
			return param
		if isinstance(param,(tuple,list)):
			# One indexable in another indexable doesn't make sence, so treat the inner indexable as outer indexable
			params=param
	if len(params)==4 and all(isinstance(param,int) for param in params):
		# Assume that we are converting from a tuple location.
		# To convert from a tuple rectangle, use L{toRectangle} instead.
		# To convert from a tuple location to L{Rect}, use this function and execute L{toRect} on the resulting location.
		return Location(*params)
	return toRect(*params).toLocation()

_POINT_CLASSES=(Point,POINT,textInfos.Point)
_RECT_CLASSES=(Rect,Location,RECT,SMALL_RECT,textInfos.Rect)