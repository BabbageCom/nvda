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

class _RectMixIn(object):
	"""Mix-in class for properties shared between Location and Rect classes"""

	@classmethod
	def __new__(self, *args, **kwargs):
		raise NotImplementedError("This mix-in class can only be used as such")

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
			return self.left >= other.left and self.top >= other.top and self.right <= other.right and self.bottom <= other.bottom

class Location(namedtuple("Location",("left","top","width","height")), _RectMixIn):
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

class Rect(namedtuple("Rect",("left","top","right","bottom")), _RectMixIn):
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

def toRect(param):
	if isinstance(param,Rect):
		return param
	if isinstance(param,_RECT_CLASSES):
		return Rect(param.left,param.top,param.right,param.bottom)
	if isinstance(param,_POINT_CLASSES):
		# Right and bottom edges of the resulting rectangle are considered exclusive
		x,y=point.x,point.y
		return Rect(x,y,x+1,y+1)
	if isinstance(param,tuple):
		if len(param)==4 and isinstance(param[0],int) and isinstance(param[1],int) and isinstance(param[2],int) and isinstance(param[3],int):
			# Assume that we are converting from a tuple rectangle
			return Rect(*param)
		if len(param)==2 and isinstance(param[0],int) and isinstance(param[1],int):
			x,y=param
			return Rect(x,y,x+1,y+1)
		if len(param)==2 and isinstance(param[0],_POINT_CLASSES) and isinstance(param[1],_POINT_CLASSES):
			left,top=param[0].x,param[0].y
			right,bottom=param[1].x,param[1].y
			return Rect(left,top,right,bottom)

def toLocation(param):
	if isinstance(param,Location):
		return param
	if isinstance(param,_RECT_CLASSES):
		return Location(param.left,param.top,param.right-param.left,param.bottom-param.top)
	if isinstance(param,_POINT_CLASSES):
		# Right and bottom edges of the resulting rectangle are considered exclusive
		x,y=point.x,point.y
		return Location(x,y,1,1)
	if isinstance(param,tuple):
		if len(param)==4 and isinstance(param[0],int) and isinstance(param[1],int) and isinstance(param[2],int) and isinstance(param[3],int):
			# Assume that we are converting from a tuple location
			return Location(*param)
		if len(param)==2 and isinstance(param[0],int) and isinstance(param[1],int):
			x,y=param
			return Location(x,y,1,1)
		if len(param)==2 and isinstance(param[0],_POINT_CLASSES) and isinstance(param[1],_POINT_CLASSES):
			left,top=param[0].x,param[0].y
			right,bottom=param[1].x,param[1].y
			return Location(left,top,right-left,bottom-top)

_POINT_CLASSES=(Point,POINT,textInfos.Point)
_RECT_CLASSES=(Rect,Location,RECT,SMALL_RECT,textInfos.Rect)