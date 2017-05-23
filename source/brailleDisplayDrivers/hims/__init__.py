#brailleDisplayDrivers/hims/__init__.py
#A part of NonVisual Desktop Access (NVDA)
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.
#Copyright (C) 2010-2017 Gianluca Casalino, NV Access Limited, Babbage B.V.

from logHandler import log
from ctypes import *
from ctypes.wintypes import *
import braille
import inputCore
from winUser import WNDCLASSEXW, WNDPROC, LRESULT, HCURSOR
import hwPortUtils
from brailleInput import BrailleInputGesture
from windowUtils import CustomWindow

HIMS_KEYPRESSED = 0x01
HIMS_KEYRELEASED = 0x02
HIMS_CURSORROUTING = 0x00
HIMS_CODE_DEVICES = {
	1: 'Braille Sense (2 scrolls mode)',
	2: 'Braille Sense QWERTY',
	3: 'Braille EDGE',
	4: 'Braille Sense (4 scrolls mode)',
	5: 'Smart Beetle',
}
HIMS_BLUETOOTH_NAMES = (
	"BrailleSense",
	"BrailleEDGE",
	"SmartBeetle",
)

#MAP OF KEYS

HIMS_KEYS = {
	0x01: 'dot1',
	0x02: 'dot2',
	0x04: 'dot3',
	0x08: 'dot4',
	0x010: 'dot5',
	0x020: 'dot6',
	0x040: 'dot7',
	0x080: 'dot8',
	0x100: 'space',
	0x200: 'advance1',
	0x400: {'Smart Beetle': 'advance4', 'Braille Sense': 'advance2'},
	0x800: 'advance3',
	0x1000: {'Smart Beetle': 'advance2', 'Braille Sense': 'advance4'},
	0x2000: 'leftSideScroll',
	0x4000: 'rightSideScroll',
	0x10000: 'leftSideScrollUp',
	0x20000: {'Braille EDGE': 'rightSideScrollUp', 'Braille Sense': 'leftSideScrollDown'},
	0x40000: {'Smart Beetle': 'leftSideScroll', 'Braille EDGE': 'rightSideScrollDown', 'Braille Sense': 'rightSideScrollUp'},
	0x80000: {'Smart Beetle': 'rightSideScroll', 'Braille EDGE': 'leftSideScrollDown', 'Braille Sense': 'rightSideScrollDown'},
	0x100000: 'Advance5',
	0x200000: 'Advance6',
	0x400000: 'Advance7',
	0x800000: 'Advance8',
	0x1000000: 'leftSideUpArrow',
	0x2000000: 'leftSideDownArrow',
	0x4000000: 'leftSideLeftArrow',
	0x8000000: 'leftSideRightArrow',
	0x10000000: 'rightSideUpArrow',
	0x20000000: 'rightSideDownArrow',
	0x40000000: 'rightSideLeftArrow',
	-0x80000000: 'rightSideRightArrow',
}
SPACE_KEY = 0x100

pressedKeys = set()
_ignoreKeyPresses = False
deviceFound = None

try:
	himsLib = cdll.LoadLibrary("brailleDisplayDrivers\\hims\\HanSoneConnect.dll")
except:
	himsLib = None

nvdaHIMSBrlWm=windll.user32.RegisterWindowMessageW(u"nvdaHIMSBrlWm")

class BrlHimsWindow(CustomWindow):
	className = u"nvdaHIMSBrlWndCls window"

	def windowProc(self, hwnd, msg, wParam, lParam):
		global pressedKeys, _ignoreKeyReleases
		super(BrlHimsWindow, self).windowProc(hwnd, msg, wParam, lParam)
		if msg == nvdaHIMSBrlWm and wParam==HIMS_KEYRELEASED:
			if not _ignoreKeyReleases and pressedKeys:
				try:
					inputCore.manager.executeGesture(InputGesture(pressedKeys))
				except inputCore.NoInputGestureAction:
					pass
				_ignoreKeyReleases = True
			pressedKeys.discard(lParam)
		elif msg == nvdaHIMSBrlWm and wParam == HIMS_CURSORROUTING:
			try:
				inputCore.manager.executeGesture(InputGesture(lParam))
			except inputCore.NoInputGestureAction:
				pass
		elif msg == nvdaHIMSBrlWm and  wParam == HIMS_KEYPRESSED:
			pressedKeys.add(lParam)
			_ignoreKeyReleases = False
		return windll.user32.DefWindowProcW(hwnd,msg,wParam,lParam)

class BrailleDisplayDriver(braille.BrailleDisplayDriver):
	""" HIMS Braille Sense/Braille EDGE braille displays.
	"""
	name = "hims"
	# Translators: The name of a series of braille displays.
	description = _("HIMS Braille Sense/Braille EDGE/Smart Beetle series")

	@classmethod
	def check(cls):
		return bool(himsLib)

	def __init__(self):
		global deviceFound
		super(BrailleDisplayDriver, self).__init__()
		self._messageWindow = BrlHimsWindow(u"nvdaHIMSBrlWndCls window")
		code = himsLib.Open("USB",self._messageWindow.handle,nvdaHIMSBrlWm)
		if  code == 0:
			for portInfo in sorted(hwPortUtils.listComPorts(onlyAvailable=True), key=lambda item: "bluetoothName" in item):
				port = portInfo["port"].lower()
				btName = portInfo.get("bluetoothName")
				if btName and any(btName.startswith(prefix) for prefix in HIMS_BLUETOOTH_NAMES):
					try:
						if int(port.split("com")[1]) > 8:
							port = "\\\\.\\"+port
					except (IndexError, ValueError):
						pass
					code = himsLib.Open(str(port),self._messageWindow.handle,nvdaHIMSBrlWm)
		if himsLib.GetBSCellCount() == 14:
			# Ugly, but effective
			code = 5
		if code >= 1:
			deviceFound = HIMS_CODE_DEVICES[code]
			log.info("Hims %s device found"%deviceFound)
			return
		raise RuntimeError("No display found")

	def terminate(self):
		super(BrailleDisplayDriver, self).terminate()
		himsLib.Close()
		self._messageWindow.destroy()

	def _get_numCells(self):
		return himsLib.GetBSCellCount()

	def display(self, cells):
		cells = "".join([chr(x) for x in cells])
		himsLib.SendData(cells)

	gestureMap = inputCore.GlobalGestureMap({
		"globalCommands.GlobalCommands": {
			"kb:leftAlt": ("br(hims):advance4","br(hims):dot1+dot3+dot4+space",),
			"kb:capsLock": ("br(hims):dot1+dot3+dot6+space",),
			"kb:control": ("br(hims):advance3",),
			"kb:shift": ("br(hims):advance5",),
			"kb:insert": ("br(hims):advance6",),
			"kb:windows": ("br(hims):advance7",),
			"kb:applicationS": ("br(hims):advance8",),
			"kb:tab": ("br(hims):advance2","br(hims):dot4+dot5+space",),
			"kb:shift+alt+tab": ("br(hims):advance2+advance3+advance1",),
			"kb:alt+tab": ("br(hims):advance2+advance3",),
			"kb:shift+tab": ("br(hims):dot1+dot2+space",),
			"kb:end": ("br(hims):dot4+dot6+space",),
			"kb:control+end": ("br(hims):dot4+dot5+dot6+space",),
			"kb:home": ("br(hims):dot1+dot3+space",),
			"kb:control+home": ("br(hims):dot1+dot2+dot3+space",),
			"kb:leftArrow": ("br(hims):leftSideLeftArrow","br(hims):dot3+space",),
			"kb:control+shift+leftArrow": ("br(hims):dot2+dot8+space+advance1",),
			"kb:control+leftArrow": ("br(hims):advance3+leftSideLeftArrow","br(hims):dot2+space",),
			"kb:shift+alt+leftArrow": ("br(hims):dot2+dot7+advance1",),
			"kb:alt+leftArrow": ("br(hims):advance4+leftSideLeftArrow","br(hims):dot2+dot7",),
			"kb:rightArrow": ("br(hims):leftSideRightArrow","br(hims):dot6+space",),
			"kb:control+shift+rightArrow": ("br(hims):dot5+dot8+space+advance1",),
			"kb:control+rightArrow": ("br(hims):advance3+leftSideRightArrow","br(hims):dot5+space",),
			"kb:shift+alt+rightArrow": ("br(hims):dot5+dot7+advance1",),
			"kb:alt+rightArrow": ("br(hims):advance4+leftSideRightArrow","br(hims):dot5+dot7",),
			"kb:pageUp": ("br(hims):dot1+dot2+dot6+space",),
			"kb:control+pageUp": ("br(hims):dot1+dot2+dot6+dot8+space",),
			"kb:upArrow": ("br(hims):leftSideUpArrow","br(hims):dot1+space",),
			"kb:control+shift+upArrow": ("br(hims):dot2+dot3+dot8+space+advance1",),
			"kb:control+upArrow": ("br(hims):advance3+leftSideUpArrow","br(hims):dot2+dot3+space",),
			"kb:shift+alt+upArrow": ("br(hims):dot2+dot3+dot7+advance1",),
			"kb:alt+upArrow": ("br(hims):advance4+leftSideUpArrow","br(hims):dot2+dot3+dot7",),
			"kb:shift+upArrow": ("br(hims):leftSideScrollDown+space",),
			"kb:pageDown": ("br(hims):dot3+dot4+dot5+space",),
			"kb:control+pageDown": ("br(hims):dot3+dot4+dot5+dot8+space",),
			"kb:downArrow": ("br(hims):leftSideDownArrow","br(hims):dot4+space",),
			"kb:control+shift+downArrow": ("br(hims):dot5+dot6+dot8+space+advance1",),
			"kb:control+downArrow": ("br(hims):advance3+leftSideDownArrow","br(hims):dot5+dot6+space",),
			"kb:shift+alt+downArrow": ("br(hims):dot5+dot6+dot7+advance1",),
			"kb:alt+downArrow": ("br(hims):advance4+leftSideDownArrow","br(hims):dot5+dot6+dot7",),
			"kb:shift+downArrow": ("br(hims):space+rightSideScrollDown",),
			"kb:backspace": ("br(hims):dot7",),
			"kb:enter": ("br(hims):dot8",),
			"kb:escape": ("br(hims):advance1","br(hims):dot1+dot5+space",),
			"kb:delete": ("br(hims):dot1+dot3+dot5+space",),
			"kb:f1": ("br(hims):dot1+dot2+dot5+space",),
			"kb:f3": ("br(hims):dot1+dot2+dot4+dot8",),
			"kb:f4": ("br(hims):dot7+advance3",),
			"kb:windows+b": ("br(hims):dot1+dot2+advance1",),
			"kb:windows+d": ("br(hims):dot1+dot4+dot5+advance1",),
			"braille_routeTo": ("br(hims):routing",),
			"braille_previousLine": ("br(hims):leftSideScrollUp",),
			"braille_nextLine": ("br(hims):rightSideScrollUp",),
			"braille_scrollBack": ("br(hims):leftSideScrollDown","br(hims):leftSideScroll",),
			"braille_scrollForward": ("br(hims):rightSideScrollDown","br(hims):rightSideScroll",),
			"review_previousLine": ("br(hims):rightSideUpArrow",),
			"review_nextLine": ("br(hims):rightSideDownArrow",),
			"review_previousCharacter": ("br(hims):rightSideLeftArrow",),
			"review_nextCharacter": ("br(hims):rightSideRightArrow",),
		}
	})

class InputGesture(braille.BrailleDisplayGesture, BrailleInputGesture):
	source = BrailleDisplayDriver.name
	def __init__(self, keys):
		super(InputGesture, self).__init__()
		if isinstance(keys,int):  
			self.routingIndex = keys
			self.id = "routing"
			return
		self.keyCodes = set(keys)
		names = set()
		isBrailleInput = True
		for value in self.keyCodes: 
			if isBrailleInput:
				if 0xff & value:
					self.dots |= value
				elif value == SPACE_KEY:
					self.space = True
				else:
					# This is not braille input.
					isBrailleInput = False
					self.dots = 0
					self.space = False
			try:
				name = HIMS_KEYS[value]
				if isinstance(name, dict):
					try:
						name = name[deviceFound]
					except KeyError:
						name = name['Braille Sense']
				names.add(name)
			except KeyError:
				log.debugWarning("Unknown key %x" % value)
		self.id = "+".join(names)