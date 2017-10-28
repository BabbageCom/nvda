#A part of NonVisual Desktop Access (NVDA)
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.
#Copyright (C) 2011-2017 NV Access Limited, Sebastian Kruber <sebastian.kruber@hedo.de>, Leonard de Ruijter
# This file represents the braille display driver for
# hedo MobilLine USB, a product from hedo Reha-Technik GmbH
# see www.hedo.de for more details

import time
import hwIo
import serial
import braille
import inputCore
import hwPortUtils
from logHandler import log

HEDO_MOBIL_USBID = "VID_0403&PID_DE58"
HEDO_MOBIL_TIMEOUT = 0.2
HEDO_MOBIL_BAUDRATE = 9600
HEDO_MOBIL_READ_INTERVAL = 50
HEDO_MOBIL_ACK = 0x30
HEDO_MOBIL_INIT = 0x01
HEDO_MOBIL_CR_BEGIN = 0x40
HEDO_MOBIL_CR_END = 0x67
HEDO_MOBIL_CELL_COUNT = 40
HEDO_MOBIL_STATUS_CELL_COUNT = 2

class BrailleDisplayDriver(braille.BrailleDisplayDriver):
	name = "hedoMobilLine"
	description = "hedo MobilLine USB"
	isThreadSafe = True
	numCells = HEDO_MOBIL_CELL_COUNT

	@classmethod
	def check(cls):
		return True

	def __init__(self):
		super(BrailleDisplayDriver, self).__init__()

		self._deviceFound = False
		for portInfo in hwPortUtils.listComPorts(onlyAvailable=True):
			port = portInfo["port"]
			hwID = portInfo["hardwareID"]
			# log.info("Found port {port} with hardwareID {hwID}".format(port=port, hwID=hwID))
			if not hwID.startswith(r"FTDIBUS\COMPORT"):
				continue
			if HEDO_MOBIL_USBID not in hwID:
				continue
			# At this point, a port bound to this display has been found.
			# Try talking to the display.
			try:
				self._ser = hwIo.Serial(port, baudrate=HEDO_MOBIL_BAUDRATE, timeout=HEDO_MOBIL_TIMEOUT, writeTimeout=HEDO_MOBIL_TIMEOUT, parity=serial.PARITY_ODD, bytesize=serial.EIGHTBITS, stopbits=serial.STOPBITS_ONE, onReceive=self._onReceive)
			except EnvironmentError:
				continue

			# Prepare a blank line
			cells = chr(HEDO_MOBIL_INIT) + chr(0) * (HEDO_MOBIL_CELL_COUNT + HEDO_MOBIL_STATUS_CELL_COUNT)

			# Send the blank line twice
			self._ser.write(cells)
			self._ser.flush()
			self._ser.write(cells)
			self._ser.flush()
			for i in xrange(3):
				# An expected response hasn't arrived yet, so wait for it.
				self._ser.waitForRead(TIMEOUT)
				if self._deviceFound:
					break
			if self._deviceFound:
				# A display responded.
				log.info("Found hedo MobilLine connected via {port}".format(port=port))
				break
			self._dev.close()
		else:
			raise RuntimeError("No display found")

		self._keysDown = set()
		self._released_keys = set()

	def terminate(self):
		try:
			super(BrailleDisplayDriver, self).terminate()
		finally:
			# We absolutely must close the Serial object, as it does not have a destructor.
			# If we don't, we won't be able to re-open it later.
			self._ser.close()

	def display(self, cells):
		# every transmitted line consists of the preamble HEDO_MOBIL_INIT, the statusCells and the Cells
		line = chr(HEDO_MOBIL_INIT) + chr(0) * HEDO_MOBIL_STATUS_CELL_COUNT + "".join(chr(cell) for cell in cells)
		# cells are already padded up numCells
		# thus the expected length of the line is 1 + HEDO_MOBIL_STATUS_CELL_COUNT + HEDO_MOBIL_CELL_COUNT
		# ... just how it should be
		self._ser.write(line)

	def _onReceive(self, data):
		if data == chr(HEDO_MOBIL_ACK):
			if not self._deviceFound:
				self._deviceFound=True
		else:
			self.handleData(ord(data))

	def handleData(self, data):
		if data >= HEDO_MOBIL_CR_BEGIN and data <= HEDO_MOBIL_CR_END:
			# Routing key is pressed
			try:
				inputCore.manager.executeGesture(InputGestureRouting(data - HEDO_MOBIL_CR_BEGIN))
			except inputCore.NoInputGestureAction:
				log.debug("No Action for routing index " + index)
				pass
			return

		# On every keypress or keyrelease information about all keys is sent
		# There are three groups of keys thus three bytes will be sent on
		# each keypress or release
		# The 4 MSB of each byte mark the group
		# Bytes of the form 0x0? include information for B1 to B3
		# Bytes of the form 0x1? include information for B4 to B6
		# Bytes of the form 0x2? include information for K1 to K3
		# The 4 LSB mark the pressed buttons in the group
		# Are all buttons of one group released, the 4 LSB are zero
		if data & 0xF0 == 0x00:
			# B1..B3
			if data & 0x01:
				self._keysDown.add("B1")
			if data & 0x02:
				self._keysDown.add("B2")
			if data & 0x04:
				self._keysDown.add("B3")
			if data == 0x00:
				self._released_keys.add("B1")

		elif data & 0xF0 == 0x10:
			# B4..B6
			if data & 0x01:
				self._keysDown.add("B4")
			if data & 0x02:
				self._keysDown.add("B5")
			if data & 0x04:
				self._keysDown.add("B6")
			if data == 0x10:
				self._released_keys.add("B4")

		elif data & 0xF0 == 0x20:
			# K1..K3
			if data & 0x01:
				self._keysDown.add("K1")
			if data & 0x02:
				self._keysDown.add("K2")
			if data & 0x04:
				self._keysDown.add("K3")
			if data == 0x20:
				self._released_keys.add("K1")

		if "B1" in self._released_keys and "B4" in self._released_keys and "K1" in self._released_keys:
			# all keys are released
			keys = "+".join(self._keysDown)
			self._keysDown = set()
			self._released_keys = set()
			try:
				inputCore.manager.executeGesture(InputGestureKeys(keys))
			except inputCore.NoInputGestureAction:
				log.debug("No Action for keys " + keys)
				pass

	gestureMap = inputCore.GlobalGestureMap({
		"globalCommands.GlobalCommands": {
			"braille_scrollBack": ("br(hedoMobilLine):K1",),
			"braille_toggleTether": ("br(hedoMobilLine):K2",),
			"braille_scrollForward": ("br(hedoMobilLine):K3",),
			"braille_previousLine": ("br(hedoMobilLine):B2",),
			"braille_nextLine": ("br(hedoMobilLine):B5",),
			"sayAll": ("br(hedoMobilLine):B6",),
			"braille_routeTo": ("br(hedoMobilLine):routing",),
		},
	})

class InputGestureKeys(braille.BrailleDisplayGesture):

	source = BrailleDisplayDriver.name

	def __init__(self, keys):
		super(InputGestureKeys, self).__init__()

		self.id = keys

class InputGestureRouting(braille.BrailleDisplayGesture):

	source = BrailleDisplayDriver.name

	def __init__(self, index):
		super(InputGestureRouting, self).__init__()

		self.id = "routing"
		self.routingIndex = index