#NVDAObjects/UIA/sysListView32.py
#A part of NonVisual Desktop Access (NVDA)
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.
#Copyright (C) 2018 NV Access Limited, Babbage B.V.

import UIAHandler
from . import UIA, GridRow
from . import ListItem as UIAListItem
import watchdog
from NVDAObjects.IAccessible.sysListView32 import LVM_GETITEMCOUNT


class List(UIA):

	def _get_name(self):
		name=super(List,self)._get_name()
		if not name:
			name=super(UIA,self)._get_name()
		return name

class ListItem(UIAListItem):
	ignoreFirstColumnHeader = True

	def _get_table(self):
		# UIA's SysListView32 implementation doesn't implement the GridItem pattern for the list items.
		if isinstance(self, GridRow):
			return self.parent
		return super(ListItem, self).table

	def _get_rowNumber(self):
		# UIA's SysListView32 implementation doesn't implement the GridItem pattern for the list items.
		# Therefore, we use the row number from the first child
		if isinstance(self, GridRow):
			return self.firstChild.rowNumber
		return super(ListItem, self).rowNumber

	def _get_positionInfo(self):
		automationID = self.UIAElement.cachedAutomationID
		if not automationID.startswith("ListViewItem"):
			return super(ListItem, self).positionInfo
		index=int(automationID.split("-")[1])+1
		totalCount=watchdog.cancellableSendMessage(self.windowHandle,LVM_GETITEMCOUNT,0,0)
		return dict(indexInGroup=index,similarItemsInGroup=totalCount) 
