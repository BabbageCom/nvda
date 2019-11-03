# A part of NonVisual Desktop Access (NVDA)
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.
# Copyright (C) 2019 NV Access Limited

import vision
import driverHandler
import wx
from autoSettingsUtils.utils import StringParameterInfo
from vision.providerBase import VisionEnhancementProviderSettings, SupportedSettingType
from typing import Optional, Type, Any, List

"""Example provider, which demonstrates using the automatically constructed GUI.

For examples of overriding the GUI and using a custom implementation, see NVDAHighlighter or ScreenCurtain.

This example imagines that some settings are "always available", while the availability of others is unknown
until "runtime".
This might be because the provider must interface with an external application or device.
"""


class AutoGuiTestSettings(VisionEnhancementProviderSettings):

	#: dictionary of the setting id's available when provider is running.
	_availableRuntimeSettings = [
	]

	# The following settings can be configured prior to runtime in this example
	shouldDoX: bool
	shouldDoY: bool
	amountOfZ: int
	nameOfSomething: str

	availableNameofsomethings = {
		"n1": StringParameterInfo(id="n1", displayName="name one"),
		"n2": StringParameterInfo(id="n2", displayName="name two"),
		"n3": StringParameterInfo(id="n3", displayName="name three"),
		"n4": StringParameterInfo(id="n4", displayName="name four"),
	}

	# The following settings are runtime only in this example
	runtimeOnlySetting_externalValueLoad: int
	runtimeOnlySetting_localDefault: int

	@classmethod
	def getId(cls) -> str:
		return "exampleProvider_autoGui"

	@classmethod
	def getTranslatedName(cls) -> str:
		return "Example Provider with Auto Gui"  # Should normally be translated with _() method.

	@classmethod
	def _get_preInitSettings(cls) -> SupportedSettingType:
		return [
			driverHandler.BooleanDriverSetting(
				"shouldDoX",  # value stored in matching property name on class
				"Should Do X",
				defaultVal=True
			),
			driverHandler.BooleanDriverSetting(
				"shouldDoY",  # value stored in matching property name on class
				"Should Do Y",
				defaultVal=False
			),
			driverHandler.NumericDriverSetting(
				"amountOfZ",  # value stored in matching property name on class
				"Amount of Z",
				defaultVal=11
			),
			driverHandler.DriverSetting(
				# options for this come from a property with name generated by
				# f"available{settingID.capitalize()}s"
				# Note:
				#   First letter of Id becomes capital, the rest lowercase.
				#   the 's' character on the end.
				# result: 'availableNameofsomethings'
				"nameOfSomething",  # value stored in matching property name on class
				"Name of something",
			)
		]

	def clearRuntimeSettingAvailability(self):
		self._availableRuntimeSettings = []

	def addRuntimeSettingsAvailibility(self, settingIDs: List[str]):
		self._availableRuntimeSettings.extend(settingIDs)
		# ensure any previously saved settings are loaded from config file:
		self._initSpecificSettings(self, self._getAvailableRuntimeSettings())

	def _hasFeature(self, settingID: str) -> bool:
		return settingID in self._availableRuntimeSettings

	def _getAvailableRuntimeSettings(self) -> SupportedSettingType:
		settings = []
		if self._hasFeature("runtimeOnlySetting_externalValueLoad"):
			settings.extend([
				driverHandler.NumericDriverSetting(
					"runtimeOnlySetting_externalValueLoad",  # value stored in matching property name on class
					"Runtime Only amount, external value load",
					# no GUI default
				),
			])
		if self._hasFeature("runtimeOnlySetting_localDefault"):
			settings.extend([
				driverHandler.NumericDriverSetting(
					"runtimeOnlySetting_localDefault",  # value stored in matching property name on class
					"Runtime Only amount, local default",
					defaultVal=50,
				),
			])
		return settings

	def _get_supportedSettings(self) -> SupportedSettingType:
		settings = []
		settings.extend(self.preInitSettings)
		settings.extend(self._getAvailableRuntimeSettings())
		return settings


class AutoGuiTestProvider(vision.providerBase.VisionEnhancementProvider):
	_settings = AutoGuiTestSettings()

	@classmethod
	def canStart(cls):
		return True  # Check any dependencies (Windows version, Hardware access, Installed applications)

	@classmethod
	def getSettingsPanelClass(cls) -> Optional[Type]:
		"""Returns the instance to be used in order to construct a settings panel for the provider.
		@return: Optional[SettingsPanel]
		@remarks: When None is returned, L{gui.settingsDialogs.VisionProviderSubPanel_Wrapper} is used.
		"""
		return None  # No custom GUI

	@classmethod
	def getSettings(cls) -> AutoGuiTestSettings:
		return cls._settings

	def __init__(self):
		super().__init__()
		self._initRuntimeOnlySettings()
		self._showCurrentConfig()

	def _initRuntimeOnlySettings(self):
		""" This method might query another application for its capabilities and initialise these configuration
			options.
		"""
		settings = self.getSettings()
		settings.addRuntimeSettingsAvailibility([
			"runtimeOnlySetting_localDefault",
			"runtimeOnlySetting_externalValueLoad"
		])

		# load and set values from the external source, this will override values loaded from config.
		settings.runtimeOnlySetting_externalValueLoad = self._getValueFromDeviceOrOtherApplication(
				"runtimeOnlySetting_externalValueLoad"
			)

	def _getValueFromDeviceOrOtherApplication(self, settingId: str) -> Any:
		""" This method might connect to another application / device and fetch default values."""
		if settingId == "runtimeOnlySetting_externalValueLoad":
			return 75
		return None

	def _showCurrentConfig(self):
		"""Simple mechanism to test updating values."""
		result = (
			f"AutoGuiTestProvider:\n"
			f"x: {self._settings.shouldDoX}\n"
			f"y: {self._settings.shouldDoY}\n"
			f"z: {self._settings.amountOfZ}\n"
			f"name: {self._settings.nameOfSomething}\n"
			f"runtimeOnlySetting_externalValueLoad: {self._settings.runtimeOnlySetting_externalValueLoad}\n"
			f"runtimeOnlySetting_localDefault: {self._settings.runtimeOnlySetting_localDefault}\n"
		)
		wx.MessageBox(result, caption="started")

	def terminate(self):
		self._settings.clearRuntimeSettingAvailability()
		super().terminate()

	def registerEventExtensionPoints(self, extensionPoints):
		pass


VisionEnhancementProvider = AutoGuiTestProvider
