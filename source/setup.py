# -*- coding: UTF-8 -*-
#setup.py
#A part of NonVisual Desktop Access (NVDA)
#Copyright (C) 2006-2018 NV Access Limited, Peter Vágner, Joseph Lee
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.

import os
import copy
import gettext
gettext.install("nvda", unicode=True)
from glob import glob
import fnmatch
from versionInfo import *
from cx_Freeze import setup, build_exe, Executable
import wx
import imp

MAIN_MANIFEST_EXTRA = r"""
<file name="brailleDisplayDrivers\handyTech\HtBrailleDriverServer.dll">
	<comClass
		description="HtBrailleDriver Class"
		clsid="{209445BA-92ED-4AB2-83EC-F24ACEE77EE0}"
		threadingModel="Apartment"
		progid="HtBrailleDriverServer.HtBrailleDriver"
		tlbid="{33257EFB-336F-4680-B94E-F5013BA6B9B3}" />
</file>
<file name="brailleDisplayDrivers\handyTech\HtBrailleDriverServer.tlb">
	<typelib tlbid="{33257EFB-336F-4680-B94E-F5013BA6B9B3}"
		version="1.0"
		helpdir="" />
</file>
<comInterfaceExternalProxyStub
	name="IHtBrailleDriverSink"
	iid="{EF551F82-1C7E-421F-963D-D9D03548785A}"
	proxyStubClsid32="{00020420-0000-0000-C000-000000000046}"
	baseInterface="{00000000-0000-0000-C000-000000000046}"
	tlbid="{33257EFB-336F-4680-B94E-F5013BA6B9B3}" />
<comInterfaceExternalProxyStub
	name="IHtBrailleDriver"
	iid="{43A71F9B-58EE-42D4-B58E-0F9FBA28D995}"
	proxyStubClsid32="{00020424-0000-0000-C000-000000000046}"
	baseInterface="{00000000-0000-0000-C000-000000000046}"
	tlbid="{33257EFB-336F-4680-B94E-F5013BA6B9B3}" />
<compatibility xmlns="urn:schemas-microsoft-com:compatibility.v1">
	<application>
		<!-- Windows Vista -->
		<supportedOS Id="{e2011457-1546-43c5-a5fe-008deee3d3f0}"/>
		<!-- Windows 7 -->
		<supportedOS Id="{35138b9a-5d96-4fbd-8e2d-a2440225f93a}"/>
		<!-- Windows 8 -->
		<supportedOS Id="{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}"/>
		<!-- Windows 8.1 -->
		<supportedOS Id="{1f676c76-80e1-4239-95bb-83d0f6d0da78}"/>
		<!-- Windows 10 -->
		<supportedOS Id="{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}"/>
	</application> 
</compatibility>
"""

def getModuleExtention(thisModType):
	for ext,mode,modType in imp.get_suffixes():
		if modType==thisModType:
			return ext
	raise ValueError("unknown mod type %s"%thisModType)

class build_nvda_exe(build_exe):
	"""Overridden cx_Freeze command to:
		* Add a command line option --enable-uiAccess to enable uiAccess for the main executable
		* Add extra info to the manifest
	"""

	user_options = build_exe.user_options + [
		("enable-uiAccess", "u", "enable uiAccess for the main executable"),
	]

	def initialize_options(self):
		build_exe.initialize_options(self)
		self.enable_uiAccess = False

	def _run(self):
		dist = self.distribution
		if self.enable_uiAccess:
			# Add a target for nvda_uiAccess, using nvda_noUIAccess as a base.
			target = Executable(**dist.executables[0].__dict__)
			target.targetName = "nvda_uiAccess.exe"
			#target["uac_info"] = (target["uac_info"][0], True)
			dist.executables.insert(1, target)
			# nvda_eoaProxy should have uiAccess.
			target = Executable(**dist.executables[3].__dict__)
			#target["uac_info"] = (target["uac_info"][0], True)

		build_exe.run(self)

	def build_manifest(self, target, template):
		mfest, rid = build_exe.py2exe.build_manifest(self, target, template)
		if getattr(target, "script", "").endswith(".pyw"):
			# This is one of the main application executables.
			mfest = mfest[:mfest.rindex("</assembly>")]
			mfest += MAIN_MANIFEST_EXTRA + "</assembly>"
		return mfest, rid

def getLocaleDataFiles():
	wxDir=wx.__path__[0]
	localeMoFiles=set()
	for f in glob("locale/*/LC_MESSAGES"):
		localeMoFiles.add((f, (os.path.join(f,"nvda.mo"),)))
		wxMoFile=os.path.join(wxDir,f,"wxstd.mo")
		if os.path.isfile(wxMoFile):
			localeMoFiles.add((f,(wxMoFile,))) 
		lang=os.path.split(os.path.split(f)[0])[1]
		if '_' in lang:
				lang=lang.split('_')[0]
				f=os.path.join('locale',lang,'lc_messages')
				wxMoFile=os.path.join(wxDir,f,"wxstd.mo")
				if os.path.isfile(wxMoFile):
					localeMoFiles.add((f,(wxMoFile,))) 
	localeDicFiles=[(os.path.dirname(f), (f,)) for f in glob("locale/*/*.dic")]
	NVDALocaleGestureMaps=[(os.path.dirname(f), (f,)) for f in glob("locale/*/gestures.ini")]
	return list(localeMoFiles)+localeDicFiles+NVDALocaleGestureMaps

def getRecursiveDataFiles(dest,source,excludes=()):
	rulesList=[]
	rulesList.append((dest,
		[f for f in glob("%s/*"%source) if not any(fnmatch.fnmatch(f,exclude) for exclude in excludes) and os.path.isfile(f)]))
	[rulesList.extend(getRecursiveDataFiles(os.path.join(dest,dirName),os.path.join(source,dirName),excludes=excludes)) for dirName in os.listdir(source) if os.path.isdir(os.path.join(source,dirName)) and not dirName.startswith('.')]
	return rulesList

compiledModExtention = getModuleExtention(imp.PY_COMPILED)
sourceModExtention = getModuleExtention(imp.PY_SOURCE)
setup(
	name = name,
	version="%s.%s.%s.%s"%(version_year,version_major,version_minor,version_build),
	description=description,
	url=url,
	classifiers=[
'Development Status :: 3 - Alpha',
'Environment :: Win32 (MS Windows)',
'Topic :: Adaptive Technologies'
'Intended Audience :: Developers',
'Intended Audience :: End Users/Desktop',
'License :: OSI Approved :: GNU General Public License (GPL)',
'Natural Language :: English',
'Programming Language :: Python',
'Operating System :: Microsoft :: Windows',
],
	cmdclass={"build_exe": build_nvda_exe},
	executables=[
		Executable(
			script="nvda.pyw",
			targetName="nvda_noUIAccess.exe",
			#uac_info=("asInvoker", False),
			icon="images/nvda.ico",
			#version=version,
			#description="NVDA application",
			#product_version=version,
			copyright=copyright,
			#company_name=publisher,
		),
		# The nvda_uiAccess target will be added at runtime if required.
		Executable(
			script="nvda_slave.pyw",
			icon="images/nvda.ico",
			#version="%s.%s.%s.%s"%(version_year,version_major,version_minor,version_build),
			#description=name,
			#product_version=version,
			copyright=copyright,
			#company_name=publisher,
		),
		Executable(
			script="nvda_eoaProxy.pyw",
			# uiAccess will be enabled at runtime if appropriate.
			#uac_info=("asInvoker", False),
			icon="images/nvda.ico",
			#version="%s.%s.%s.%s"%(version_year,version_major,version_minor,version_build),
			#description="NVDA Ease of Access proxy",
			#product_version=version,
			copyright=copyright,
			#company_name=publisher,
		),
	],
	options = {"build_exe": {
		"excludes": ["Tkinter",
			"serial.loopback_connection", "serial.rfc2217", "serial.serialcli", "serial.serialjava", "serial.serialposix", "serial.socket_connection"],
		"packages": ["NVDAObjects","virtualBuffers","appModules","comInterfaces","brailleDisplayDrivers","synthDrivers"],
		# #3368: bisect was implicitly included with Python 2.7.3, but isn't with 2.7.5.
		# Also, the service executable used win32api, which some add-ons use for various purposes.
		# Explicitly include them so we don't break some add-ons.
		"includes": ["nvdaBuiltin", "bisect", "win32api"],
		"zip_include_packages": "*",
		"zip_exclude_packages": "",
	}},
	data_files=[
		(".",glob("*.dll")+glob("*.manifest")+["builtin.dic"]),
		("documentation", ['../copying.txt', '../contributors.txt']),
		("lib/%s"%version, glob("lib/*.dll")),
		("lib64/%s"%version, glob("lib64/*.dll") + glob("lib64/*.exe")),
		("waves", glob("waves/*.wav")),
		("images", glob("images/*.ico")),
		("louis/tables",glob("louis/tables/*")),
		("COMRegistrationFixes", glob("COMRegistrationFixes/*.reg")),
		(".", ['message.html' ])
	] + (
		getLocaleDataFiles()
		+ getRecursiveDataFiles("synthDrivers", "synthDrivers",
			excludes=("*%s" % sourceModExtention, "*%s" % compiledModExtention, "*.exp", "*.lib", "*.pdb"))
		+ getRecursiveDataFiles("brailleDisplayDrivers", "brailleDisplayDrivers", excludes=("*%s"%sourceModExtention,"*%s"%compiledModExtention))
		+ getRecursiveDataFiles('documentation', '../user_docs', excludes=('*.t2t', '*.t2tconf', '*/developerGuide.*'))
	),
)
