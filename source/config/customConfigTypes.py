# -*- coding: UTF-8 -*-
#A part of NonVisual Desktop Access (NVDA)
#Copyright (C) 2016 NV Access Limited, Babbage B.V.
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.

"""Implements validator functions for custom configObj types."""

import validate
from logHandler import log

def is_fixed_togglable_string_list(value, default=None):
	"""
	Check that L{value} is a list of strings.
	Every string should either start with "+" (enable) or "-" (disable)
	Every list item, regardless whether supposed to enable or disable, should be part of L{default}.
	"""
	if isinstance(value, basestring):
		raise validate.VdtTypeError(value)
	try:
		valueLength = len(value)
	except TypeError:
		raise validate.VdtTypeError(value)
	# validate check functions do not check the default, so this one doesn't either.
	defaultLength = len(default)
	if valueLength < defaultLength:
		raise validate.VdtValueTooShortError(value)
	if valueLength > defaultLength:
		raise validate.VdtValueTooLongError(value)
	for item in value:
		if not isinstance(item, basestring):
			raise validate.VdtTypeError(item)
		if not item.startswith(("+", "-")) or item[1:] not in (s[1:] for s in default):
			raise validate.VdtValueError(item)
	return list(value)
