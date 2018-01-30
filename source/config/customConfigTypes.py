# -*- coding: UTF-8 -*-
#A part of NonVisual Desktop Access (NVDA)
#Copyright (C) 2016 NV Access Limited, Babbage B.V.
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.

"""Implements validator functions for custom configObj types."""

import validate
from logHandler import log

def is_fixed_string_list(value, default=None):
	"""
	Check that L{value} is a list of strings,
	that every list item is part of L{default},
	and that every item in L{default} is part of L{value}
	"""
	if isinstance(value, basestring):
		raise validate.VdtTypeError(value)
	log.error(default)
	try:
		valueLenght = len(value)
	except TypeError:
		raise validate.VdtTypeError(value)
	return validate.is_string_list(value)
