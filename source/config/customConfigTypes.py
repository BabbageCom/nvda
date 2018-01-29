# -*- coding: UTF-8 -*-
#A part of NonVisual Desktop Access (NVDA)
#Copyright (C) 2016 NV Access Limited, Babbage B.V.
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.

"""Implements validator functions for custom configObj types."""

from validate import *
from logHandler import log

def is_fixed_string_list(value, options=None):
	"""
	Check that L{value} is a list of strings,
	that every list item is part of L{options},
	and that every item in L{options} is part of L{value}
	"""
	if isinstance(value, basestring):
		raise VdtTypeError(value)
	log.error(options)
	try:
		valueLenght = len(value)
	except TypeError:
		raise VdtTypeError(value)
	return is_string_list(value)
