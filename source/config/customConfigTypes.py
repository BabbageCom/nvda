# -*- coding: UTF-8 -*-
#A part of NonVisual Desktop Access (NVDA)
#Copyright (C) 2016 NV Access Limited, Babbage B.V.
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.

"""Implements validator functions for custom configObj types."""

import validate
from logHandler import log

def is_default_list(value, default=None, min=None, max=None):
	log.error(value)
	log.error(default)
	return validate.is_string_list(value, min=min, max=max)
