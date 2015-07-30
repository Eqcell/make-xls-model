# -*- coding: utf-8 -*-
"""
Created on Fri Jul 31 01:37:09 2015

@author: Евгений
"""

import win32api
e_msg = win32api.FormatMessage(-2146827284)
print(e_msg)