# GUI Application automation and testing library
# Copyright (C) 2006-2018 Mark Mc Mahon and Contributors
# https://github.com/pywinauto/pywinauto/graphs/contributors
# http://pywinauto.readthedocs.io/en/latest/credits.html
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# * Redistributions of source code must retain the above copyright notice, this
#   list of conditions and the following disclaimer.
#
# * Redistributions in binary form must reproduce the above copyright notice,
#   this list of conditions and the following disclaimer in the documentation
#   and/or other materials provided with the distribution.
#
# * Neither the name of pywinauto nor the names of its
#   contributors may be used to endorse or promote products derived from
#   this software without specific prior written permission.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
# FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
# DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
# SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
# CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
# OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
# OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

from __future__ import print_function

try:
  from pywinauto import application
except ImportError:
  import os.path
  pywinauto_path = os.path.abspath(__file__)
  pywinauto_path = os.path.split(os.path.split(pywinauto_path)[0])[0]
  import sys
  sys.path.append(pywinauto_path)
  from pywinauto import application

from pywinauto.controls.hwndwrapper import HwndWrapper
from pywinauto import WindowAmbiguousError

import sys
import time
import re
import clipboard

FIREFOX_PATH = r"C:\Program Files\Mozilla Firefox\firefox.exe {}"
WORD_PATH = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"

def screenshot(filename):
	web_addresses = []
	with open(filename, 'r', encoding='utf-8') as f:	
		web_addresses = f.readlines()

	word_app = application.Application().start(WORD_PATH)
	word_app.Word.type_keys("%NL")
	word = word_app['Document1 - Word']
	time.sleep(1)
	
	for web_address in web_addresses:
		if not re.search(r"http", web_address):
			continue
			
		web_address = web_address.strip()
		clipboard.copy(web_address)
		word.type_keys("^v+{ENTER}")
		time.sleep(0.5)
		
		browser_app = application.Application().start(FIREFOX_PATH.format(web_address))
		time.sleep(4)					# wait for webpage to load
			
		if browser_app.windows():
			mozilla = browser_app.window(title_re=".*Mozilla Firefox")
		else:			
			browser_app = application.Application().connect(title_re=".*Mozilla Firefox")
			mozilla = browser_app.window(title_re=".*Mozilla Firefox")
      
		mozilla.type_keys("{PRTSC}")	# take screenshot
		time.sleep(0.5)
		mozilla.type_keys("%FC")		# exit firefox

		word.type_keys("^v{ENTER}{ENTER}")
		time.sleep(1)

if __name__ == "__main__":
	screenshot("linkfile.txt")
