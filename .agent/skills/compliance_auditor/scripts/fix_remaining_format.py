# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import os

# fix bare excepts
fpath = 'addon/appModules/boa_enhancements/excel_enhancements/formula_auditor.py'
with open(fpath, 'r', encoding='utf-8') as f:
	content = f.read()
content = content.replace('except:\n', 'except Exception:\n')
content = content.replace('else: break\n', 'else:\n\t\t\t\t\t\t\t\tbreak\n')
with open(fpath, 'w', encoding='utf-8') as f:
	f.write(content)

# fix sheet layout analyzer
fpath2 = 'addon/appModules/boa_enhancements/excel_enhancements/sheet_layout_analyzer.py'
with open(fpath2, 'r', encoding='utf-8') as f:
	content = f.read()

content = content.replace('else: break\n', 'else:\n\t\t\t\t\t\t\tbreak\n')
content = content.replace('else: hidden_borders.append', 'else:\n\t\t\t\t\t\t\thidden_borders.append')

with open(fpath2, 'w', encoding='utf-8') as f:
	f.write(content)

print("Fixed the last few formatting issues.")
