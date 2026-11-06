# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import os
import re

replacements = {
	r'\b_is_renaming_sheet\b': '_isRenamingSheet',
	r'\b_monitoring_active\b': '_isMonitoringActive',
	r'\bis_slotted\b': 'isSlotted',
	r'\b_drift_timer_running\b': '_isDriftTimerRunning',
	r'\bis_match\b': 'isMatch',
	r'\bnative_success\b': 'isNativeSuccess',
	r'\bis_merged\b': 'isMerged',
	r'\bis_hiding\b': 'isHiding'
}

count = 0
for root, dirs, files in os.walk('addon/'):
	for file in files:
		if file.endswith('.py'):
			filepath = os.path.join(root, file)
			with open(filepath, 'r', encoding='utf-8') as f:
				content = f.read()
			
			new_content = content
			for pattern, replacement in replacements.items():
				new_content = re.sub(pattern, replacement, new_content)
				
			if new_content != content:
				with open(filepath, 'w', encoding='utf-8') as f:
					f.write(new_content)
				count += 1
				
print(f"Refactored boolean variables in {count} files.")
