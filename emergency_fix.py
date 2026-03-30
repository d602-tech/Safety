import os

file_path = r'd:\AI\GI01SafetyWalk\frontend\js\app.js'

with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# 1. Remove redundant submitEditCaseInfo (lines 1480-1504)
# Note: indices are 0-based, so line 1480 is index 1479
start_redundant = 1479
end_redundant = 1503 # line 1504

# Double check content before deleting
if 'submitEditCaseInfo: async (caseId) => {' in lines[start_redundant]:
    print(f"Removing redundant submitEditCaseInfo at lines {start_redundant+1}-{end_redundant+1}")
    del lines[start_redundant:end_redundant+1]
else:
    print("WARNING: Could not find redundant submitEditCaseInfo at expected lines.")

# 2. Fix the premature closing brace and move it to the end
# The premature brace was around line 2339 (but shifted due to deletion above)
# Deletion removed 25 lines.
premature_idx = 2339 - 25 - 1 # Adjusted index

target_line = '};'
found_premature = False
# Search around the expected area
for i in range(premature_idx - 10, premature_idx + 10):
    if i < len(lines) and lines[i].strip() == '};':
        # Check if it's the one before window.onload or the premature one
        # The premature one is before "/** ======================== 第 10 次優化"
        if i + 1 < len(lines) and '第 10 次優化' in lines[i+1]:
            print(f"Removing premature brace at line {i+1}")
            del lines[i]
            found_premature = True
            break

if not found_premature:
    print("WARNING: Could not find premature brace.")

# 3. Add closing brace before window.onload
# Search for window.onload
onload_idx = -1
for i in range(len(lines)-1, -1, -1):
    if 'window.onload = () => app.initAuth();' in lines[i]:
        onload_idx = i
        break

if onload_idx != -1:
    # Ensure there's a }; before it
    if lines[onload_idx-1].strip() != '};':
        print(f"Adding closing brace before window.onload at line {onload_idx+1}")
        lines.insert(onload_idx, '};\n\n')
else:
    print("WARNING: Could not find window.onload.")

with open(file_path, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Emergency repair applied.")
