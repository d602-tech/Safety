import os

file_path = r'd:\AI\GI01SafetyWalk\frontend\js\app.js'

with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Search for the premature brace specifically before the new section
found = False
for i in range(len(lines)):
    if lines[i].strip() == '};' and i + 3 < len(lines) and '第 10 次優化' in lines[i+3]:
        print(f"Removing premature brace at line {i+1}")
        del lines[i]
        found = True
        break

if not found:
    # Try another way: find submitRegisterDeptAccount and look for }; after it
    for i in range(len(lines)):
        if 'submitRegisterDeptAccount: async () => {' in lines[i]:
            # Look for the next }; which is NOT followed by a comma
            for j in range(i+1, len(lines)):
                if lines[j].strip() == '};':
                    print(f"Removing premature brace at line {j+1}")
                    del lines[j]
                    found = True
                    break
            if found: break

if found:
    with open(file_path, 'w', encoding='utf-8') as f:
        f.writelines(lines)
    print("Fix applied successfully.")
else:
    print("Could not find the target brace.")
