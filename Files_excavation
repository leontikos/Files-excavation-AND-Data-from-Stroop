import os
import shutil

# choose path
PTH = r"C:\Wytykowska_badanie"
os.chdir(PTH)
print(os.getcwd())

# excavate multiple xlsx files from multiple folders and copy them with attributes to one folder
dst = 'C:\Wytyk_ALL'
for root, dirs, files in os.walk(PTH):
    for file in files:
        if file.endswith(".xlsx"):
            print(os.path.join(root, file) + ' - DONE')
            shutil.copy2(os.path.join(root, file), dst)
