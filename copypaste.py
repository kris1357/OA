import os
import shutil


originalPath = "D:\\Script\\VBA\\"
file = "test.xlsx"
originalFile = originalPath + file
targetPath = "D:\\Script\\VBA\\dir2\\"
targetFile = targetPath + file

if os.path.isfile(originalFile):
    if not os.path.isfile(targetFile):
        shutil.copyfile(originalFile, targetFile)
        print("Done")
    else:
        print("Already exist")
else:
    print("Not exists")


