import zipfile
import shutil
import os
import re
import glob


def unlock(filename):
    with zipfile.ZipFile(filename) as file:
        file.extractall(path="temp")
        with open("temp/ppt/presentation.xml", "r+") as textfile:
            textinside = textfile.read()
            try:
                fixedtextiinside = textinside.replace(re.search("(?:<p:modifyVerifier).*?(/>)",textinside).group(0),"")
            except AttributeError as e:
                shutil.rmtree("temp")
                print("Already Converted: ",filename)
                return 1
            except e:
                shutil.rmtree("temp")
                print("INTERNAL ERROR ",filename)
                return 2
            textfile.seek(0)
            textfile.write(fixedtextiinside)

    shutil.make_archive(filename,"zip","temp")
    os.rename(filename+".zip",filename)
    shutil.rmtree("temp")
    print("Converted: ",filename)


for filepath in glob.iglob('convert/*.pptx'):
    print("Converting: ",filepath)
    unlock(filepath)
