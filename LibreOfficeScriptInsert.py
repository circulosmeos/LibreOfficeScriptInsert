# Add a python script to a LibreOffice Calc document
#
# Use: 
# $ python LibreOfficeScriptInsert.py <document.ods> <script.py>
#
# by circulosmeos, 2017-04
# https://github.com/circulosmeos/LibreOfficeScriptInsert
# licensed under GPLv3
#

import zipfile
import tempfile
import shutil
import os
import sys
import re

# http://stackoverflow.com/questions/4653768/overwriting-file-in-ziparchive
def remove_from_zip(zipfname, *filenames):
    tempdir = tempfile.mkdtemp()
    # print(tempdir)
    try:
        tempname = os.path.join(tempdir, 'new.zip')
        with zipfile.ZipFile(zipfname, 'r') as zipread:
            with zipfile.ZipFile(tempname, 'w') as zipwrite:
                for item in zipread.infolist():
                    if item.filename not in filenames:
                        data = zipread.read(item.filename)
                        zipwrite.writestr(item, data)
        shutil.move(tempname, zipfname)
    finally:
        shutil.rmtree(tempdir)

# .................................................

if len(sys.argv) != 3:
    exit("Use: $ python LibreOfficeScriptInsert.py <document.ods> <script.py>")

ODS_FILE    = sys.argv[1]
PYTHON_FILE = sys.argv[2]

filename = ODS_FILE
filename = re.sub(".ods", ".with_script.ods", filename, flags=re.I)
if (len(filename) == len(ODS_FILE)):
    exit("Aborted: Libreoffice file do not have .ods extension.")

if os.path.isfile(filename)==True:
    exit("Aborted: target file already exists: " + filename)

print("LibreOffice file: " + ODS_FILE)
shutil.copyfile(ODS_FILE, filename)

doc = zipfile.ZipFile(filename, 'a')
doc.write(PYTHON_FILE, "Scripts/python/" + PYTHON_FILE)
manifest = []
for line in doc.open('META-INF/manifest.xml'):
  if '</manifest:manifest>' in line.decode('utf-8'):
    for path in ['Scripts/', 'Scripts/python/', 'Scripts/python/' + PYTHON_FILE]:
      manifest.append(' <manifest:file-entry manifest:media-type="application/binary" manifest:full-path="%s"/>' % path)
      manifest.append("\n")
  manifest.append(line.decode('utf-8'))
doc.close()

remove_from_zip(filename, 'META-INF/manifest.xml')

doc = zipfile.ZipFile(filename, 'a')
doc.writestr('META-INF/manifest.xml', ''.join(manifest))

print("File created: " + filename)

