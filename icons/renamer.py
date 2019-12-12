import os
li = []

for filename in os.listdir(os.getcwd()):
    if 'icons8' in filename:
        name = filename.replace('icons8-','')
        name = name.replace('-50','')
        li.append('<file>icons/' + name + '</file>')
        os.rename(filename,name)

    with open('qrc.txt', 'w') as f:
        for item in li:
            f.write("%s\n" % item)