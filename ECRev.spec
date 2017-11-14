# -*- mode: python -*-


app_directory = 'C:\Users\sdb25\Desktop\Mydocs\New Trieal'

import sys

sys.path.append(app_directory)

import ECRev

app_name = "%05.1f"%float(ECRev.version)
app_name = app_name.replace('.', '') + ".ECRev.exe"

print app_name


def Datafiles(*filenames, **kw):
    import os
    
    def datafile(path, strip_path=True):
        parts = path.split('/')
        path = name = os.path.join(*parts)
        if strip_path:
            name = os.path.basename(path)
        return name, path, 'DATA'

    strip_path = kw.get('strip_path', True)
    return TOC(
        datafile(filename, strip_path=strip_path)
        for filename in filenames
        if os.path.isfile(filename))



a = Analysis(['{}\\ECRev.py'.format(app_directory)],
             pathex=[app_directory, 'C:\Python27\Lib\site-packages'],
             hiddenimports=['pyodbc'],
             hookspath=None)

a.datas += [('interface.xrc', 'interface.xrc', 'DATA')]
a.datas += [('CommitteeECRs.xlsm', 'CommitteeECRs.xlsm', 'DATA')]
a.datas += [('icons\\internet-news-reader.png', 'icons\\internet-news-reader.png', 'DATA')]
a.datas += [('icons\\preferences-desktop.png', 'icons\\preferences-desktop.png', 'DATA')]
a.datas += [('icons\\software-update-available.png', 'icons\\software-update-available.png', 'DATA')]
a.datas += [('icons\\system-log-out.png', 'icons\\system-log-out.png', 'DATA')]


#docfiles = Datafiles('interface.xrc', 'icons\\internet-news-reader.png', 'icons\\preferences-desktop.png', 'icons\\software-update-available.png', 'icons\\system-log-out.png')


pyz = PYZ(a.pure)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name=app_name,
          debug=False,
          strip=None,
          upx=True,
          icon='ECRev.ico',
          console=False )

'''
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               docfiles,
               name='distFinal')
'''

