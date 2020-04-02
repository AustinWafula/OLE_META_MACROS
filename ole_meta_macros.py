import olefile
import os
import sys
import shutil

def detect_Macros():
    for root,dir,files in os.walk(sys.argv[1]):
        for file in files :
            try:
                with olefile.OleFileIO(file,raise_defects=olefile.DEFECT_INCORRECT) as ole_win:
                    if ole_win.exists('VBA') or ole_win.exists('Macro') :
                        print(file)
                        print(file+":macro/vba detected")
                        print(os.path.abspath(file))
                        meta=ole_win.get_metadata()
                        if sys.argv[2] != None and sys.argv[2] !='meta':
                            shutil.copy(file,sys.argv[2])
                        print('Macrcopied to',sys.argv[2])
                        print(meta.dump())
                        print('')
            except :
                pass
def obtain_meta():
        for root,dir,files in os.walk(sys.argv[1]):
            for file in files :
                try:
                    ole_win=olefile.OleFileIO(file,raise_defects=olefile.DEFECT_INCORRECT)
                    print(file)
                    print(os.path.abspath(file))
                    meta=ole_win.get_metadata()
                    print(meta.dump())
                    print('')
                except :
                    pass
    
if sys.argv[2]=='meta':
    obtain_meta()
else:
    detect_Macros()
