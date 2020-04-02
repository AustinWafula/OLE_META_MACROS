import olefile
import os
import sys
import shutil
import hashlib

BUFFER_SIZE=8899
file_l=[]
def detect_Macros():
    for root,dir,files in os.walk(sys.argv[1]):
        for file in files :
            try:
                with olefile.OleFileIO(file,raise_defects=olefile.DEFECT_INCORRECT) as ole_win:
                    if ole_win.exists('VBA') or ole_win.exists('Macro') :
                        if file  not in file_l:
                            file_l.append(file)
                            file_hash_SHA256= hashlib.sha256()
                            file_hash_MD5=hashlib.md5()
                            with open(file,'rb') as f:
                                file_bs=f.read(BUFFER_SIZE)
                                while len(file_bs)>0:
                                    file_hash_SHA256.update(file_bs)
                                    file_hash_MD5.update(file_bs)
                                    file_bs=f.read(BUFFER_SIZE)
                            print(file)
                            print(file+":macro/vba detected")
                            print(os.path.abspath(file))
                            print('SHA-256:',file_hash_SHA256.hexdigest())
                            print('MD5:',file_hash_MD5.hexdigest())
                            meta=ole_win.get_metadata()
                            if sys.argv[2] != None and sys.argv[2] !='meta':
                                shutil.copy(file,sys.argv[2])
                            print('Macrcopied to',sys.argv[2])
                            print(meta.dump())
                            print('')
            except:
                pass
def obtain_meta():
        for root,dir,files in os.walk(sys.argv[1]):
            for file in files :
                try:
                    if file not in file_l:
                        file_l.append(file)
                        ole_win=olefile.OleFileIO(file,raise_defects=olefile.DEFECT_INCORRECT)
                        file_hash_SHA256= hashlib.sha256()
                        file_hash_MD5=hashlib.md5()
                        with open(file,'rb') as f:
                            file_bs=f.read(BUFFER_SIZE)
                            while len(file_bs)>0:
                                file_hash_SHA256.update(file_bs)
                                file_hash_MD5.update(file_bs)
                                file_bs=f.read(BUFFER_SIZE)
                        print(file)
                        print(os.path.abspath(file))
                        print("SHA-256:",file_hash_SHA256.hexdigest())
                        print('MD5:',file_hash_MD5.hexdigest())
                        meta=ole_win.get_metadata()
                        print(meta.dump())
                        print('')
                except:
                    pass
                
if sys.argv[2]=='meta':
    obtain_meta()
else:
    detect_Macros()
            
            
            
    
