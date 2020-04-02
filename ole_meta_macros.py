import olefile
import os
import sys
import shutil
import hashlib
from virustotal_python import Virustotal
import json
from pprint import pprint

BUFFER_SIZE=8899
file_l=[]
virus_total = Virustotal("#virus_total api key")
def hash_f():
    global ole_win
    file_hash_SHA256= hashlib.sha256()
    file_hash_MD5=hashlib.md5()
    for file in file_l:
        with open(file,'rb') as f:
            file_bs=f.read(BUFFER_SIZE)
            while len(file_bs)>0:
                file_bs=f.read(BUFFER_SIZE)
                file_hash_SHA256.update(file_bs)
                file_hash_MD5.update(file_bs)
        print(file)
        print(file+":macro/vba detected")
        print(os.path.abspath(file))
        print('SHA-256:',file_hash_SHA256.hexdigest())
        print('MD5:',file_hash_MD5.hexdigest())
        meta=ole_win.get_metadata()
        #if sys.argv[2] != None and sys.argv[2] !='meta':
        shutil.copy(file,sys.argv[2])
        print('Macro copied to',sys.argv[2])
        print(meta.dump())
        print('')
        v_resp = virus_total.file_report([file_hash_SHA256.hexdigest()])
        pprint(v_resp)
        with open(file+'_.json', 'w') as j_file:
            json.dump(v_resp, j_file,indent=1,ensure_ascii=False)
        shutil.copy(file+'_.json',sys.argv[2])
def detect_Macros():
    global ole_win
    for root,dir,files in os.walk(sys.argv[1]):
        for file in files :
            try:
                ole_win=olefile.OleFileIO(file,raise_defects=olefile.DEFECT_INCORRECT)
                if ole_win.exists('VBA') or ole_win.exists('WordDocument') and file not in file_l:
                    file_l.append(file)
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
    hash_f()
            
            
            
    
