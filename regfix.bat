reg delete "HKCR\CLSID\{6EC97137-BE18-44B9-BB5B-92240A8D3481}\InprocServer32\4.5.5.1556" /f >> "C:\TEMP\HPRM_LOG\regfix.txt"
reg delete "HKEY_CLASSES_ROOT\CLSID\{6EC97137-BE18-44B9-BB5B-92240A8D3481}\InprocServer32\4.5.5.1556" /f >> "C:\TEMP\HPRM_LOG\regfix.txt"
regedit /s delete_old_ke.reg

