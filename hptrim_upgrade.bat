:MSB_PRESTART
rem @echo off
cls
setlocal ENABLEDELAYEDEXPANSION
REM ______________________________________________________________
REM __
REM __  HPE Content Manager Deployment Script
REM __
REM __  HPE Content Manager v9.10.914
REM __
REM __  Jeremy Bloomstrom <jeremy.bloomtrom@matsugov.us>
REM __
REM __  version 2017.11.06
REM __


REM Set your server list here, separated by a space
set SERVERLIST=dsjkb
rem goto :GetBestServer
goto :Commands


REM Commands go here
:Commands
REM SET INSTALL PATH
rem if !bestping! equ dsjkb set InstallPath=\\dsjkb\desoft$\HPRecordsManager\Deployment
REM if !bestping! equ drlogos7sql set InstallPath=\\drlogos7sql\Software\TRIM
REM if !bestping! equ pwsw7918 set InstallPath=\\PWSW7918\c$\temp\trim

set InstallPath=\\dsjkb\desoft$\HPRecordsManager\Deployment91
echo InstallPath %InstallPath%


REM MAP FIRST AVAILABLE DRIVE LETTER
for %%a in (Z Y X W V U T S R Q P O N M L K J I H G F E D C B A) DO (
 set driveletter=%%a
 if not exist %%a:\. goto mapdrive
)

:mapdrive
echo Mapping %driveletter% to %InstallPath%
net use %driveletter%: %InstallPath% /y
%driveletter%:


REM -- KAPISH SCRIPT STARTS HERE --
REM --------------------------------------------------


:START_SCRIPT
                ECHO ---------------------------------------------------------------
                ECHO ----- TRIM Context to HP Records Manager Upgrade Starting -----
                ECHO ---------------------------------------------------------------
                ECHO.
                GOTO INITIALIZE


:INITIALIZE
                ECHO ------------------------------------------------------
                ECHO ----- Initializing Script Variables and Folders ------
                ECHO ------------------------------------------------------
                ECHO.

                SET INSTALLROOT=%ProgramFiles%
                SET LOGPATH=C:\TEMP\HPRM_LOG
                IF NOT EXIST %LOGPATH% MKDIR %LOGPATH%

                REM Apply VCRedist Fix
                REG ADD "HKLM\SOFTWARE\Wow6432Node\Microsoft\DevDiv\VC\Servicing\8.0" /v "SP" /t REG_DWORD /d 1 /f

				REM --- Cleanup Old Dataset Registry Settings ---
				rem cscript "Source\Config\SetAllUserHKCUDatasetSettings.vbs"

                GOTO UNINSTALL_TRIMCONTEXT

:UNINSTALL_TRIMCONTEXT
                ECHO -----------------------------------------------------------
                ECHO ----- Initiating Uninstall of HP TRIM Components      -----
                ECHO -----------------------------------------------------------
                ECHO.
                ECHO  HP TRIM x64 7.34.5739
                MSIEXEC /X{D8B2D69F-7FA0-4BC8-8E31-C675162229D1} /passive /norestart
                ECHO  HP TRIM x86 7.34.5739
                MSIEXEC /X{B480C3B2-9432-41B9-BD4A-421A4A6AB4C6} /passive /norestart
                ECHO  HP TRIM x86 7.12.2017
                MSIEXEC /X{8E5A092F-4EAA-4511-B6AF-0D285902EDBC} /passive /norestart
                ECHO  HP TRIM x64 7.12.2017
                MSIEXEC /X{42D7AD02-7270-4110-8512-E908122E961F} /passive /norestart
                ECHO  HP TRIM x86 7.12.1879
                MSIEXEC /X{B0B145AF-079B-43D0-80D8-C56339930964} /passive /norestart
                ECHO  HP TRIM x64 7.12.1879
                MSIEXEC /X{B9702CB4-1E85-4503-B8FF-E46255C2A37B} /passive /norestart
                ECHO  HP TRIM x86 7.11.1828
                MSIEXEC /X{AAB1C0BB-C8D6-4D70-9FF8-00856F6BB4D9} /passive /norestart
                ECHO  HP TRIM x86 7.11.1821
                MSIEXEC /X{0AB912BE-4DE4-4740-840C-A831E7242DCA} /passive /norestart
                ECHO  HP TRIM Microsoft runtimes for VS2008 x64 7.0.0.0
                MSIEXEC /X{753CC7F1-69E0-4C81-9F55-481DA54ADACE} /passive /norestart
                ECHO  HP TRIM Microsoft VB runtimes 7.0.0.0
                MSIEXEC /X{112268A9-B0FB-421C-BEDB-A08B32E84207} /passive /norestart
                ECHO  HP TRIM Microsoft runtimes for VS2008 x86 7.0.0.0
                MSIEXEC /X{B822B0A8-7145-474C-AE0E-FB7BC3B38E94} /passive /norestart

                REM -- VERSION 8.2 STUFF --
                ECHO    HP Records Manager Microsoft runtimes for VS2010 x64 8.20.7591
                MsiExec /X{D38E5627-C8B6-4204-BC83-0776266FFC6C} /passive /norestart
                ECHO    HP Records Manager Microsoft runtimes for VS2010 x86 8.20.7591
                MsiExec /X{AC789DFC-FBF6-49B1-A64C-EFF770809C53} /passive /norestart
                ECHO    HP Records Manager Microsoft runtimes for VS2013 x64 8.20.7591
                MsiExec /X{9D33A9C9-AEF0-45A9-A895-D20D772ECE4F} /passive /norestart
                ECHO    HP Records Manager Microsoft runtimes for VS2013 x86 8.20.7591
                MsiExec /X{BEBA8756-E029-4804-860E-648FE7B176A1} /passive /norestart

                ECHO    HP Records Manager x64  8.20.7591
                MsiExec /X{B822B0A8-7145-474C-AE0E-FB7BC3B38E94} /passive /norestart
                ECHO    HP Records Manager x64  8.20.7804
                MsiExec /X{A0CA7FD0-7CC5-4667-B357-72490A14F0F3} /passive /norestart
                ECHO    HP Records Manager x86  8.20.7591
                MsiExec /X{926E096A-30BA-420F-A7D1-AC923B44F9DD} /passive /norestart
                ECHO    HP Records Manager x86  8.20.7804
                MsiExec /X{8EE56206-600D-425A-B30F-B0A601F2FB39} /passive /norestart

:UNINSTALL_KAPISH
                ECHO -----------------------------------------------------------
                ECHO ----- Initiating Uninstall of Kapish Addons           -----
                ECHO -----------------------------------------------------------
                ECHO.



                ECHO  Kapish Explorer x64 4.55.1556
                MsiExec /X{4C9BD660-4294-41D9-AFF9-C7FC0B67A53D} /passive /norestart
                ECHO  Kapish Explorer x86 4.55.1556
                MsiExec /X{4FB1E602-27A6-4147-B726-347EE80549CE} /passive /norestart
                ECHO  Kapish Explorer x86 4.50.1536
                MsiExec /X{536F96E6-BD50-4E28-8B26-5736B861E1C7} /passive /norestart
                ECHO  Kapish Explorer x64 4.50.1536
                MsiExec /X{47C367E3-A076-4C71-A17F-3704EC7C54F3} /passive /norestart
                ECHO  Kapish Explorer x64 4.52.1546
                MsiExec /X{39A78900-31AA-4B2F-8D9C-7CDDCB8F69D7} /passive /norestart
                ECHO  Kapish Explorer x64 4.43.1354
                MSIEXEC /X{8A45DBB4-46B2-4129-9149-53EA32482E7E} /passive /norestart
                ECHO  Kapish Explorer x86 4.43.1354
                MSIEXEC /X{14DB2EC1-16DE-4BA6-B636-808A11FC056E} /passive /norestart
                ECHO  Kapish Explorer x64 4.40.1274
                MsiExec /X{75E0FB22-2205-4957-BE08-93FB517E6429} /passive /norestart
                ECHO  Kapish Explorer x86 4.40.1274
                MsiExec /X{928CA651-8B83-4D01-A528-BF98BEF35931} /passive /norestart
                ECHO  Kapish Explorer x64 4.40.1272
                MSIEXEC /X{75E0FB22-2205-4957-BE08-93FB517E6429} /passive /norestart
                ECHO  Kapish Explorer x64 4.35.1228
                MSIEXEC /X{823F96A8-CDC5-4F1D-B5F0-BE2AE49D36D6} /passive /norestart
                ECHO  Kapish Explorer x86 4.35.1228
                MSIEXEC /X{3F17465C-441B-4D60-9917-7B6D12F780CB} /passive /norestart
                ECHO  Kapish Explorer x86 4.30.1112
                MSIEXEC /X{D86A51BC-8B38-4F59-9F1B-1F029EC45FB8} /passive /norestart
                ECHO  Kapish Explorer x64 4.30.1112
                MSIEXEC /X{5882234A-5834-4C04-AEA7-C0FE27714A58} /passive /norestart
                ECHO  Kapish TRIM Explorer x86 for TRIM7.x 4.12.01057
                MSIEXEC /X{488204A3-017A-4A50-ACD4-D335096DF481} /passive /norestart
                ECHO  Kapish TRIM Explorer x64 for TRIM7.x 4.12.01057
                MSIEXEC /X{67E53865-1CDC-40D7-88C3-AE16CBC99C86} /passive /norestart
                ECHO  Kapish TRIM Explorer x86 for TRIM7.1 4.12.01051
                MSIEXEC /X{D27B40F8-4FC7-41D5-9576-C32F4A3A5C69} /passive /norestart
                ECHO  Kapish TRIM Explorer x64 for TRIM7.1 4.12.01051
                MSIEXEC /X{C0A1CAE1-4BE0-4341-A91F-B12E158C6A9C} /passive /norestart

                ECHO  Kapish TRIM Folder Wizard x64 3.00.1003
                MSIEXEC /X{93229831-E2D7-4C40-AE56-AB2EB098F243} /passive /norestart
                ECHO  Kapish TRIM Folder Wizard x86 3.00.1003
                MSIEXEC /X{A0FCB4FE-E4FB-4DF9-A131-1701CE7DF328} /passive /norestart
                ECHO  Kapish TRIM Folder Wizard x64 3.30.1108
                MSIEXEC /X{60F5031F-1C5C-4279-ACC2-9480E348B535} /passive /norestart
                ECHO  Kapish TRIM Folder Wizard x86 3.30.1108
                MSIEXEC /X{80ED09A5-5FAF-4FD1-BF49-95052789ECF7} /passive /norestart

                ECHO  Kapish Metrics Utility x86 2.00.2000
                MSIEXEC /X{BCE513B3-5F7E-46A8-8E52-3DE9B8849431} /passive /norestart
                ECHO  Kapish Metrics Utility x86 2.00.1000
                MSIEXEC /X{9608B7A0-070E-49FC-9C34-BF9B2D299270} /passive /norestart
                ECHO  Kapish Metrics Utility x86 2.00.0000
                MSIEXEC /X{3CBC0722-C842-4F9E-8884-9122B7638E38} /passive /norestart
                ECHO  Kapish Metrics Utility 1.1.0.1
                MSIEXEC /X{DE6875D8-C6CC-4CB1-AC7C-071130C561A5} /passive /norestart
                ECHO  Kapish Metrics Utility 1.0.0.1
                MSIEXEC /X{DE6875D8-C6CC-4CB1-AC7C-071130C561A5} /passive /norestart

                ECHO  Kapish TRIM Record Remover x86 1.12.1007
                MSIEXEC /X{B48DAAC9-C4B6-4857-9A63-E7DB331957E1} /passive /norestart
                ECHO  Kapish Record Remover x64 1.40.1106
                MSIEXEC /X{F56EFBE0-1239-4F06-AE70-AE75E2B4A5FE} /passive /norestart
                ECHO  Kapish Record Remover x86 1.40.1106
                MSIEXEC /X{BA7006FF-79F4-4D04-A006-29F0E40E57D4} /passive /norestart

                ECHO  Kapish TRIM Rename Addin 1.0.0
                MSIEXEC /X37DA07C3-CA53-4B6D-B1BC-170FCC711C6D /passive /norestart
                ECHO  Kapish TRIM Rename Addin 1.0.4
                MSIEXEC /X{E869857B-1F43-4FD1-A619-91D9A5967AFD} /passive /norestart
                ECHO  Kapish TRIM Rename Addin 1.0.3
                MSIEXEC /X{1F9619C1-3896-424C-B85E-3D60DD718DF3} /passive /norestart
                ECHO  Kapish TRIM Rename Addin 1.0.2
                MSIEXEC /X{2C91BAED-FB84-4756-ABCF-A2F4AB4871A7} /passive /norestart
                ECHO  Kapish TRIM Rename Addin 1.0.1
                MSIEXEC /X{E2F709B4-3339-4CAB-895E-EEC88AC155B9} /passive /norestart

                ECHO  Container Part Creator 1.0.0
                MSIEXEC /X{6EFF98F1-16C8-497A-A21F-BF56F13AE1AC} /passive /norestart

                ECHO  Kapish Tree Size Automation Server 1.0.0
                MSIEXEC /X{97719B65-6EE7-4A35-8BB8-CCB1D16D3277} /passive /norestart

				ECHO Unregister trimrenameaddin.dll
				C:\Windows\Microsoft.net\framework64\v2.0.50727\regasm -u "C:\Program Files\Kapish\Rename Addin\TRIMRenameAddin.dll"

				ECHO Kapish Rename Addin x64 1.1.0.1683
				MsiExec.exe /X{82714193-A748-41D0-9F33-5D3C87EE2FC2} /passive /norestart


:INSTALL_HPRM_64BITWORKSTATION
                ECHO -------------------------------------------------------
                ECHO ----- Installing Components for 64bit and 32bit   -----
                ECHO -------------------------------------------------------
                ECHO.

                ECHO 01. Install vcredist_x64.exe
                "Source\vcredist_x64.exe" /Q

                ECHO 02. Install HP Records Manager VS2010_Runtimes_x64.msi
                MsiExec.exe /I "Source\VS2010_Runtimes_x64.msi" /passive /norestart /log "%logpath%\02-Install-VS2010_Runtimes_x64.log"

                ECHO 02. Install HP Records Manager VS2010_Runtimes_x86.msi
                MsiExec.exe /I "Source\VS2010_Runtimes_x86.msi" /passive /norestart /log "%logpath%\02-Install-VS2010_Runtimes_x86.log"

                ECHO 03. Install HP Records Manager VS2013_Runtimes_x64.msi
                MsiExec.exe /I "Source\VS2013_Runtimes_x64.msi" /passive /norestart /log "%logpath%\03-Install-VS2013_Runtimes_x64.log"

                ECHO 03. Install HP Records Manager VS2013_Runtimes_x86.msi
                MsiExec.exe /I "Source\VS2013_Runtimes_x86.msi" /passive /norestart /log "%logpath%\03-Install-VS2013_Runtimes_x86.log"

                ECHO 04. Install HP Records Manager v9.10.914 x64 Client Software
				MsiExec.exe /I "Source\HPE_CM_x64.msi" /passive /norestart /l*vx "%logpath%\04-Install-HPRM91Client_x64.log" INSTALLDIR="C:\Program Files\Hewlett-Packard Enterprise\Content Manager\" ADDLOCAL="Client" ALLUSERS="1" AUTOGG="1" DEFAULTDB="C2" DEFAULTDBNAME="HPE CM Test" LANG="US" ODMA="None" ODMALOCAL="1" OUTLOOK_ON="1" PORTNO="1137" POWERPOINT_ON="1" PROJECT_ON="0" STARTMENU_NAME="HPE Content Manager" TRIM_DSK="1" TRIMREF="TRIM" PRIMARYURL="RMAPP02TST" WORD_ON="1" EXCEL_ON="1" TRIMUserSetup_On="1"

                ECHO 04. Install HP Records Manager v9.10.914 x86 Client Software
                MsiExec.exe /I "Source\HPE_CM_x86.msi" /passive /norestart /l*vx "%logpath%\04-Install-HPRM91Client_x86.log" INSTALLDIR="C:\Program Files (x86)\Hewlett-Packard Enterprise\Content Manager\" ADDLOCAL="Client" ALLUSERS="1" AUTOGG="1" DEFAULTDB="C2" DEFAULTDBNAME="HPE CM Test" LANG="US" ODMA="None" ODMALOCAL="1" OUTLOOK_ON="1" PORTNO="1137" POWERPOINT_ON="1" PROJECT_ON="0" STARTMENU_NAME="HPE Content Manager" TRIM_DSK="1" TRIMREF="TRIM" PRIMARYURL="RMAPP02TST" WORD_ON="1" EXCEL_ON="1" TRIMUserSetup_On="1"

                ECHO 05. Install Kapish Folder Wizard-x64-3.50.1904
                MsiExec.exe /I "Source\Kapish Folder Wizard-x64-3.50.1904.msi" /passive /norestart /log "%logpath%\05-Install-Kapish Folder Wizard-x64-3.30.1108.log"

                ECHO 05. Install Kapish Folder Wizard-x86-3.50.1904
                MsiExec.exe /I "Source\Kapish Folder Wizard-x86-3.50.1904.msi" /passive /norestart /log "%logpath%\05-Install-Kapish Folder Wizard-x86-3.30.1108.log"

                ECHO 06. Install Kapish Record Remover-x64-1.60.1400
                MsiExec.exe /I "Source\Kapish Record Remover-x64-1.60.1400.msi" /passive /norestart /log "%logpath%\06-Install-Kapish Record Remover-x64-1.40.1106.log"

                ECHO 06. Install Kapish Record Remover-x86-1.60.1400
                MsiExec.exe /I "Source\Kapish Record Remover-x86-1.60.1400.msi" /passive /norestart /log "%logpath%\06-Install-Kapish Record Remover-x86-1.40.1106.log"

				ECHO 07. Install Kapish_Explorer_4.50.1536_x64
				MsiExec.exe /I "Source\Kapish_Explorer_4.50.1536_x64.msi" /passive /norestart /log "%logpath%\07-Install-Kapish_Explorer_4.40.1274_x64.log"

                ECHO 07. Install Kapish_Explorer_4.50.1536_x86
                MsiExec.exe /I "Source\Kapish_Explorer_4.50.1536_x86.msi" /passive /norestart /log "%logpath%\07-Install-Kapish_Explorer_4.43.1354_x86.log"

				ECHO 08. Install Kapish_Rename_Addin_x64_1.1.1.3149.msi
				MsiExec.exe /I "Source\Kapish_Rename_Addin_x64_1.1.1.3149.msi" /passive /norestart /log "%logpath%\08-Install-Kapish_Rename_Addin_x64_1.1.1.3149.log"

                ECHO 08. Install Kapish_Rename_Addin_x86_1.1.1.3149.msi
                MsiExec.exe /I "Source\Kapish_Rename_Addin_x86_1.1.1.3149.msi" /passive /norestart /log "%logpath%\08-Install-Kapish_Rename_Addin_x86_1.1.1.3149.log"

                GOTO CLEAN_UP_WORKSTATION


:CLEAN_UP_WORKSTATION

                ECHO -------------------------------------------------------
                ECHO ----- Applying Workstation Preferences            -----
                ECHO -------------------------------------------------------
                ECHO.
				if not exist "C:\Program Files\Hewlett-Packard Enterprise\Content Manager\Backup" mkdir "C:\Program Files\Hewlett-Packard Enterprise\Content Manager\Backup"
                erase /Q "C:\Users\Public\Desktop\HPRM Desktop.lnk"
                erase /Q "C:\Users\Public\Desktop\HPE CM Desktop.lnk"
                erase /Q "C:\Users\Public\Desktop\HPRM Queue Processor.lnk"
                erase /Q "C:\Users\Public\Desktop\HPE CM Queue Processor.lnk"
                xcopy /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\HP Records Manager\*.*" "C:\Program Files\Hewlett-Packard Enterprise\Content Manager\Backup"
                erase /Q "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\HP Records Manager\HP TRIM Desktop.lnk"
				erase /Q "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\HPE Content Manager\HPE Content Manager Desktop.lnk"
                erase /Q "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\HP Records Manager\HP TRIM Queue Processor.lnk"
                erase /Q "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\HPE Content Manager\HPE Content Manager Queue Processor.lnk"

                REM ENSURE THE USER HAS THE RIGHT DESKTOP SHORTCUT
                xcopy /Y "HPE Content Manager.lnk" "C:\Users\Public\Desktop\"

                setlocal enableextensions
                for /F "tokens=2,*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" /v ProfileImagePath /s ^| find "REG_EXPAND_SZ" ^| findstr /v /i "\\windows\\ \\system32\\"') do (
                                call :removeLinks "%%b"
                )
                endlocal


                GOTO WORKSTATION_CONFIGURATION



:WORKSTATION_CONFIGURATION
                ECHO --------------------------------------------
                ECHO ----- Import Workstation Registry Keys -----
                ECHO --------------------------------------------
                ECHO.

				REM REG IMPORT "Source\Config\eml_64bit.reg"

                REG IMPORT "tr5_handler.reg"

                GOTO END_SCRIPT


:END_SCRIPT
                ECHO ---------------------------------------------------------------
                ECHO ----- TRIM Context to HP Records Manager Upgrade Complete -----
                ECHO ---------------------------------------------------------------
                ECHO.


:removeLinks
                erase /Q "%~1\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu\HP TRIM Desktop.lnk"
                erase /Q "%~1\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu\HP TRIM Queue Processor.lnk"
                erase /Q "%~1\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\HP TRIM Desktop.lnk"
                erase /Q "%~1\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\HP TRIM Queue Processor.lnk"
                erase /Q "%~1\AppData\Local\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu\HP TRIM Desktop.lnk"
                erase /Q "%~1\AppData\Local\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu\HP TRIM Queue Processor.lnk"
                erase /Q "%~1\AppData\Local\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\HP TRIM Desktop.lnk"
                erase /Q "%~1\AppData\Local\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\HP TRIM Queue Processor.lnk"
                erase /Q "%~1\Desktop\HP TRIM Desktop.lnk"
                erase /Q "%~1\Desktop\HP TRIM Queue Processor.lnk"


:FINISH_AND_EXIT
ECHO ALL DONE
net use %driveletter%: /delete /yes

exit

:GetBestServer
set maxping=10000
set maxhops=100
set bestping=
for %%s in (%SERVERLIST%) do (
 ping -n 1 %%s | FIND /i "could not find host" > nul
 if !errorlevel! equ 0 (
  echo.
 ) else (
  ping -n 2 %%s | FIND /i "100%% loss" > nul
  if !errorlevel! equ 0 (
   echo.
  ) else (
   set count=0
   TRACERT %%s > test.txt
   for /f "usebackq tokens=* delims=:" %%h in (`TYPE test.txt`) do call :addhop
   PING -n 3 %%s | FIND /i "Maximum" > test.txt
   for /f "usebackq tokens=2 delims=," %%i in (`TYPE test.txt`) do echo %%i > test.txt
   for /f "usebackq tokens=3 delims= " %%j in (`TYPE test.txt`) do echo %%j > test2.txt
   for /f "usebackq delims=" %%k in (`TYPE test2.txt`) do (call :pingparse %%k %%s !count!)
  )
 )
)
goto :Commands

:pingparse
 set max=%1
 set max=%max:ms=%
 set /a max=%max%*1
 set /a numhops=%3*1
 if %max% LEQ %maxping% (
 if %numhops% LSS %maxhops% (
  set bestping=%2
  set maxping=%max%
  set maxhops=%numhops%
 )
)
goto :eof

:addhop
 set /a count=!count!+1
goto :eof
