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

    GOTO UNINSTALL_TRIMCONTEXT

:UNINSTALL_TRIMCONTEXT

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



    ECHO 00. .NET Framework 3.5
    DISM /Online /Enable-Feature /FeatureName:NetFx3

    ECHO 07. Install Kapish_Explorer_4.50.1536_x64
    MsiExec.exe /I "Source\Kapish_Explorer_4.50.1536_x64.msi" /passive /norestart /log "%logpath%\07-Install-Kapish_Explorer_4.40.1274_x64.log"

    ECHO 07. Install Kapish_Explorer_4.50.1536_x86
    MsiExec.exe /I "Source\Kapish_Explorer_4.50.1536_x86.msi" /passive /norestart /log "%logpath%\07-Install-Kapish_Explorer_4.43.1354_x86.log"


:FINISH_AND_EXIT
ECHO ALL DONE
net use %driveletter%: /delete /yes

exit
