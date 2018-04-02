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


:INSTALL_HPRM_64BITWORKSTATION
                ECHO -------------------------------------------------------
                ECHO ----- Installing Components for 64bit and 32bit   -----
                ECHO -------------------------------------------------------
                ECHO.

                msiexec /p "Source\Patch\HPE_CM_x64.msp" /norestart

                msiexec /p "Source\Patch\HPE_CM_x86.msp" /norestart


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
