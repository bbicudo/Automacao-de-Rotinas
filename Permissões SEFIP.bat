@echo off

REM #############################################################
REM ################### ESCRITO POR #############################
REM ############### BRUNO BRANCO BICUDO #########################
REM ################# 11/01/2021 ################################
REM #############################################################
REM #############################################################

set ARCH=wmic OS get OSArchitecture | findstr /I bits

IF /i "%ARCH%"=="64 bits" (

set PROGFILE=Program Files

) ELSE (

set "PROGFILE=Program Files (x86)"

)

set /P USER=Insira o nome do usuario:

set "FOLDERNAME=C:\%PROGFILE%\CAIXA"

echo HKEY_LOCAL_MACHINE\SOFTWARE\Caixa [1 7 17] >> %TEMP%\regini.txt

set REGNAME=%TEMP%\regini.txt

icacls "%FOLDERNAME%" /grant %USER%:(OI)(CI)F /T

regini %REGNAME%

@pause