@echo off

call :Check
goto %version%
goto :eof

:install
for /f "skip=1 delims=" %%a in (
     'wmic product where "Name like '%%adobe acrobat%%'" get name'
) do @for /f "delims=" %%b in ("%%a") do set SOFTWARE=%%b

IF "%SOFTWARE%" == "Adobe Acrobat Reader DC - Portuguˆs  " (
	copy "\\addc\scripts$\Scripts GPO\UninstallSoftware\AdobeReader.exe" C:\
	wmic product where "Name like '%%Adobe Acrobat%%'" call uninstall
	start /wait C:\AdobeReader.exe /sAll
	echo done >> C:\Admin\adobecheck.txt
	del C:\AdobeReader.exe
	goto :eof
) ELSE (
	copy "\\addc\scripts$\Scripts GPO\UninstallSoftware\AdobeReader.exe" C:\
	wmic product where "Name like '%%Adobe Acrobat%%'" call uninstall
	start /wait C:\AdobeReader.exe /sAll
	echo done >> C:\Admin\adobecheck.txt
	del C:\AdobeReader.exe
	goto :eof
)

:done
goto :eof

:Check
REM Verifica se o arquivo de controle existe, e caso exista lˆ o texto escrito nele para verificar em que passo estamos
if exist C:\Admin\adobecheck.txt (
    set /p version=<C:\Admin\adobecheck.txt
) else (
	mkdir C:\Admin
    set version=install
)