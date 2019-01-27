@echo off
chcp 65001>nul

net session>nul 2>nul
if %errorLevel% GEQ 1 goto :startAsAdmin

%~d0
cd %~dp0

set officeSetupURL=https://onedrive.live.com/download?cid=D3AF852448CB4BF6^&resid=D3AF852448CB4BF6%%21259^&authkey=AAK3Qw80R8to-VE
set officeSetupAdditionalURL=https://public.dm.files.1drv.com/y4mTqNAebstFsw9p507h2xqKwivr_pHN6OwyaEAA3-xavLhFr_9HmsF-bF931oFmOZ-ynEy53Blug8XG1FLTmT0VT36kjGfbT1a_tItImyjwJqqKSTp1qCXBdPbKmlI5uNy0P6tkSMicg32ddWL3Z91nyoXV8SXymCpC_Bwp1SoqzBjBNAV4CXfr5t-QtlkJapj/Microsoft%%20Office%%20Professional%%20Plus%%202016.iso?access_token=EwAIA61DBAAUcSSzoTJJsy%%2bXrnQXgAKO5cj4yc8AAdNI1D0Km20nFjkwjZJAiQrksgJ3Bpa5AYk%%2fVPN9VGXuBitjIC6LhGh3WQcX%%2fE%%2f0V9IPo7%%2f2JLzjJnJ9%%2bSwX%%2bNm37S8I6zXYsDfy7AervE2iGE%%2bSJ901s1sjMHULB%%2btCGYvsUIEHNQTPA4dAn8gCmlrpp%%2f%%2f6cGuJnBlc2jysi1%%2bxKUcREdO8tfwpLvXaR9W%%2btDp5kKiLXvKuG9H0gCLpbknzFMkyaeeGemUTzGRglwqTTPlp94%%2fEmaMW9O5qg2STAFqKV6H%%2f%%2flNtevRIoCctJgU9dXcOfbc5YdRhySjbBGJxDLReJJk4X2zeRvq62G3ITD25jEOwYufL7POHXJOe47kDZgAACEMQTepMithw2AEdh0sQB%%2bLFCpxLdVafSfaeStp31%%2fHUPqg7TeINPS7DuEP3Ga%%2fqOPNX6CtkWzkrodHWyXsQj5eSV6ZMFZdZa2zrxSntXJs%%2bkaVAMLvGtXN8lwMXjyCZw8yhboCdwEqR8IzbgZsTR5DOXGLAcq%%2fRt81DQzUnnsHdnsuDO%%2ffELmE8ccu3eBp3ntqzz9MqxpsLotGpmwL5y72QWnmFM4UnCEhTYo1QzYoxyELtavpBik5y2%%2fSLUthnrXtxUGLuj9xAHcXfewJmGbhA3DVSnKdx9RqckzYjqBBISzqYQVmbWJeYsZIQaQrhcOkudEbpVTUplF4I%%2bYOJqiOCSI6W9lL6fTWdLuMYgsXTnnMtFMNPYeTTaYDQoZj1GqAZckKcdscy%%2b%%2foNZXkSNlPaJEZdZoozvuEFgRzt%%2fmWM9YvS7aCfia6kwDRxY9VEYwLvPQNhFpg3DGpTI%%2brrKokLUs6q9TIUBUfD1SUbXMTnN8cB1Jpsveic9wAfhg837RZVdBvfWZOYnv4myviNwqtXjgaxtzwpb6atb4EEOy6KQLAhqbZwHBdWIQhypIqFfRcATwpSENEP%%2b2hF7T878znu3rE%%2fJYijcuk%%2fH8GlzFi7y7y9%%2bl3hsW4L7eb6ybZD%%2fy7JEAI%%3d
set officeSetupISO=Microsoft Office Professional Plus 2016.iso





call :logo
echo.^(^i^) Office Setup is running...
timeout /nobreak /t 1 >nul

echo.^(^?^) Are you sure^? ^(Enter or close^)
pause>nul





:start
call :logo
if not exist "%officeSetupISO%" (
  echo.^(^i^) Downloading Microsoft Office Professional Plus 2016 Setup
  wget.exe --quiet --show-progress --progress=bar:force:noscroll --no-check-certificate --tries=3 "%officeSetupURL%" --output-document="%officeSetupISO%"
  if not exist "%officeSetupISO%" if "%officeSetupURL%" NEQ "%officeSetupAdditionalURL%" (
    set officeSetupURL=%officeSetupAdditionalURL%
    goto :start
  )
  timeout /nobreak /t 1 >nul
)



echo.^(^i^) Mounting iso file...
powershell.exe "Mount-DiskImage ""%~dp0%officeSetupISO%"""
timeout /nobreak /t 1 >nul



echo.^(^i^) Running setup...
for /f "skip=3" %%i in ('powershell.exe "Get-DiskImage """%~dp0%officeSetupISO%""" | Get-Volume | Select-Object {$_.DriveLetter}"') do start /wait %%i:\O16Setup.exe
:question
set /p answer=^(^>^) Setup is completed^? ^(y/n^) ^> 
if "%answer%" == "n" goto :start
if "%answer%" NEQ "y" goto :question



echo.^(^i^) Unmounting iso file...
powershell.exe "Dismount-DiskImage ""%~dp0%officeSetupISO%"""
timeout /nobreak /t 1 >nul





call :logo
echo.^(^i^) The work is completed^!
timeout /nobreak /t 1 >nul

echo.^(^?^) Reload now^? ^(Enter or close^)
pause>nul

echo.^(^!^) Reboot^!
shutdown /r /t 5
timeout /t 4 >nul
exit







:logo
title [MikronT] Office Setup
color 0b
cls
echo.
echo.
echo.    [MikronT] ==^> Office Setup
echo.   ==========================================
echo.     See other here:
echo.         github.com/MikronT
echo.
echo.                  Will no longer be updated^!
echo.                  Merged with Ten Tweaker
echo.
echo.
echo.
exit /b





:startAsAdmin
echo.^(^!^) Please, run as Admin^!
timeout /nobreak /t 3 >nul
exit