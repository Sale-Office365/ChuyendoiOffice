# XÓA KEY OFFICE, CHUYỂN ĐỔI OFFICE VÀ SAO LƯU OFFICE #

## XÓA KEY OFFICE ##

**Bạn có thể dùng AIO Tools V3.1.3 để xóa, cũng có thể dùng file cmd tạo từ đoạn mã sau:**

```php
chcp 65001 >nul
@echo off
Title XOA KEY OFFICE
mode con: cols=96 lines=35
chcp 65001 >nul
@echo.
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo  Run CMD as Administrator...
    goto goUAC 
) else (
 goto goADMIN )

:goUAC
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:goADMIN
    pushd "%CD%"
    CD /D "%~dp0"
	
:main
cls
color f0
@echo. 
echo        XOA KEY OFFICE
echo     Chon Phien Ban Office Can Xoa Key
echo =========================================
echo [  1. Office 2010     : Nhan phim so 1  ]
echo [  2. Office 2013     : Nhan phim so 2  ]
echo [  3. Office 2016     : Nhan phim so 3  ]
echo [  4. Office 2019     : Nhan phim so 4  ]
echo [  5. Office 2021     : Nhan phim so 5  ]
echo [  6. Office 365      : Nhan phim so 6  ]
echo =========================================
Choice /N /C 123456 /M "* Nhap lua chon : 
if %errorlevel% == 6 ( set "xx=16" & goto vogia)
if %errorlevel% == 5 ( set "xx=16" & goto vogia)
if %errorlevel% == 4 ( set "xx=16" & goto vogia)
if %errorlevel% == 3 ( set "xx=16" & goto vogia)
if %errorlevel% == 2 ( set "xx=15" & goto vogia)
if %errorlevel% == 1 ( set "xx=14" & goto vogia)

:vogia
if exist "%ProgramFiles%\Microsoft Office\Office%xx%\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office%xx%"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office%xx%\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office%xx%"
cscript ospp.vbs /dstatus >dstatus.txt
start dstatus.txt
goto office
)

:office
set /p key= * NHAP 5 KY TU CUOI CUA KEY : 
@echo  ...DANG XOA KEY OFFICE...
cscript OSPP.VBS /unpkey:%key%
@echo =========================================
@echo      DA XOA KEY OFFICE THANH CONG !
@echo =========================================
goto office
)
```

**Ngoài ra chúng ta có thể Download, cài đặt và kích hoạt Office từ Office Tool Plus [bấm vào đây](https://otp.landian.vip/en-us/) hoặc [tại đây](https://1drv.ms/u/s!AkwSBX-xWiVhg3bKuI5HGHa_nUB7?e=4lsbfR) nếu không chạy được là do thiếu runtime, download về cài đặt bổ sung [tại đây](https://bsthanh-my.sharepoint.com/:u:/g/personal/0914678254_bsthanh_tk/Ebuo4utXHOhGncmFJ8phrZcB0sEldAucovhYOdDQ6SmwkQ?e=l7KLVp)**

![1](https://user-images.githubusercontent.com/82578024/163676849-0c17b2f4-0316-4e02-a712-cb48914046e6.jpg)
Chọn Office sau đó intall licenses, bấm Yes
![2](https://user-images.githubusercontent.com/82578024/163676923-384d2e00-6f0d-4585-aeec-cdb22e5b08cd.jpg)

**Hoàn thành Kích hoạt!**

## Ghi chú ##

Bạn có thể chuyển đổi qua lại giữa các office sau:

- Office 2016
- Office 2019
- Office 2021
- Office 365

Mà không cần cài đặt lại, giả sử bạn muốn sử dụng Office 2021 mà có một bộ cài đặt Office 2016 hoặc Office 2019 thì phải làm sao? Câu trả lời cứ cài theo bộ cài có sẵn sau đó dùng thủ thuật chuyển đổi trong vài nốt nhạc!

Bạn dùng **xóa key Office** như đã nói ở phần trên (file cmd), để xóa sạch toàn bộ key Office

Sau đó bạn chạy file kích hoạt office (2016, 2019, 2021) bằng cmd là OK! 

Trường hợp bạn muốn sử dụng Office 365 thì sao? Sau khi xóa sạch key Office, bạn dùng **Office Tool Plus** để gán giấy phép Office 365 xong, dùng tài khoản để kích hoạt, mình có một số tài khoản để ở phần trên!

https://user-images.githubusercontent.com/82578024/170656193-daa7c7b7-7aed-477f-9b88-1a4bd58eb018.mp4

https://user-images.githubusercontent.com/82578024/170853349-a0b1b25a-1f19-454b-8dd0-6a45532ac558.mp4

## SAO LƯU OFFICE ##

Mở **NotePad** copy đoạn mã sau vào và bấm **Save As** với tên **SaoluuOfficeVaWindows.cmd** rồi Run file này dưới quyền **Run Administrator**, làm theo hướng dẫn.

```php
@echo off&set local&color 0f&mode con cols=64 lines=25&title  Backup Restore Activations 1.1
::--------------------------------------------------------------------------------------------------------------------------------------------------------
::--------------------------------------------------------------------------------------------------------------------------------------------------------
:: Elevating UAC Administrator Privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if "%errorlevel%" NEQ "0" (
	echo: Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
	echo: UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
	"%temp%\getadmin.vbs" &	exit 
)
if exist "%temp%\getadmin.vbs" del /f /q "%temp%\getadmin.vbs"
) else if "%errorlevel%" NEQ "1" (
cd /d "%~dp0" && ( if exist "%temp%\getadmin.vbs" del "%temp%\getadmin.vbs" ) && fsutil dirty query %systemdrive% 1>nul 2>nul || (  cmd /u /c echo Set UAC = CreateObject^("Shell.Application"^) : UAC.ShellExecute "cmd.exe", "/k cd ""%~sdp0"" && ""%~s0""", "", "runas", 1 >> "%temp%\getadmin.vbs" && "%temp%\getadmin.vbs" && exit /B )
)
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" set folder="%ProgramFiles%\Microsoft Office\Office14"& set /a ver=4
if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" set folder="%ProgramFiles(x86)%\Microsoft Office\Office14"& set /a ver=4
if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" set folder="%ProgramFiles%\Microsoft Office\Office15"& set /a ver=5
if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" set folder="%ProgramFiles(x86)%\Microsoft Office\Office15"& set /a ver=5
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" set folder="%ProgramFiles%\Microsoft Office\Office16"& set /a ver=6
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" set folder="%ProgramFiles(x86)%\Microsoft Office\Office16"& set /a ver=6
::--------------------------------------------------------------------------------------------------------------------------------------------------------
:start
cls&color 0f&mode con cols=66 lines=25
echo  ================================================================
echo  ^|                                                              ^|
echo  ^|                BACKUP AND RESTORE ACTIVATION                 ^|
echo  ^|                                                              ^|
echo  ================================================================
echo  ^|                              ^|                               ^|
echo  ^|    [1] BACKUP ACTIVATION     ^|    [2] RESTORE ACTIVATION     ^|
echo  ^|                              ^|                               ^|
echo  ================================================================
echo  ^|                                                              ^|
echo  ^|           [3] EXIT AND THANK YOU FOR USING THIS TOOL         ^|
echo  ^|                                                              ^|
echo  ================================================================
echo.&echo.
echo  ==========================Coded by Kaz==========================
echo     Thanks to Nguyen Viet Hoang, Le Quang Dat, Tran Vinh Trung
echo                     support me to make this tool
echo                        My Name: Huynh Danh Dat
echo    Any problems please inbox this facebook:fb.com/dat.huynhdanh  
echo  ================================================================
echo.&CHOICE /C 123 /N /M " YOUR CHOICE: "
IF ERRORLEVEL == 3 exit
IF ERRORLEVEL == 2 GOTO:RESTORE
IF ERRORLEVEL == 1 GOTO:BACKUP

:BACKUP
cls
echo This Tool will delete all old backup !!!. Do you want to continue ?? 
CHOICE /C yn /N /M "Your choice (Y/N): "
IF ERRORLEVEL == 2 GOTO:start
IF ERRORLEVEL == 1 GOTO:BACKUP1
:BACKUP1
for /f "tokens=6 delims=[.] " %%a in ('ver') do set ver1=%%a
if %ver1% LEQ 7601 (
XCOPY C:\Windows\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\* Backup\SoftwareProtectionPlatform /s /i /y >nul
if %ver% LEQ 4 (
goto:start
goto:BACKUP
) else (
XCOPY C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform\* Backup\OfficeSoftwareProtectionPlatform /s /i /y>nul
goto:start
goto:BACKUP
)
) else (
attrib -s -h “C:\Windows\System32\spp\store\2.0\data.dat” /s /d 
XCOPY C:\Windows\System32\spp\store\* Backup\store /s /i /y>nul
if %ver% LEQ 4 (
XCOPY C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform Backup\OfficeSoftwareProtectionPlatform /s /i /y>nul
goto:start
goto:BACKUP
) else (
goto:start
goto:BACKUP
)
)

:RESTORE
cls
echo STOPPING SOME SERVICES FOR RESTORE ACTIVATION ...
net stop sppsvc>nul 2>nul 
net stop osppsvc>nul 2>nul
echo Done.
for /f "tokens=6 delims=[.] " %%a in ('ver') do set ver1=%%a
if %ver1% LEQ 7601 (
echo RESTORING WINDOWS AND OFFICE LICENSE FILES ...
XCOPY Backup\SoftwareProtectionPlatform\* C:\Windows\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform /s /i /y
if %ver% LEQ 4 (
echo Done.
goto:restore1
) else (
XCOPY Backup\OfficeSoftwareProtectionPlatform\* C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform  /s /i /y
echo Done.
goto:restore1
)
) else (
echo RESTORING WINDOWS AND OFFICE LICENSE FILES ...
if %ver% LEQ 4 (
XCOPY Backup\store\* C:\Windows\System32\spp\store /s /i /y
XCOPY Backup\OfficeSoftwareProtectionPlatform\* C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform  /s /i /y 
echo Done.
goto:restore1
) else (
XCOPY Backup\store\* C:\Windows\System32\spp\store /s /i /y
echo Done.
goto:restore1
) 
)

:restore1
echo ACTIVATING WINDOWS AND OFFICE ...
sc config sppsvc start= auto >nul 2>nul& net start sppsvc >nul 2>nul
sc config osppsvc  start= auto >nul 2>nul& net start osppsvc >nul 2>nul
sc config wuauserv start= auto >nul 2>nul& net start wuauserv >nul 2>nul
sc config LicenseManager start= auto >nul 2>nul& net start LicenseManager >nul 2>nul
cscript /nologo %windir%\system32\slmgr.vbs -rilc >nul 2>nul
cscript /nologo %windir%\system32\slmgr.vbs -dli >nul 2>nul
cscript /nologo %windir%\system32\slmgr.vbs -ato 
cd C:\ >nul 2>nul
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" set folder="%ProgramFiles%\Microsoft Office\Office14"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" set folder="%ProgramFiles(x86)%\Microsoft Office\Office14"
if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" set folder="%ProgramFiles%\Microsoft Office\Office15"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" set folder="%ProgramFiles(x86)%\Microsoft Office\Office15"
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" set folder="%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" set folder="%ProgramFiles(x86)%\Microsoft Office\Office16"
cscript //Nologo %folder%\ospp.vbs /act 
cscript //Nologo %folder%\ospp.vbs /dstatus 
echo Done.
echo Press any button to go restart...
pause>nul
shutdown.exe /r /t 00
goto:start
```



