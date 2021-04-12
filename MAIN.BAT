::Code by  : LEO
::功    能 : 采集本机所有网卡MAC地址,过滤虚拟网卡
::DATE : 2021/4/9

@echo off
COLOR 0A
title MAC采集维护工具 POWERBY LEO
setlocal EnableDelayedExpansion
cls
echo=
echo=
echo=
echo=

rem 个性化设置参数
set dns1="x.x.x.x"
rem 备用DNS为电信线路
rem set dns2="114.114.114.114"
rem 备用DNS为移动线路
set dns2="211.136.192.6"
set mailad=******@**.com

rem 主DNS检测
for /f "tokens=2 delims=:" %%a in ('ipconfig ^/all ^|findstr /IC:"DNS 服务器"') do (echo %%a|find %dns1% >nul||(set erro=1&echo    您的主DNS为 %%a))

rem 备DNS检测
FOR /F "tokens=2* delims=适配器" %%i IN ('ipconfig ^/all^|find "适配器"') DO (
FOR /F "tokens=1* delims=:" %%i in ("%%i") do (
set var=%%i
if "!var:~0,1!"==" " (set var=!var:~1!)
for /f "tokens=1 delims=" %%m in ('netsh int ip sh dns "!var!"') do (call :t %%m)))
goto MAC

:t
set a=%1
if not "%a:.=%"=="%a%" (echo %a%|find %dns2% >nul||(set erro=1&echo    您的备DNS为  %a%))
goto EOF

:MAC
rem DNS修复模块
if "%erro%"=="1" (
echo=
echo    DNS需要修复为^:
echo    主^:广州中心10.26.15.40/41^(优选^)
echo       深圳中心10.27.24.1^/2
echo    备^:电信114.114.114.114
echo       移动211.136.192.6
echo=
echo strGatewayMetric = Array^(1^) >dns.vbs
echo arrDNSServers = Array^(%dns1%,%dns2%^) >>dns.vbs

echo strComputer = ".">>dns.vbs
echo Set objWMIService = GetObject^("winmgmts:" ^& "{impersonationLevel=impersonate}!\\" ^& strComputer ^& "\root\cimv2"^)>>dns.vbs
echo Set colNetAdapters = objWMIService.ExecQuery ^("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE"^)>>dns.vbs
echo For Each objNetAdapter in colNetAdapters>>dns.vbs
echo errDNS = objNetAdapter.SetDNSServerSearchOrder^(arrDNSServers^)>>dns.vbs
echo Next>>dns.vbs
echo=
echo=
echo=
set /p temp=    即将自动修复DNS配置,请回车确认:
echo=
echo    修复成功&echo=
call dns.vbs
del dns.vbs
set erro=
cls
echo=
echo=
echo=
echo=
)

rem MAC采集模块
ipconfig /all >temp.txt
set cs=0
for /f "tokens=2 delims=:" %%a in ('systeminfo^|find "系统型号:"') do (echo    主机型号^:%%a & echo    主机型号^:%%a>result.txt)

for /f "delims=" %%i in (temp.txt) do (
set a=%%i
set b=!a:. =!
if !cs! equ 0  (
echo !b! |findstr /IC:"主机名"&& (echo !b!>>result.txt)
echo !b!|findstr /IC:"主 DNS 后缀" && (
set cs=1
echo !b!>>result.txt
echo !b!|findstr "citics" >nul &&(echo    已加域&echo    已加域>>result.txt)||(echo    未加域&echo    未加域>>result.txt)
)
)

echo !b! |findstr /IC:"描述" >nul && (echo !b!| findstr "VPN Virtua" >nul &&(set cs=1)||(set cs=2&echo=&echo=>>result.txt& echo !b! & echo !b!>>result.txt))

if !cs! equ 2  (
echo !b! |findstr /IC:"物理地址"&& (echo !b!>>result.txt)
echo !b! |findstr /IC:"IPv4"&& (echo !b!>>result.txt)
echo !b! |findstr /IC:"子网掩码"&& (echo !b!>>result.txt)
echo !b! |findstr /IC:"默认网关"&& (echo !b!>>result.txt)
echo !b! |findstr /IC:"DNS 服务器" && (echo !b!>>result.txt)
)

)

del temp.txt
echo=
echo=
echo    %date% %time%
echo=
echo=
set /p name=    1/4 (请标记员工姓名):&echo=
echo=>>result.txt
echo    员工姓名:%name%>>result.txt

set /p tm=    2/4 (请固定资产标签):&echo=
echo    固定资产标签:%tm%>>result.txt
echo=>>result.txt

set /p gn=    3/4 是否联机启动邮件上报(Y/N)?:&echo=
if "%gn%" == "Y" (goto MAIL)
if "%gn%" == "y" (goto MAIL) else (goto PASS)

:MAIL
echo NameSpace="http://schemas.microsoft.com/cdo/configuration/">mail.vbs
echo set Email = CreateObject("CDO.Message")>>mail.vbs
echo Email.From = "118072212@qq.com" >>mail.vbs

echo     即将Mail To^:%mailad%&echo=
set /p mailad=    4/4 (是否输入新邮件地址):&echo=
echo Email.To = "%mailad%">>mail.vbs

for /f "tokens=2 delims=:" %%a in ('ipconfig ^/all ^|find "主机名"') do (set host=%%a)
echo Email.Subject = "MAC地址采集:%name%-%host%">>mail.vbs

echo x="result.txt" >>mail.vbs
echo Set fso=CreateObject("Scripting.FileSystemObject")>>mail.vbs
echo Set myfile=fso.OpenTextFile(x,1,Ture)>>mail.vbs
echo c=myfile.readall>>mail.vbs
echo myfile.Close>>mail.vbs
echo Email.Textbody = c>>mail.vbs
echo with Email.Configuration.Fields>>mail.vbs
echo .Item(NameSpace^&"sendusing") = 2 >>mail.vbs
echo .Item(NameSpace^&"smtpserver") = "pop.qq.com">>mail.vbs
echo .Item(NameSpace^&"smtpserverport") = 465>>mail.vbs
echo .Item(NameSpace^&"smtpusessl") = True>>mail.vbs
echo .Item(NameSpace^&"smtpauthenticate") = 1 >>mail.vbs
echo .Item(NameSpace^&"sendusername") = "********">>mail.vbs
echo .Item(NameSpace^&"sendpassword") = "********">>mail.vbs
echo .Update>>mail.vbs
echo end with>>mail.vbs
echo Email.Send>>mail.vbs
echo Set Email=Nothing>>mail.vbs
call mail.vbs
echo     发送邮件成功&echo=
del mail.vbs
goto PASS

:PASS
echo     程序结束
echo=
echo=
echo=
type result.txt>>alldata.txt
echo=>>alldata.txt
echo -----------------------------------------%date% %time%----------------->>alldata.txt
echo=>>alldata.txt
notepad result.txt
del result.txt
pause

:EOF
