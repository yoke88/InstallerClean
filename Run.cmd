@echo off
echo  =========================================================================
echo       1. 请以管理员身份执行该脚本，否则可能无权限执行操作。
echo       2. 请选择要清理的文件，你可以通过弹出的窗口进行筛选。
echo       3. 该程序实际上是执行powershell脚本，需要powershell 版本2，
echo          且out-gridview 可用
echo  =========================================================================
powershell  -noprofile  -ExecutionPolicy remotesigned -file %~dp0InstallerClean.ps1 
