@echo off
echo  =========================================================================
echo       1. 请以管理员身份执行该脚本，否则可能无权限执行操作。
echo       2. 该静默版本默认将清理所有可安全删除的microsoft的旧补丁。   
echo  =========================================================================
powershell  -noprofile  -ExecutionPolicy remotesigned -file %~dp0InstallerClean.ps1 -quiet

