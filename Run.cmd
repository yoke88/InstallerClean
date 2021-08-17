@echo off
echo  =========================================================================
echo       1. 请以管理员身份执行该脚本，否则可能无权限执行操作。
echo       2. 请在弹出的窗口中选择要清理的文件，建议筛选microsoft的补丁或者office补丁。
echo       3. 如果您是win7等低版本系统,默认将清理所有可安全删除的microsoft的旧补丁。
echo  =========================================================================
powershell  -noprofile  -ExecutionPolicy remotesigned -file %~dp0InstallerClean.ps1 
