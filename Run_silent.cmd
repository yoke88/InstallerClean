@echo off
echo  =========================================================================
echo       1. ���Թ���Ա���ִ�иýű������������Ȩ��ִ�в�����
echo       2. �þ�Ĭ�汾Ĭ�Ͻ��������пɰ�ȫɾ����microsoft�ľɲ�����   
echo  =========================================================================
powershell  -noprofile  -ExecutionPolicy remotesigned -file %~dp0InstallerClean.ps1 -quiet

