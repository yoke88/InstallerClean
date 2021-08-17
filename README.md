# 简介 Introduction

installerclean.exe 是一个自解压的exe 文件，实际上最终执行installerclean.ps1。该程序主要用于清理windows installer 目录。我公司的几个桌面计算机的Installer目录占了60-70GB，虽然我采用了很多办法，比如压缩，清理磁盘空间（该项目下的autocleanmgr.cmd),效果都不大，我后面参考了一些流行的工具，分析了的他们的底层原理，我觉得我可以写一个脚本来完成这个事情，而不要安装依赖包，比如.net 4.0 。考虑到大部分powershell 已经内置在2.0以上，所以我写了这个工具。你可以完全按照自己的意愿删除installer缓存，所有代码和逻辑都是可见的。

如果你想了解这个程序的产生的一些分析过程，欢迎访问[我的博客文章](http://blog.51cto.com/yoke88/2116798)

## 一些问题点
1. 由于一些系统可能powershell 可用，但是out-gridview命令不可用（需要.net的一些支持，或者版本较低不支持passthrou 开关，该开关可以弹出一个图形窗口来过滤项目）~~这种情况下我会在桌面上生成一个包含可以清理的installer的文件的相关信息的CSV文件，你可以用这些信息选择去清理哪些Installer~~. 更新后的版本,使用了一个函数来判断out-gridview及passthrou 开关是否可用,如果不可用,默认清理所有可以清理的microsoft 补丁文件,生成的csv文件在%temp%目录下,在控制台会有提示该路径.
2. 同时发布了一个静默版本,带silent的exe 文件,可以用来静默的删除所有可清理的windows 补丁文件,方便批量或者在后台执行.
