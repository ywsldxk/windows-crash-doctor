# Windows Crash Doctor

[English](README.md) | [简体中文](README.zh-CN.md)

`Windows Crash Doctor` 是一个面向 Windows 的排障小工具，专门用来处理这种很烦的情况：

- 不是某一个软件坏了，而是很多软件一起随机崩
- 事件查看器总是指向 `KERNELBASE.dll`、`MSVCP140.dll`、`VCRUNTIME140.dll`、`ucrtbase.dll`
- 你怀疑是 VPN、输入法、主板服务、驱动、网络钩子之类的“共因问题”

这类 DLL 往往只是“出事地点”，不一定是真正根因。这个工具的目标就是把“共享根因”尽量找出来。

## 这个名字为什么更合适

我把项目名定成 `Windows Crash Doctor`，原因很简单：

- 好懂，不拐弯
- 英文搜索友好
- 用户很容易用自然语言搜到它

常见能搜到它的英文关键词包括：

- `windows app crash`
- `KERNELBASE.dll crash`
- `MSVCP140.dll crash`
- `random software crashes on Windows`
- `many apps crashing after VPN install`

## 功能

- 读取最近的 `Application Error` 事件日志
- 汇总最常崩溃的程序和模块
- 检查已安装软件、服务、进程、启动项、Winsock 提供程序和驱动包
- 对可疑目标打分，并标成 `low`、`medium`、`high`
- 输出 Markdown 和 JSON 报告
- 对少数已知目标提供受限的 `apply-fix` 自动修复流程

## 当前支持的目标

已实现自动修复：

- `Sangfor`

已实现检测和建议：

- `ROGLiveService`
- `ChineseImeStack`

## 主要文件

- `WindowsCrashDoctor.ps1`：主脚本
- `AppCrashDoctor.ps1`：旧脚本名兼容入口
- `README.md`：英文说明

## 快速开始

先做只读扫描：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\WindowsCrashDoctor.ps1 -Mode scan
```

生成 Markdown 和 JSON 报告：

```powershell
.\WindowsCrashDoctor.ps1 `
  -Mode suggest-fix `
  -Days 14 `
  -ReportPath .\reports\last-scan.md `
  -JsonPath .\reports\last-scan.json
```

先预演修复，不真正改系统：

```powershell
.\WindowsCrashDoctor.ps1 `
  -Mode apply-fix `
  -Target Sangfor `
  -ResetWinsock `
  -WhatIf
```

需要真正执行修复时，建议用管理员 PowerShell：

```powershell
Start-Process powershell -Verb RunAs -ArgumentList @(
  '-ExecutionPolicy', 'Bypass',
  '-File', '.\WindowsCrashDoctor.ps1',
  '-Mode', 'apply-fix',
  '-Target', 'Sangfor',
  '-ResetWinsock'
)
```

## 模式说明

### `scan`

只读扫描，快速看当前系统有没有明显嫌疑项。

### `suggest-fix`

只读扫描，并输出更明确的修复建议。

### `apply-fix`

对指定目标执行修复流程。

## 当前 `Sangfor` 自动修复会做什么

- 停掉已知深信服相关进程
- 尝试停止相关服务
- 调用常见安装目录里的官方卸载器
- 用 `pnputil` 尝试卸载匹配驱动
- 尝试删除 `SangforVnic`
- 可选执行 `netsh winsock reset`
- 尝试删除残留安装目录

## 安全说明

- 建议先跑 `scan` 或 `suggest-fix`
- 真正执行修复前先加 `-WhatIf`
- 有些动作需要管理员权限
- 有些驱动删除后仍然需要重启
- 公司 VPN、EDR、远程运维或 OEM 软件可能是刚需，删除前请先确认

## 适合什么场景

- 装了 VPN 之后很多软件开始一起崩
- 事件查看器老是指向 `KERNELBASE.dll`
- 某个 OEM 后台服务在疯狂崩溃
- 想先得到一份结构化排查报告，再决定卸载什么

## 许可证

MIT
