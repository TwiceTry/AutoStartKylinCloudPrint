# AutoStartKylinCloudPrint 项目

该项目包含两个脚本文件，用于自动配置和启动 麒麟云打印。以下是项目的详细说明。

## 文件结构

```
AutoSetupKylinCloudServer.ps1
GenerateStartupScript.vbs
```

### 文件说明

#### 1. `AutoSetupKylinCloudServer.ps1`

这是一个 PowerShell 脚本，用于自动启动 Kylin Cloud 打印服务并进行必要的配置。

- **功能**：
  - 启动 `kylin-cloud-printer-server.exe` 程序。
  - 自动定位目标窗口并模拟用户操作（如输入密码、点击按钮）。
  - 支持通过参数传递密码和工作目录。
  - 提供窗口截图功能。

#### 2. GenerateStartupScript.vbs

这是一个 VBScript 脚本，用于添加开机启动项并生成的快捷方式脚本。

- **功能**：
  - 提示用户输入密码，并将其传递给 PowerShell 脚本。
  - 生成一个 `.vbs` 脚本，用于隐藏窗口运行 PowerShell 脚本。
  - 将生成的 `.vbs` 脚本注册到 Windows 启动项中。
  - 在桌面创建快捷方式，方便用户手动启动。

## 使用步骤

1. 将这两个文件复制到麒麟云打印服务端安装目录中
2. 双击运行 GenerateStartupScript.vbs，根据提示输入密码。

## 注意事项

- **杀软提醒**：由于 VBScript 脚本涉及文件生成和注册表操作，可能会被杀毒软件（如 360）误报为风险并自动清除。运行脚本前，请暂时关闭杀毒软件。
- **兼容性**：脚本适用于 Windows 7 及以上版本的 Windows 系统。
- **操作影响**：由于脚本会模拟鼠标和键盘操作，运行脚本时请避免使用鼠标和键盘，以免干扰脚本执行。
- **分辨率适配**：脚本基于相对程序窗口的点击位置进行操作，在不同分辨率或缩放设置下，可能导致操作偏移，从而影响正常运行。
- **风险提示**：运行脚本时，请避免同时操作其他文本或输入框，否则可能导致文本被覆盖的风险。

## 实现目标

简化麒麟云打印打印服务的配置和启动流程，减少用户手动操作。

## 附
### [麒麟云打印1.1.3版本服务端及客户端](https://pan.baidu.com/s/1B2554iMGwxDVTNcoRwg8ew?pwd=fgbr)
