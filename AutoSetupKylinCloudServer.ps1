param (
    [string]$password,
    [string]$WorkingDirectory
)

# 如果没有传递密码参数，则使用默认密码
if (-not $password) {
    $password = "123456" # 最长12个字符
}

# 定义密码框和设置按钮的坐标
$passwordBoxX = 155
$passwordBoxY = 160
$settingsButtonX = 200
$settingsButtonY = 255
Add-Type @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public class Win32API {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll", SetLastError=true)]
    public static extern IntPtr SetFocus(IntPtr hWnd);
    
    [DllImport("user32.dll", SetLastError=true)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    
    
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT {
        public int X;
        public int Y;
    }


    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    public const int INPUT_MOUSE = 0;
    public const uint MOUSEEVENTF_MOVE = 0x0001;
    public const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    public const uint MOUSEEVENTF_LEFTUP = 0x0004;
    public const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    public const uint MOUSEEVENTF_RIGHTUP = 0x0010;
    public const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
    public const uint MOUSEEVENTF_MIDDLEUP = 0x0040;
    public const uint MOUSEEVENTF_ABSOLUTE = 0x8000;
    
    public struct MOUSEINPUT {
    public int dx;
    public int dy;
    public uint mouseData;
    public uint dwFlags;
    public uint time;
    public IntPtr dwExtraInfo;
}

public struct INPUT {
    public int type;
    public MOUSEINPUT mi;
}
}


"@

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Get-WindowsAndProcessIdsByPocessNameClassName {
    param(
        $processName, $className
    )
    $windowsInfo = New-Object 'System.Collections.Generic.List[object]'  # 兼容PowerShell 2.0

    # 定义枚举窗口的回调函数
    $enumFunc = [Win32API+EnumWindowsProc] {
        param($hWnd, $lParam)
        
        # 在回调函数内部定义局部变量以接收进程ID
        $processId = 0
        [void][Win32API]::GetWindowThreadProcessId($hWnd, [ref]$processId)
        $process = Get-Process -Id $processId
        $thisProcessName = $process.Name
        # 获取窗口标题
        $title = New-Object System.Text.StringBuilder 1024
        [void][Win32API]::GetWindowText($hWnd, $title, $title.Capacity)

        # 获取窗口类名
        $thisClassName = New-Object System.Text.StringBuilder 1024
        [void][Win32API]::GetClassName($hWnd, $thisClassName, $thisClassName.Capacity)


        if ($thisProcessName -eq $processName -and $thisClassName.ToString() -eq $className) {
            
            
            $windowsInfo.Add(@{
                    WindowHandle = $hWnd
                    ProcessId    = $processId
                    Title        = $title.ToString()
                    ClassNmae    = $thisClassName
                })
            return $false #找到目标窗口后返回false，停止枚举
        }
        return $true
    }

    # 调用EnumWindows并传入回调函数
    $result = [Win32API]::EnumWindows($enumFunc, [IntPtr]::Zero)

    if (-not $result) {
        Write-Host "EnumWindows failed."
        return $windowsInfo
    }

    return $windowsInfo
}

function Capture-Window {
    param (
        [IntPtr]$hWnd
    )
    $rect = New-Object Win32API+RECT
    [Win32API]::GetWindowRect($hWnd, [ref]$rect)

    $bounds = [Drawing.Rectangle]::FromLTRB($rect.Left, $rect.Top, $rect.Right, $rect.Bottom)
    $bitmap = New-Object Drawing.Bitmap($bounds.width, $bounds.height)
    $graphics = [Drawing.Graphics]::FromImage($bitmap)

    $graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.size)

    $bitmap.Save("C:\" + $rect.Left.ToString() + $rect.Top.ToString() + ".png") # 修改为您希望保存图片的位置
    $graphics.Dispose()
    $bitmap.Dispose()
}

function Click-InnerWindow {
    param (
        $hWnd, $x, $y
    )
    $windowRect = New-Object Win32API+RECT
    $result = [Win32API]::GetWindowRect($hWnd, [ref]$windowRect)
    $windowRect
    $clientRect = New-Object Win32API+RECT
    $result = [Win32API]::GetClientRect($hWnd, [ref]$clientRect)
    $clientRect
    $pointInClient = New-Object Win32API+POINT
    $pointInClient.X = $x  
    $pointInClient.Y = $y  

    $result = [Win32API]::ClientToScreen($hWnd, [ref]$pointInClient)

    if ($hWnd -ne [System.IntPtr]::Zero) {
        # 创建MOUSEINPUT结构体实例用于移动鼠标到指定位置
        $mouseMove = New-Object -TypeName Win32API+MOUSEINPUT
        $mouseMove.dx = [int]($pointInClient.X * (65536 / [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Width))
        $mouseMove.dy = [int]($pointInClient.Y * (65536 / [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Height))
        $mouseMove.dwFlags = [Win32API]::MOUSEEVENTF_MOVE -bor [Win32API]::MOUSEEVENTF_ABSOLUTE

        # 创建一个MOUSEINPUT结构体实例用于模拟左键点击
        $mouseDown = New-Object -TypeName Win32API+MOUSEINPUT
        $mouseDown.dwFlags = [Win32API]::MOUSEEVENTF_LEFTDOWN

        $mouseUp = New-Object -TypeName Win32API+MOUSEINPUT
        $mouseUp.dwFlags = [Win32API]::MOUSEEVENTF_LEFTUP

        # 创建INPUT数组并将MOUSEINPUT结构体添加进去
        $inputMouseMove = New-Object -TypeName Win32API+INPUT
        $inputMouseMove.type = [Win32API]::INPUT_MOUSE
        $inputMouseMove.mi = $mouseMove

        $inputMouseDown = New-Object -TypeName Win32API+INPUT
        $inputMouseDown.type = [Win32API]::INPUT_MOUSE
        $inputMouseDown.mi = $mouseDown

        $inputMouseUp = New-Object -TypeName Win32API+INPUT
        $inputMouseUp.type = [Win32API]::INPUT_MOUSE
        $inputMouseUp.mi = $mouseUp

        # 发送鼠标移动事件
        [Win32API]::SendInput(1, [Win32API+INPUT[]]@($inputMouseMove), [System.Runtime.InteropServices.Marshal]::SizeOf([Type][Win32API+INPUT]))

        # 短暂等待以模仿真实的用户操作
        Start-Sleep -Milliseconds 50

        # 发送鼠标左键按下和释放事件
        [Win32API]::SendInput(1, [Win32API+INPUT[]]@($inputMouseDown), [System.Runtime.InteropServices.Marshal]::SizeOf([Type][Win32API+INPUT]))
        [Win32API]::SendInput(1, [Win32API+INPUT[]]@($inputMouseUp), [System.Runtime.InteropServices.Marshal]::SizeOf([Type][Win32API+INPUT]))
    
    }
    return $pointInClient
}


# 定义最大等待时间和每次检查间隔
$maxWaitTimeMilliseconds = 10000
$sleepIntervalMilliseconds = 50

$startTime = [System.DateTime]::Now

$targetClassName = "Qt5QWindowIcon" # 根据窗口类名查找窗口

# 获取脚本所在目录
if (-not $WorkingDirectory) {
    if (-not $PSScriptRoot) {
        $scriptPath = $MyInvocation.MyCommand.Path # 兼容PowerShell 2.0
        if (-not $scriptPath) {
            Write-Error "无法确定脚本路径。请确保脚本是从文件运行的。"
            exit 1
        }
        $scriptDirectory = Split-Path -Parent $scriptPath
    }
    else {
        $scriptDirectory = $PSScriptRoot
    }
}
else {
    $scriptDirectory = $WorkingDirectory
}

$workingDirectory = $scriptDirectory # 脚本所在目录作工作目录

Set-Location -Path $workingDirectory

$exeName = "kylin-cloud-printer-server.exe" # 服务端程序名
$BaseName = (Get-Item $exeName).BaseName # 进程名


$exePath = Join-Path -Path $workingDirectory -ChildPath $exeName # 完整路径
# 检查目标程序是否存在
if (-not (Test-Path $exePath)) {
    Write-Error "目标程序不存在: $exePath"
    exit 1
}

# 启动目标程序
Start-Process -FilePath $exePath -WorkingDirectory $workingDirectory # -PassThru
# 等待程序启动
Start-Sleep -Milliseconds 500
# 循环查找目标窗口，直到找到或超时
do {
    # 获取指定进程名和类名的窗口
    $windows = Get-WindowsAndProcessIdsByPocessNameClassName -processName $BaseName -className $targetClassName

    
    if ($windows) {
        break
    }
    Start-Sleep -Milliseconds $sleepIntervalMilliseconds
} while ([System.DateTime]::Now.Subtract($startTime).TotalMilliseconds -lt $maxWaitTimeMilliseconds)


$windows | ForEach-Object {
    $hWnd = $_.WindowHandle
    $result = [Win32API]::SetForegroundWindow($hWnd) # 激活窗口
    $result = [Win32API]::SetForegroundWindow($hWnd) # 激活窗口
    $result = Click-InnerWindow -hWnd $hWnd -x $passwordBoxX -y $passwordBoxY # 点击密码框 
    $result = [Win32API]::SetForegroundWindow($hWnd) # 激活窗口
    [System.Windows.Forms.SendKeys]::SendWait("^a") # 全选密码框文本
    $result = [Win32API]::SetForegroundWindow($hWnd) # 激活窗口
    [System.Windows.Forms.SendKeys]::SendWait($password) # 输入密码
    $result = [Win32API]::SetForegroundWindow($hWnd) # 激活窗口
    
    $result = Click-InnerWindow -hWnd $hWnd -x $settingsButtonX -y $settingsButtonY # 点击设置按钮 
    
    $result = [Win32API]::ShowWindow($hWnd, 6) # 最小化窗口
}