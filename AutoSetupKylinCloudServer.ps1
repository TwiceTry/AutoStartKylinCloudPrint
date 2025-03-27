param (
    [string]$password,
    [string]$WorkingDirectory
)

# ���û�д��������������ʹ��Ĭ������
if (-not $password) {
    $password = "123456" # �12���ַ�
}

# �������������ð�ť������
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
    $windowsInfo = New-Object 'System.Collections.Generic.List[object]'  # ����PowerShell 2.0

    # ����ö�ٴ��ڵĻص�����
    $enumFunc = [Win32API+EnumWindowsProc] {
        param($hWnd, $lParam)
        
        # �ڻص������ڲ�����ֲ������Խ��ս���ID
        $processId = 0
        [void][Win32API]::GetWindowThreadProcessId($hWnd, [ref]$processId)
        $process = Get-Process -Id $processId
        $thisProcessName = $process.Name
        # ��ȡ���ڱ���
        $title = New-Object System.Text.StringBuilder 1024
        [void][Win32API]::GetWindowText($hWnd, $title, $title.Capacity)

        # ��ȡ��������
        $thisClassName = New-Object System.Text.StringBuilder 1024
        [void][Win32API]::GetClassName($hWnd, $thisClassName, $thisClassName.Capacity)


        if ($thisProcessName -eq $processName -and $thisClassName.ToString() -eq $className) {
            
            
            $windowsInfo.Add(@{
                    WindowHandle = $hWnd
                    ProcessId    = $processId
                    Title        = $title.ToString()
                    ClassNmae    = $thisClassName
                })
            return $false #�ҵ�Ŀ�괰�ں󷵻�false��ֹͣö��
        }
        return $true
    }

    # ����EnumWindows������ص�����
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

    $bitmap.Save("C:\" + $rect.Left.ToString() + $rect.Top.ToString() + ".png") # �޸�Ϊ��ϣ������ͼƬ��λ��
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
        # ����MOUSEINPUT�ṹ��ʵ�������ƶ���굽ָ��λ��
        $mouseMove = New-Object -TypeName Win32API+MOUSEINPUT
        $mouseMove.dx = [int]($pointInClient.X * (65536 / [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Width))
        $mouseMove.dy = [int]($pointInClient.Y * (65536 / [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Height))
        $mouseMove.dwFlags = [Win32API]::MOUSEEVENTF_MOVE -bor [Win32API]::MOUSEEVENTF_ABSOLUTE

        # ����һ��MOUSEINPUT�ṹ��ʵ������ģ��������
        $mouseDown = New-Object -TypeName Win32API+MOUSEINPUT
        $mouseDown.dwFlags = [Win32API]::MOUSEEVENTF_LEFTDOWN

        $mouseUp = New-Object -TypeName Win32API+MOUSEINPUT
        $mouseUp.dwFlags = [Win32API]::MOUSEEVENTF_LEFTUP

        # ����INPUT���鲢��MOUSEINPUT�ṹ����ӽ�ȥ
        $inputMouseMove = New-Object -TypeName Win32API+INPUT
        $inputMouseMove.type = [Win32API]::INPUT_MOUSE
        $inputMouseMove.mi = $mouseMove

        $inputMouseDown = New-Object -TypeName Win32API+INPUT
        $inputMouseDown.type = [Win32API]::INPUT_MOUSE
        $inputMouseDown.mi = $mouseDown

        $inputMouseUp = New-Object -TypeName Win32API+INPUT
        $inputMouseUp.type = [Win32API]::INPUT_MOUSE
        $inputMouseUp.mi = $mouseUp

        # ��������ƶ��¼�
        [Win32API]::SendInput(1, [Win32API+INPUT[]]@($inputMouseMove), [System.Runtime.InteropServices.Marshal]::SizeOf([Type][Win32API+INPUT]))

        # ���ݵȴ���ģ����ʵ���û�����
        Start-Sleep -Milliseconds 50

        # �������������º��ͷ��¼�
        [Win32API]::SendInput(1, [Win32API+INPUT[]]@($inputMouseDown), [System.Runtime.InteropServices.Marshal]::SizeOf([Type][Win32API+INPUT]))
        [Win32API]::SendInput(1, [Win32API+INPUT[]]@($inputMouseUp), [System.Runtime.InteropServices.Marshal]::SizeOf([Type][Win32API+INPUT]))
    
    }
    return $pointInClient
}


# �������ȴ�ʱ���ÿ�μ����
$maxWaitTimeMilliseconds = 10000
$sleepIntervalMilliseconds = 50

$startTime = [System.DateTime]::Now

$targetClassName = "Qt5QWindowIcon" # ���ݴ����������Ҵ���

# ��ȡ�ű�����Ŀ¼
if (-not $WorkingDirectory) {
    if (-not $PSScriptRoot) {
        $scriptPath = $MyInvocation.MyCommand.Path # ����PowerShell 2.0
        if (-not $scriptPath) {
            Write-Error "�޷�ȷ���ű�·������ȷ���ű��Ǵ��ļ����еġ�"
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

$workingDirectory = $scriptDirectory # �ű�����Ŀ¼������Ŀ¼

Set-Location -Path $workingDirectory

$exeName = "kylin-cloud-printer-server.exe" # ����˳�����
$BaseName = (Get-Item $exeName).BaseName # ������


$exePath = Join-Path -Path $workingDirectory -ChildPath $exeName # ����·��
# ���Ŀ������Ƿ����
if (-not (Test-Path $exePath)) {
    Write-Error "Ŀ����򲻴���: $exePath"
    exit 1
}

# ����Ŀ�����
Start-Process -FilePath $exePath -WorkingDirectory $workingDirectory # -PassThru
# �ȴ���������
Start-Sleep -Milliseconds 500
# ѭ������Ŀ�괰�ڣ�ֱ���ҵ���ʱ
do {
    # ��ȡָ���������������Ĵ���
    $windows = Get-WindowsAndProcessIdsByPocessNameClassName -processName $BaseName -className $targetClassName

    
    if ($windows) {
        break
    }
    Start-Sleep -Milliseconds $sleepIntervalMilliseconds
} while ([System.DateTime]::Now.Subtract($startTime).TotalMilliseconds -lt $maxWaitTimeMilliseconds)


$windows | ForEach-Object {
    $hWnd = $_.WindowHandle
    $result = [Win32API]::SetForegroundWindow($hWnd) # �����
    $result = [Win32API]::SetForegroundWindow($hWnd) # �����
    $result = Click-InnerWindow -hWnd $hWnd -x $passwordBoxX -y $passwordBoxY # �������� 
    $result = [Win32API]::SetForegroundWindow($hWnd) # �����
    [System.Windows.Forms.SendKeys]::SendWait("^a") # ȫѡ������ı�
    $result = [Win32API]::SetForegroundWindow($hWnd) # �����
    [System.Windows.Forms.SendKeys]::SendWait($password) # ��������
    $result = [Win32API]::SetForegroundWindow($hWnd) # �����
    
    $result = Click-InnerWindow -hWnd $hWnd -x $settingsButtonX -y $settingsButtonY # ������ð�ť 
    
    $result = [Win32API]::ShowWindow($hWnd, 6) # ��С������
}