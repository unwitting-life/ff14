Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;
public static class user32 {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern bool MessageBox(IntPtr hWnd, String text, String caption, int options);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName,int nMaxCount);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetKeyboardState(byte[] lpKeyState);

    [DllImport("USER32.dll")]
    public static extern short GetKeyState(Keys vKey);

    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(Keys vKey);
    public const int KEYEVENTF_EXTENDEDKEY = 1;
    public const int KEYEVENTF_KEYUP = 2;
}

public static class utils {
    public static void KeyDown(Keys vKey) {
        user32.keybd_event((byte)vKey, 0, user32.KEYEVENTF_EXTENDEDKEY, 0);
        
    }
    public static void KeyUp(Keys vKey) {
        user32.keybd_event((byte)vKey, 0, user32.KEYEVENTF_EXTENDEDKEY | user32.KEYEVENTF_KEYUP, 0);
    }
    public static void KeyPress(Keys vKey) {
        utils.KeyDown(vKey);
        utils.KeyUp(vKey);
    }
    public static bool IsKeyDown(Keys vKey) {
        return (user32.GetAsyncKeyState(vKey) & 0x8000) == 0x8000;
    }
    public static byte GetKeyStatus(Keys vKey) {
        byte[] keys = new byte[256];
        user32.GetKeyboardState(keys);
        return keys[(byte)vKey];
    }
    public static String GetForegroundClassName() {
        StringBuilder className = new StringBuilder(256);
        user32.GetClassName(user32.GetForegroundWindow(), className, className.Capacity);
        return className.ToString();
    }
}
"@ -ReferencedAssemblies "System.Windows.Forms"
    
$DEFAULT_TRANSITION_DELAY = 1000
$DEFAULT_KEY_PRESS_DELAY = 200

function isFFXI() {
  return [utils]::GetForegroundClassName() -eq "FFXIVGAME"
}

function employees($index) {
    <# ʹ�÷���������һ�¹�Ա�б������Ȼ��NumPad1��ʾ��1�Ź�Ա��ʼ�� #>
    <# ���ĳһ������û����Ч��˵��2������֮����ʱ̫�� #>
    for ($e = $index; $e -le 9; $e = $e + 1) {
        for($i = 1; $i -lt 3; $i = $i + 1) {
            Write-Host "��Ա{$e} �� δѡ��״̬"
            if (-not(isFFXI)) { return }
            [user32]::SetCursorPos(0, 0) | Out-Null
            Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

            Write-Host "δѡ��״̬ �� ��Ա{1}: [System.Windows.Forms.Keys]::Left"
            if (-not(isFFXI)) { return }
            [utils]::KeyPress([System.Windows.Forms.Keys]::Left);
            Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY
        }

        <# ��Ա�б� �� ��Ա$index #>
        for($i = 1; $i -lt $e; $i = $i + 1) {
            Write-Host "��Ա�б�{$i} �� ��Ա{$i + 1} [System.Windows.Forms.Keys]::Down"
            if (-not(isFFXI)) { return }
            [utils]::KeyPress([System.Windows.Forms.Keys]::Down);
            Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY
        }

        Write-Host "��Ա{$e} �� ��Ա��Ӵ��������: [System.Windows.Forms.Keys]::NumPad0"
        if (-not(isFFXI)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY

        Write-Host "��Ա{$e} �� Ӵ��������: ��ʲôָʾ�� [System.Windows.Forms.Keys]::NumPad0"
        if (-not(isFFXI)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY

        for($i = 1; $i -lt 3; $i = $i + 1) {
            Write-Host "��Ա{$e}: ��ʲôָʾ�� �� δѡ��״̬"
            if (-not(isFFXI)) { return }
            [user32]::SetCursorPos(0, 0) | Out-Null
            Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

            Write-Host "��Ա{$e}: ��ʲôָʾ�� �� ���߹���: [System.Windows.Forms.Keys]::Left"
            if (-not(isFFXI)) { return }
            [utils]::KeyPress([System.Windows.Forms.Keys]::Left)
            Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY
        }

        Write-Host "��Ա{$e}: ���߹��� �� ��ҹ���: [System.Windows.Forms.Keys]::Down"
        if (-not(isFFXI)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Down)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        Write-Host "��Ա{$e}: ��ҹ��� �� ���ۣ����������Ʒ��: [System.Windows.Forms.Keys]::Down"
        if (-not(isFFXI)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Down)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        Write-Host "��Ա{$e}: ���ۣ����������Ʒ�� �� ���ۣ���Ա������Ʒ��: [System.Windows.Forms.Keys]::Down"
        if (-not(isFFXI)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Down)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        Write-Host "��Ա{$e}: ���ۣ���Ա������Ʒ�� �� �鿴���ۼ�¼: [System.Windows.Forms.Keys]::Down"
        if (-not(isFFXI)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Down)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        Write-Host "��Ա{$e}: �鿴���ۼ�¼ �� �鿴��Ա̽�����: [System.Windows.Forms.Keys]::Down"
        if (-not(isFFXI)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Down)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        Write-Host "��Ա{$e}: �鿴��Ա̽����� �� ̽�����: [System.Windows.Forms.Keys]::NumPad0"
        if (-not(isFFXI)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY

        Write-Host "��Ա{$e}: ̽����� �� ����ί��: [System.Windows.Forms.Keys]::Left"
        if (-not(isFFXI)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Left)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        Write-Host "��Ա{$e}: ����ί�� �� ̽������: [System.Windows.Forms.Keys]::NumPad0"
        if (-not(isFFXI)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY

        Write-Host "��Ա{$e}: ̽������ �� ί��: [System.Windows.Forms.Keys]::Left"
        if (-not(isFFXI)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Left)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        Write-Host "��Ա{$e}: ί�� �� ��֪���� [System.Windows.Forms.Keys]::NumPad0"
        if (-not(isFFXI)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY

        Write-Host "��Ա{$e}: ��֪���� �� ��ʲôָʾ��[System.Windows.Forms.Keys]::Escape"
        if (-not(isFFXI)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY

        Write-Host "��Ա{$e}: ��ʲôָʾ�� �� �ټ���: [System.Windows.Forms.Keys]::Escape"
        if (-not(isFFXI)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Escape)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY

        Write-Host "��Ա{$e}: �ټ��� �� ��Ա�б�: [System.Windows.Forms.Keys]::Escape"
        if (-not(isFFXI)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Escape)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY
    }
}

while (1) {
    <# Run with administrator! #>
    if ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad1)) {
        employees(1)
    } elseif ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad2)) {
        employees(2)
    } elseif ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad3)) {
        employees(3)
    } elseif ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad4)) {
        employees(4)
    } elseif ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad5)) {
        employees(5)
    } elseif ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad6)) {
        employees(6)
    } elseif ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad7)) {
        employees(7)
    } elseif ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad8)) {
        employees(8)
    } elseif ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad9)) {
        employees(9)
    }
    Start-Sleep -Milliseconds 50
}