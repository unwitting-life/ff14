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
$DEFAULT_SHORT_TRANSITION_DELAY = 500
$DEFAULT_KEY_PRESS_DELAY = 100

<# ����Ա�� #>
$MAX_EMPLOYEE = 9

function isFFXIV() {
  return [utils]::GetForegroundClassName() -eq "FFXIVGAME"
}

function employees($index) {
    <# ʹ�÷������������Ա�б���Ա�б��Ȼ��NumPad1��ʾ��1�Ź�Ա��ʼ�ջ��� #>
    <# NumPad2�����2�Ź�Ա��ʼ�ջ����Դ����� #>
    <# ���ĳһ������û����Ч��˵��2������֮����ʱ̫�� #>
    for ($e = $index; $e -le $MAX_EMPLOYEE; $e = $e + 1) {
        <# ��Ա{$e} �� δѡ��״̬ #>
        if (-not(isFFXIV)) { return }
        [user32]::SetCursorPos(0, 0) | Out-Null
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# δѡ��״̬ �� ��Ա{1} #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Left);
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        for($i = 1; $i -lt $e; $i = $i + 1) {
            <# ��Ա�б�{$i} �� ��Ա{$i + 1} #>
            if (-not(isFFXIV)) { return }
            [utils]::KeyPress([System.Windows.Forms.Keys]::Down);
            Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY
        }

        <# ��Ա{$e} �� ��Ա��Ӵ�������� #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY

        <# ��Ա{$e} �� Ӵ��������: ��ʲôָʾ�� #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# ��Ա{$e}: ��ʲôָʾ�� �� δѡ��״̬ #>
        if (-not(isFFXIV)) { return }
        [user32]::SetCursorPos(0, 0) | Out-Null
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# ��Ա{$e}: ��ʲôָʾ�� �� ���߹��� #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Left)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# ��Ա{$e}: ���߹��� �� ��ҹ��� #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Down)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# ��Ա{$e}: ��ҹ��� �� ���ۣ����������Ʒ��#>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Down)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# ��Ա{$e}: ���ۣ����������Ʒ�� �� ���ۣ���Ա������Ʒ��#>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Down)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# ��Ա{$e}: ���ۣ���Ա������Ʒ�� �� �鿴���ۼ�¼ #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Down)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# ��Ա{$e}: �鿴���ۼ�¼ �� �鿴��Ա̽����� #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Down)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# ��Ա{$e}: �鿴��Ա̽����� �� ̽����� #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_SHORT_TRANSITION_DELAY

        <# ��Ա{$e}: ̽����� �� ����ί�� #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Left)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# ��Ա{$e}: ����ί�� �� ̽������ #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_SHORT_TRANSITION_DELAY

        <# ��Ա{$e}: ̽������ �� ί�� #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Left)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# ��Ա{$e}: ί�� �� ��֪���� #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY

        <# ��Ա{$e}: ��֪���� �� ��ʲôָʾ�� #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY

        <# ��Ա{$e}: ��ʲôָʾ�� �� �ټ��� #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Escape)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY

        <# ��Ա{$e}: �ټ��� �� ��Ա�б� #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Escape)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY
    }
}

while (1) {
    <# Run with administrator! #>
    if ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad1)) {
        employees(1)
        break
    } elseif ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad2)) {
        employees(2)
        break
    } elseif ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad3)) {
        employees(3)
        break
    } elseif ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad4)) {
        employees(4)
        break
    } elseif ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad5)) {
        employees(5)
        break
    } elseif ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad6)) {
        employees(6)
        break
    } elseif ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad7)) {
        employees(7)
        break
    } elseif ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad8)) {
        employees(8)
        break
    } elseif ([utils]::IsKeyDown([System.Windows.Forms.Keys]::NumPad9)) {
        employees(9)
        break
    }
    Start-Sleep -Milliseconds 50
}