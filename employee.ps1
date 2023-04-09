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

<# 最大雇员数 #>
$MAX_EMPLOYEE = 9

function isFFXIV() {
  return [utils]::GetForegroundClassName() -eq "FFXIVGAME"
}

function employees($index) {
    <# 使用方法：鼠标点击雇员列表将雇员列表激活，然后NumPad1表示从1号雇员开始收货， #>
    <# NumPad2代表从2号雇员开始收货，以此类推 #>
    <# 如果某一个按键没有生效，说明2个按键之间延时太短 #>
    for ($e = $index; $e -le $MAX_EMPLOYEE; $e = $e + 1) {
        <# 雇员{$e} → 未选中状态 #>
        if (-not(isFFXIV)) { return }
        [user32]::SetCursorPos(0, 0) | Out-Null
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# 未选中状态 → 雇员{1} #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Left);
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        for($i = 1; $i -lt $e; $i = $i + 1) {
            <# 雇员列表{$i} → 雇员{$i + 1} #>
            if (-not(isFFXIV)) { return }
            [utils]::KeyPress([System.Windows.Forms.Keys]::Down);
            Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY
        }

        <# 雇员{$e} → 雇员：哟！我来了 #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY

        <# 雇员{$e} → 哟！我来了: 有什么指示？ #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# 雇员{$e}: 有什么指示？ → 未选中状态 #>
        if (-not(isFFXIV)) { return }
        [user32]::SetCursorPos(0, 0) | Out-Null
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# 雇员{$e}: 有什么指示？ → 道具管理 #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Left)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# 雇员{$e}: 道具管理 → 金币管理 #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Down)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# 雇员{$e}: 金币管理 → 出售（玩家所持物品）#>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Down)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# 雇员{$e}: 出售（玩家所持物品） → 出售（雇员所持物品）#>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Down)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# 雇员{$e}: 出售（雇员所持物品） → 查看出售记录 #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Down)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# 雇员{$e}: 查看出售记录 → 查看雇员探险情况 #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Down)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# 雇员{$e}: 查看雇员探险情况 → 探险情况 #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_SHORT_TRANSITION_DELAY

        <# 雇员{$e}: 探险情况 → 重新委托 #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Left)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# 雇员{$e}: 重新委托 → 探险详情 #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_SHORT_TRANSITION_DELAY

        <# 雇员{$e}: 探险详情 → 委托 #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Left)
        Start-Sleep -Milliseconds $DEFAULT_KEY_PRESS_DELAY

        <# 雇员{$e}: 委托 → 我知道了 #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY

        <# 雇员{$e}: 我知道了 → 有什么指示？ #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::NumPad0)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY

        <# 雇员{$e}: 有什么指示？ → 再见了 #>
        if (-not(isFFXIV)) { return }
        [utils]::KeyPress([System.Windows.Forms.Keys]::Escape)
        Start-Sleep -Milliseconds $DEFAULT_TRANSITION_DELAY

        <# 雇员{$e}: 再见了 → 雇员列表 #>
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