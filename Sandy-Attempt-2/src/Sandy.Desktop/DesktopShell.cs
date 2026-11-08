using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using Sandy.Blazor;
using Sandy.Core;
namespace Sandy.Desktop;

internal sealed class DesktopShell(Window window,Workspace state) : IDisposable
{
    private const int HotkeyId=0x5341,TrayId=1,TrayMessage=0x8000+65;
    private IntPtr handle;private HwndSource? source;private bool icon,hotkey,exiting;private ShellOptions options=new();
    internal bool TrayRegistered=>icon;
    internal bool HotkeyRegistered=>hotkey;
    internal void Initialize(ShellOptions initial)
    {
        handle=new WindowInteropHelper(window).Handle;source=HwndSource.FromHwnd(handle);source?.AddHook(Message);Apply(initial);
    }
    internal void Apply(ShellOptions value)
    {
        if(value.VirtualKey is <0x30 or >0x7B||(value.Modifiers&~7U)!=0||value.GlobalHotkey&&value.Modifiers==0)throw new ArgumentException("Choose a Ctrl/Alt/Shift shortcut with a letter or function key.");
        var old=options;
        try{ApplyNative(value);options=value;}catch{ApplyNative(old);throw;}
    }
    private void ApplyNative(ShellOptions value)
    {
        if(handle==IntPtr.Zero){options=value;return;}
        if(hotkey){UnregisterHotKey(handle,HotkeyId);hotkey=false;}
        if(value.GlobalHotkey){if(!RegisterHotKey(handle,HotkeyId,value.Modifiers|0x4000,value.VirtualKey))throw new Win32Exception(Marshal.GetLastWin32Error(),"That global shortcut is already in use.");hotkey=true;}
        if(icon&&!value.Tray){var data=Data();Shell_NotifyIcon(2,ref data);icon=false;}
        if(!icon&&value.Tray){var data=Data();if(!Shell_NotifyIcon(0,ref data))throw new Win32Exception(Marshal.GetLastWin32Error(),"Could not create the Sandy tray icon.");icon=true;}
    }
    internal bool HideOnClose(CancelEventArgs e){if(!options.Tray||exiting)return false;e.Cancel=true;window.Hide();return true;}
    private void Show(){window.Show();if(window.WindowState==WindowState.Minimized)window.WindowState=WindowState.Normal;window.Activate();}
    private IntPtr Message(IntPtr hwnd,int message,IntPtr wParam,IntPtr lParam,ref bool handled)
    {
        if(message==0x0312&&wParam.ToInt32()==HotkeyId){Show();handled=true;}
        if(message==TrayMessage)
        {
            if(lParam.ToInt32()==0x0203)Show();
            else if(lParam.ToInt32()==0x0205)
            {
                var menu=new ContextMenu();var open=new MenuItem{Header="Open Sandy"};open.Click+=(_,_)=>Show();menu.Items.Add(open);
                foreach(var template in state.Library.Templates.Where(state.IsFavorite).Take(50)){var item=new MenuItem{Header=template.Name};item.Click+=(_,_)=>{state.Select(template);Show();};menu.Items.Add(item);}
                menu.Items.Add(new Separator());var exit=new MenuItem{Header="Exit Sandy"};exit.Click+=(_,_)=>{exiting=true;window.Close();if(window.IsVisible)exiting=false;};menu.Items.Add(exit);
                SetForegroundWindow(handle);menu.IsOpen=true;
            }
            handled=true;
        }
        return IntPtr.Zero;
    }
    private NotifyIconData Data()=>new(){Size=(uint)Marshal.SizeOf<NotifyIconData>(),Window=handle,Id=TrayId,Flags=1|2|4,CallbackMessage=TrayMessage,Icon=LoadIcon(IntPtr.Zero,new IntPtr(32512)),Tip="Sandy - Slice and Dice",Info="",InfoTitle=""};
    public void Dispose(){if(hotkey)UnregisterHotKey(handle,HotkeyId);if(icon){var data=Data();Shell_NotifyIcon(2,ref data);}source?.RemoveHook(Message);}
    [StructLayout(LayoutKind.Sequential,CharSet=CharSet.Unicode)]private struct NotifyIconData
    {
        public uint Size;public IntPtr Window;public uint Id,Flags,CallbackMessage;public IntPtr Icon;
        [MarshalAs(UnmanagedType.ByValTStr,SizeConst=128)]public string Tip;public uint State,StateMask;
        [MarshalAs(UnmanagedType.ByValTStr,SizeConst=256)]public string Info;public uint TimeoutOrVersion;
        [MarshalAs(UnmanagedType.ByValTStr,SizeConst=64)]public string InfoTitle;public uint InfoFlags;public Guid Guid;public IntPtr BalloonIcon;
    }
    [DllImport("shell32.dll",CharSet=CharSet.Unicode,SetLastError=true)][return:MarshalAs(UnmanagedType.Bool)]private static extern bool Shell_NotifyIcon(uint message,ref NotifyIconData data);
    [DllImport("user32.dll",SetLastError=true)]private static extern bool RegisterHotKey(IntPtr window,int id,uint modifiers,uint key);
    [DllImport("user32.dll")]private static extern bool UnregisterHotKey(IntPtr window,int id);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)]private static extern IntPtr LoadIcon(IntPtr instance,IntPtr icon);
    [DllImport("user32.dll")]private static extern bool SetForegroundWindow(IntPtr window);
}
