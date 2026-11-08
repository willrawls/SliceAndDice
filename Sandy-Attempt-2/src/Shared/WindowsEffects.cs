using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Sandy.Core;

namespace Sandy.Windows;

/// <summary>Runs only after a user accepts an evaluated plan. Evaluation itself has no side effects.</summary>
internal static class WindowsEffects
{
    internal static bool IsFile(ExternalEffect effect) => effect.Kind is "FileAppend" or "FileOverwrite" or "DeleteFile" or "Rename" or "CreateDirectory" or "DeleteDirectory";
    internal static void Validate(IEnumerable<ExternalEffect> effects)
    {
        foreach(var effect in effects)
            if(!IsFile(effect) && effect.Kind is not ("Message" or "Clipboard" or "Setting" or "DeleteSetting" or "Process" or "Activate" or "SendKeys" or "Pause" or "WaitForFile" or "EditorClose" or "EditorCloseAll"))
                throw new NotSupportedException("This host cannot apply " + effect.Kind + ": " + effect.Target);
    }
    internal static string? ReadSetting(string name)
    {
        var (section,key)=SettingPath(name);
        using var registry=Registry.CurrentUser.OpenSubKey(@"Software\VB and VBA Program Settings\Slice and Dice\"+section);
        return registry?.GetValue(key) as string;
    }
    private static (string Section,string Key) SettingPath(string name)
    {
        var slash=name.LastIndexOf('/');if(slash<1)throw new FormatException("Setting requires section/key.");
        return(name[..slash].Replace('/','\\'),name[(slash+1)..]);
    }
    internal static string? Apply(ExternalEffect effect)
    {
        switch(effect.Kind)
        {
            case "Clipboard": SetClipboard(effect.Value);break;
            case "Setting":
                var (section,key)=SettingPath(effect.Target);
                using(var registry=Registry.CurrentUser.CreateSubKey(@"Software\VB and VBA Program Settings\Slice and Dice\"+section))registry.SetValue(key,effect.Value);
                break;
            case "DeleteSetting":
                var (removeSection,removeKey)=SettingPath(effect.Target);
                using(var registry=Registry.CurrentUser.OpenSubKey(@"Software\VB and VBA Program Settings\Slice and Dice\"+removeSection,true))registry?.DeleteValue(removeKey,false);
                break;
            case "Process": return Microsoft.VisualBasic.Interaction.Shell(effect.Target,Microsoft.VisualBasic.AppWinStyle.NormalFocus).ToString(System.Globalization.CultureInfo.InvariantCulture);
            case "Activate": if(int.TryParse(effect.Target,out var processId))Microsoft.VisualBasic.Interaction.AppActivate(processId);else Microsoft.VisualBasic.Interaction.AppActivate(effect.Target);break;
            case "SendKeys":
                OnSta(()=> {
                    var type=Type.GetTypeFromProgID("WScript.Shell")??throw new NotSupportedException("Windows Script Host is unavailable.");
                    var shell=Activator.CreateInstance(type)!;
                    try{type.InvokeMember("SendKeys",System.Reflection.BindingFlags.InvokeMethod,null,shell,[effect.Target]);}
                    finally{Marshal.FinalReleaseComObject(shell);}
                });break;
            // These two commands were no-ops in Sandy/sadFile.cls. Keep that behavior.
            case "Pause":case "WaitForFile":break;
            case "Message":break;
            // No native IDE code windows exist in the standalone host. The VSIX handles these itself.
            case "EditorClose":case "EditorCloseAll":break;
            default:if(!IsFile(effect))throw new NotSupportedException("Unknown effect: "+effect.Kind);break;
        }
        return null;
    }
    private static void OnSta(Action action)
    {
        Exception? failure=null;var thread=new Thread(()=>{try{action();}catch(Exception e){failure=e;}});
        thread.SetApartmentState(ApartmentState.STA);thread.Start();thread.Join();if(failure!=null)System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }
    internal static string ReadClipboard()
    {
        if(!OpenClipboard(IntPtr.Zero))throw new System.ComponentModel.Win32Exception();
        try{var data=GetClipboardData(13);if(data==IntPtr.Zero)return "";var pointer=GlobalLock(data);if(pointer==IntPtr.Zero)throw new System.ComponentModel.Win32Exception();try{return Marshal.PtrToStringUni(pointer)??"";}finally{GlobalUnlock(data);}}
        finally{CloseClipboard();}
    }
    private static void SetClipboard(string text)
    {
        var owner=CreateWindowEx(0,"STATIC","Sandy Clipboard",0,0,0,0,0,new IntPtr(-3),IntPtr.Zero,IntPtr.Zero,IntPtr.Zero);
        if(owner==IntPtr.Zero)throw new System.ComponentModel.Win32Exception();
        var data=GlobalAlloc(0x0042,(nuint)((text.Length+1)*2));if(data==IntPtr.Zero)throw new OutOfMemoryException();
        try
        {
            var pointer=GlobalLock(data);if(pointer==IntPtr.Zero)throw new System.ComponentModel.Win32Exception();
            try{Marshal.Copy((text+'\0').ToCharArray(),0,pointer,text.Length+1);}finally{GlobalUnlock(data);}
            if(!OpenClipboard(owner))throw new System.ComponentModel.Win32Exception();
            try{if(!EmptyClipboard()||SetClipboardData(13,data)==IntPtr.Zero)throw new System.ComponentModel.Win32Exception();data=IntPtr.Zero;}
            finally{CloseClipboard();}
        }
        finally{if(data!=IntPtr.Zero)GlobalFree(data);DestroyWindow(owner);}
    }
    [DllImport("user32.dll",SetLastError=true)] private static extern bool OpenClipboard(IntPtr owner);
    [DllImport("user32.dll")] private static extern bool CloseClipboard();
    [DllImport("user32.dll",SetLastError=true)] private static extern bool EmptyClipboard();
    [DllImport("user32.dll",SetLastError=true)] private static extern IntPtr SetClipboardData(uint format,IntPtr data);
    [DllImport("user32.dll")] private static extern IntPtr GetClipboardData(uint format);
    [DllImport("kernel32.dll",SetLastError=true)] private static extern IntPtr GlobalAlloc(uint flags,nuint size);
    [DllImport("kernel32.dll",SetLastError=true)] private static extern IntPtr GlobalLock(IntPtr memory);
    [DllImport("kernel32.dll")] private static extern bool GlobalUnlock(IntPtr memory);
    [DllImport("kernel32.dll")] private static extern IntPtr GlobalFree(IntPtr memory);
    [DllImport("user32.dll",CharSet=CharSet.Unicode,SetLastError=true)] private static extern IntPtr CreateWindowEx(uint extended,string className,string windowName,uint style,int x,int y,int width,int height,IntPtr parent,IntPtr menu,IntPtr instance,IntPtr parameter);
    [DllImport("user32.dll")] private static extern bool DestroyWindow(IntPtr window);
}
