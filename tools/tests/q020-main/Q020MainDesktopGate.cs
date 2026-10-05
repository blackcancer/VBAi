using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;

// Test-only main-desktop transport. It never attaches to or closes an existing user IDE.
public static class Q020MainDesktopGate
{
    static string sid, visualStudio, solution, evidenceRoot, solidWorks;
    static int protectedVs;
    static readonly Dictionary<int, Generation> owned = new Dictionary<int, Generation>();
    // Preserve original handles even when a post-CreateProcess identity check refuses.
    static readonly List<NativeChild> retainedChildren = new List<NativeChild>();
    sealed class Generation { internal string Birth, Image; internal NativeChild Original; }
    sealed class NativeIdentity { internal string Birth, Image; }
    delegate bool Visitor(IntPtr w, IntPtr unused);
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)] struct Startup {
        internal uint Size; internal string Reserved, Desktop, Title;
        internal uint X,Y,Width,Height,XChars,YChars,Fill,Flags;
        internal ushort Show,ReservedSize; internal IntPtr ReservedPointer,Input,Output,Error;
    }
    [StructLayout(LayoutKind.Sequential)] struct Child { internal IntPtr Process,Thread; internal uint Pid,Tid; }
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("kernel32.dll")] static extern IntPtr GetCurrentProcess();
    [DllImport("advapi32.dll",SetLastError=true)] static extern bool OpenProcessToken(IntPtr p,uint access,out IntPtr token);
    [DllImport("advapi32.dll",SetLastError=true)] static extern bool GetTokenInformation(IntPtr token,int type,out int value,uint size,out uint needed);
    [DllImport("user32.dll")] static extern IntPtr GetThreadDesktop(uint tid);
    [DllImport("user32.dll")] static extern IntPtr GetProcessWindowStation();
    [DllImport("user32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern bool GetUserObjectInformation(IntPtr h,int type,StringBuilder value,uint size,out uint needed);
    [DllImport("user32.dll",SetLastError=true)] static extern IntPtr OpenInputDesktop(uint flags,bool inherit,uint access);
    [DllImport("user32.dll",SetLastError=true)] static extern bool CloseDesktop(IntPtr h);
    [DllImport("user32.dll",SetLastError=true)] static extern bool EnumDesktopWindows(IntPtr desktop,Visitor callback,IntPtr unused);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr w,out uint pid);
    [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr w,uint type);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr w);
    [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern bool CreateProcessW(string exe,StringBuilder command,IntPtr pa,IntPtr ta,bool inherit,uint flags,IntPtr env,string directory,ref Startup startup,out Child child);
    [DllImport("kernel32.dll",SetLastError=true)] static extern uint WaitForSingleObject(IntPtr h,uint ms);
    [DllImport("kernel32.dll",SetLastError=true)] static extern bool GetExitCodeProcess(IntPtr h,out uint code);
    [DllImport("kernel32.dll",SetLastError=true)] static extern bool CloseHandle(IntPtr h);
    [DllImport("kernel32.dll",SetLastError=true)] static extern IntPtr OpenProcess(uint access,bool inherit,uint pid);
    [DllImport("kernel32.dll",SetLastError=true)] static extern uint GetProcessId(IntPtr handle);
    [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern bool QueryFullProcessImageNameW(IntPtr handle,uint flags,StringBuilder image,ref uint size);
    [DllImport("kernel32.dll",SetLastError=true)] static extern bool GetProcessTimes(IntPtr handle,out long creation,out long exit,out long kernel,out long user);
    static string Name(IntPtr h) { var b=new StringBuilder(256);uint n;if(h==IntPtr.Zero||!GetUserObjectInformation(h,2,b,512,out n))throw new Win32Exception(Marshal.GetLastWin32Error());return b.ToString(); }
    static string Input() { IntPtr h=OpenInputDesktop(0,false,1);if(h==IntPtr.Zero)throw new Win32Exception(Marshal.GetLastWin32Error());try{return Name(h);}finally{if(!CloseDesktop(h))throw new Win32Exception(Marshal.GetLastWin32Error());} }
    internal static void Configure(string userSid,string selectedVs,string selectedSolution,string campaignRoot,int protectedVsPid) {
        if(sid!=null)throw new InvalidOperationException("Main-desktop configuration is immutable.");
        if(string.IsNullOrWhiteSpace(userSid)||!Path.IsPathRooted(selectedVs)||!File.Exists(selectedVs)||Path.GetFileName(selectedVs)!="devenv.exe"||
            !Path.IsPathRooted(selectedSolution)||!File.Exists(selectedSolution)||Path.GetExtension(selectedSolution)!=".sln"||
            !Path.IsPathRooted(campaignRoot)||!Directory.Exists(campaignRoot)||protectedVsPid<=0)
            throw new ArgumentException("Pinned exact selected utility profile, actor and protected IDE required.");
        sid=userSid;visualStudio=Path.GetFullPath(selectedVs);solution=Path.GetFullPath(selectedSolution);evidenceRoot=Path.GetFullPath(campaignRoot).TrimEnd(Path.DirectorySeparatorChar);protectedVs=protectedVsPid;
        RequireCurrent("Default");
    }
    internal static void RequireCurrent(string desktop) {
        if(sid==null||desktop!="Default"||!Environment.Is64BitProcess||WindowsIdentity.GetCurrent().User.Value!=sid||
            Name(GetProcessWindowStation())!="WinSta0"||Name(GetThreadDesktop(GetCurrentThreadId()))!="Default"||Input()!="Default")
            throw new InvalidOperationException("Exact same-user x64 interactive Default context required.");
        IntPtr token;if(!OpenProcessToken(GetCurrentProcess(),8,out token))throw new Win32Exception(Marshal.GetLastWin32Error());
        try{int type,elevated;uint n;if(!GetTokenInformation(token,18,out type,4,out n)||!GetTokenInformation(token,20,out elevated,4,out n))throw new Win32Exception(Marshal.GetLastWin32Error());if(type!=3||elevated!=0)throw new InvalidOperationException("Limited original actor required.");}
        finally{if(!CloseHandle(token))throw new Win32Exception(Marshal.GetLastWin32Error());}
    }
    static NativeIdentity ReadIdentity(IntPtr handle,int pid) {
        if(handle==IntPtr.Zero||pid<=0)throw new ArgumentException("Retained process handle and PID required.");
        uint actualPid=GetProcessId(handle);if(actualPid==0)throw new Win32Exception(Marshal.GetLastWin32Error());
        if(actualPid!=(uint)pid)throw new InvalidOperationException("Retained handle PID differs.");
        uint wait=WaitForSingleObject(handle,0);if(wait==0xFFFFFFFF)throw new Win32Exception(Marshal.GetLastWin32Error());
        if(wait!=258)throw new InvalidOperationException("Original generation is not active.");
        var image=new StringBuilder(32768);uint size=(uint)image.Capacity;
        if(!QueryFullProcessImageNameW(handle,0,image,ref size))throw new Win32Exception(Marshal.GetLastWin32Error());
        if(size==0||size>=image.Capacity||!Path.IsPathRooted(image.ToString()))throw new InvalidOperationException("Native executable identity unavailable.");
        long creation,exit,kernel,user;
        if(!GetProcessTimes(handle,out creation,out exit,out kernel,out user))throw new Win32Exception(Marshal.GetLastWin32Error());
        string birth=DateTime.FromFileTimeUtc(creation).ToString("o");
        wait=WaitForSingleObject(handle,0);if(wait==0xFFFFFFFF)throw new Win32Exception(Marshal.GetLastWin32Error());
        if(wait!=258)throw new InvalidOperationException("Original generation exited during identity read.");
        return new NativeIdentity{Birth=birth,Image=image.ToString()};
    }
    static void Remember(int pid,string birth,string image) {
        Generation previous;owned.TryGetValue(pid,out previous);
        NativeIdentity identity;
        if(previous!=null&&previous.Original!=null)identity=ReadIdentity(previous.Original.ProcessHandle,pid);
        else {
            // Explicitly adopted generations use a temporary query handle, never a launch handle.
            IntPtr query=OpenProcess(0x00101000,false,(uint)pid);
            if(query==IntPtr.Zero)throw new Win32Exception(Marshal.GetLastWin32Error());
            try{identity=ReadIdentity(query,pid);}finally{if(!CloseHandle(query))throw new Win32Exception(Marshal.GetLastWin32Error());}
        }
        if(identity.Birth!=birth||!string.Equals(identity.Image,image,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("Owned original generation differs.");
        owned[pid]=new Generation{Birth=birth,Image=image,Original=previous==null?null:previous.Original};
    }
    internal static void RegisterHost(int pid,string birth,string selectedSw) {
        RequireCurrent("Default");if(pid<=0||pid==protectedVs||!Path.IsPathRooted(selectedSw)||!File.Exists(selectedSw)||!string.Equals(Path.GetFileName(selectedSw),"SLDWORKS.exe",StringComparison.OrdinalIgnoreCase))throw new ArgumentException("Exact owned SOLIDWORKS generation required.");
        if(owned.ContainsKey(pid))throw new InvalidOperationException("Generation registration cannot be replayed.");
        Remember(pid,birth,selectedSw);
    }
    internal static void ConfigureSolidWorksExecutable(string executable) {
        RequireCurrent("Default");
        if(solidWorks!=null||!Path.IsPathRooted(executable)||!File.Exists(executable)||!string.Equals(Path.GetFileName(executable),"SLDWORKS.exe",StringComparison.OrdinalIgnoreCase))throw new ArgumentException("One pinned selected SOLIDWORKS executable required.");
        solidWorks=Path.GetFullPath(executable);
    }
    static string GenerationDirectory(string workingDir) {
        string directory=Path.GetFullPath(workingDir).TrimEnd(Path.DirectorySeparatorChar);
        if(!Directory.Exists(directory)||!string.Equals(Path.GetDirectoryName(directory),evidenceRoot,StringComparison.OrdinalIgnoreCase)||(Path.GetFileName(directory)!="generation-1"&&Path.GetFileName(directory)!="generation-2"))throw new ArgumentException("Exact direct qualification generation directory required.");
        return directory;
    }
    internal static NativeChild LaunchSolidWorks(string executable,string workingDir,string desktop) {
        RequireCurrent(desktop);string directory=GenerationDirectory(workingDir);
        if(solidWorks==null||!string.Equals(Path.GetFullPath(executable),solidWorks,StringComparison.OrdinalIgnoreCase)||executable.IndexOf('"')>=0)throw new ArgumentException("Pinned selected SOLIDWORKS executable required.");
        var existing=Process.GetProcessesByName("SLDWORKS");try{if(existing.Length!=0)throw new InvalidOperationException("An existing SOLIDWORKS process prevents direct qualification launch.");}finally{foreach(var p in existing)p.Dispose();}
        var startup=new Startup{Size=(uint)Marshal.SizeOf(typeof(Startup)),Desktop="WinSta0\\Default"};Child child;
        if(!CreateProcessW(executable,new StringBuilder("\""+executable+"\""),IntPtr.Zero,IntPtr.Zero,false,0,IntPtr.Zero,directory,ref startup,out child))throw new Win32Exception(Marshal.GetLastWin32Error());
        var retained=new NativeChild(child.Process,child.Thread,(int)child.Pid,child.Tid);retainedChildren.Add(retained);
        // Return ownership before every fallible post-creation identity operation.
        // The caller must write its raw creation receipt before ValidateSolidWorksChild.
        return retained;
    }
    internal static void ValidateSolidWorksChild(NativeChild child) {
        RequireCurrent("Default");
        if(child==null||!retainedChildren.Contains(child)||solidWorks==null||owned.ContainsKey(child.ProcessId))throw new InvalidOperationException("One retained unregistered original SOLIDWORKS child required.");
        child.ValidateIdentity(solidWorks);
        owned.Add(child.ProcessId,new Generation{Birth=child.BirthUtc,Image=child.ImagePath,Original=child});
    }
    internal static void RequireOfficeWindowInventory(string desktop,uint pid,bool requireOwned,IntPtr required) {
        RequireCurrent(desktop);Generation expected;if(pid==protectedVs||!owned.TryGetValue((int)pid,out expected))throw new InvalidOperationException("Only launched IDE or explicitly registered owned host accepted.");
        Remember((int)pid,expected.Birth,expected.Image);
        int visited=0,ownedCount=0;bool found=required==IntPtr.Zero;Exception error=null;
        Visitor callback=(w,s)=>{try{
            if(++visited>8192)throw new InvalidOperationException("Default window inventory bound exceeded.");
            uint p;uint tid=GetWindowThreadProcessId(w,out p);
            if(w==required&&(p!=pid||tid==0||GetAncestor(w,2)!=w))throw new InvalidOperationException("Required exact root identity unavailable.");
            // Foreign and transient unrelated windows carry no admission evidence.
            if(p!=pid)return true;
            if(tid==0)throw new InvalidOperationException("Owned window thread identity unavailable.");
            ownedCount++;if(w==required)found=true;return true;
        }catch(Exception e){error=e;return false;}};
        bool complete=EnumDesktopWindows(GetThreadDesktop(GetCurrentThreadId()),callback,IntPtr.Zero);GC.KeepAlive(callback);
        if(error!=null)throw error;if(!complete||!found||(requireOwned&&ownedCount==0))throw new InvalidOperationException("Complete owned Default root inventory required.");
        if(required!=IntPtr.Zero&&!IsWindowVisible(required))throw new InvalidOperationException("Required owned root is not visible.");
        RequireCurrent(desktop);Remember((int)pid,expected.Birth,expected.Image);
    }
    internal static NativeChild Launch(string executable,string[] arguments,string workingDir,string desktop) {
        RequireCurrent(desktop);
        string directory=GenerationDirectory(workingDir);
        if(!string.Equals(executable,visualStudio,StringComparison.OrdinalIgnoreCase)||arguments==null||arguments.Length!=1||
            !string.Equals(Path.GetFullPath(arguments[0]),solution,StringComparison.OrdinalIgnoreCase)||!Directory.Exists(directory)||
            !directory.StartsWith(evidenceRoot+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase)||
            (Path.GetFileName(directory)!="generation-1"&&Path.GetFileName(directory)!="generation-2"))throw new ArgumentException("Only exact selected qualification IDE profile/generation directory may launch.");
        if(executable.IndexOf('"')>=0||arguments[0].IndexOf('"')>=0)throw new ArgumentException("Unexpected quote in pinned path.");
        var startup=new Startup{Size=(uint)Marshal.SizeOf(typeof(Startup)),Desktop="WinSta0\\Default"};Child child;
        if(!CreateProcessW(executable,new StringBuilder("\""+executable+"\" \""+arguments[0]+"\""),IntPtr.Zero,IntPtr.Zero,false,0,IntPtr.Zero,directory,ref startup,out child))throw new Win32Exception(Marshal.GetLastWin32Error());
        var retained=new NativeChild(child.Process,child.Thread,(int)child.Pid,child.Tid);
        retainedChildren.Add(retained);
        if(child.Pid==(uint)protectedVs)throw new InvalidOperationException("Protected user IDE cannot be adopted.");
        return retained;
    }
    internal static void ValidateVisualStudioChild(NativeChild child) {
        RequireCurrent("Default");
        if(child==null||!retainedChildren.Contains(child)||visualStudio==null||owned.ContainsKey(child.ProcessId))throw new InvalidOperationException("One retained unregistered original IDE child required.");
        child.ValidateIdentity(visualStudio);
        owned.Add(child.ProcessId,new Generation{Birth=child.BirthUtc,Image=child.ImagePath,Original=child});
    }
    internal sealed class NativeChild : IDisposable {
        internal IntPtr ProcessHandle{get;private set;}internal int ProcessId{get;private set;}internal uint ThreadId{get;private set;}IntPtr thread;
        internal IntPtr ThreadHandle{get{return thread;}}
        internal string BirthUtc{get;private set;}internal string ImagePath{get;private set;}
        internal bool IdentityAttempted{get;private set;}internal bool IdentityVerified{get;private set;}
        internal NativeChild(IntPtr p,IntPtr t,int pid,uint tid){ProcessHandle=p;thread=t;ProcessId=pid;ThreadId=tid;}
        internal void ValidateIdentity(string expectedImage) {
            if(IdentityAttempted)throw new InvalidOperationException("Original identity admission already attempted; no automatic retry.");
            IdentityAttempted=true;
            var identity=ReadIdentity(ProcessHandle,ProcessId);
            BirthUtc=identity.Birth;ImagePath=identity.Image;
            if(!Path.IsPathRooted(expectedImage)||!string.Equals(Path.GetFullPath(expectedImage),ImagePath,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("Original native executable identity differs.");
            IdentityVerified=true;
        }
        internal bool Wait(int milliseconds){if(milliseconds<0||ProcessHandle==IntPtr.Zero)throw new ArgumentOutOfRangeException("milliseconds");uint result=WaitForSingleObject(ProcessHandle,(uint)milliseconds);if(result==0xFFFFFFFF)throw new Win32Exception(Marshal.GetLastWin32Error());if(result!=0&&result!=258)throw new InvalidOperationException("Unknown original wait outcome.");return result==0;}
        internal uint ExitCode(){if(!Wait(0))throw new InvalidOperationException("Original process exit not observed.");uint code;if(!GetExitCodeProcess(ProcessHandle,out code))throw new Win32Exception(Marshal.GetLastWin32Error());return code;}
        public void Dispose(){if(!Wait(0))throw new InvalidOperationException("Live original owner handle must be retained.");if(thread!=IntPtr.Zero){IntPtr h=thread;thread=IntPtr.Zero;if(!CloseHandle(h))throw new Win32Exception(Marshal.GetLastWin32Error());}if(ProcessHandle!=IntPtr.Zero){IntPtr h=ProcessHandle;ProcessHandle=IntPtr.Zero;if(!CloseHandle(h))throw new Win32Exception(Marshal.GetLastWin32Error());}}
    }
}
