;@Ahk2Exe-AddResource OpenDafYomi.ico, 5
;@Ahk2Exe-AddResource images/close.ico, 15
;@Ahk2Exe-AddResource images/reload.ico, 25
;@Ahk2Exe-AddResource images/pdf.ico, 35

#Persistent
#SingleInstance, Force
#WinActivateForce
#Include JSON.ahk
#Include Jxon.ahk
#Include GetSwitchParams.ahk

global OpenWebOnly:=0
If (GetSwitchParams("web"))
{
	OpenWebOnly:=1
}

CoordMode, ToolTip,Screen

SplitPath, A_ScriptFullPath,,,, A_ScriptNameNoExtension
Menu, Tray, NoStandard
Menu, Tray, Add, &Abrir PDF, OpenPDF
Menu, Tray, Add, &Abrir en Jabrutouch, OpenWeb
Menu, Tray, Add
Menu, Tray, Add, &Recargar, ReloadApplication
Menu, Tray, Add, &Salir, ExitApplication
If (A_IsCompiled)
{
    Menu, Tray, Icon, &Salir, %A_ScriptFullPath%, -15
    Menu, Tray, Icon, &Recargar, %A_ScriptFullPath%, -25
    Menu, Tray, Icon, &Abrir PDF, %A_ScriptFullPath%, -35
    Menu, Tray, Icon, &Abrir en Jabrutouch, %A_ScriptFullPath%, -5
}

database=%A_ScriptDir%\cms3926896145652424982.csv
FileInstall cms3926896145652424982.csv,%database%,1

vlcPath:=0

pf64 := StrReplace(A_ProgramFiles, " (x86)","")
pf32 := pf64 . " (x86)"

if(FileExist(pf32 . "\VideoLAN\VLC\vlc.exe"))
    vlcPath:=pf32 . "\VideoLAN\VLC\vlc.exe"
if(FileExist(pf64 . "\VideoLAN\VLC\vlc.exe"))
    vlcPath:=pf64 . "\VideoLAN\VLC\vlc.exe"

speedVLC:="1.25"
IniRead, speedVLC, %A_ScriptDir%\%A_ScriptNameNoExtension%.ini, VLC, DefaultSpeed, %speedVLC%
IniWrite, %speedVLC%, %A_ScriptDir%\%A_ScriptNameNoExtension%.ini, VLC, DefaultSpeed

Current:=A_YYYY "-" A_MM "-" A_DD
found:=0
Loop, Read, %database%
{
    Loop, Parse, A_LoopReadLine, CSV
    {
        If A_Index=1
        {
		    if(A_LoopField==Current)
            {
                found:=1
            }
        }
        If (found=1 && A_Index=5)
        {
		    VideoID:=A_LoopField
            break
        }
    }
    if(found)
        break
}

try
{
oWhr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
oWhr.SetTimeouts("600000", "600000", "600000", "600000")
oWhr.Open("GET", "https://api.jabrutouch.com/api/delivery/spaces/1/entries?language=es&contentType=lesson&number=" VideoID, false)
oWhr.SetRequestHeader("x-api-key", "a63p1n4ujf13hsd4mf8mgfkncp")
oWhr.SetRequestHeader("Content-Type", "application/json")
oWhr.Send(payload)
if(oWhr.Status==200)
    response:=oWhr.ResponseText
}
catch
{
    msgbox, No se pudo contactar el servidor de Jabrutouch, se cerrará el programa ahora.
    ExitApp
}

data := Jxon_Load(response)

for i, d in data
{
    VideoIDLong := d["id"]
    if(OpenWebOnly)
    {
        GoSub, OpenWeb
        ExitApp
    }
    for j, f in d["fields"]
    {
        if(f["key"]="video")
        {
            for k, f0 in f["value"]["fields"]
            {
                if(f0["key"]="url")
                {
                    downloadVideoFile:=0
                    IniRead, downloadVideoFile, %A_ScriptDir%\%A_ScriptNameNoExtension%.ini, Settings, DownloadVideoFile, %downloadVideoFile%
                    IniWrite, %downloadVideoFile%, %A_ScriptDir%\%A_ScriptNameNoExtension%.ini, Settings, DownloadVideoFile
                    if(downloadVideoFile)
                    {
                        localVideoPath:=A_ScriptDir . "\dafyomi.mp4"
                        ModifiedDate:=0
                        if(FileExist(localVideoPath))
                        {
                            FileGetTime, ModifiedTime, %localVideoPath%, M
                            FormatTime, ModifiedDate, %ModifiedTime%, yyyy-MM-dd
                        }
                        downloaded:=0
                        if(!FileExist(localVideoPath) || ModifiedDate!=Current)
                        {
                            downloaded:=1
                            if(FileExist(localVideoPath))
                                FileDelete, %localVideoPath%
                            DownloadFile(f0["value"], localVideoPath)
                        }
                        if(FileExist(vlcPath))
                        {
                            if(downloaded)
                            {
                                Run, % """" . vlcPath . """ --start-time 1 --no-start-paused --repeat --no-play-and-pause --rate=" . speedVLC . " dafyomi.mp4", %A_ScriptDir%,,VLCPID
                            }
                            else
                            {
                                Run, % """" . vlcPath . """ --repeat --no-play-and-pause --rate=" . speedVLC . " dafyomi.mp4", %A_ScriptDir%,,VLCPID
                            }
                            WinWait, ahk_pid %VLCPID%,,3
                        }
                        else
                            Run, % localVideoPath
                    }
                    else
                    {
                        localVideoPath:=f0["value"]
                        if(FileExist(vlcPath))
                        {
                            Run, % """" . vlcPath . """ --repeat --no-play-and-pause --rate=" . speedVLC . " """ . localVideoPath . """", %A_ScriptDir%,,VLCPID
                            WinWait, ahk_pid %VLCPID%,,3
                            Sleep, 10000
                        }
                        else
                            Run, % localVideoPath
                    }
                }
            }
        }
        if(f["key"]="pdf")
        {
            pdfUrl:=""
            for k, f0 in f["value"]["fields"]
            {
                if(f0["key"]="url")
                {
                    localPDFPath:=A_ScriptDir . "\dafyomi.pdf"
                    ModifiedDate:=0
                    if(FileExist(localPDFPath))
                    {
                        FileGetTime, ModifiedTime, %localPDFPath%, M
                        FormatTime, ModifiedDate, %ModifiedTime%, yyyy-MM-dd
                    }
                    pdfDownloaded:=0
                    if(!FileExist(localPDFPath) || ModifiedDate!=Current)
                    {
                        pdfDownloaded:=1
                        pdfUrl:=f0["value"]
                        if(FileExist(localPDFPath))
                            FileDelete, %localPDFPath%
                        DownloadFile(f0["value"], localPDFPath)
                    }
                }
            }
        }
        if(f["key"]="name")
        {
            dafName:=f["value"]
        }
        if(f["key"]="rabbi")
        {
            for k, f0 in f["value"]["fields"]
            {
                if(f0["key"]="name")
                {
                    rabbiName:=f0["value"]   
                }
            }
        }
    }
}
if(localPDFPath)
{
    Run, %localPDFPath%,,,PDFPID
    newTitle:=dafName . " - " . rabbiName
    WinWait, ahk_pid %PDFPID%,,3
    WinSetTitle, ahk_pid %PDFPID%,,%newTitle%
}
Sleep, 3000
WinGet, CtrlList, ControlList, ahk_class Qt5QWindowIcon ahk_pid %VLCPID%
Loop, Parse, CtrlList, `n
{
    if (RegExMatch(A_LoopField, "^VLC video output [0-9a-fA-F]{9,}$"))
    {
        VLCVidCtrl:=A_LoopField
    }
}
WorkerW:=PinToDesktop("ahk_exe vlc.exe")
ControlSend,%VLCVidCtrl%, f, ahk_id %WorkerW%

if(PDFPID)
    WinActivate, ahk_pid %PDFPID%

OnExit, ExitApplication
return

OpenPDF:
if(localPDFPath)
{
    if(WinExist("ahk_pid " . PDFPID))
        WinActivate, ahk_pid %PDFPID%
    else
    {
        Run, %localPDFPath%,,,PDFPID
        newTitle:=dafName . " - " . rabbiName
        WinWait, ahk_pid %PDFPID%,,3
        WinSetTitle, ahk_pid %PDFPID%,,%newTitle%
    }
}
else
{
    msgbox, No se pudo descargar el PDF.
}
return

OpenWeb:
    Run, https://www.jabrutouch.com/lesson/%VideoIDLong%
return

PinToDesktop(title="A", OnTop=0)
{
	; Obtain Program Manager Handle
	Progman := DllCall("FindWindowW", "Str", "Progman", "UPtr", 0, "Ptr")

	; Send Message to Program Manager
	; Post-Creator's Update Windows 10. WM_SPAWN_WORKER = 0x052C
	DllCall("SendMessage", "ptr", WinExist("ahk_class Progman"), "uint", 0x052C, "ptr", 0x0000000D, "ptr", 0)
	DllCall("SendMessage", "ptr", WinExist("ahk_class Progman"), "uint", 0x052C, "ptr", 0x0000000D, "ptr", 1)
	; Obtain Handle to Newly Created Window
	WinGet List, List, ahk_class WorkerW
	Loop % List
	{
		If (Found)
		{
			WorkerW := List%A_Index%
			Break
		}
		If (DllCall("FindWindowExW", "Ptr", List%A_Index%, "Ptr", 0, "Str", "SHELLDLL_DefView", "UPtr", 0, "Ptr"))
			Found := TRUE
	}
    WinShow, ahk_id %WorkerW%
	DllCall("SetParent", "Ptr", WinExist(title), "Ptr", WorkerW, "Ptr")
    return WorkerW
}

#If (WinActive("ahk_pid " . PDFPID) && VLCVidCtrl)

^Left::
ControlSend,%VLCVidCtrl%, {Left}, ahk_id %WorkerW%
return

^Right::
ControlSend,%VLCVidCtrl%, {Right}, ahk_id %WorkerW%
return

^Space::
ControlSend,%VLCVidCtrl%, {Space}, ahk_id %WorkerW%
return

^t::
if(!transparency)
	transparency:=3
else
    transparency:=0
if(transparency)
    WinSet, Transparent, % (255*transparency/5), % "ahk_pid " . PDFPID
else
    WinSet, Transparent, Off, % "ahk_pid " . PDFPID
return

^a::
^n::
    apuntesFolder:=A_ScriptDir . "\apuntes"
    if(!FileExist(apuntesFolder))
        FileCreateDir, %apuntesFolder%
    filePath:=apuntesFolder . "\" . dafName . ".md"
    if(!FileExist(filePath))
    {
        contents=
(
# %dafName%
_%rabbiName%_



)
        FileAppend, %contents%, %filePath%
    }
    Run, %filePath%
return

#If

ReloadApplication:
Reload
return

ExitApplication:
WinHide, ahk_id %WorkerW%
KillChildProcesses("OpenDafYomi.exe")
ExitApp
return

isDafYomiPDF()
{
    global PDFPID
    return WinActive("ahk_id " PDFPID)
}

isInvalidFileName(_fileName, _isLong=true)
{
    forbiddenChars := _isLong ? "[<>|""\\/:*?]" : "[;=+<>|""\]\[\\/']"
    ErrorLevel := RegExMatch( _fileName , forbiddenChars )
    Return ErrorLevel
}

cleanFileName(_fileName, replace="-", _isLong=true)
{
    forbiddenChars := _isLong ? "[<>|""\\/:*?]" : "[;=+<>|""\]\[\\/']"
    _fileName := RegExReplace( _fileName , forbiddenChars, replace)
    Return _fileName
}

DownloadFile(UrlToFile, SaveFileAs){
    try
    {
    SplitPath, SaveFileAs, name, dir, ext, name_no_ext, drive
	if(isInvalidFileName(name))
	{
		name:=cleanFileName(name)
		SaveFileAs:=dir . "\" . name
	}
    ToolTip, Downloading file "%name%",0,0
	VarSetCapacity(WinHttpObj,10240000)
	VarSetCapacity(ADODBObj,10240000)
	Overwrite := False
	If (FileExist(SaveFileAs)) {
		If (Overwrite)
			FileDelete, %SaveFileAs%
		Else
		{
			Tooltip
			Return
		}
	}
	WinHttpObj := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	WinHttpObj.Open("GET", UrlToFile)
	WinHttpObj.Send()
	ADODBObj := ComObjCreate("ADODB.Stream")
	ADODBObj.Type := 1
	ADODBObj.Open()
	ADODBObj.Write(WinHttpObj.ResponseBody)
	ADODBObj.SaveToFile(SaveFileAs, Overwrite ? 2:1)
	ADODBObj.Close()
	WinHttpObj := ""
	VarSetCapacity(WinHttpObj,0)
	ADODBObj := ""
	VarSetCapacity(ADODBObj,0)
    ToolTip
    }
    catch
    {
        msgbox, Could not download files from Jabrutouch server, the application will now exit.
        ExitApp
    }
}

KillChildProcesses(ParentPidOrExe){
	static Processes, i
	ParentPID:=","
	If !(Processes)
		Processes:=ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
	i++
	for Process in Processes
		If (Process.Name=ParentPidOrExe || Process.ProcessID=ParentPidOrExe)
			ParentPID.=process.ProcessID ","
	for Process in Processes
		If InStr(ParentPID,"," Process.ParentProcessId ","){
			KillChildProcesses(process.ProcessID)
            WinClose, % "ahk_pid " process.ProcessID
            WinWaitClose, % "ahk_pid " process.ProcessID,,5
			Process,Close,% process.ProcessID 
		}
	i--
	If !i
		Processes=
}