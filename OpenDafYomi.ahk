#Persistent
#SingleInstance, Force
#WinActivateForce
#Include JSON.ahk
#Include Jxon.ahk
CoordMode, ToolTip,Screen

database=%A_ScriptDir%\cms3926896145652424982.csv
FileInstall cms3926896145652424982.csv,%database%,1

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
oWhr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
oWhr.SetTimeouts("600000", "600000", "600000", "600000")
oWhr.Open("GET", "https://api.jabrutouch.com/api/delivery/spaces/1/entries?language=es&contentType=lesson&number=" VideoID, false)
oWhr.SetRequestHeader("x-api-key", "a63p1n4ujf13hsd4mf8mgfkncp")
oWhr.SetRequestHeader("Content-Type", "application/json")
oWhr.Send(payload)
if(oWhr.Status==200)
    response:=oWhr.ResponseText

data := Jxon_Load(response)
Clipboard:=response

for i, d in data
{
    for j, f in d["fields"]
    {
        if(f["key"]="video")
        {
            for k, f0 in f["value"]["fields"]
            {
                if(f0["key"]="url")
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
                        DownloadFile(f0["value"], localVideoPath)
                    }
                    if(FileExist("C:\Program Files (x86)\VideoLAN\VLC\vlc.exe"))
                    {
                        if(downloaded)
                            Run, % """C:\Program Files (x86)\VideoLAN\VLC\vlc.exe"" --start-time 1 --no-start-paused --repeat dafyomi.mp4", %A_ScriptDir%,,VLCPID
                        else
                            Run, % """C:\Program Files (x86)\VideoLAN\VLC\vlc.exe"" --repeat dafyomi.mp4", %A_ScriptDir%,,VLCPID
                    }
                    else
                        Run, % localVideoPath
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
                    pdfUrl:=f0["value"]
                    DownloadFile(f0["value"], localPDFPath)
                    Run, %localPDFPath%,,,PDFPID
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
if(PDFPID)
{
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

OnExit, ExitApplication
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

#If (WinActive("ahk_pid " . PDFPID))

^Left::
ControlSend,%VLCVidCtrl%, {Left}, ahk_id %WorkerW%
return

^Right::
ControlSend,%VLCVidCtrl%, {Right}, ahk_id %WorkerW%
return

#If

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
        msgbox, Could not contact Jabrutouch server, the application will now exit.
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
			Process,Close,% process.ProcessID 
		}
	i--
	If !i
		Processes=
}