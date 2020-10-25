#Persistent
#SingleInstance, Force
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
                            Run, % "vlc.exe --start-time=0.0 --repeat " localVideoPath, C:\Program Files (x86)\VideoLAN\vlc
                        else
                            Run, % "vlc.exe --repeat " localVideoPath, C:\Program Files (x86)\VideoLAN\vlc
                    }
                    else
                        Run, % localVideoPath
                }
            }
        }
    }
}
if(FileExist("C:\Users\langh\Utilities\Autohotkey\AttachVLCToDesktop\AttachVLCToDesktop.exe"))
{
    Sleep, 3000
    Run, AttachVLCToDesktop.exe, C:\Users\langh\Utilities\Autohotkey\AttachVLCToDesktop
}

OnExit, ExitApplication
return

ExitApplication:
KillChildProcesses("OpenDafYomi.exe")
ExitApp
return

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