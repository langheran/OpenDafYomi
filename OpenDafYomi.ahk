#Persistent
#SingleInstance, Force
#Include JSON.ahk
#Include Jxon.ahk

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
                    if(FileExist("C:\Program Files (x86)\VideoLAN\VLC\vlc.exe"))
                        Run, % "vlc.exe --repeat " f0["value"], C:\Program Files (x86)\VideoLAN\vlc
                    else
                        Run, % f0["value"]
                }
            }
        }
    }
}
if(FileExist("C:\Users\langh\Utilities\Autohotkey\AttachVLCToDesktop\AttachVLCToDesktop.exe"))
{
    Sleep, 5000
    Run, AttachVLCToDesktop.exe, C:\Users\langh\Utilities\Autohotkey\AttachVLCToDesktop
}

OnExit, ExitApplication
return

ExitApplication:
KillChildProcesses("OpenDafYomi.exe")
ExitApp
return

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