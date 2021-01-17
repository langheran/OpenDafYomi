taskkill /f /im "OpenDafYomi.exe"
taskkill /f /im "vlc.exe"
taskkill /f /im "SumatraPDF.exe"
"C:\Program Files\AutoHotkey\Compiler\Ahk2Exe.exe" /cp 65001 /icon OpenDafYomi.ico /in OpenDafYomi.ahk /out OpenDafYomi.exe
echo f | xcopy "OpenDafYomi.exe" "C:\Users\langh\Utilities\Autohotkey\Helper\OpenDafYomi.exe" /H/Y
pushd "C:\Users\langh\Utilities\Autohotkey\Helper"
OpenDafYomi.exe
popd