Set WshShell = CreateObject("WScript.Shell")
dim list
Set list = CreateObject("System.Collections.ArrayList")
Set WshShellExec = WshShell.Exec("adb devices")
Do
    strFromProc = WshShellExec.StdOut.ReadLine()
    list.Add replace(strFromProc, "device", "")
    Loop While Not WshShellExec.Stdout.atEndOfStream


list.RemoveAt(0)
For Each Arg In list
        WshShell.Run "cmd /c scrcpy.exe -s" & Arg, 0, false
Next
