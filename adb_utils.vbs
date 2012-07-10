'キーコード
Public Const VK_HOME = 3
Public Const VK_BACK = 4
Public Const VK_MENU =82 
Public Const VK_SEARCH = 84

Public Const VK_UP = 19
Public Const VK_DOWN = 20
Public Const VK_LEFT = 21
Public Const VK_RIGHT = 22
Public Const VK_CENTER = 23

Public Const VK_CAMERA = 27
Public Const VK_ENTER = 66
Public Const VK_DELETE = 67

Public Const VK_VOL_UP = 24
Public Const VK_VOL_DOWN = 25
Public Const VK_VOL_MUTE = 91

Dim objWShell
Set objWShell = CreateObject("WScript.Shell")

'コマンド実行
Private Function run(cmd)
    objWShell.Run cmd, 0, True
End Function

'コマンド実行
Function exec(cmd)
    Set objExec = objWShell.Exec(cmd)
    strLine = ""
    Do Until objExec.StdOut.AtEndOfStream
        strLine = strLine & objExec.StdOut.ReadLine & vbCrLf
    Loop
    exec = strLine
End Function

'adb install を実行する
Function install(filename)
    run("adb install """ & filename & """")
End Function

'adb uninstall を実行する
Function uninstall(packagename)
    run("adb uninstall " & packagename)
End Function

' adb shell input textを実行する
Function sendText(text)
    run("adb shell input text " & text)
    sendKey(VK_ENTER)
End Function

' adb shell input keyeventを実行する
Function sendKey(keycode)
    run("adb shell input keyevent " & keycode)
End Function

' Android端末の列挙
Function listDevices()
    Dim devices()
    Dim devs
    result = exec("adb devices")
    devs = Split(result, vbCrLf, -1)
    cnt = UBound(devs)
    If cnt > 2 Then
        Redim devices(cnt - 2)
        For i = 1 To cnt - 2
            devAndState = Split(devs(i), vbTab, -1)
            If UBound(devAndState) > 0 Then
                devices(i-1) = devAndState(0)
            End If
        Next
    End If
    listDevices = devices
End Function

