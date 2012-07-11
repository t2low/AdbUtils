Option Explicit

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

Class AdbUtli

    Private objWShell
    Private device

    'コンストラクタ
    Private Sub Class_Initialize
        Set objWShell = CreateObject("WScript.Shell")
        device = ""
    End Sub

    'デストラクタ
    Private Sub Class_Terminate
        objWShell = Null
    End Sub

    'コマンド実行
    Private Sub run(cmd)
        objWShell.Run cmd, 0, True
    End Sub 

    'コマンド実行
    Private Function exec(cmd)
        Dim objExec
        Dim strLine
        Set objExec = objWShell.Exec(cmd)
        strLine = ""
        Do Until objExec.StdOut.AtEndOfStream
            strLine = strLine & objExec.StdOut.ReadLine & vbCrLf
        Loop
        exec = strLine
    End Function

    'adb コマンドの作成
    Private Function createCmd(cmd)
        If Len(device) = 0 Then
            createCmd = "adb " & cmd & " "
        Else
            createCmd = "adb -s " & device & " " & cmd & " "
        End If
    End Function

    'adb install を実行する
    Public Function install(filename)
        run(createCmd("install") & """" & filename & """")
    End Function

    'adb uninstall を実行する
    Public Function uninstall(packagename)
        run(createCmd("uninstall") & packagename)
    End Function

    ' adb shell input textを実行する
    Public Function sendText(text)
        run(createCmd("shell input text") & text)
        sendKey(VK_ENTER)
    End Function

    ' adb shell input keyeventを実行する
    Public Function sendKey(keycode)
        run(createCmd("shell input keyevent") & keycode)
    End Function

    ' Android端末の列挙
    Public Function listDevices()
        Dim result
        Dim devices()
        Dim devs
        Dim devAndState
        Dim cnt, i

        Redim devices(0)
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

    ' パッケージの列挙
    Public Function listPackages()
        Dim result
        Dim packages()
        Dim pkgs
        Dim pkg
        Dim cnt, i

        Redim packages(0)
        result = exec(createCmd("shell pm list package"))
        pkgs = Split(result, vbCrLf, -1)
        cnt = UBound(pkgs)
        If cnt > 1 Then
            Redim packages(cnt - 1)
            For i = 0 To cnt - 1
                pkg = Split(pkgs(i), ":", -1)
                If UBound(pkg) >= 1 Then
                    packages(i) = pkg(1)
                End If
            Next
        End If
        listPackages = packages
    End Function

End Class
