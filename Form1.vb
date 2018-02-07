Public Class Form1
    Dim ComSFT(16) As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim PathFtp, PathArh, NameFile, PathLog, login, password, server, port, hostkey As String
        'Const pr As Char = "%"
        'Const dcav As Char = Chr(34)

        PathFtp = "c:\testFtp"   '"c:\testFtp\test1.txt"
        NameFile = "test1.txt"
        PathArh = "c:\testFtp\arh\%sdat%.txt"
        PathLog = "C:\testFtp\WinSCP1.log"
        login = "prog"
        password = "3RBkJLzW"
        server = "spb.root.autocrm.ru"
        port = "14872"
        hostkey = "ssh-rsa 2048 30:31:44:d4:00:2c:f3:7d:70:09:83:58:68:c9:68:c8"

        Dim dd = New CreateBatFile
        ComSFT = dd.CreateSM(PathFtp, NameFile, PathArh, PathLog, login, password, server, port, hostkey)




        'ComSFT(0) = "@echo On"
        'ComSFT(1) = "Set " & dcav & "sdat=%date% %time~0,-3%" & dcav
        'ComSFT(2) = "Set " & dcav & "sdat=%date% %time~0,-3%" & dcav
        'ComSFT(3) = "Set " & dcav & "sdat=%sdat:=.%" & dcav
        'ComSFT(4) = "Set " & dcav & "sdat=%" & "sdat:=_%" & dcav
        'ComSFT(5) = "copy /-y " & PathFtp & " " & PathArh   'c:\testFtp\test1.txt c:\testFtp\arh\%sdat%.txt
        'ComSFT(6) = dcav & "C:\Program Files (x86)\WinSCP\WinSCP.com" & dcav & "^"
        'ComSFT(7) = "/log=" & dcav & "C:\Program Files (x86)\WinSCP\WinSCP1.log" & dcav & " /ini=nul ^"
        'ComSFT(8) = "/command ^"
        'ComSFT(9) = dcav & "open sftp://prog:3RBkJLzW@spb.root.autocrm.ru:14872/ -hostkey=" & dcav & dcav & "ssh-rsa 2048 30:31:44:d4:00:2c:f3:7d:70:09:83:58:68:c9:68:c8" & dcav & dcav & "^"
        'ComSFT(10) = dcav & "lcd c:\testftp" & dcav & "^"
        'ComSFT(11) = dcav & "put *.txt" & dcav & "^"
        'ComSFT(12) = dcav & "exit" & dcav
        'ComSFT(13) = ""
        'ComSFT(14) = "set WINSCP_RESULT=%ERRORLEVEL%"
        'ComSFT(15) = "If %WINSCP_RESULT% equ 0 ( echo Success ) Else ( echo Error )"
        'ComSFT(16) = "Exit /b %WINSCP_RESULT%"
        ''If %WINSCP_RESULT% equ 0 (
        ''  echo Success
        '') else (
        ''  echo Error
        '')

        ''Exit /b %WINSCP_RESULT%
        ''"
        ''        RichTextBox1.Text = $"Ghbdtn dfcz
        ''{TextBox1.Text }^ 
        ''{TextBox2.Text }^

        ''"
        ''" {0} {1}", TextBox1.Text,TextBox2.Text
        For Each i As String In ComSFT

            RichTextBox1.Text = RichTextBox1.Text & i & vbCrLf
        Next



    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim PathFtpBat As String
        PathFtpBat = "c:\testFtp\testbat2.bat"   '"c:\testFtp\test1.txt"
        IO.File.WriteAllLines(PathFtpBat, ComSFT)
        If IO.File.Exists(PathFtpBat) Then
            MsgBox("Ghjjjj jjjj")

        End If

    End Sub
End Class
Public Class CreateBatFile
    Public Function CreateSM(ByVal _PathFtp As String,
                             ByVal _NameFile As String,
                             ByVal _PathArh As String,
                             ByVal _PathLog As String,
                             ByVal _login As String,
                             ByVal _password As String,
                             ByVal _server As String,
                             ByVal _port As String,
                             ByVal _hostkey As String
                             ) As String()

        Dim ComSFT(16) As String
        'Dim PathFtp, PathArh As String
        'Const pr As Char = "%"
        Const dcav As Char = Chr(34)

        'PathFtp = "c:\testFtp\test1.txt"
        'PathArh = "c:\testFtp\arh\%sdat%.txt"

        ComSFT(0) = "@echo Off"
        ComSFT(1) = "set " & dcav & "sdat=%date% %time~0,-3%" & dcav
        ComSFT(2) = "set " & dcav & "sdat=%date% %time~0,-3%" & dcav
        ComSFT(3) = "set " & dcav & "sdat=%sdat:=.%" & dcav
        ComSFT(4) = "set " & dcav & "sdat=%" & "sdat:=_%" & dcav
        ComSFT(5) = "copy /-y " & _PathFtp & " " & _PathArh   'c:\testFtp\test1.txt c:\testFtp\arh\%sdat%.txt
        ComSFT(6) = dcav & "C:\Program Files (x86)\WinSCP\WinSCP.com" & dcav & "^"
        ComSFT(7) = "/log=" & dcav & _PathLog & dcav & " /ini=nul ^" '"C:\Program Files (x86)\WinSCP\WinSCP1.log
        ComSFT(8) = "/command ^"
        'ComSFT(9) = dcav & "open sftp://prog:3RBkJLzW@spb.root.autocrm.ru:14872/ -hostkey=" & dcav & dcav & "ssh-rsa 2048 30:31:44:d4:00:2c:f3:7d:70:09:83:58:68:c9:68:c8" & dcav & dcav & & dcav "^"
        ComSFT(9) = dcav & "open sftp://" & _login & ":" & _password & "@" & _server & ":" & _port & "/ -hostkey=" & dcav & dcav & _hostkey & dcav & dcav & dcav & "^"
        ComSFT(10) = dcav & "lcd " & _PathFtp & dcav & "^"
        ComSFT(11) = dcav & "put " & _NameFile & dcav & "^"
        ComSFT(12) = dcav & "exit" & dcav
        ComSFT(13) = ""
        ComSFT(14) = "set WINSCP_RESULT=%ERRORLEVEL%"
        ComSFT(15) = "If %WINSCP_RESULT% equ 0 ( echo Success ) Else ( echo Error )"
        ComSFT(16) = "Exit /b %WINSCP_RESULT%"

        Return ComSFT

    End Function

End Class

