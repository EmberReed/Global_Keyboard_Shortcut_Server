Imports System.Net.Sockets
Imports System.Text
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.Net

Public Class Form1
    Private RunServer As Boolean = False
    Private CTS As CancellationTokenSource
    Private ConnectionAlive As Boolean = False
    Private SendString As String = ""
    Public Const MOD_ALT As Integer = &H1 'Alt key
    Public Const MOD_WIN As Integer = &H8
    Public Const WM_SetRedraw = &HB
    Public GroupSetting As Integer = 0
    Public ActiveColor As Color = Color.FromArgb(255, 255, 128, 0)
    Public PassiveColor As Color = Color.FromArgb(255, 65, 65, 65)
    Public Const WM_HOTKEY As Integer = &H312

    Sub msg(ByVal mesg As String)
        mesg.Trim()
        Console.WriteLine(" >> " + mesg)
    End Sub

    Private Sub ChangeGroupSetting(NewSelection As Integer)
        If GroupSetting <> NewSelection Then
            Dim MyButton As Button = Controls("Group" & GroupSetting)
            MyButton.BackColor = PassiveColor
            GroupSetting = NewSelection
            MyButton = Controls("Group" & GroupSetting)
            MyButton.BackColor = ActiveColor
        End If
    End Sub

    Private Sub SendNumPress(NumPressed As Integer)
        If SendString = "" And ConnectionAlive = True Then
            SendString = GroupSetting & "_" & NumPressed
        End If
    End Sub

    'Private Sub Button1_Click(sender As Object, e As EventArgs)
    '    If RunServer = False Then
    '        RunServer = True
    '        Button1.Text = "Stop"
    '        Dim RunMyServer As Task = RuntcpServer()
    '    Else
    '        RunServer = False
    '    End If
    'End Sub

    Declare Function SetDrawing Lib "user32" Alias "SendMessageA" (ByVal hwndLock As IntPtr, ByVal Msg As Integer, ByVal wParam As Boolean, ByVal lParam As IntPtr) As Integer

    Private Sub ConsoleText(NewText As String)
        'Updates the console text-box and handles the drawing state to avoid flicker (May or may not work as intended)
        If Visible = True Then SetDrawing(TextBox1.Handle, WM_SetRedraw, False, 0)
        TextBox1.Text = NewText
        If Visible = True Then
            SetDrawing(TextBox1.Handle, WM_SetRedraw, True, 0)
            TextBox1.Refresh()
            Application.DoEvents()
        End If
    End Sub



    Private Sub ConsoleAppend(AppendText As String)
        'Appends the console text-box and handles the drawing state to avoid flicker (May or may not work as intended)
        If Visible = True Then SetDrawing(TextBox1.Handle, WM_SetRedraw, False, 0)
        TextBox1.AppendText(AppendText)
        TextBox1.SelectionStart = Len(TextBox1.Text)
        TextBox1.ScrollToCaret()
        If Visible = True Then
            SetDrawing(TextBox1.Handle, WM_SetRedraw, True, 0)
            TextBox1.Refresh()
            Application.DoEvents()
        End If
    End Sub

    Private Sub ConsoleBackspace()
        'Removes the last character from the console text-box and handles the drawing state to avoid flicker
        If Visible = True Then SetDrawing(TextBox1.Handle, WM_SetRedraw, False, 0)
        TextBox1.SelectionStart = Len(TextBox1.Text) - 1
        TextBox1.SelectionLength = 1
        TextBox1.SelectedText = ""
        If Visible = True Then
            SetDrawing(TextBox1.Handle, WM_SetRedraw, True, 0)
            TextBox1.Refresh()
            Application.DoEvents()
        End If
    End Sub

    Private Sub ConsoleClearLine()
        'Removes the last line from the console text-box and handles the drawing state to avoid flicker
        If Visible = True Then SetDrawing(TextBox1.Handle, WM_SetRedraw, False, 0)

        Dim SelectionStart As Long = InStrRev(TextBox1.Text, Chr(10))
        If SelectionStart <> 0 Then
            Dim SelectionLength As Long = Len(TextBox1.Text)
            If SelectionLength <> SelectionStart Then
                SelectionLength = SelectionLength - SelectionStart
                TextBox1.SelectionStart = SelectionStart
                TextBox1.SelectionLength = SelectionLength
                TextBox1.SelectedText = ""
                Application.DoEvents()
            End If
        End If

        If Visible = True Then
            SetDrawing(TextBox1.Handle, WM_SetRedraw, True, 0)
            TextBox1.Refresh()
            Application.DoEvents()
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CTS = New CancellationTokenSource
        'kbHook = New KeyboardHook
        RegisterHotKey(Me.Handle, 100, MOD_ALT, &H60)
        RegisterHotKey(Me.Handle, 101, MOD_ALT, &H61)
        RegisterHotKey(Me.Handle, 102, MOD_ALT, &H62)
        RegisterHotKey(Me.Handle, 103, MOD_ALT, &H63)
        RegisterHotKey(Me.Handle, 104, MOD_ALT, &H64)
        RegisterHotKey(Me.Handle, 105, MOD_ALT, &H65)
        RegisterHotKey(Me.Handle, 106, MOD_ALT, &H66)
        RegisterHotKey(Me.Handle, 107, MOD_ALT, &H67)
        RegisterHotKey(Me.Handle, 108, MOD_ALT, &H68)
        RegisterHotKey(Me.Handle, 109, MOD_ALT, &H69)

        RegisterHotKey(Me.Handle, 200, MOD_WIN, &H60)
        RegisterHotKey(Me.Handle, 201, MOD_WIN, &H61)
        RegisterHotKey(Me.Handle, 202, MOD_WIN, &H62)
        RegisterHotKey(Me.Handle, 203, MOD_WIN, &H63)
        RegisterHotKey(Me.Handle, 204, MOD_WIN, &H64)
        RegisterHotKey(Me.Handle, 205, MOD_WIN, &H65)
        RegisterHotKey(Me.Handle, 206, MOD_WIN, &H66)
        RegisterHotKey(Me.Handle, 207, MOD_WIN, &H67)
        RegisterHotKey(Me.Handle, 208, MOD_WIN, &H68)
        RegisterHotKey(Me.Handle, 209, MOD_WIN, &H69)

        'RegisterHotKey(Me.Handle, 300, 0, Keys.T)

    End Sub

    <DllImport("User32.dll")> Public Shared Function RegisterHotKey(ByVal hwnd As IntPtr,
                                                                    ByVal id As Integer, ByVal fsModifiers As Integer,
                                                                    ByVal vk As Integer) As Integer
    End Function

    <DllImport("User32.dll")> Public Shared Function UnregisterHotKey(ByVal hwnd As IntPtr,
                                                                      ByVal id As Integer) As Integer
    End Function

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = WM_HOTKEY Then
            Dim id As IntPtr = m.WParam
            Select Case (id.ToString)
                Case "100"
                    SendNumPress(0)
                Case "101"
                    SendNumPress(1)
                Case "102"
                    SendNumPress(2)
                Case "103"
                    SendNumPress(3)
                Case "104"
                    SendNumPress(4)
                Case "105"
                    SendNumPress(5)
                Case "106"
                    SendNumPress(6)
                Case "107"
                    SendNumPress(7)
                Case "108"
                    SendNumPress(8)
                Case "109"
                    SendNumPress(9)
                Case "200"
                    ChangeGroupSetting(0)
                Case "201"
                    ChangeGroupSetting(1)
                Case "202"
                    ChangeGroupSetting(2)
                Case "203"
                    ChangeGroupSetting(3)
                Case "204"
                    ChangeGroupSetting(4)
                Case "205"
                    ChangeGroupSetting(5)
                Case "206"
                    ChangeGroupSetting(6)
                Case "207"
                    ChangeGroupSetting(7)
                Case "208"
                    ChangeGroupSetting(8)
                Case "209"
                    ChangeGroupSetting(9)
                    'Case "300"
                    'TextBox1.Text = TextBox1.Text & vbCrLf & "Rock And Stone" & vbCrLf
                    'MessageBox.Show("W9")
            End Select
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As System.Object,
                        ByVal e As System.Windows.Forms.FormClosingEventArgs) _
                        Handles MyBase.FormClosing
        UnregisterHotKey(Me.Handle, 100)
        UnregisterHotKey(Me.Handle, 101)
        UnregisterHotKey(Me.Handle, 102)
        UnregisterHotKey(Me.Handle, 103)
        UnregisterHotKey(Me.Handle, 104)
        UnregisterHotKey(Me.Handle, 105)
        UnregisterHotKey(Me.Handle, 106)
        UnregisterHotKey(Me.Handle, 107)
        UnregisterHotKey(Me.Handle, 108)
        UnregisterHotKey(Me.Handle, 109)
        UnregisterHotKey(Me.Handle, 200)
        UnregisterHotKey(Me.Handle, 201)
        UnregisterHotKey(Me.Handle, 202)
        UnregisterHotKey(Me.Handle, 203)
        UnregisterHotKey(Me.Handle, 204)
        UnregisterHotKey(Me.Handle, 205)
        UnregisterHotKey(Me.Handle, 206)
        UnregisterHotKey(Me.Handle, 207)
        UnregisterHotKey(Me.Handle, 208)
        UnregisterHotKey(Me.Handle, 209)
        If RunServer = True Then
            RunServer = False
            e.Cancel = True
        End If

        'UnregisterHotKey(Me.Handle, 300)
    End Sub


    Public Async Function RuntcpServer() As Task
        Dim PortNumber As Integer = 8888
        Dim serverSocket As New TcpListener(Net.IPAddress.Any, PortNumber)
        Dim hostname As String = Dns.GetHostName()
        Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(hostname)
        Dim STRipaddress As String = ""


        For Each ipheal As System.Net.IPAddress In iphe.AddressList
            If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                STRipaddress = ipheal.ToString()
            End If
        Next

        Dim clientSocket As TcpClient
        Dim MynetworkStream As NetworkStream
        Dim bytesFrom As Byte
        Dim dataFromClient As String
        Dim SendBytes() As Byte

        serverSocket.Start()
        ConsoleAppend("-Server Started: " & STRipaddress & ":" & PortNumber & vbCrLf)
        Application.DoEvents()
ReconnectClient:
        Do Until serverSocket.Pending = True
            If RunServer = False Then GoTo StopServer
            Await Task.Delay(20)
        Loop
        clientSocket = Await serverSocket.AcceptTcpClientAsync()
        ConsoleAppend("-Connected to: " & clientSocket.Client.RemoteEndPoint.ToString & vbCrLf)
        Application.DoEvents()
        Try
            MynetworkStream = clientSocket.GetStream()
            ConnectionAlive = True
        Catch ex As Exception
            MsgBox("-Connection fail: " & ex.ToString)
            GoTo didntwork
        End Try

        While RunServer And MynetworkStream.CanRead And serverSocket.Pending = False 'And serverSocket.Pending = False And MynetworkStream.CanRead
            Try
                If MynetworkStream.DataAvailable = True Then
                    bytesFrom = MynetworkStream.ReadByte
                    dataFromClient = Chr(bytesFrom)
                    dataFromClient = Replace(dataFromClient, Chr(13), vbCrLf)
                    ConsoleAppend(dataFromClient)
                    Application.DoEvents()
                Else
                    If SendString <> "" Then
                        SendBytes = Encoding.ASCII.GetBytes(SendString & vbCrLf)
                        MynetworkStream.Write(SendBytes, 0, SendBytes.Length)
                        'MsgBox(SendBytes.Length)
                        ConsoleAppend("-Sent: " & SendString & vbCrLf)
                        SendString = ""
                        Application.DoEvents()
                    End If
                    Await Task.Delay(20)
                End If
            Catch ex As Exception
                MsgBox("socket fail: " & ex.ToString)
                GoTo didntwork
            End Try
        End While

        If RunServer = True Then
            ConnectionAlive = False
            clientSocket.Close()
            ConsoleAppend("-Disconnected" & vbCrLf)
            Application.DoEvents()
            GoTo ReconnectClient
        End If

didntwork:
        ConnectionAlive = False
        clientSocket.Close()
        ConsoleAppend("-Disconnected" & vbCrLf)
        Application.DoEvents()
StopServer:
        serverSocket.Stop()
        ConsoleAppend("-Server Stopped")
        Await Task.Delay(100)
        Close()
        'Button1.Text = "Start"
    End Function

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        RunServer = True
        Dim RunMyServer As Task = RuntcpServer()

    End Sub



    Private Sub Group0_Click(sender As Object, e As EventArgs) Handles Group0.Click
        ChangeGroupSetting(0)
    End Sub

    Private Sub Group1_Click(sender As Object, e As EventArgs) Handles Group1.Click
        ChangeGroupSetting(1)
    End Sub

    Private Sub Group2_Click(sender As Object, e As EventArgs) Handles Group2.Click
        ChangeGroupSetting(2)
    End Sub

    Private Sub Group3_Click(sender As Object, e As EventArgs) Handles Group3.Click
        ChangeGroupSetting(3)
    End Sub

    Private Sub Group4_Click(sender As Object, e As EventArgs) Handles Group4.Click
        ChangeGroupSetting(4)
    End Sub

    Private Sub Group5_Click(sender As Object, e As EventArgs) Handles Group5.Click
        ChangeGroupSetting(5)
    End Sub

    Private Sub Group6_Click(sender As Object, e As EventArgs) Handles Group6.Click
        ChangeGroupSetting(6)
    End Sub

    Private Sub Group7_Click(sender As Object, e As EventArgs) Handles Group7.Click
        ChangeGroupSetting(7)
    End Sub

    Private Sub Group8_Click(sender As Object, e As EventArgs) Handles Group8.Click
        ChangeGroupSetting(8)
    End Sub

    Private Sub Group9_Click(sender As Object, e As EventArgs) Handles Group9.Click
        ChangeGroupSetting(9)
    End Sub

    Private Sub numpad0_Click(sender As Object, e As EventArgs) Handles numpad0.Click
        SendNumPress(0)
    End Sub
    Private Sub numpad1_Click(sender As Object, e As EventArgs) Handles numpad1.Click
        SendNumPress(1)
    End Sub
    Private Sub numpad2_Click(sender As Object, e As EventArgs) Handles numpad2.Click
        SendNumPress(2)
    End Sub
    Private Sub numpad3_Click(sender As Object, e As EventArgs) Handles numpad3.Click
        SendNumPress(3)
    End Sub
    Private Sub numpad4_Click(sender As Object, e As EventArgs) Handles numpad4.Click
        SendNumPress(4)
    End Sub
    Private Sub numpad5_Click(sender As Object, e As EventArgs) Handles numpad5.Click
        SendNumPress(5)
    End Sub
    Private Sub numpad6_Click(sender As Object, e As EventArgs) Handles numpad6.Click
        SendNumPress(6)
    End Sub
    Private Sub numpad7_Click(sender As Object, e As EventArgs) Handles numpad7.Click
        SendNumPress(7)
    End Sub
    Private Sub numpad8_Click(sender As Object, e As EventArgs) Handles numpad8.Click
        SendNumPress(8)
    End Sub
    Private Sub numpad9_Click(sender As Object, e As EventArgs) Handles numpad9.Click
        SendNumPress(9)
    End Sub


End Class




