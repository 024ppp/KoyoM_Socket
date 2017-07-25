' VB.NETのソケット通信サンプル
Imports System.Threading
Imports System.Net
Imports SZKDLL_Orc.DBACCESS

Public Class Form1
    ' ソケット・リスナー
    Private myListener As Sockets.TcpListener
    ' クライアント送受信
    Private myClient As ClientTcpIp() = New ClientTcpIp(3) {}

    ' フォームロード時のソケット接続処理
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles MyBase.Load

        ' IPアドレス＆ポート番号設定
        Dim myPort As Integer = CInt(clsConnect.GetConnectionString("myPort"))
        Dim myEndPoint As New IPEndPoint(IPAddress.Any, myPort)

        'Dim ipString = clsConnect.GetConnectionString("ipAdd")
        'Dim ipAdd As System.Net.IPAddress = System.Net.IPAddress.Parse(ipString)
        'Dim myEndPoint As New IPEndPoint(ipAdd, myPort)

        Try
            ' リスナー開始
            myListener = New Sockets.TcpListener(myEndPoint)
            myListener.Start()

            ' クライアント接続待ち開始
            Dim myServerThread As New Thread(New ThreadStart(AddressOf ServerThread))
            myServerThread.Start()

        Catch ex As Exception
            TextBox1.Text = TextBox1.Text & vbCrLf & ex.Message
        End Try
    End Sub

    ' フォームクローズ時のソケット切断処理
    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) _
        Handles Me.FormClosed

        ' リスナー終了
        myListener.Stop()
        ' クライアント切断
        For i As Integer = 0 To myClient.GetLength(0) - 1
            If myClient(i) Is Nothing = False AndAlso _
                   myClient(i).objSck.Connected = True Then
                ' ソケットクローズ
                myClient(i).objStm.Close()
                myClient(i).objSck.Close()
            End If
        Next
    End Sub

    ' クライアント接続待ちスレッド
    Private Sub ServerThread()
        Try
            Dim intNo As Integer
            While True
                ' ソケット接続待ち
                Dim myTcpClient As Sockets.TcpClient = myListener.AcceptTcpClient()
                ' クライアントから接続有り
                For intNo = 0 To myClient.GetLength(0) - 1
                    If myClient(intNo) Is Nothing Then
                        Exit For
                    ElseIf myClient(intNo).objSck.Connected = False Then
                        Exit For
                    End If
                Next
                If intNo < myClient.GetLength(0) Then
                    ' クライアント送受信オブジェクト生成
                    myClient(intNo) = New ClientTcpIp(Me)
                    myClient(intNo).intNo = intNo + 1
                    myClient(intNo).objSck = myTcpClient
                    myClient(intNo).objStm = myTcpClient.GetStream()
                    ' クライアントとの送受信開始
                    Dim myClientThread As New Thread( _
                        New ThreadStart(AddressOf myClient(intNo).ReadWrite))
                    myClientThread.Start()
                Else
                    ' 接続拒否
                    myTcpClient.Close()
                End If
            End While
        Catch ex As Exception
        End Try
    End Sub

    'テキストボックスの値設定（Delegateするメソッド）
    Friend Sub SetTextBox1(ByVal Value As String)
        TextBox1.Text = Value
    End Sub

    'Delegate Sub SetTextDelegate(ByVal msg As String)

    'Sub SetText(ByVal msg As String)
    '    If InvokeRequired Then
    '        ' 別スレッドから呼び出された場合
    '        Dim callback As New SetTextDelegate(AddressOf SetText)
    '        Invoke(callback, msg)
    '        'Invoke(New SetTextDelegate(AddressOf SetText), New Object() {msg})
    '        Return
    '    End If

    '    TextBox1.AppendText(msg & vbCrLf)

    '    Me.TextBox2.Focus()
    '    'Me.Refresh()

    'End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TextBox1.Text = ""
    End Sub
End Class

