' VB.NETのソケット通信サンプル
Imports System
Imports System.Threading
Imports System.Net
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports SZKDLL_Orc
Imports SZKDLL_Orc.Class
Imports SZKDLL_Orc.DBACCESS

' クライアント送受信クラス
Public Class ClientTcpIp 
    Public intNo As Integer
    Public objSck As Sockets.TcpClient
    Public objStm As Sockets.NetworkStream

    Delegate Sub SetTextBox1Delegate(ByVal Value As String)     'コントロールを扱うためのデリゲート宣言
    Private TextBox1Delegate As SetTextBox1Delegate             'デリゲート宣言をデータ型とした変数を作成
    Private _Form1 As Form1                                     'フォームの参照を保持する
    Private dbAction As New DBAction

    Private ecUni As Encoding = Encoding.GetEncoding("UTF-8")
    Private ecSjis As Encoding = Encoding.GetEncoding("shift-jis")

    ' クライアント送受信スレッド
    Public Sub ReadWrite()
        Try
            While True
                ' ソケット受信
                Dim rdat As Byte() = New Byte(1023) {}
                Dim ldat As Int32 = objStm.Read(rdat, 0, rdat.GetLength(0))

                Dim sCmdText As String                  '受信したコマンド文字
                Dim sExcludeCmdText As String           'コマンド文字を除いた文字
                Dim sSendData As String = "Nothing"     'クライアントへの送信用文字
                Dim sErr As String = ProcessCommand.pc.ERR.ToString

                If ldat > 0 Then
                    ' クライアントからの受信データ有り
                    Dim sdat As Byte() = New Byte(ldat - 1) {}
                    Array.Copy(rdat, sdat, ldat)

                    Dim msg As String = System.Text.Encoding.GetEncoding("UTF-8").GetString(sdat)
                    '改行文字削除
                    msg = msg.Replace(Environment.NewLine, "")

                    'コマンド文字と、それ以降とを分離
                    Dim cmdLen As Integer = ProcessCommand.pc.COMMAND_LENGTH
                    sCmdText = Left(msg, cmdLen)
                    sExcludeCmdText = Mid(msg, cmdLen + 1)

                    'クライアントから受信した値をもとに、処理を選択する
                    Select Case sCmdText
                        Case ProcessCommand.pc.SAG.ToString
                            sSendData = sCmdText & dbAction.getSagyoName()

                        Case ProcessCommand.pc.KIK.ToString
                            If dbAction.checkKikai(sExcludeCmdText) Then
                                'Vコンが存在している場合、受信データを送り返す
                                sSendData = sCmdText & sExcludeCmdText
                            Else
                                sSendData = sErr & "Vコンが存在しません。"
                            End If

                        Case ProcessCommand.pc.UPD.ToString
                            If dbAction.Update(sExcludeCmdText) Then
                                sSendData = sCmdText & "更新に成功しました。"
                            Else
                                sSendData = sCmdText & "更新失敗！"
                            End If

                        Case Else
                            sSendData = "case else"
                    End Select

                    'フォームに書き出し
                    Call WriteTextBox(msg & "-" & sSendData)

                    sdat = System.Text.Encoding.GetEncoding("UTF-8").GetBytes(sSendData & vbCrLf)
                    ' ソケット送信
                    objStm.Write(sdat, 0, sdat.GetLength(0))

                Else
                    ' ソケット切断有り
                    ' ソケットクローズ
                    objStm.Close()
                    objSck.Close()
                    Return
                End If
            End While
        Catch ex As Exception
        End Try
    End Sub

    'コンストラクタ
    Public Sub New(ByVal frm As Form)
        _Form1 = CType(frm, Form1)
        TextBox1Delegate = New SetTextBox1Delegate(AddressOf _Form1.SetTextBox1)
    End Sub

    'テキストボックスに値を設定する
    Public Sub WriteTextBox(ByVal msg As String)
        Dim sTextBox1 As String = _Form1.TextBox1.Text

        If sTextBox1 <> "" Then
            sTextBox1 = sTextBox1 & vbCrLf
        End If

        '書き込み
        _Form1.Invoke(TextBox1Delegate, New Object() {sTextBox1 & msg})
    End Sub

End Class