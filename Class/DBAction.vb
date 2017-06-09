Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports SZKDLL_Orc
Imports SZKDLL_Orc.Class
Imports SZKDLL_Orc.DBACCESS

'TODO クラス名変更 DBに対する要求まとめクラスらしく
Public Class DBAction
    'クラス名
    Private Const C_CLASSNAME As String = "DBAction.vb"
    Private m_DT As DataTable = Nothing
    Private m_ErrMsg As String = ""
    Dim m_clsCom As New clsCom

    'データテーブルの取得
    Public Property rtnDT() As DataTable
        Get
            Return m_DT
        End Get
        Set(ByVal Value As DataTable)
        End Set
    End Property

    'エラーメッセージ
    Public Property ErrMsg() As String
        Get
            Return m_ErrMsg
        End Get
        Set(ByVal Value As String)
        End Set
    End Property

    '作業者取得
    Public Function getSagyoName() As String
        Dim strSQL As String = ""
        Dim sSagyoName As String = ""
        Dim ds As New clsDsCtrl

        Try
            strSQL = " SELECT * "
            strSQL &= " FROM MS02 "
            strSQL &= " WHERE DELKBN = 0 "
            strSQL &= " ORDER BY TNTCOD "

            Call ds.Connect()
            Call ds.ExecuteSQL(strSQL, "")
            'm_DT = ds.DataTable

            For Each dtRow As DataRow In ds.DataTable.Rows
                If Not sSagyoName.Equals("") Then
                    sSagyoName &= ","
                End If
                sSagyoName &= dtRow.Item("TNTNAM").ToString
            Next

            Return sSagyoName
        Catch ex As SqlException
            m_ErrMsg = ex.Message
            Return ""
        Finally
            Call ds.Dispose()
        End Try
    End Function

    'Vコンが存在するかチェックして結果を返す
    Public Function checkKikai(ByVal strKikai As String) As Boolean
        Dim strSQL As String = ""
        Dim ds As New clsDsCtrl
        Dim result As Boolean = False

        Try
            strSQL = " SELECT * "
            strSQL &= " FROM MM52K@MTRSLINK "
            strSQL &= " WHERE 1 = 1 "
            If strKikai <> "" Then
                strSQL &= " AND KIKCOD = '" & strKikai & "' "
            End If

            Call ds.Connect()
            Call ds.ExecuteSQL(strSQL, "")

            If ds.DataTable.Rows.Count > 0 Then
                result = True
            End If

            'For Each dtRow As DataRow In ds.DataTable.Rows
            '    sVkonNo = dtRow.Item("VKONNO").ToString
            'Next

            Return result
        Catch ex As SqlException
            m_ErrMsg = ex.Message
            Return result
        Finally
            Call ds.Dispose()
        End Try
    End Function

    'MV01更新
    Public Function Update(ByVal strUpdText As String) As Boolean
        Dim strSQL As String = ""
        Dim strArrUpdText As String() = Nothing
        Dim sVal As String = ""
        Dim sWhere As String = ""

        Dim ds As New clsDsCtrl

        Try
            strArrUpdText = strUpdText.Split(",")
            sVal = createSqlText(strArrUpdText, 0)
            sWhere = createSqlText(strArrUpdText, 1)

            If sVal.Equals("") Or sWhere.Equals("") Then
                Return False
            End If

            'SQL文作成
            strSQL &= " UPDATE MV01 SET "
            strSQL &= " UPDYMD = '" & Now.ToString("yyyyMMdd") & "' "
            strSQL &= " ,UPDHMS = '" & Now.ToString("HHmmss") & "' "
            strSQL &= sVal
            strSQL &= " WHERE "
            strSQL &= sWhere

            Call ds.Connect()
            Call ds.ExecuteSQL(strSQL, "")
            Return True
        Catch ex As Exception
            m_ErrMsg = ex.Message
            Return False
        End Try
    End Function

    'iCmd = 0 : SET句
    'iCmd = 1 : WHERE句
    Private Function createSqlText(ByVal arr As String(), ByVal iCmd As Integer) As String
        Dim txt As String = ""
        Dim arrCmdText As New ArrayList                    '受信したコマンド文字
        Dim arrExcludeCmdText As New ArrayList             'コマンド文字を除いた文字
        Dim cmdLen As Integer = ProcessCommand.pc.COMMAND_LENGTH

        Try
            '引数を分解
            For i As Integer = 0 To arr.Length - 1
                arrCmdText.Add(Left(arr(i), cmdLen))
                arrExcludeCmdText.Add(Mid(arr(i), cmdLen + 1))
            Next

            '分解した引数を使い、SQL文を作成
            For i As Integer = 0 To arrCmdText.Count - 1
                '該当するコマンド文字を持つ部分を抽出
                Select Case iCmd
                    Case 0
                        If arrCmdText(i).Equals(ProcessCommand.pc.AM1.ToString) Then
                            'SET句は無条件でカンマ付け
                            'addDelimiter(txt, iCmd)
                            txt &= " , "
                            txt &= arrCmdText(i) & " = " & arrExcludeCmdText(i)
                        End If

                    Case 1
                        If arrCmdText(i).Equals(ProcessCommand.pc.KIK.ToString) Then
                            addDelimiter(txt, iCmd)
                            txt &= arrCmdText(i) & " = '" & arrExcludeCmdText(i) & "'"
                        End If

                    Case Else
                        Return ""
                End Select
            Next

            'コマンド文字を、テーブルのフィールド名に置き換える
            txt = txt.Replace(ProcessCommand.pc.AM1.ToString, "MESY")
            txt = txt.Replace(ProcessCommand.pc.KIK.ToString, "KIKNAM")

            Return txt
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Sub addDelimiter(ByRef txt As String, ByVal iCmd As Integer)
        Select Case iCmd
            Case 0
                If Not txt.Equals("") Then
                    txt &= " , "
                End If
            Case 1
                If Not txt.Equals("") Then
                    txt &= " AND "
                End If
        End Select
    End Sub

    '未使用
    Private Function replaceCmdText(ByVal txt As String) As String
        Dim sBuf As String = ""
        Try
            sBuf = txt.Replace(ProcessCommand.pc.AM1.ToString, "MESY")
            sBuf = txt.Replace(ProcessCommand.pc.KIK.ToString, "VKONNO")

            Return sBuf
        Catch ex As Exception
            Return ""
        End Try
    End Function

    '枠網取得
    Public Function getWakuAmi(ByVal strSyoriFunn As String) As Integer
        Dim strSQL As String = ""
        Dim ds As New clsDsCtrl

        Try
            strSQL = " SELECT * "
            strSQL &= " FROM MM11K@MTRSLINK "
            strSQL &= " WHERE SUTBAN = 1 "
            If strSyoriFunn <> "" Then
                strSQL &= " AND SFCOD = '" & strSyoriFunn & "' "
            End If

            Call ds.Connect()
            Call ds.ExecuteSQL(strSQL, "")
            m_DT = ds.DataTable
            Return 0
        Catch ex As SqlException
            m_ErrMsg = ex.Message
            Return -9
        Finally
            Call ds.Dispose()
        End Try
    End Function

    'MM52削除
    Public Function Delete(ByVal DataUpd As clsDataUpd, ByVal strKikaiCD As String)
        Dim strSQL As String = ""
        Try
            strSQL &= " DELETE MM52K@MTRSLINK "
            strSQL &= " WHERE KIKCOD = '" & strKikaiCD & "' "
            Call DataUpd.ExecuteSQL(strSQL)
            Return 0
        Catch ex As Exception
            DataUpd.RollBackTran()
            m_ErrMsg = ex.Message
            Return -9
        End Try
    End Function

    'MM52追加
    Public Function Insert(ByVal DataUpd As clsDataUpd _
                         , ByVal strKikaiCD As String _
                         , ByVal strMESY1 As String _
                         , ByVal strMESY2 As String _
                         , ByVal strMESY3 As String _
                         , ByVal strWAKU1 As String _
                         , ByVal strWAKU2 As String _
                         , ByVal strWAKU3 As String _
                         , ByVal strWAKU4 As String)
        Dim strSQL As String = ""
        Try
            strSQL &= " INSERT INTO MM52K@MTRSLINK ( "
            strSQL &= "  KIKCOD "
            strSQL &= " ,MESY1 "
            strSQL &= " ,MESY2 "
            strSQL &= " ,MESY3 "
            strSQL &= " ,WAKU1 "
            strSQL &= " ,WAKU2 "
            strSQL &= " ,WAKU3 "
            strSQL &= " ,WAKU4 "
            strSQL &= " ,ADDYMD "
            strSQL &= " ,ADDHMS "
            strSQL &= " )VALUES( "
            strSQL &= "  '" & strKikaiCD & "' "
            strSQL &= " ," & strMESY1 & " "
            strSQL &= " ," & strMESY2 & " "
            strSQL &= " ," & strMESY3 & " "
            strSQL &= " ," & strWAKU1 & " "
            strSQL &= " ," & strWAKU2 & " "
            strSQL &= " ," & strWAKU3 & " "
            strSQL &= " ," & strWAKU4 & " "
            strSQL &= " ,'" & Now.ToString("yyyyMMdd") & "' "
            strSQL &= " ,'" & Now.ToString("HHmmss") & "' "
            strSQL &= " ) "
            Call DataUpd.ExecuteSQL(strSQL)
            Return 0
        Catch ex As Exception
            DataUpd.RollBackTran()
            m_ErrMsg = ex.Message
            Return -9
        End Try
    End Function


End Class
