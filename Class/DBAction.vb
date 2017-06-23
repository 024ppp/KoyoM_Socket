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
    Public Function getSagyoName(ByRef names As String) As Boolean
        Dim strSQL As String = ""
        Dim ds As New clsDsCtrl

        Try
            strSQL = " SELECT * "
            strSQL &= " FROM MS02 "
            strSQL &= " WHERE DELKBN = 0 "
            strSQL &= " ORDER BY TNTCOD "

            Call ds.Connect()
            Call ds.ExecuteSQL(strSQL, "")

            If ds.DataTable.Rows.Count <= 0 Then
                m_ErrMsg = "作業者名取得エラー。"
                Return False
            End If

            For Each dtRow As DataRow In ds.DataTable.Rows
                If Not names.Equals("") Then
                    names &= ","
                End If
                names &= dtRow.Item("TNTNAM").ToString
            Next

            Return True
        Catch ex As Exception
            m_ErrMsg = ex.Message
            Return False
        Finally
            Call ds.Dispose()
        End Try
    End Function

    'Vコンが存在するかチェックして結果を返す
    Public Function checkKikai(ByVal strKikai As String) As Boolean
        Dim strSQL As String = ""
        Dim ds As New clsDsCtrl

        Try
            strSQL = " SELECT * "
            strSQL &= " FROM MM52K@MTRSLINK "
            strSQL &= " WHERE 1 = 1 "
            If strKikai <> "" Then
                strSQL &= " AND KIKCOD = '" & strKikai & "' "
            End If

            Call ds.Connect()
            Call ds.ExecuteSQL(strSQL, "")

            If ds.DataTable.Rows.Count <= 0 Then
                m_ErrMsg = "機械Noが存在しません。"
                Return False
            End If

            Return True
        Catch ex As Exception
            m_ErrMsg = ex.Message
            Return False
        Finally
            Call ds.Dispose()
        End Try
    End Function

    '処理粉情報を取得
    Public Function getSyoriInfo(ByVal sKokban As String _
                               , ByRef sSfcd As String _
                               , ByRef sZainmk As String) As Boolean
        Dim strSQL As String = ""
        Dim ds As New clsDsCtrl

        Try
            If Not IsNumeric(sKokban) Then
                Return ""
            End If

            strSQL = " SELECT SFCD,ZAINMK "
            strSQL &= " FROM MD01@MTRSLINK "
            strSQL &= " WHERE 1 = 1 "
            If sKokban <> "" Then
                strSQL &= " AND KOKBAN = '" & sKokban & "' "
            End If

            Call ds.Connect()
            Call ds.ExecuteSQL(strSQL, "")

            If ds.DataTable.Rows.Count <= 0 Then
                m_ErrMsg = "工管番号が存在しません。"
                Return False
            End If

            For Each dtRow As DataRow In ds.DataTable.Rows
                sSfcd = dtRow.Item("SFCD").ToString
                sZainmk = dtRow.Item("ZAINMK").ToString
            Next

            Return True
        Catch ex As SqlException
            m_ErrMsg = ex.Message
            Return False
        Finally
            Call ds.Dispose()
        End Try
    End Function

    '処理粉情報を取得
    Public Function getWakuAmi(ByVal strSfcd As String, ByRef sWakuami As String) As Boolean
        Dim strSQL As String = ""
        Dim ds As New clsDsCtrl

        Try
            strSQL = " SELECT * "
            strSQL &= " FROM MM11K@MTRSLINK "
            strSQL &= " WHERE SUTBAN = 1 "
            If strSfcd <> "" Then
                strSQL &= " AND SFCOD = '" & strSfcd & "' "
            End If

            Call ds.Connect()
            Call ds.ExecuteSQL(strSQL, "")

            If ds.DataTable.Rows.Count <= 0 Then
                m_ErrMsg = "枠網情報が存在しません。"
                Return False
            End If

            For Each dtRow As DataRow In ds.DataTable.Rows
                sWakuami = dtRow.Item("WAKU1").ToString & "," & _
                           dtRow.Item("MESY1").ToString & "," & _
                           dtRow.Item("WAKU2").ToString & "," & _
                           dtRow.Item("MESY2").ToString & "," & _
                           dtRow.Item("WAKU3").ToString & "," & _
                           dtRow.Item("MESY3").ToString & "," & _
                           dtRow.Item("WAKU4").ToString
            Next

            Return True
        Catch ex As SqlException
            m_ErrMsg = ex.Message
            Return False
        Finally
            Call ds.Dispose()
        End Try
    End Function

    'MM52K更新
    Public Function RegisterMM52K(ByVal strUpdText As String) As Boolean
        Dim strSQL As String = ""
        Dim arrUpdText As String() = Nothing
        Dim ds As New clsDsCtrl

        Try
            arrUpdText = strUpdText.Split(",")

            'SQL文作成
            'Delete
            strSQL = " DELETE MM52K@MTRSLINK "
            strSQL &= " WHERE KIKCOD = '" & arrUpdText(0) & "' "

            Call ds.Connect()
            Call ds.ExecuteSQL(strSQL, "")

            'Insert
            strSQL = " INSERT INTO MM52K@MTRSLINK ( "
            strSQL &= "  KIKCOD "
            strSQL &= " ,WAKU1 "
            strSQL &= " ,MESY1 "
            strSQL &= " ,WAKU2 "
            strSQL &= " ,MESY2 "
            strSQL &= " ,WAKU3 "
            strSQL &= " ,MESY3 "
            strSQL &= " ,WAKU4 "
            strSQL &= " ,ADDYMD "
            strSQL &= " ,ADDHMS "
            strSQL &= " )VALUES( "
            For Each updText As String In arrUpdText
                strSQL &= " '" & updText & "', "
            Next
            strSQL &= " '" & Now.ToString("yyyyMMdd") & "', "
            strSQL &= " '" & Now.ToString("HHmmss") & "' "
            strSQL &= " ) "

            Call ds.ExecuteSQL(strSQL, "")
            Return True
        Catch ex As Exception
            m_ErrMsg = ex.Message
            Return False
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
