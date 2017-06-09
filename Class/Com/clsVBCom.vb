Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports SZKDLL_Orc
Imports SZKDLL_Orc.Class
Imports SZKDLL_Orc.DBACCESS

Public Class clsVBCom
    'クラス名
    Private Const C_CLASSNAME As String = "clsVBCom.vb"
    Private m_DT As DataTable = Nothing
    Private m_ErrMsg As String = ""

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

    '画面クリア
    Public Sub PanelClear(ByVal Obj As Form)
        Dim Cnt As Control
        For Each Cnt In Obj.Controls
            If Not TypeOf Cnt Is Label _
             And Not TypeOf Cnt Is CheckBox _
             And Not TypeOf Cnt Is RadioButton _
             And Not TypeOf Cnt Is ComboBox _
             And Not TypeOf Cnt Is Button Then
                Cnt.Text = ""
            End If
        Next
    End Sub

    '画面クリア
    Public Sub PanelClear(ByVal panlObj As Panel)
        Dim Cnt As Control
        For Each Cnt In panlObj.Controls
            If Not TypeOf Cnt Is Label _
             And Not TypeOf Cnt Is CheckBox _
             And Not TypeOf Cnt Is RadioButton _
             And Not TypeOf Cnt Is ComboBox _
             And Not TypeOf Cnt Is Button Then
                Cnt.Text = ""
            End If
        Next
    End Sub

    '画面ロック
    Public Sub PanelEnable(ByVal panlObj As Panel, ByVal bolMode As Boolean)
        Dim Cnt As Control
        For Each Cnt In panlObj.Controls
            If Not TypeOf Cnt Is Label _
            And Not TypeOf Cnt Is Panel Then
                Cnt.Enabled = bolMode
            End If
        Next
    End Sub

    'Dialog処理
    Public Function bolDialog(ByVal strMsg As String, Optional ByVal strKengen As String = "") As Boolean
        Dim dialRet As DialogResult
        dialRet = MessageBox.Show(strMsg, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
        If dialRet <> DialogResult.Yes Then
            Return False
        End If
        Return True
    End Function

    '金額カマン表示
    Public Function setComma(ByVal strValue As String, Optional ByVal strDefValue As String = "") As String
        strValue = strValue.Replace(",", "")
        If IsNumeric(strValue) Then
            Return Convert.ToDecimal(strValue).ToString("N0")
        Else
            Return strDefValue
        End If
    End Function

    '金額カマン表示
    Public Function setCommaZero(ByVal strValue As String, Optional ByVal bolFlag As Boolean = True) As String
        strValue = strValue.Replace(",", "")
        If IsNumeric(strValue) Then
            If Val(strValue) = "0" Then
                Return ""
            End If
            If Not bolFlag Then
                Return Val(strValue).ToString
            End If
            Return Convert.ToDecimal(strValue).ToString("N0")
        End If
        Return ""
    End Function

    '金額カマン非表示
    Public Function delComma(ByVal strValue As String, Optional ByVal strDefValue As String = "") As String
        If strValue = "" Then
            Return strDefValue
        End If
        Return strValue.Replace(",", "")
    End Function

    'コンボボックスSelectedIndex設定
    Public Sub setComboSelectedIndex(ByVal cmbObj As ComboBox, ByVal index As Integer)
        If cmbObj.Items.Count > 0 And index < cmbObj.Items.Count Then
            cmbObj.SelectedIndex = index
        End If
    End Sub

    'コンボボックスSelectedValue設定
    Public Sub setComboSelectedValue(ByVal cmbObj As ComboBox, ByVal strValue As String)
        Try
            If Not cmbObj.SelectedValue Is Nothing Then
                cmbObj.SelectedValue = strValue
            Else
                cmbObj.Text = strValue
            End If
        Catch ex As Exception
            cmbObj.SelectedIndex = 0
        End Try
    End Sub

    'コンボボックスSelectedValue取得
    Public Function getComboSelectedValue(ByVal cmbObj As ComboBox, ByVal strValue As String) As String
        If Not cmbObj.SelectedValue Is Nothing Then
            Return cmbObj.SelectedValue.ToString
        End If
        Return strValue
    End Function

    '四捨五入、切り捨て、切り上げ
    Public Function FormatKingaku(ByVal decValue As Decimal, ByVal strMode As String, Optional ByVal intKeta As Integer = 0) As Decimal
        Select Case strMode
            Case clsSysParameter.gKirisute
                '切り捨て
                Return Math.Truncate(decValue)
            Case clsSysParameter.gKiriage
                '切り上げ
                Return Math.Ceiling(decValue)
            Case Else
                '四捨五入
                Return Math.Round(decValue * (10 ^ intKeta) + 0.00000001, 0) / (10 ^ intKeta)
        End Select
    End Function

    'Gridのスタイル設定
    Public Sub setGrdStyle(ByVal grdObj As DataGridView, Optional ByVal FontSize As Single = 9.0F)
        grdObj.EnableHeadersVisualStyles = False
        grdObj.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(128, 128, 255)
        grdObj.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        grdObj.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False
        grdObj.ColumnHeadersDefaultCellStyle.Font = New System.Drawing.Font("MS UI Gothic", FontSize, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        grdObj.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grdObj.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
        grdObj.DefaultCellStyle.Font = New System.Drawing.Font("MS UI Gothic", FontSize, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        grdObj.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        grdObj.AllowUserToResizeRows = False
        grdObj.AllowUserToAddRows = False
        grdObj.AllowUserToDeleteRows = False
        grdObj.CellBorderStyle = DataGridViewCellBorderStyle.Single
        grdObj.RowsDefaultCellStyle.BackColor = Color.White
        grdObj.BackgroundColor = System.Drawing.Color.FromName("AliceBlue")
        grdObj.GridColor = Color.Black
        grdObj.RowHeadersVisible = False
        grdObj.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        grdObj.MultiSelect = False
        grdObj.ReadOnly = True
        grdObj.AllowUserToAddRows = False
        grdObj.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(192, 255, 255)
        grdObj.RowTemplate.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Red
        grdObj.BackgroundColor = System.Drawing.Color.FromArgb(224, 224, 224)
        grdObj.Cursor = Cursors.Hand
        grdObj.ClearSelection()
    End Sub

    '時間日付表示
    Public Function getDataTime(ByVal strData As String, ByVal strTime As String) As String
        If strData.Length <> 8 Then
            Return ""
        End If
        Return Mid(strData, 1, 4) & "/" & Mid(strData, 5, 2) _
            & "/" & Mid(strData, 7, 2) & " " _
            & Mid(strTime, 1, 2) & ":" & Mid(strTime, 3)
    End Function

    '日付表示
    Public Function getData(ByVal strData As String) As String
        If strData.Length <> 8 Then
            Return ""
        End If
        Return Mid(strData, 1, 4) & "/" & Mid(strData, 5, 2) _
            & "/" & Mid(strData, 7, 2) 
    End Function

    '材料区分
    Public Function getZaiKBN(ByVal strData As String) As String
        Select Case strData
            Case "1"
                Return "銅粉"
            Case Else
                Return ""
        End Select
    End Function
End Class
