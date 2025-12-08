Imports System
Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Net
Imports System.Data.SqlClient
Public Class FormControl

    Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Integer)
    'unicodeのEcodingクラスに作成
    Dim encUni As Encoding = Encoding.GetEncoding("utf-16")
    's-jisのEncodingクラスの作成
    Dim encSjis As Encoding = Encoding.GetEncoding("shift-jis")

    Private Sub FormControl_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '日本語・英語表記の切り替えを行う
        Translation_FormControl()
        'グラフフォーマット候補を表示する==========
        ComboBox_Format.Items.Clear()
        For i As Integer = 0 To UBound(gType, 1)
            ComboBox_Format.Items.Add(gType(i))
        Next

        DTP_Dstartdate.Value = strStartDate
        DTP_Astartdate.Value = strAlarmStartDate
        '==========================================
        '初期値を0とする===========================
        'Me.TextX_CL.Text = 0
        'Me.TextX_UCL.Text = 0
        'Me.TextX_LCL.Text = 0
        'Me.TextX_S.Text = 0
        'Me.TextR_CL.Text = 0
        'Me.TextR_UCL.Text = 0
        'Me.TextR_S.Text = 0
        'Me.TextBox1.Text = 0
        ''Me.TextBox2.Text = 0
        'Me.Text_Upper.Text = 0
        'Me.Text_Lower.Text = 0
        '==========================================

        Me.CheckBox1.Checked = True
        Me.CheckBox2.Checked = True
        Me.CheckBox3.Checked = True
        Get_ProcessInfo()
        Get_TreeInfo()

    End Sub

    'サーバーよりツリー一覧を取得する
    Public Sub Get_TreeInfo()
        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim Adapter As New SqlDataAdapter
        Dim strSQL As String = ""
        Dim n As Integer
        Dim Tree1Buf(vv), Tree2Buf(vv), Tree3Buf(vv), Tree4Buf(vv), Tree5Buf(vv), Tree6Buf(vv), Tree7Buf(vv), Tree8Buf(vv), Tree9Buf(vv), Tree10Buf(vv) As String
        Dim c1, c2, c3, c4, c5, c6, c7, c8, c9, c10 As Integer
        Dim Exist1, Exist2, Exist3, Exist4, Exist5, Exist6, Exist7, Exist8, Exist9, Exist10 As Boolean
        Dim table As New DataTable
        Try

            ComboBox_Tree1.Items.Clear()
            ComboBox_Tree2.Items.Clear()
            ComboBox_Tree3.Items.Clear()
            ComboBox_Tree4.Items.Clear()
            ComboBox_Tree5.Items.Clear()

            Cn.ConnectionString = StrServerConnection
            table.Clear()

            strSQL = "SELECT cTreeName1,"
            strSQL &= " cTreeName2,"
            strSQL &= " cTreeName3,"
            strSQL &= " cTreeName4,"
            strSQL &= " cTreeName5,"
            strSQL &= " cTreeName6,"
            strSQL &= " cTreeName7,"
            strSQL &= " cTreeName8,"
            strSQL &= " cTreeName9,"
            strSQL &= " cTreeName10"
            strSQL &= " FROM SPC_Property"
            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(table)
            n = table.Rows.Count

            Adapter.Dispose()
            Cn.Dispose()

            If n = 0 Then
                table.Dispose()
                Exit Sub
            End If
            c1 = 0
            c2 = 0
            c3 = 0
            c4 = 0
            c5 = 0
            c6 = 0
            c7 = 0
            c8 = 0
            c4 = 0
            c10 = 0
            For i = 0 To n - 1
                'ツリー1のリストを表示-----------------------------------
                Exist1 = False
                For j = 0 To c1
                    If Tree1Buf(j) = table.Rows(i)("cTreeName1") Then
                        Exist1 = True
                    End If
                Next
                If Exist1 = False Then
                    Tree1Buf(c1) = table.Rows(i)("cTreeName1")
                    ComboBox_Tree1.Items.Add(table.Rows(i)("cTreeName1"))
                    c1 += 1
                End If
                '--------------------------------------------------------

                'ツリー2のリストを表示-----------------------------------
                Exist2 = False
                For j = 0 To c2
                    If Tree2Buf(j) = table.Rows(i)("cTreeName2") Then
                        Exist2 = True
                    End If
                Next
                If Exist2 = False Then
                    Tree2Buf(c2) = table.Rows(i)("cTreeName2")
                    ComboBox_Tree2.Items.Add(table.Rows(i)("cTreeName2"))
                    c2 += 1
                End If
                '--------------------------------------------------------


                'ツリー3のリストを表示-----------------------------------
                Exist3 = False
                For j = 0 To c3
                    If Tree3Buf(j) = table.Rows(i)("cTreeName3") Then
                        Exist3 = True
                    End If
                Next
                If Exist3 = False Then
                    Tree3Buf(c3) = table.Rows(i)("cTreeName3")
                    ComboBox_Tree3.Items.Add(table.Rows(i)("cTreeName3"))
                    c3 += 1
                End If
                '--------------------------------------------------------

                'ツリー4のリストを表示-----------------------------------
                Exist4 = False
                For j = 0 To c4
                    If Tree4Buf(j) = table.Rows(i)("cTreeName4") Then
                        Exist4 = True
                    End If
                Next
                If Exist4 = False Then
                    Tree4Buf(c4) = table.Rows(i)("cTreeName4")
                    ComboBox_Tree4.Items.Add(table.Rows(i)("cTreeName4"))
                    c4 += 1
                End If
                '--------------------------------------------------------


                'ツリー5のリストを表示-----------------------------------
                Exist5 = False
                For j = 0 To c5
                    If Tree5Buf(j) = table.Rows(i)("cTreeName5") Then
                        Exist5 = True
                    End If
                Next
                If Exist5 = False Then
                    Tree5Buf(c5) = table.Rows(i)("cTreeName5")
                    ComboBox_Tree5.Items.Add(table.Rows(i)("cTreeName5"))
                    c5 += 1
                End If
                '--------------------------------------------------------

                'ツリー6のリストを表示-----------------------------------
                Exist6 = False
                For j = 0 To c6
                    If Tree6Buf(j) = table.Rows(i)("cTreeName6") Then
                        Exist6 = True
                    End If
                Next
                If Exist6 = False Then
                    Tree6Buf(c6) = table.Rows(i)("cTreeName6")
                    ComboBox_Tree6.Items.Add(table.Rows(i)("cTreeName6"))
                    c6 += 1
                End If
                '--------------------------------------------------------
                'ツリー7のリストを表示-----------------------------------
                Exist7 = False
                For j = 0 To c7
                    If Tree7Buf(j) = table.Rows(i)("cTreeName7") Then
                        Exist7 = True
                    End If
                Next
                If Exist7 = False Then
                    Tree7Buf(c7) = table.Rows(i)("cTreeName7")
                    ComboBox_Tree7.Items.Add(table.Rows(i)("cTreeName7"))
                    c7 += 1
                End If
                '--------------------------------------------------------
                'ツリー8のリストを表示-----------------------------------
                Exist8 = False
                For j = 0 To c8
                    If Tree8Buf(j) = table.Rows(i)("cTreeName8") Then
                        Exist8 = True
                    End If
                Next
                If Exist8 = False Then
                    Tree8Buf(c8) = table.Rows(i)("cTreeName8")
                    ComboBox_Tree8.Items.Add(table.Rows(i)("cTreeName8"))
                    c8 += 1
                End If
                '--------------------------------------------------------
                'ツリー9のリストを表示-----------------------------------
                Exist9 = False
                For j = 0 To c9
                    If Tree9Buf(j) = table.Rows(i)("cTreeName9") Then
                        Exist9 = True
                    End If
                Next
                If Exist9 = False Then
                    Tree9Buf(c9) = table.Rows(i)("cTreeName9")
                    ComboBox_Tree9.Items.Add(table.Rows(i)("cTreeName9"))
                    c9 += 1
                End If
                '--------------------------------------------------------
                'ツリー10のリストを表示-----------------------------------
                Exist10 = False
                For j = 0 To c10
                    If Tree10Buf(j) = table.Rows(i)("cTreeName10") Then
                        Exist10 = True
                    End If
                Next
                If Exist10 = False Then
                    Tree10Buf(c10) = table.Rows(i)("cTreeName10")
                    ComboBox_Tree10.Items.Add(table.Rows(i)("cTreeName10"))
                    c10 += 1
                End If
                '--------------------------------------------------------
            Next


        Catch ex As System.Exception
            Adapter.Dispose()
            Cn.Dispose()

            StrErrMes = "ツリー一覧取得エラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Sub
        End Try
    End Sub

    'サーバーよりプロセス一覧を取得する
    Public Sub Get_ProcessInfo()
        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim Adapter As New SqlDataAdapter
        Dim strSQL As String = ""
        Dim n As Integer
        Dim table As New DataTable
        Try

            ComboBox_Process.Items.Clear()

            Cn.ConnectionString = StrServerConnection
            table.Clear()

            strSQL = "SELECT DISTINCT cProcessName"

            strSQL &= " FROM SPC_Master"
            strSQL &= " ORDER BY cProcessName"
            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(table)
            n = table.Rows.Count

            Adapter.Dispose()
            Cn.Dispose()

            If n = 0 Then
                table.Dispose()
                Exit Sub
            End If
            For i = 0 To n - 1
                ComboBox_Process.Items.Add(table.Rows(i)("cProcessName"))
            Next


        Catch ex As System.Exception
            Adapter.Dispose()
            Cn.Dispose()

            StrErrMes = "プロセス一覧取得エラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Sub
        End Try
    End Sub


    'サーバーより設備No一覧を取得する
    Public Sub Get_McNoInfo(ByVal _Process As String)
        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim Adapter As New SqlDataAdapter
        Dim strSQL As String = ""
        Dim n As Integer
        Dim table As New DataTable
        Try

            ComboBox_McNo.Items.Clear()

            Cn.ConnectionString = StrServerConnection
            table.Clear()

            strSQL = "SELECT DISTINCT cMachineNo"
            strSQL &= " FROM SPC_Master"
            strSQL &= " WHERE cProcessName = '" & _Process & "'"
            strSQL &= " ORDER BY cMachineNo"
            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(table)
            n = table.Rows.Count

            Adapter.Dispose()
            Cn.Dispose()

            If n = 0 Then
                table.Dispose()
                Exit Sub
            End If
            For i = 0 To n - 1
                ComboBox_McNo.Items.Add(table.Rows(i)("cMachineNo"))
            Next


        Catch ex As System.Exception
            Adapter.Dispose()
            Cn.Dispose()

            StrErrMes = "設備No一覧取得エラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Sub
        End Try
    End Sub

    'サーバーより管理項目一覧を取得する
    Public Sub Get_ControlItemInfo(ByVal _McNo As String)
        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim Adapter As New SqlDataAdapter
        Dim strSQL As String = ""
        Dim n As Integer
        Dim table As New DataTable
        Try

            ComboBox_Item.Items.Clear()

            Cn.ConnectionString = StrServerConnection
            table.Clear()

            strSQL = "SELECT DISTINCT cControlItem"
            strSQL &= " FROM SPC_Master"
            strSQL &= " WHERE cMachineNo = '" & _McNo & "'"
            strSQL &= " ORDER BY cControlItem"
            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(table)
            n = table.Rows.Count

            Adapter.Dispose()
            Cn.Dispose()

            If n = 0 Then
                table.Dispose()
                Exit Sub
            End If
            For i = 0 To n - 1
                ComboBox_Item.Items.Add(table.Rows(i)("cControlItem"))
            Next


        Catch ex As System.Exception
            Adapter.Dispose()
            Cn.Dispose()

            StrErrMes = "管理項目一覧取得エラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Sub
        End Try
    End Sub

    'サーバーより機種一覧を取得する
    Public Sub Get_DeviceInfo(ByVal _McNo As String, ByVal _ControlItem As String)
        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim Adapter As New SqlDataAdapter
        Dim strSQL As String = ""
        Dim n As Integer
        Dim table As New DataTable
        Try

            ComboBox_Device.Items.Clear()

            Cn.ConnectionString = StrServerConnection
            table.Clear()

            strSQL = "SELECT DISTINCT cDeviceName"
            strSQL &= " FROM SPC_Master"
            strSQL &= " WHERE cMachineNo = '" & _McNo & "'"
            strSQL &= " AND cControlItem = '" & _ControlItem & "'"
            strSQL &= " ORDER BY cDeviceName"
            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(table)
            n = table.Rows.Count

            Adapter.Dispose()
            Cn.Dispose()

            If n = 0 Then
                table.Dispose()
                Exit Sub
            End If
            For i = 0 To n - 1
                ComboBox_Device.Items.Add(table.Rows(i)("cDeviceName"))
            Next


        Catch ex As System.Exception
            Adapter.Dispose()
            Cn.Dispose()

            StrErrMes = "機種一覧取得エラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Sub
        End Try
    End Sub

    'サーバーより測定条件を取得する
    Public Sub Get_FilterInfo(ByVal _McNo As String, ByVal _ControlItem As String)
        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim Adapter As New SqlDataAdapter
        Dim strSQL As String = ""
        Dim n As Integer
        Dim table As New DataTable
        Dim Tree1Buf(vv), Tree2Buf(vv), Tree3Buf(vv), Tree4Buf(vv), Tree5Buf(vv), Tree6Buf(vv), Tree7Buf(vv), Tree8Buf(vv), Tree9Buf(vv), Tree10Buf(vv) As String
        Dim c1, c2, c3, c4, c5, c6, c7, c8, c9, c10 As Integer
        Dim Exist1, Exist2, Exist3, Exist4, Exist5, Exist6, Exist7, Exist8, Exist9, Exist10 As Boolean
        Try

            ComboBox_Filter1.Items.Clear()
            ComboBox_Filter2.Items.Clear()
            ComboBox_Filter3.Items.Clear()
            ComboBox_Filter4.Items.Clear()
            ComboBox_Filter5.Items.Clear()

            Cn.ConnectionString = StrServerConnection
            table.Clear()

            strSQL = "SELECT cFilter_1,"
            strSQL &= " cFilter_2,"
            strSQL &= " cFilter_3,"
            strSQL &= " cFilter_4,"
            strSQL &= " cFilter_5,"
            strSQL &= " cFilter_6,"
            strSQL &= " cFilter_7,"
            strSQL &= " cFilter_8,"
            strSQL &= " cFilter_9,"
            strSQL &= " cFilter_10"
            strSQL &= " FROM SPC_Master"
            strSQL &= " WHERE cMachineNo = '" & _McNo & "'"
            strSQL &= " AND cControlItem = '" & _ControlItem & "'"

            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(table)
            n = table.Rows.Count

            Adapter.Dispose()
            Cn.Dispose()

            If n = 0 Then
                table.Dispose()
                Exit Sub
            End If


            c1 = 0
            c2 = 0
            c3 = 0
            c4 = 0
            c5 = 0
            c6 = 0
            c7 = 0
            c8 = 0
            c9 = 0
            c10 = 0
            For i = 0 To n - 1
                'フィルター1のリストを表示-----------------------------------
                Exist1 = False
                If (Not IsDBNull(table.Rows(i)("cFilter_1"))) Then
                    If Not table.Rows(i)("cFilter_1") = "" Then
                        For j = 0 To c1
                            If Tree1Buf(j) = table.Rows(i)("cFilter_1") Then
                                Exist1 = True
                            End If
                        Next
                        If Exist1 = False Then
                            Tree1Buf(c1) = table.Rows(i)("cFilter_1")
                            ComboBox_Filter1.Items.Add(table.Rows(i)("cFilter_1"))
                            c1 += 1
                        End If
                    End If
                End If
                '--------------------------------------------------------
                'フィルター2のリストを表示-----------------------------------
                Exist2 = False
                If (Not IsDBNull(table.Rows(i)("cFilter_2"))) Then
                    If Not table.Rows(i)("cFilter_2") = "" Then
                        For j = 0 To c2
                            If Tree2Buf(j) = table.Rows(i)("cFilter_2") Then
                                Exist2 = True
                            End If
                        Next
                        If Exist2 = False Then
                            Tree2Buf(c2) = table.Rows(i)("cFilter_2")
                            ComboBox_Filter2.Items.Add(table.Rows(i)("cFilter_2"))
                            c2 += 1
                        End If
                    End If
                End If
                '--------------------------------------------------------
                'フィルター3のリストを表示-----------------------------------
                Exist3 = False
                If (Not IsDBNull(table.Rows(i)("cFilter_3"))) Then
                    If Not table.Rows(i)("cFilter_3") = "" Then
                        For j = 0 To c3
                            If Tree3Buf(j) = table.Rows(i)("cFilter_3") Then
                                Exist3 = True
                            End If
                        Next
                        If Exist3 = False Then
                            Tree3Buf(c3) = table.Rows(i)("cFilter_3")
                            ComboBox_Filter3.Items.Add(table.Rows(i)("cFilter_3"))
                            c3 += 1
                        End If
                    End If
                End If
                '--------------------------------------------------------
                'フィルター4のリストを表示-----------------------------------
                Exist4 = False
                If (Not IsDBNull(table.Rows(i)("cFilter_4"))) Then
                    If Not table.Rows(i)("cFilter_4") = "" Then
                        For j = 0 To c4
                            If Tree4Buf(j) = table.Rows(i)("cFilter_4") Then
                                Exist4 = True
                            End If
                        Next
                        If Exist4 = False Then
                            Tree4Buf(c4) = table.Rows(i)("cFilter_4")
                            ComboBox_Filter4.Items.Add(table.Rows(i)("cFilter_4"))
                            c4 += 1
                        End If
                    End If
                End If
                '--------------------------------------------------------
                'フィルター5のリストを表示-----------------------------------
                Exist5 = False
                If (Not IsDBNull(table.Rows(i)("cFilter_5"))) Then
                    If Not table.Rows(i)("cFilter_5") = "" Then
                        For j = 0 To c5
                            If Tree5Buf(j) = table.Rows(i)("cFilter_5") Then
                                Exist5 = True
                            End If
                        Next
                        If Exist5 = False Then
                            Tree5Buf(c5) = table.Rows(i)("cFilter_5")
                            ComboBox_Filter5.Items.Add(table.Rows(i)("cFilter_5"))
                            c5 += 1
                        End If
                    End If
                End If
                '--------------------------------------------------------
                'フィルター6のリストを表示-----------------------------------
                Exist6 = False
                If (Not IsDBNull(table.Rows(i)("cFilter_6"))) Then
                    If Not table.Rows(i)("cFilter_6") = "" Then
                        For j = 0 To c6
                            If Tree6Buf(j) = table.Rows(i)("cFilter_6") Then
                                Exist6 = True
                            End If
                        Next
                        If Exist6 = False Then
                            Tree6Buf(c6) = table.Rows(i)("cFilter_6")
                            ComboBox_Filter6.Items.Add(table.Rows(i)("cFilter_6"))
                            c6 += 1
                        End If
                    End If
                End If
                '--------------------------------------------------------
                'フィルター7のリストを表示-----------------------------------
                Exist7 = False
                If (Not IsDBNull(table.Rows(i)("cFilter_7"))) Then
                    If Not table.Rows(i)("cFilter_7") = "" Then
                        For j = 0 To c7
                            If Tree7Buf(j) = table.Rows(i)("cFilter_7") Then
                                Exist7 = True
                            End If
                        Next
                        If Exist7 = False Then
                            Tree7Buf(c7) = table.Rows(i)("cFilter_7")
                            ComboBox_Filter7.Items.Add(table.Rows(i)("cFilter_7"))
                            c7 += 1
                        End If
                    End If
                End If
                '--------------------------------------------------------
                'フィルター8のリストを表示-----------------------------------
                Exist8 = False
                If (Not IsDBNull(table.Rows(i)("cFilter_8"))) Then
                    If Not table.Rows(i)("cFilter_8") = "" Then
                        For j = 0 To c8
                            If Tree8Buf(j) = table.Rows(i)("cFilter_8") Then
                                Exist8 = True
                            End If
                        Next
                        If Exist8 = False Then
                            Tree8Buf(c8) = table.Rows(i)("cFilter_8")
                            ComboBox_Filter8.Items.Add(table.Rows(i)("cFilter_8"))
                            c8 += 1
                        End If
                    End If
                End If
                '--------------------------------------------------------
                'フィルター9のリストを表示-----------------------------------
                Exist9 = False
                If (Not IsDBNull(table.Rows(i)("cFilter_9"))) Then
                    If Not table.Rows(i)("cFilter_9") = "" Then
                        For j = 0 To c9
                            If Tree9Buf(j) = table.Rows(i)("cFilter_9") Then
                                Exist9 = True
                            End If
                        Next
                        If Exist9 = False Then
                            Tree9Buf(c9) = table.Rows(i)("cFilter_9")
                            ComboBox_Filter9.Items.Add(table.Rows(i)("cFilter_9"))
                            c9 += 1
                        End If
                    End If
                End If
                '--------------------------------------------------------
                'フィルター10のリストを表示-----------------------------------
                Exist10 = False
                If (Not IsDBNull(table.Rows(i)("cFilter_10"))) Then
                    If Not table.Rows(i)("cFilter_10") = "" Then
                        For j = 0 To c10
                            If Tree10Buf(j) = table.Rows(i)("cFilter_10") Then
                                Exist10 = True
                            End If
                        Next
                        If Exist10 = False Then
                            Tree10Buf(c10) = table.Rows(i)("cFilter_10")
                            ComboBox_Filter10.Items.Add(table.Rows(i)("cFilter_10"))
                            c10 += 1
                        End If
                    End If
                End If
                '--------------------------------------------------------
            Next


        Catch ex As System.Exception
            Adapter.Dispose()
            Cn.Dispose()

            StrErrMes = "測定条件取得エラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Sub
        End Try
    End Sub
    'SPCプロパティのMAXIDを取得する
    Public Function Get_PropertyNo() As String
        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim Adapter As New SqlDataAdapter
        Dim strSQL As String = ""
        Dim n As Integer
        Dim table As New DataTable
        Try
            Get_PropertyNo = ""
            ComboBox_Process.Items.Clear()

            Cn.ConnectionString = StrServerConnection
            table.Clear()

            strSQL = "SELECT MAX(iGraphNo)"
            strSQL &= " FROM SPC_Property"

            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(table)
            n = table.Rows.Count

            Adapter.Dispose()
            Cn.Dispose()

            If n = 0 Then
                table.Dispose()
                Get_PropertyNo = 1
                Exit Function
            End If

            Get_PropertyNo = table.Rows(0)(0) + 1

            'Form2.Show()
            'Form2.DataGridView1.DataSource = table


        Catch ex As System.Exception
            Get_PropertyNo = ""
            Adapter.Dispose()
            Cn.Dispose()

            StrErrMes = "プロセス一覧取得エラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Function
        End Try
    End Function
    '*******************************************************************
    'SPCプロパティデータをサーバーにインサートする
    '*******************************************************************
    Private Sub INSERT_PropertyInfo(ByVal _ID As String, ByVal _Na As String)

        Dim Cn As New SqlConnection
        Dim strSQL As String
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        Dim trans As SqlTransaction 'トランザクション定義

        Try

            Cn.ConnectionString = StrServerConnection

            Cn.Open()
            trans = Cn.BeginTransaction
            SQLCm.Transaction = trans

            strSQL = ""
            strSQL = "INSERT INTO SPC_Property VALUES ('"
            strSQL &= _ID & "','" '01 iGraphNo
            strSQL &= ComboBox_Process.Text & "','" '02 cProcessName
            strSQL &= ComboBox_McNo.Text & "','" '03 cMachineNo
            strSQL &= ComboBox_Item.Text & "','" '04 cControlItem
            strSQL &= ComboBox_Device.Text & "','" '05 cDeviceName
            strSQL &= ComboBox_Filter1.Text & "','" '06 cFilter_1
            strSQL &= ComboBox_Filter2.Text & "','" '07 cFilter_2
            strSQL &= ComboBox_Filter3.Text & "','" '08 cFilter_3
            strSQL &= ComboBox_Filter4.Text & "','" '09 cFilter_4
            strSQL &= ComboBox_Filter5.Text & "','" '10 cFilter_5
            strSQL &= ComboBox_Filter6.Text & "','" '11 cFilter_6
            strSQL &= ComboBox_Filter7.Text & "','" '12 cFilter_7
            strSQL &= ComboBox_Filter8.Text & "','" '13 cFilter_8
            strSQL &= ComboBox_Filter9.Text & "','" '14 cFilter_9
            strSQL &= ComboBox_Filter10.Text & "','" '15 cFilter_10
            strSQL &= ComboBox_Tree1.Text & "','" '16 cTreeName1
            strSQL &= ComboBox_Tree2.Text & "','" '17 cTreeName2
            strSQL &= ComboBox_Tree3.Text & "','" '18 cTreeName3
            strSQL &= ComboBox_Tree4.Text & "','" '19 cTreeName4
            strSQL &= ComboBox_Tree5.Text & "','" '20 cTreeName5
            strSQL &= ComboBox_Tree6.Text & "','" '21 cTreeName6
            strSQL &= ComboBox_Tree7.Text & "','" '22 cTreeName7
            strSQL &= ComboBox_Tree8.Text & "','" '23 cTreeName8
            strSQL &= ComboBox_Tree9.Text & "','" '24 cTreeName9
            strSQL &= ComboBox_Tree10.Text & "','" '25 cTreeName10
            strSQL &= TextBox5.Text & "','" '26 cUnit


            For i As Integer = 0 To UBound(gType, 1)
                If ComboBox_Format.Text = gType(i) Then
                    If i = 0 Then
                        strSQL &= "UpperLower" & "','" '27 cLimitType
                        strSQL &= (CSng(Text_Upper.Text) + CSng(Text_Lower.Text)) / 2 & "','" '28 cScl
                        strSQL &= (Text_Upper.Text - Text_Lower.Text) / 2 & "','" '29 cTolerance
                    ElseIf i = 1 Then
                        strSQL &= "Upper" & "','" '27 cLimitType
                        strSQL &= TextX_CL.Text & "','" '28 cScl
                        strSQL &= (Text_Upper.Text - TextX_CL.Text) & "','" '29 cTolerance
                    ElseIf i = 2 Then
                        strSQL &= "Lower" & "','" '27 cLimitType
                        strSQL &= TextX_CL.Text & "','" '28 cScl
                        strSQL &= (TextX_CL.Text - Text_Lower.Text) & "','" '29 cTolerance
                    End If
                End If
            Next


            strSQL &= Text_Upper.Text & "','" '30 cUsl
            strSQL &= Text_Lower.Text & "','" '31 cLsl
            strSQL &= TextX_CL.Text & "','" '32cXcl
            strSQL &= TextX_S.Text & "','" '33 cXdev
            strSQL &= TextX_UCL.Text & "','" '34 cXucl
            strSQL &= TextX_LCL.Text & "','" '35 cXlcl
            strSQL &= TextR_CL.Text & "','" '36 cRcl
            strSQL &= TextR_S.Text & "','" '37 cRdev
            strSQL &= TextR_UCL.Text & "','" '38 cRucl
            strSQL &= "0" & "','" '39 cRlcl
            strSQL &= "0" & "','" '40 cMRcl
            strSQL &= "0" & "','" '41 cMRdev
            strSQL &= "0" & "','" '42 cMRucl
            strSQL &= "0" & "','" '43 cMRlcl

            If CheckBox9.Checked = True Then
                strSQL &= "1" & "','" '44 cMR
            Else
                strSQL &= "0" & "','" '44 cMR
            End If
            If CheckBox1.Checked = True Then
                strSQL &= "1" & "','" '45 cSpcrule1
            Else
                strSQL &= "0" & "','" '45 cSpcrule1
            End If
            If CheckBox2.Checked = True Then
                strSQL &= "1" & "','" '46 cSpcrule2
            Else
                strSQL &= "0" & "','" '46 cSpcrule2
            End If
            If CheckBox3.Checked = True Then
                strSQL &= "1" & "','" '47 cSpcrule3
            Else
                strSQL &= "0" & "','" '47 cSpcrule3
            End If
            If CheckBox4.Checked = True Then
                strSQL &= "1" & "','" '48 cSpcrule4
            Else
                strSQL &= "0" & "','" '48 cSpcrule4
            End If
            If CheckBox5.Checked = True Then
                strSQL &= "1" & "','" '49 cSpcrule5
            Else
                strSQL &= "0" & "','" '49 cSpcrule5
            End If
            If CheckBox6.Checked = True Then
                strSQL &= "1" & "','" '50 cSpcrule6
            Else
                strSQL &= "0" & "','" '50 cSpcrule6
            End If
            If CheckBox7.Checked = True Then
                strSQL &= "1" & "','" '51 cSpcrule7
            Else
                strSQL &= "0" & "','" '51 cSpcrule7
            End If
            If CheckBox8.Checked = True Then
                strSQL &= "1" & "','" '52 cSpcrule8
            Else
                strSQL &= "0" & "','" '52 cSpcrule8
            End If

            strSQL &= DTP_Rdate.Value & "','" '53 cUpdateDate
            strSQL &= "Initial setting" & "','" '54 cUpdateContent
            strSQL &= "Initial setting" & "','" '55 cUpdateReason
            strSQL &= _Na & "','" '56 cIncharge
            strSQL &= DTP_Rdate.Value & "','" '57 cApprovalDate
            strSQL &= "Unknown" & "','" '58 cApproverName
            strSQL &= DTP_Dstartdate.Value & "','" '59 dStartDate
            strSQL &= DTP_Astartdate.Value & "','" '60 aStartDate
            strSQL &= Text_PRODUCT.Text & "','" '61 PRODUCT
            strSQL &= Text_SamQuantity.Text & "')" '62 Sample Quantity


            'strSQL &= _Category & "','"
            'strSQL &= TreeName1 & "','"
            'strSQL &= TreeName2 & "','"
            'strSQL &= TreeName3 & "','"
            'strSQL &= TreeName4 & "','"
            'strSQL &= TreeName5 & "','"

            'If _AlarmNo = "1" Then
            '    strSQL &= "True" & "','"
            'Else
            '    strSQL &= "False" & "','"
            'End If
            'If _AlarmNo = "2" Then
            '    strSQL &= "True" & "','"
            'Else
            '    strSQL &= "False" & "','"
            'End If
            'If _AlarmNo = "3" Then
            '    strSQL &= "True" & "','"
            'Else
            '    strSQL &= "False" & "','"
            'End If
            'If _AlarmNo = "4" Then
            '    strSQL &= "True" & "','"
            'Else
            '    strSQL &= "False" & "','"
            'End If
            'If _AlarmNo = "5" Then
            '    strSQL &= "True" & "','"
            'Else
            '    strSQL &= "False" & "','"
            'End If
            'If _AlarmNo = "6" Then
            '    strSQL &= "True" & "','"
            'Else
            '    strSQL &= "False" & "','"
            'End If
            'If _AlarmNo = "7" Then
            '    strSQL &= "True" & "','"
            'Else
            '    strSQL &= "False" & "','"
            'End If
            'If _AlarmNo = "8" Then
            '    strSQL &= "True" & "','"
            'Else
            '    strSQL &= "False" & "','"
            'End If

            'strSQL &= "','','','','','','',''" & ")"

            SQLCm.CommandText = strSQL
            SQLCm.ExecuteNonQuery()

            trans.Commit()
            Cn.Close()

        Catch ex As Exception
            If IsNothing(trans) = False Then
                trans.Rollback()
            End If
            StrErrMes = "SPCアラームデータ更新エラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Sub
        End Try

    End Sub
    'Private Function Check_Password(ByRef Na As String) As Boolean
    '    Dim strPassword As String = ""
    '    If StrLanguage = "Japanese" Then
    '        strPassword = InputBox("SPCを新規追加します。管理者パスワードを入力して下さい。")
    '    ElseIf StrLanguage = "English" Then
    '        strPassword = InputBox("Create a new SPC. Please enter administrator password")
    '    End If

    '    Dim Cn As New System.Data.SqlClient.SqlConnection
    '    Dim Adapter As New SqlDataAdapter
    '    Dim strSQL As String = ""
    '    Dim n As Integer
    '    Dim table As New DataTable


    '    Cn.ConnectionString = StrServerConnection
    '    table.Clear()

    '    strSQL = "SELECT *"
    '    strSQL &= " FROM SPC_User"

    '    Adapter = New SqlDataAdapter()
    '    Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
    '    Adapter.SelectCommand.CommandType = CommandType.Text
    '    Adapter.Fill(table)
    '    n = table.Rows.Count

    '    Adapter.Dispose()
    '    Cn.Dispose()
    '    table.Dispose()

    '    Check_Password = False
    '    For i As Integer = 0 To n - 1
    '        If strPassword = table.Rows(i)("cUserID") & " " & table.Rows(i)("cPassword") Then
    '            Check_Password = True
    '            Na = table.Rows(i)("cName")
    '            Exit Function
    '        End If
    '    Next

    'End Function
    Public Function Check_Textdata() As Boolean

        Check_Textdata = True

        Dim eMas As String = ""
        If ComboBox_Process.Text = "" Then
            eMas &= " ・Process not entered." & Environment.NewLine
        End If
        If ComboBox_McNo.Text = "" Then
            eMas &= " ・Equipment No. not entered." & Environment.NewLine
        End If
        If ComboBox_Item.Text = "" Then
            eMas &= " ・Mesure item not entered." & Environment.NewLine
        End If
        If ComboBox_Tree1.Text = "" Then
            eMas &= " ・Tree1 not entered." & Environment.NewLine
        End If
        If ComboBox_Format.Text = "" Then
            eMas &= " ・Graph type not entered." & Environment.NewLine
        End If
        If IsNumeric(Text_Upper.Text) = False Then
            eMas &= " ・No value entered for Upper." & Environment.NewLine
        End If
        If IsNumeric(Text_Lower.Text) = False Then
            eMas &= " ・No value entered for Lower." & Environment.NewLine
        End If
        If IsNumeric(TextX_S.Text) = False Then
            eMas &= " ・No value entered for XBar_σ." & Environment.NewLine
        Else
            If TextX_S.Text <= 0 Then
                eMas &= " ・0 or less entered for XBar_σ." & Environment.NewLine
            End If
        End If
        If IsNumeric(TextX_CL.Text) = False Then
            eMas &= " ・No value entered for XBar_CL." & Environment.NewLine
        End If
        If IsNumeric(TextR_S.Text) = False Then
            eMas &= " ・No value entered for R_σ." & Environment.NewLine
        Else
            If TextR_S.Text <= 0 Then
                eMas &= " ・0 or less entered for R_σ." & Environment.NewLine
            End If
        End If
        If IsNumeric(TextR_CL.Text) = False Then
            eMas &= " ・No value entered for R_CL." & Environment.NewLine
        Else
            If TextR_CL.Text <= 0 Then
                eMas &= " ・0 or less entered for R_CL." & Environment.NewLine
            End If
        End If
        If IsNumeric(TextR_UCL.Text) = False Then
            eMas &= " ・No value entered for R_UCL." & Environment.NewLine
        Else
            If TextR_UCL.Text <= 0 Then
                eMas &= " ・0 or less entered for R_UCL." & Environment.NewLine
            End If
        End If
        If Text_PRODUCT.Text = "" Then
            eMas &= " ・PRODUCT not entered." & Environment.NewLine
        End If
        If IsNumeric(Text_SamQuantity.Text) = False Then
            eMas &= " ・No value entered for Sample Quantity." & Environment.NewLine
        Else
            If Text_SamQuantity.Text <= 0 Then
                eMas &= " ・0 or less entered for Sample Quantity." & Environment.NewLine
            End If
        End If
        If Not eMas = "" Then
            Check_Textdata = False
            MsgBox("<Error>" & Environment.NewLine & eMas)
        End If


    End Function
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        Dim MaxID As String

        If Check_Textdata() = False Then
            Exit Sub
        End If

        UserName = ""
        JP_Message = "管理者ID,パスワードを入力して下さい。"
        EN_Message = "Enter administrator ID and password."
        UserName = FormAlarmInput.Input_Pass_to_Get_UserName(JP_Message, EN_Message, 0, 1) '0:IDPass合ってれば名前取得　1:QC承認者であるかも判定 


        If Not UserName = "" Then

            MaxID = Get_PropertyNo()
            INSERT_PropertyInfo(MaxID, UserName)

            Form1.GetTreeList_Server()

        End If


    End Sub


    '日本語・英語表記の切り替えを行う
    Public Sub Translation_FormControl()

        'If StrLanguage = "Japanese" Then '日本語表記の場合
        '    Label1.Text = "SPC初期設定画面"
        '    Label2.Text = "大項目"
        '    Label3.Text = "Monitor種類"
        '    Label4.Text = "装置名"
        '    Label5.Text = "管理項目"
        '    Label6.Text = "設備・機種"
        '    Label11.Text = "検索条件1"
        '    Label7.Text = "検索条件2"
        '    Label8.Text = "検索条件3"
        '    Label9.Text = "検索条件4"
        '    Label10.Text = "検索条件5"
        '    Label14.Text = "規格センター"
        '    Label15.Text = "規格公差"
        '    Label16.Text = "上限"
        '    Label17.Text = "下限"
        '    Label18.Text = "単位"
        '    Label25.Text = "登録日"

        'ElseIf StrLanguage = "English" Then '英語表記の場合
        '    Label1.Text = "SPC initial setting registration"
        '    Label2.Text = "Top item"
        '    Label3.Text = "Monitor"
        '    Label4.Text = "Machine name"
        '    Label5.Text = "Management item"
        '    Label6.Text = "Machine No/Device"
        '    Label11.Text = "Search item1"
        '    Label7.Text = "Search item2"
        '    Label8.Text = "Search item3"
        '    Label9.Text = "Search item4"
        '    Label10.Text = "Search item5"
        '    Label14.Text = "SCL"
        '    Label15.Text = "Tolerance"
        '    Label16.Text = "USL"
        '    Label17.Text = "LSL"
        '    Label18.Text = "Unit"
        '    Label25.Text = "Display strat date"
        'End If

    End Sub


    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub Panel3_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel3.Paint

    End Sub

    Private Sub ComboBox_Process_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox_Process.SelectedIndexChanged
        Get_McNoInfo(ComboBox_Process.Text)
    End Sub

    Private Sub ComboBox_McNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox_McNo.SelectedIndexChanged
        Get_ControlItemInfo(ComboBox_McNo.Text)
    End Sub

    Private Sub ComboBox_Item_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox_Item.SelectedIndexChanged
        Get_DeviceInfo(ComboBox_McNo.Text, ComboBox_Item.Text)
        Get_FilterInfo(ComboBox_McNo.Text, ComboBox_Item.Text)
    End Sub

    Private Sub ComboBox_Format_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox_Format.TextChanged
        For i As Integer = 0 To UBound(gType, 1)
            If ComboBox_Format.Text = gType(i) Then
                If i = 0 Then
                    Text_Upper.Enabled = True
                    Text_Lower.Enabled = True
                ElseIf i = 1 Then
                    Text_Upper.Enabled = True
                    Text_Lower.Enabled = False
                    Text_Lower.Text = "0"
                ElseIf i = 2 Then
                    Text_Upper.Enabled = False
                    Text_Lower.Enabled = True
                    Text_Upper.Text = "0"
                End If
            End If
        Next
    End Sub

    Private Sub TextX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextX_S.TextChanged, TextX_CL.TextChanged
        UCLLCU("X")
    End Sub
    Private Sub TextR_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextR_S.TextChanged, TextR_CL.TextChanged
        UCLLCU("R")
    End Sub

    Private Sub UCLLCU(ByVal Mode As String)
        If Mode = "X" Then
            If IsNumeric(TextX_S.Text) = True And IsNumeric(TextX_CL.Text) = True Then
                TextX_UCL.Text = TextX_CL.Text + 3 * TextX_S.Text
                TextX_LCL.Text = TextX_CL.Text - 3 * TextX_S.Text
            Else
                TextX_UCL.Text = ""
                TextX_LCL.Text = ""
            End If
        ElseIf Mode = "R" Then
            If IsNumeric(TextR_S.Text) = True And IsNumeric(TextR_CL.Text) = True Then
                TextR_UCL.Text = TextR_CL.Text + 3 * TextR_S.Text
            Else
                TextR_UCL.Text = ""
            End If
        End If
    End Sub

End Class