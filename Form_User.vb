'Dec.10.2020        Panatsakorn O.      Security Password
Option Strict On                'Dec.10.2020

Imports System.Data.SqlClient

Public Class Form_User
    Dim Old_Password As String = Nothing
    Dim New_Password As String = Nothing

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If TextBox1.Text = "" Then
            MsgBox("UseID not entered.")
            Exit Sub
        End If
        If TextBox1.Text = "Stan_No" Then
            MsgBox("This UseID is not available.")
            Exit Sub
        End If
        If TextBox2.Text = "" Then
            MsgBox("Password not entered.")
            Exit Sub
        End If
        INSERT_UserInfo(TextBox1.Text, TextBox2.Text)
        TextBox1.Text = ""
        TextBox2.Text = ""
        RadioButton1.Checked = False
        RadioButton2.Checked = False
    End Sub

    '*******************************************************************
    'ユーザー情報をサーバーにインサートする
    '*******************************************************************
    Public Sub INSERT_UserInfo(ByVal _UserID As String, ByVal _Password As String)

        Dim Cn As New SqlConnection
        Dim strSQL As String
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        Dim trans As SqlTransaction 'トランザクション定義
        Dim Adapter As New SqlDataAdapter
        Dim table As New DataTable
        Dim n As Integer
        Try

            Cn.ConnectionString = StrServerConnection

            strSQL = "SELECT *"
            strSQL &= " FROM SPC_User"
            strSQL &= " WHERE cUserID = '" & _UserID & "'"

            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(table)
            n = table.Rows.Count

            Adapter.Dispose()
            If Not n = 0 Then
                MsgBox("UserID:" & _UserID & " is already registered.")
                Exit Sub
            End If


            Cn.Open()
            trans = Cn.BeginTransaction
            SQLCm.Transaction = trans

            strSQL = ""
            strSQL = "INSERT INTO SPC_User VALUES ('"
            strSQL &= _UserID & "','"
            strSQL &= _Password & "','"

            If RadioButton1.Checked = True Then
                strSQL &= "True" & "'"
            Else
                strSQL &= "False" & "'"
            End If

            strSQL &= ")"

            SQLCm.CommandText = strSQL
            SQLCm.ExecuteNonQuery()

            trans.Commit()
            Cn.Close()
            MsgBox(_UserID & " registered.")

        Catch ex As Exception
            If IsNothing(trans) = False Then
                trans.Rollback()
            End If
            StrErrMes = "ユーザー情報更新エラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now().ToString, StrErrMes)
            Exit Sub
        End Try

    End Sub


    Private Sub Form_User_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox4.Text = Form1.txtStan_No.Text
        If TextBox4.Text = "No Data" Then
            TextBox4.Text = ""
        End If

        RadioButton2.Checked = True
        Form1.GetUserList()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        'Form1.GetUserInfo(ComboBox1.Text)          'Dec.10.2020
        '>>> Dec.10.2020
        Dim res As String = GetUserInfo2(ComboBox1.SelectedItem.ToString)
        If Not String.IsNullOrEmpty(res) Then
            Old_Password = res
        End If
        '<<< Dec.10.2020
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ComboBox1.Text = "" Then
            MsgBox("UseID not entered.")
            Exit Sub
        End If
        If TextBox3.Text = "" Then
            MsgBox("Old Password not entered.")             'Dec.10.2020
            Exit Sub
        End If

        '>>> Dec.10.2020
        If TextBox5.Text = "" Then
            MsgBox("New Password not entered.")
            Exit Sub
        End If

        If Old_Password = TextBox3.Text Then
            UPDATE_UserInfo(ComboBox1.Text, TextBox5.Text)
            ComboBox1.Text = ""
            RadioButton3.Checked = False
            RadioButton4.Checked = False
            Form1.GetUserList()
        Else
            MsgBox("Password incorrect !!")
        End If
        TextBox3.Text = ""
        TextBox5.Text = ""
        '<<< Dec.10.2020
    End Sub
    '*******************************************************************
    'ユーザー情報をアップデートする
    '*******************************************************************
    Public Sub UPDATE_UserInfo(ByVal _UserID As String, ByVal _Password As String)

        Dim Cn As New SqlConnection
        Dim strSQL As String
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        Dim trans As SqlTransaction 'トランザクション定義
        Dim temp() As String
        Dim Updateflag As Boolean = False

        Try


            Cn.ConnectionString = StrServerConnection

            Cn.Open()
            trans = Cn.BeginTransaction
            SQLCm.Transaction = trans


            strSQL = ""
            strSQL = "UPDATE SPC_User SET "
            strSQL &= " cPassword= '" & _Password & "',"
            If RadioButton4.Checked = True Then
                strSQL &= " bApprover= '" & "True" & "'"
            Else
                strSQL &= " bApprover= '" & "False" & "'"
            End If
            strSQL &= " WHERE "
            strSQL &= " cUserID ='" & _UserID & "'"

            SQLCm.CommandText = strSQL
            SQLCm.ExecuteNonQuery()

            trans.Commit()
            Cn.Close()

            MsgBox(_UserID & " data updated.")

        Catch ex As Exception
            If IsNothing(trans) = False Then
                trans.Rollback()
            End If
            StrErrMes = "ユーザー情報アップデートエラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now().ToString, StrErrMes)
            Exit Sub
        End Try

    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox4.Text = "" Then
            MsgBox("Standard No. not entered.")
            Exit Sub
        End If
        INSERT_StandardNo("Stan_No", TextBox4.Text)
    End Sub
    Public Sub INSERT_StandardNo(ByVal _UserID As String, ByVal _Password As String)


        Dim Cn As New SqlConnection
        Dim strSQL As String
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        Dim trans As SqlTransaction 'トランザクション定義
        Dim Adapter As New SqlDataAdapter
        Dim table As New DataTable
        Dim n As Integer
        Try

            Cn.ConnectionString = StrServerConnection

            strSQL = "SELECT *"
            strSQL &= " FROM SPC_User"
            strSQL &= " WHERE cUserID = '" & _UserID & "'"

            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(table)
            n = table.Rows.Count
            Adapter.Dispose()

            Cn.Open()
            trans = Cn.BeginTransaction
            SQLCm.Transaction = trans

            If Not n = 0 Then '一旦削除
                SQLCm.CommandText = "DELETE FROM SPC_User WHERE cUserID = '" & _UserID & "'"
                SQLCm.ExecuteNonQuery()
            End If



            strSQL = ""
            strSQL = "INSERT INTO SPC_User VALUES ('"
            strSQL &= _UserID & "','"
            strSQL &= _Password & "','"
            strSQL &= "False" & "')"

            SQLCm.CommandText = strSQL
            SQLCm.ExecuteNonQuery()

            trans.Commit()
            Cn.Close()
            Form1.txtStan_No.Text = _Password
            FormMiddle.txtStan_No.Text = _Password
            FormSmall.txtStan_No.Text = _Password
            MsgBox(_Password & " registered.")

        Catch ex As Exception
            If IsNothing(trans) = False Then
                trans.Rollback()
            End If
            StrErrMes = "ユーザー情報更新エラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now().ToString, StrErrMes)
            Exit Sub
        End Try

    End Sub
    '>>> Dec.10.2020
    Private Function GetUserInfo2(ByVal _UserID As String) As String

        Using con As New SqlConnection(StrServerConnection)
            Dim strSQL As String = "SELECT [cUserID],[cPassword],[bApprover] FROM SPC_User WHERE cUserID = @User"
            Try
                con.Open()
            Catch ex As Exception
                Call SaveLog(Now().ToString, "Cannot Open Connection !!, " + ex.Message & ex.StackTrace)
                Return Nothing
            End Try

            Using cmd As New SqlCommand(strSQL, con)
                cmd.Parameters.Add("User", SqlDbType.VarChar).Value = _UserID

                Using read As SqlDataReader = cmd.ExecuteReader
                    If read.HasRows Then
                        Dim table As New DataTable
                        table.Load(read)
                        If table.Rows.Count > 0 Then
                            If CBool(table.Rows(0)("bApprover")) Then
                                RadioButton3.Checked = False
                                RadioButton4.Checked = True
                            Else
                                RadioButton3.Checked = True
                                RadioButton4.Checked = False
                            End If

                            Return table.Rows(0)("cPassword").ToString
                        Else
                            Call SaveLog(Now().ToString, "Not Found User [row is zero] !!, ")
                            Exit Function
                        End If
                    Else
                        Call SaveLog(Now().ToString, "Not Found User !!, ")
                        Exit Function
                    End If
                End Using
            End Using
        End Using
    End Function
    '<<< Dec.10.2020
End Class