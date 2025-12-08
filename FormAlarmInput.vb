Imports System.Data.SqlClient
Public Class FormAlarmInput

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        UPDATE_AlarmComment()
        GraphDisp()

        Me.Close()
    End Sub
    '*******************************************************************
    'アラームコメントをアップデートする
    '*******************************************************************
    Public Sub UPDATE_AlarmComment()

        Dim Cn As New SqlConnection
        Dim strSQL As String
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        Dim trans As SqlTransaction 'トランザクション定義
        Dim temp() As String
        Dim Updateflag As Boolean = False

        Dim strID As String = readMaster(M_Data(SerectPoint), _id)



        Try


            Cn.ConnectionString = StrServerConnection

            Cn.Open()
            trans = Cn.BeginTransaction
            SQLCm.Transaction = trans


            strSQL = ""
            strSQL = "UPDATE SPC_Alarm SET "

            If TextPerson.Text <> "" Then
                strSQL &= " cSurveyIncharge= '" & TextPerson.Text & "'"
                Updateflag = True
            End If
            If TextResult.Text <> "" Then
                If Updateflag = True Then
                    strSQL &= ","
                End If
                strSQL &= " cSurveyResult= '" & TextResult.Text & "'"
                Updateflag = True
            End If
            If TextPerson2.Text <> "" Then
                If Updateflag = True Then
                    strSQL &= ","
                End If
                strSQL &= " cTreatIncharge= '" & TextPerson2.Text & "'"
                Updateflag = True
            End If
            If TextAction.Text <> "" Then
                If Updateflag = True Then
                    strSQL &= ","
                End If
                strSQL &= " cTreatResult= '" & TextAction.Text & "'"
                Updateflag = True
            End If
            If TextCheck.Text <> "" Then
                If Updateflag = True Then
                    strSQL &= ","
                End If
                strSQL &= " cTreatEffect= '" & TextCheck.Text & "'"
                Updateflag = True
            End If
            If TextQC.Text <> "" Then
                If Updateflag = True Then
                    strSQL &= ","
                End If
                strSQL &= " cApproverName= '" & TextQC.Text & "'"
                Updateflag = True
            End If
            If TextAnother.Text <> "" Then
                If Updateflag = True Then
                    strSQL &= ","
                End If
                strSQL &= " cMaintenanceID= '" & TextAnother.Text & "'"
                Updateflag = True
            End If

            strSQL &= " WHERE "
            strSQL &= " iID ='" & strID & "'"
            strSQL &= " AND cGraphFormat = '" & TextMode.Text & "'"
            For i As Integer = 0 To UBound(TreeName, 1)
                strSQL &= " AND"
                strSQL &= " cTreeName" & i + 1 & " = '" & TreeName(i) & "'"
            Next

            SQLCm.CommandText = strSQL
            SQLCm.ExecuteNonQuery()

            trans.Commit()
            Cn.Close()

            If Updateflag = True Then
                Dim p As Integer = 0
                If TextMode.Text = "X" Then
                    p = 0
                ElseIf TextMode.Text = "R" Then
                    p = 1
                ElseIf TextMode.Text = "MR" Then
                    p = 2
                End If
                M_Alarm(SerectPoint)(p) = "2" & M_Alarm(SerectPoint)(p).Substring(1)
                If Not TextQC.Text = "" Then
                    M_Alarm(SerectPoint)(p) = "3" & M_Alarm(SerectPoint)(p).Substring(1)
                End If

            End If


        Catch ex As Exception
            If IsNothing(trans) = False Then
                trans.Rollback()
            End If
            StrErrMes = "アラームコメントアップデートエラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Sub
        End Try

    End Sub
    Private Sub FormAlarmInput_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        '日本語・英語表記の切り替えを行う
        Translation_AlarmInput()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click


        UserName = ""
        JP_Message = "QCID、パスワードを入力して下さい。"
        EN_Message = "Enter QCID and Password"
        UserName = Input_Pass_to_Get_UserName(JP_Message, EN_Message, 1, 1) '0:IDPass合ってれば名前取得　1:QC承認者であるかも判定 

        If UserName <> "" Then
            TextQC.Text = UserName
            UPDATE_AlarmComment_QC(UserName)
            GraphDisp()
            Me.Close()
        End If

    End Sub
    '*******************************************************************
    'アラームコメントをアップデートする(QC承認)
    '*******************************************************************
    Public Sub UPDATE_AlarmComment_QC(ByVal User As String)

        Dim Cn As New SqlConnection
        Dim strSQL As String
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        Dim trans As SqlTransaction 'トランザクション定義
        Dim temp() As String
        Dim Updateflag As Boolean = False

        Dim strID As String = readMaster(M_Data(SerectPoint), _id)

        Try


            Cn.ConnectionString = StrServerConnection

            Cn.Open()
            trans = Cn.BeginTransaction
            SQLCm.Transaction = trans


            strSQL = ""
            strSQL = "UPDATE SPC_Alarm SET "
            strSQL &= " cApprovalDate= '" & DateTime.Now & "'"
            strSQL &= ","
            strSQL &= " cApproverName= '" & User & "'"


            strSQL &= " WHERE "
            strSQL &= " iID ='" & strID & "'"
            strSQL &= " AND cGraphFormat = '" & TextMode.Text & "'"
            For i As Integer = 0 To UBound(TreeName, 1)
                strSQL &= " AND"
                strSQL &= " cTreeName" & i + 1 & " = '" & TreeName(i) & "'"
            Next

            SQLCm.CommandText = strSQL
            SQLCm.ExecuteNonQuery()

            trans.Commit()
            Cn.Close()


            Dim p As Integer = 0
            If TextMode.Text = "X" Then
                p = 0
            ElseIf TextMode.Text = "R" Then
                p = 1
            ElseIf TextMode.Text = "MR" Then
                p = 2
            End If
            M_Alarm(SerectPoint)(p) = "3" & M_Alarm(SerectPoint)(p).Substring(1)

        Catch ex As Exception
            If IsNothing(trans) = False Then
                trans.Rollback()
            End If
            StrErrMes = "アラームコメントアップデートエラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Sub
        End Try

    End Sub
    Public Function Input_Pass_to_Get_UserName(ByVal J_Message As String, ByVal E_Message As String, ByVal m As Integer, ByVal l As Integer) As String
        Dim strPassword As String = ""
        Dim temp() As String

        Input_Pass_to_Get_UserName = ""

        If StrLanguage = "Japanese" Then
            strPassword = InputBox(J_Message)
        ElseIf StrLanguage = "English" Then
            strPassword = InputBox(E_Message)
        End If

        If strPassword = "" Then
            MsgBox("Enter the ID and Password")
            Exit Function
        End If

        temp = Split(strPassword, " ") 'ID 半角スペース Pass で入力
        If Not UBound(temp, 1) = 1 Then
            MsgBox("Enter the ID + " & Chr(34) & " " & Chr(34) & " + Password")
            Exit Function
        End If
        '名前が帰ってくればパス照合OK
        Input_Pass_to_Get_UserName = Form1.Check_UserPass(temp(0), temp(1), m, l)
    End Function
    '日本語・英語表記の切り替えを行う
    Public Sub Translation_AlarmInput()

        If StrLanguage = "Japanese" Then '日本語表記の場合
            Label1.Text = "発生日："
            Label3.Text = "異常内容："
            Label6.Text = "調査結果"
            Label7.Text = "担当者"
            Label8.Text = "結果"
            Label12.Text = "処置内容と効果確認"
            Label10.Text = "担当者"
            Label9.Text = "処置内容"
            Label13.Text = "効果確認"
            Label18.Text = "QC承認"
            Label15.Text = "保全記録ID"
            Label16.Text = "調査・処置記入表"
            Button1.Text = "登録"
            Button3.Text = "承認"
        ElseIf StrLanguage = "English" Then '英語表記の場合
            Label1.Text = "Date："
            Label3.Text = "Alarm content："
            Label6.Text = "Survey results"
            Label7.Text = "Surveyor"
            Label8.Text = "Result"
            Label12.Text = "Treatment contents and effect confirmation"
            Label10.Text = "Person"
            Label9.Text = "Action"
            Label13.Text = "Effect"
            Label18.Text = "QC approval"
            Label15.Text = "Maintenance record"
            Label16.Text = "Alarm Comment"
            Button1.Text = "Registration"
            Button3.Text = "Password"
        End If

    End Sub


End Class