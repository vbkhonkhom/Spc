Public Class FormAlarmDisp

    Private Sub FormAlarmDisp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim po As Integer
        'Dim temp() As String
        'Dim strAlm As String
        'Dim uc As Integer
        'Dim AlarmCommentFlag As Boolean
        ''日本語・英語表記の切り替えを行う
        'Translation_AlarmDisp()

        'AlarmCommentFlag = False
        'strAlm = ""
        'po = SerectPoint
        'temp = Split(tmpBufMain(po), ",")
        'uc = UBound(temp)

        'LabDate.Text = SPCDateBuf(po)

        'strAlm = ""
        'If GraphMode = "X" Then
        '    If StrLanguage = "Japanese" Then
        '        If SPCAlarmBuf(po) = "1" Then
        '            strAlm &= "①1点が3σ制限を越える"
        '        ElseIf SPCAlarmBuf(po) = "2" Then
        '            strAlm &= "②8点連続で片側に出現"
        '        ElseIf SPCAlarmBuf(po) = "3" Then
        '            strAlm &= "③3点のうち2点が2σ制限を越える"
        '        ElseIf SPCAlarmBuf(po) = "4" Then
        '            strAlm &= "④5点のうち4点が1σ制限を越える"
        '        ElseIf SPCAlarmBuf(po) = "5" Then
        '            strAlm &= "⑤15点連続で1σ制限内に出現"
        '        ElseIf SPCAlarmBuf(po) = "6" Then
        '            strAlm &= "⑥8点連続で1σ制限を越える"
        '        ElseIf SPCAlarmBuf(po) = "7" Then
        '            strAlm &= "⑦7点連続上昇or下降"
        '        ElseIf SPCAlarmBuf(po) = "8" Then
        '            strAlm &= "⑧14点連続で交互に上下する"
        '        End If
        '    ElseIf StrLanguage = "English" Then '英語表記の場合
        '        If SPCAlarmBuf(po) = "1" Then
        '            strAlm &= "①Any single data point falls outside The 3σ limit from the centerline"
        '        ElseIf SPCAlarmBuf(po) = "2" Then
        '            strAlm &= "②Eight consecutive points fall on the same side of the centerline"
        '        ElseIf SPCAlarmBuf(po) = "3" Then
        '            strAlm &= "③Two out of three consecutive points fall beyond the 2σ limit"
        '        ElseIf SPCAlarmBuf(po) = "4" Then
        '            strAlm &= "④Four out of five consecutive points fall beyond the 1σ limit"
        '        ElseIf SPCAlarmBuf(po) = "5" Then
        '            strAlm &= "⑤Fifteen consective points fall within ±1σ"
        '        ElseIf SPCAlarmBuf(po) = "6" Then
        '            strAlm &= "⑥Eight consective points fall beyond the 1σ limit"
        '        ElseIf SPCAlarmBuf(po) = "7" Then
        '            strAlm &= "⑦Seven consective points fall continuous rise or descent"
        '        ElseIf SPCAlarmBuf(po) = "8" Then
        '            strAlm &= "⑧Fourteen consective points fall alternate up and down"
        '        End If

        '    End If

        '    LabSPCAlarm.Text = strAlm

        '    LabAnother.Text = temp(24)
        '    LabQC.Text = temp(23)
        '    LabCheck.Text = temp(22)
        '    LabAction.Text = temp(21)
        '    LabPerson2.Text = temp(20)
        '    LabResult.Text = temp(19)
        '    LabPerson.Text = temp(18)

        'ElseIf GraphMode = "R" Then
        '    If SPCRAlarmBuf(po) = "1" Then
        '        If StrLanguage = "Japanese" Then
        '            strAlm &= "①1点が3σ制限を越える"
        '        ElseIf StrLanguage = "English" Then '英語表記の場合
        '            strAlm &= "①Any single data point falls outside The 3σ limit from the centerline"
        '        End If

        '    End If

        '    LabSPCAlarm.Text = strAlm

        '    LabAnother.Text = temp(31)
        '    LabQC.Text = temp(30)
        '    LabCheck.Text = temp(29)
        '    LabAction.Text = temp(28)
        '    LabPerson2.Text = temp(27)
        '    LabResult.Text = temp(26)
        '    LabPerson.Text = temp(25)
        'ElseIf GraphMode = "MR" Then
        '    If SPCMRAlarmBuf(po) = "1" Then
        '        If StrLanguage = "Japanese" Then
        '            strAlm &= "①1点が3σ制限を越える"
        '        ElseIf StrLanguage = "English" Then '英語表記の場合
        '            strAlm &= "①Any single data point falls outside The 3σ limit from the centerline"
        '        End If
        '    End If

        '        LabSPCAlarm.Text = strAlm

        '        LabAnother.Text = temp(38)
        '        LabQC.Text = temp(37)
        '        LabCheck.Text = temp(36)
        '        LabAction.Text = temp(35)
        '        LabPerson2.Text = temp(34)
        '        LabResult.Text = temp(33)
        '        LabPerson.Text = temp(32)
        '    End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    '日本語・英語表記の切り替えを行う
    Public Sub Translation_AlarmDisp()

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
        End If

    End Sub
End Class