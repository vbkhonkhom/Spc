Public Class FormSPCRule

    Private Sub FormSPCRule_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If PropertyTable.Rows.Count = 0 Then
            Me.Close()
            Exit Sub
        End If
        'SPCルールを読み込む==============

        '①1点が3σ制限を越える
        If PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cSpcRule1") = True Then
            CheckBox1.Checked = True
        Else
            CheckBox1.Checked = False
        End If
        '②8点連続で片側に出現
        If PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cSpcRule2") = True Then
            CheckBox2.Checked = True
        Else
            CheckBox2.Checked = False
        End If
        '③3点のうち2点が2σ制限を越える
        If PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cSpcRule3") = True Then
            CheckBox3.Checked = True
        Else
            CheckBox3.Checked = False
        End If
        '④5点のうち4点が1σ制限を越える
        If PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cSpcRule4") = True Then
            CheckBox4.Checked = True
        Else
            CheckBox4.Checked = False
        End If
        '⑤15点連続で1σ制限内に出現
        If PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cSpcRule5") = True Then
            CheckBox5.Checked = True
        Else
            CheckBox5.Checked = False
        End If
        '⑥8点連続で1σ制限を越える
        If PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cSpcRule6") = True Then
            CheckBox6.Checked = True
        Else
            CheckBox6.Checked = False
        End If
        '⑦7点連続上昇or下降
        If PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cSpcRule7") = True Then
            CheckBox7.Checked = True
        Else
            CheckBox7.Checked = False
        End If
        '⑧14点連続で交互に上下する
        If PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cSpcRule8") = True Then
            CheckBox8.Checked = True
        Else
            CheckBox8.Checked = False
        End If

        'MR管理図を適用するか
        If PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cMR") = True Then
            CheckBox9.Checked = True
        Else
            CheckBox9.Checked = False
        End If
        '================================

        TextX_CL.Text = PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cXcl")
        TextX_UCL.Text = PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cXucl")
        TextX_LCL.Text = PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cXlcl")
        TextX_S.Text = PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cXdev")

        TextR_CL.Text = PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cRcl")
        TextR_UCL.Text = PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cRucl")
        TextR_LCL.Text = PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cRlcl")
        TextR_S.Text = PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cRdev")

        TextMR_CL.Text = PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cMRcl")
        TextMR_UCL.Text = PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cMRucl")
        TextMR_LCL.Text = PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cMRlcl")
        TextMR_σ.Text = PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cMRdev")


    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        FormLog.Show()
    End Sub

    Public Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Width = 1110
        Button2.Visible = False
        TextX_CL_After.Text = TextX_CL.Text
        TextX_UCL_After.Text = TextX_UCL.Text
        TextX_LCL_After.Text = TextX_LCL.Text
        TextX_S_After.Text = TextX_S.Text

        TextR_CL_After.Text = TextR_CL.Text
        TextR_UCL_After.Text = TextR_UCL.Text
        'TextR_LCL_After.Text = TextR_LCL.Text
        TextR_LCL_After.Text = 0
        TextR_S_After.Text = TextR_S.Text

        TextMR_CL_After.Text = TextMR_CL.Text
        TextMR_UCL_After.Text = TextMR_UCL.Text
        TextMR_LCL_After.Text = TextMR_LCL.Text
        TextMR_S_After.Text = TextMR_σ.Text
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        If Check_Textdata_Update() = False Then
            Exit Sub
        End If

        Form1.UpdateControlLine()
        Me.Width = 850
        Button2.Visible = True
        Me.Close()
    End Sub
    Private Function Check_Textdata_Update() As Boolean

        Check_Textdata_Update = True

        Dim eMas As String = ""
        If IsNumeric(TextX_CL_After.Text) = False Then
            eMas &= " ・No value entered for XBar_CL." & Environment.NewLine
        End If
        If IsNumeric(TextX_S_After.Text) = False Then
            eMas &= " ・No value entered for XBar_σ." & Environment.NewLine
        Else
            If TextX_S_After.Text <= 0 Then
                eMas &= " ・0 or less entered for XBar_σ." & Environment.NewLine
            End If
        End If
        If IsNumeric(TextR_CL_After.Text) = False Then
            eMas &= " ・No value entered for R_CL." & Environment.NewLine
        Else
            If TextR_CL_After.Text <= 0 Then
                eMas &= " ・0 or less entered for R_CL." & Environment.NewLine
            End If
        End If
        If IsNumeric(TextR_S_After.Text) = False Then
            eMas &= " ・No value entered for R_σ." & Environment.NewLine
        Else
            If TextR_S_After.Text <= 0 Then
                eMas &= " ・0 or less entered for R_σ." & Environment.NewLine
            End If
        End If
        If IsNumeric(TextR_UCL_After.Text) = False Then
            eMas &= " ・No value entered for R_UCL." & Environment.NewLine
        Else
            If TextR_UCL_After.Text <= 0 Then
                eMas &= " ・0 or less entered for R_UCL." & Environment.NewLine
            End If
        End If
        If IsNumeric(TextX_UCL_After.Text) = False Then
            eMas &= " ・No value entered for XBar_UCL." & Environment.NewLine
        End If
        If IsNumeric(TextX_LCL_After.Text) = False Then
            eMas &= " ・No value entered for XBar_LCL." & Environment.NewLine
        End If
        If IsNumeric(TextR_CL_After.Text) = False Then
            eMas &= " ・No value entered for RBar_LCL." & Environment.NewLine
        End If

        If Not eMas = "" Then
            Check_Textdata_Update = False
            MsgBox("<Error>" & Environment.NewLine & eMas)
        End If


    End Function

    Private Sub TextX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextX_S_After.TextChanged, TextX_CL_After.TextChanged
        UCLLCU("X")
    End Sub
    Private Sub TextR_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextR_S_After.TextChanged, TextR_CL_After.TextChanged
        UCLLCU("R")
    End Sub

    Private Sub UCLLCU(ByVal Mode As String)
        If Mode = "X" Then
            If IsNumeric(TextX_S_After.Text) = True And IsNumeric(TextX_CL_After.Text) = True Then
                TextX_UCL_After.Text = TextX_CL_After.Text + 3 * TextX_S_After.Text
                TextX_LCL_After.Text = TextX_CL_After.Text - 3 * TextX_S_After.Text
            Else
                TextX_UCL_After.Text = ""
                TextX_LCL_After.Text = ""
            End If
        ElseIf Mode = "R" Then
            If IsNumeric(TextR_S_After.Text) = True And IsNumeric(TextR_CL_After.Text) = True Then
                TextR_UCL_After.Text = TextR_CL_After.Text + 3 * TextR_S_After.Text
            Else
                TextR_UCL_After.Text = ""
            End If
        End If
    End Sub


End Class