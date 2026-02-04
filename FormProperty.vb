Public Class FormProperty

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        strStartDate = DateTimePicker1.Value
        strAlarmStartDate = DateTimePicker2.Value
        Form1.SaveConfigData()
        Form1.UPDATE_d_a_StartDate(DateTimePicker1.Value, DateTimePicker2.Value, 1)
        Me.Close()
    End Sub

    Private Sub FormProperty_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = strStartDate
        DateTimePicker2.Value = strAlarmStartDate
        If PropertyTable.Rows.Count = 0 Then
            Button2.Enabled = False
        Else
            Button2.Enabled = True
            DateTimePicker4.Value = PropertyTable.Rows(PropertyTable.Rows.Count - 1)("dStartDate")
            DateTimePicker3.Value = PropertyTable.Rows(PropertyTable.Rows.Count - 1)("aStartDate")
        End If


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Form1.SPCData_Export()
    End Sub


    Private Sub SaveFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        FormControl.Show()
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Form1.UPDATE_d_a_StartDate(DateTimePicker4.Value, DateTimePicker3.Value, 0)     
    End Sub
End Class