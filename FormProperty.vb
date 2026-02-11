Public Class FormProperty

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        strStartDate = DateTimePicker1.Value
        strAlarmStartDate = DateTimePicker2.Value
        Form1.SaveConfigData()
        Form1.UPDATE_d_a_StartDate(DateTimePicker1.Value, DateTimePicker2.Value, 1)
        Me.Close()
    End Sub

    Private Sub FormProperty_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim tempStartDate As DateTime
        Dim tempEndDate As DateTime
        Try
            If DateTime.TryParse(strStartDate, tempStartDate) Then
                If tempStartDate >= DateTimePicker1.MinDate AndAlso tempStartDate <= DateTimePicker1.MaxDate Then
                    DateTimePicker1.Value = tempStartDate
                Else
                    DateTimePicker1.Value = DateTime.Now
                End If
            Else
                DateTimePicker1.Value = DateTime.Now
            End If
            If DateTime.TryParse(strEndDate, tempEndDate) Then
                If tempEndDate >= DateTimePicker2.MinDate AndAlso tempEndDate <= DateTimePicker2.MaxDate Then
                    DateTimePicker2.Value = tempEndDate
                Else
                    DateTimePicker2.Value = DateTime.Now
                End If
            Else
                DateTimePicker2.Value = DateTime.Now
            End If
        Catch ex As Exception

        End Try
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