Public Class FormLog

    Private Sub FormLog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DispDataGrid()
    End Sub
    Private Sub DispDataGrid()

        'データテーブルを作成()
        Dim tbl As New DataTable
        tbl.Columns.Add(New DataColumn("Date"))
        tbl.Columns.Add(New DataColumn("Change point"))
        tbl.Columns.Add(New DataColumn("Change reason"))
        tbl.Columns.Add(New DataColumn("In charge"))
        tbl.Columns.Add(New DataColumn("QC approval"))
        tbl.Columns.Add(New DataColumn("QC approval date"))
        'tbl.Columns.Add(New DataColumn("Reference material ID"))

        Dim n As Integer = PropertyTable.Rows.Count
        Dim idx As Integer

        For i As Integer = 0 To n - 1
            tbl.Rows.Add()
            tbl.Rows(idx).Item(0) = PropertyTable.Rows(i)("cUpdateDate") '日付
            tbl.Rows(idx).Item(1) = PropertyTable.Rows(i)("cUpdateContent") '変更内容
            tbl.Rows(idx).Item(2) = PropertyTable.Rows(i)("cUpdateReason") '変更理由
            tbl.Rows(idx).Item(3) = PropertyTable.Rows(i)("cIncharge") '変更者
            tbl.Rows(idx).Item(4) = PropertyTable.Rows(i)("cApproverName") 'QC承認者
            tbl.Rows(idx).Item(5) = PropertyTable.Rows(i)("cApprovalDate") 'QC承認日
            idx += 1
        Next

        Me.DataGridView1.DataSource = tbl
        tbl.Dispose()



    End Sub

    Private Sub DataGridView1_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick

        Dim c1 As Integer
        Dim r1 As Integer
        Dim Buf(0) As String

        '現在のセルの列インデックスを表示
        c1 = DataGridView1.CurrentCell.ColumnIndex
        '現在のセルの行インデックスを表示
        r1 = DataGridView1.CurrentCell.RowIndex
        If r1 >= DataGridView1.RowCount - 1 Then
            Exit Sub
        End If

        If c1 = 4 Then 'QC確認欄をクリックした場合
            If StrNetworkFolder = "L:\" Then
                Exit Sub
            End If
            If DataGridView1.Rows(r1).Cells(c1).Value.ToString = "" Then
                UserName = ""
                JP_Message = "QCID、パスワードを入力して下さい。"
                EN_Message = "Enter QCID and Password"
                UserName = FormAlarmInput.Input_Pass_to_Get_UserName(JP_Message, EN_Message, 1, 1) '0:IDPass合ってれば名前取得　1:QC承認者であるかも判定 

                If UserName <> "" Then
                    Form1.Input_Chenged_Property(Buf, "", "Control limit approval", r1)
                    PropertyTable = getProperty()
                    Me.Close()
                End If
            End If
        End If

        Exit Sub

        'Dim dnum As Integer
        'Dim str1 As String
        'Dim dt As Date
        'Dim temp() As String
        'Dim strData As String
        'Dim strSFileName As String
        'Dim strTFileName As String
        'Dim alllen As Integer
        'Dim filelen As Integer



        'dt = Now()
        ''現在のセルの列インデックスを表示
        'c1 = DataGridView1.CurrentCell.ColumnIndex
        ''現在のセルの行インデックスを表示
        'r1 = DataGridView1.CurrentCell.RowIndex
        'If r1 >= DataGridView1.RowCount - 1 Then
        '    Exit Sub
        'End If
        'If c1 = 4 Then 'QC確認欄をクリックした場合
        '    If StrNetworkFolder = "L:\" Then
        '        Exit Sub
        '    End If
        '    If DataGridView1.Rows(r1).Cells(4).Value.ToString = "" Then '
        '        str1 = InputBox("管理者パスワードを入力して下さい")

        '        If str1 = "10020" Then
        '            DataGridView1.Rows(r1).Cells(4).Value = "福田　圭"
        '            DataGridView1.Rows(r1).Cells(5).Value = Format(dt, "yyyy/MM/dd")
        '        ElseIf str1 = "10105" Then
        '            DataGridView1.Rows(r1).Cells(4).Value = "安武　龍太朗"
        '            DataGridView1.Rows(r1).Cells(5).Value = Format(dt, "yyyy/MM/dd")
        '        Else
        '            MsgBox("パスワードに誤りがあります。確認して再度入力して下さい")
        '            Exit Sub
        '        End If

        '        Dim strfilename2 As String = StrRootFolder & StrSerchHostItem & "\" & StrSerchMonitor & "\" & StrSelectMc & "\" & StrSelectMc
        '        dt = Now

        '        strSFileName = strfilename2 & "_Property.csv"
        '        Dim sr1 As New System.IO.StreamReader(strSFileName, System.Text.Encoding.Default)

        '        strTFileName = strfilename2 & "_Temp.csv"
        '        Dim sw1 As New System.IO.StreamWriter(strTFileName, True, System.Text.Encoding.Default)  'Append

        '        'ファイルの最後までループ
        '        Do Until sr1.Peek() = -1
        '            strData = sr1.ReadLine()
        '            If strData = tmpBufMainControl(r1) Then  '
        '                temp = Split(tmpBufMainControl(r1), ",")
        '                strData = ""
        '                dnum = UBound(temp)
        '                For i = 0 To dnum

        '                    If i = 38 Then
        '                        strData &= DataGridView1.Rows(r1).Cells(4).Value.ToString & ","
        '                    ElseIf i = 39 Then
        '                        strData &= DataGridView1.Rows(r1).Cells(5).Value.ToString & ","
        '                    ElseIf i = 40 Then
        '                        strData &= temp(i)
        '                    Else
        '                        strData &= temp(i) & ","
        '                    End If

        '                Next

        '            End If
        '            sw1.WriteLine(strData)
        '        Loop
        '        sr1.Close()
        '        sw1.Close()
        '        If System.IO.File.Exists(strTFileName) Then
        '            System.IO.File.Delete(strSFileName)
        '            System.IO.File.Move(strTFileName, strSFileName)
        '            If StrNetworkFolder <> "" Then
        '                System.IO.File.Copy(strSFileName, StrNetworkFolder & "SPCData\" & StrSerchHostItem & "\" & StrSerchMonitor & "\" & StrSelectMc & "\" & StrSelectMc & "_Property.csv", True)
        '            End If
        '            MsgBox("登録しました")
        '            Form1.GetAlarmData() 'アラームデータを消去する
        '            Form1.TreeView1.Nodes.Clear() 'ツリービューのノードをクリア
        '            Form1.GetHostItemList() 'ツリービュー再描画
        '        End If

        '    End If
        'ElseIf c1 = 6 Then
        '    If DataGridView1.Rows(r1).Cells(6).Value = "" And StrNetworkFolder <> "L:\" Then

        '        Dim openFileDialog1 As New OpenFileDialog

        '        With openFileDialog1
        '            .Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        '            .FilterIndex = 1
        '            .RestoreDirectory = True

        '            If openFileDialog1.ShowDialog() = DialogResult.OK Then

        '                Dim fileName As String = openFileDialog1.FileName    'ファイルのパス
        '                alllen = Len(fileName)
        '                For i = 1 To alllen
        '                    If Mid(fileName, alllen - i, 1) = "\" Then
        '                        filelen = alllen - i
        '                        Exit For
        '                    End If
        '                Next

        '                System.IO.File.Copy(fileName, StrRootFolder & Mid(fileName, filelen + 1), True)

        '                DataGridView1.Rows(r1).Cells(6).Value = Mid(fileName, filelen + 1)


        '                Dim strfilename2 As String = StrRootFolder & StrSerchHostItem & "\" & StrSerchMonitor & "\" & StrSelectMc & "\" & StrSelectMc
        '                dt = Now

        '                strSFileName = strfilename2 & "_Property.csv"
        '                Dim sr1 As New System.IO.StreamReader(strSFileName, System.Text.Encoding.Default)

        '                strTFileName = strfilename2 & "_Temp.csv"
        '                Dim sw1 As New System.IO.StreamWriter(strTFileName, True, System.Text.Encoding.Default)  'Append

        '                'ファイルの最後までループ
        '                Do Until sr1.Peek() = -1
        '                    strData = sr1.ReadLine()
        '                    If strData = tmpBufMainControl(r1) Then  '
        '                        temp = Split(tmpBufMainControl(r1), ",")
        '                        strData = ""
        '                        dnum = UBound(temp)
        '                        For i = 0 To dnum
        '                            If i = 40 Then
        '                                strData &= Mid(fileName, filelen + 1)
        '                            Else
        '                                strData &= temp(i) & ","
        '                            End If
        '                        Next

        '                    End If
        '                    sw1.WriteLine(strData)
        '                Loop
        '                sr1.Close()
        '                sw1.Close()
        '                If System.IO.File.Exists(strTFileName) Then
        '                    System.IO.File.Delete(strSFileName)
        '                    System.IO.File.Move(strTFileName, strSFileName)
        '                    If StrNetworkFolder <> "" Then
        '                        System.IO.File.Copy(strSFileName, StrNetworkFolder & "SPCData\" & StrSerchHostItem & "\" & StrSerchMonitor & "\" & StrSelectMc & "\" & StrSelectMc & "_Property.csv", True)
        '                        System.IO.File.Copy(StrRootFolder & DataGridView1.Rows(r1).Cells(6).Value, StrNetworkFolder & "SPCData\" & DataGridView1.Rows(r1).Cells(6).Value, True)
        '                    End If
        '                    MsgBox("登録しました")
        '                    Form1.GetAlarmData() 'アラームデータを消去する
        '                    Form1.TreeView1.Nodes.Clear() 'ツリービューのノードをクリア
        '                    Form1.GetHostItemList() 'ツリービュー再描画
        '                End If
        '            End If
        '        End With
        '    Else 'すでに証拠残しデータが登録されている場合、そのデータを表示する
        '        If StrNetworkFolder <> "" Then
        '            System.IO.File.Copy(StrNetworkFolder & "SPCData\" & DataGridView1.Rows(r1).Cells(6).Value, StrRootFolder & DataGridView1.Rows(r1).Cells(6).Value, True)
        '        End If
        '        Dim proc As New Process()                 ' （1）
        '        proc.StartInfo.FileName = StrRootFolder & DataGridView1.Rows(r1).Cells(6).Value    ' （2）
        '        proc.Start()

        '    End If
        'End If


    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellPainting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles DataGridView1.CellPainting
        
    End Sub
End Class