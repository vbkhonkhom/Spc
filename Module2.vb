Imports System.Data.SqlClient
Imports System.Draw
Module Module2
    '??????????
    Dim p_iID As Integer = 0
    Dim p_dWorkDate As Integer = 1
    Dim p_cUsl As Integer = 2
    Dim p_cLsl As Integer = 3
    Dim p_cXcl As Integer = 4
    Dim p_cXucl As Integer = 5
    Dim p_cXlcl As Integer = 6
    Dim p_cXdev As Integer = 7
    Dim p_cRucl As Integer = 8
    Dim p_cRlcl As Integer = 9
    Dim p_cMRucl As Integer = 10
    Dim p_cMR As Integer = 11
    Dim p_cSpcRule0 As Integer = 12
    Dim p_cSpcRule1 As Integer = 13
    Dim p_cSpcRule2 As Integer = 14
    Dim p_cSpcRule3 As Integer = 15
    Dim p_cSpcRule4 As Integer = 16
    Dim p_cSpcRule5 As Integer = 17
    Dim p_cSpcRule6 As Integer = 18
    Dim p_cSpcRule7 As Integer = 19




    Public Function readMaster(ByVal _data As String, ByVal _p As Integer) As String
        If String.IsNullOrEmpty(_data) Then Return ""
        Dim temp() As String = Split(_data, ",")
        If _p >= 0 AndAlso _p < temp.Length Then
            Return temp(_p)
        Else
            Return ""
        End If

    End Function

    Public Function treeInfo(ByVal _TreeView As TreeView) As String()

        Dim n As Integer '????????
        Dim TreeNodeBuf As TreeNode = Nothing
        Dim TreeNameBuf(10 - 1) As String
        Dim TreeName(10 - 1) As String
        For i As Integer = 0 To UBound(TreeNameBuf, 1)
            TreeNameBuf(i) = ""
            TreeName(i) = ""
        Next
        '???????????????
        For i As Integer = 0 To UBound(TreeNameBuf, 1)
            If i = 0 Then
                TreeNodeBuf = _TreeView.SelectedNode
                TreeNameBuf(i) = TreeNodeBuf.Text
                n += 1
            Else
                TreeNodeBuf = TreeNodeBuf.Parent
                If TreeNodeBuf Is Nothing Then
                    Exit For
                End If
                TreeNameBuf(i) = TreeNodeBuf.Text
                n += 1
            End If
        Next
        '??????????
        For i As Integer = 0 To n - 1
            TreeName(i) = TreeNameBuf(n - 1 - i)
        Next

        Return TreeName
    End Function

    '???????????????????
    Public Function getProperty() As DataTable

        Dim _pTable As New DataTable
        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim Adapter As New SqlDataAdapter
        Dim strSQL As String = ""

        getProperty = Nothing

        Try

            Cn.ConnectionString = StrServerConnection
            _pTable.Clear()

            strSQL = "SELECT *"
            strSQL &= " FROM SPC_Property"
            For i As Integer = 0 To UBound(TreeName, 1)
                If i = 0 Then
                    strSQL &= " WHERE"
                Else
                    strSQL &= " AND"
                End If

                strSQL &= " cTreeName" & i + 1 & " = '" & TreeName(i) & "'"
            Next


            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(_pTable)


            Adapter.Dispose()
            Cn.Dispose()
            _pTable.Dispose()

            If _pTable.Rows.Count = 0 Then
                Return Nothing
            End If



            QCNotCheckFlag = True

            For k = 0 To _pTable.Rows.Count - 1
                If IsDBNull(_pTable.Rows(k)("cApprovalDate")) Then 'QC????????
                    QCNotCheckFlag = False
                End If
            Next


            Return _pTable

        Catch ex As System.Exception
            Adapter.Dispose()
            Cn.Dispose()

            StrErrMes = "?????????????" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
        End Try
    End Function
    Public Function setFilter(ByVal _col As String) As String
        Dim _str As String = ""

        If (Not IsDBNull(PropertyTable.Rows(PropertyTable.Rows.Count - 1)(_col))) Then
            If PropertyTable.Rows(PropertyTable.Rows.Count - 1)(_col) <> "" Then
                _str = " AND " & _col & " = '" & PropertyTable.Rows(PropertyTable.Rows.Count - 1)(_col) & "'"
            End If
        End If

        Return _str
    End Function
    Public Function getData() As DataTable

        getData = Nothing

        Dim _sTable As New DataTable


        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim Adapter As New SqlDataAdapter
        Dim strSQL As String = ""
        Dim n As Integer

        Try

            Cn.ConnectionString = StrServerConnection
            _sTable.Clear()

            strSQL = "SELECT *"
            strSQL &= " FROM SPC_Master"
            strSQL &= " WHERE"
            strSQL &= " cMachineNo = '" & PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cMachineNo") & "'"
            strSQL &= " AND cControlItem = '" & PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cControlItem") & "'"

            strSQL &= setFilter("cFilter_1")
            strSQL &= setFilter("cFilter_2")
            strSQL &= setFilter("cFilter_3")
            strSQL &= setFilter("cFilter_4")
            strSQL &= setFilter("cFilter_5")
            strSQL &= setFilter("cFilter_6")
            strSQL &= setFilter("cFilter_7")
            strSQL &= setFilter("cFilter_8")
            strSQL &= setFilter("cFilter_9")
            strSQL &= setFilter("cFilter_10")
            strSQL &= setFilter("cDeviceName")

            'strSQL &= "ORDER BY dWorkDate,iID"
            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(_sTable)
            n = _sTable.Rows.Count

            Adapter.Dispose()
            Cn.Dispose()
            If n = 0 Then
                _sTable.Dispose()
                Return Nothing
            End If

            'dWorkDate?????
            Dim dv = New DataView(_sTable)
            dv.Sort = "dWorkDate,iID"
            _sTable = dv.ToTable
            n = _sTable.Rows.Count
            Return _sTable

        Catch ex As System.Exception
            Adapter.Dispose()
            Cn.Dispose()

            StrErrMes = "?????????????" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
        End Try

    End Function
    Public Sub txtclear()

        Dim str As String = "No Data"
        If Form1.Visible = True Then
            Form1.TextItem.Text = str
            Form1.PictureBox1.Image = Nothing
            Form1.PictureBox2.Image = Nothing
            Form1.PictureBox4.Image = Nothing
            Form1.PictureBox9.Image = Nothing
            Form1.LabUnit.Text = "Unit"
            Form1.LabUpCpk.Text = "0"
            Form1.LabLoCpk.Text = "0"
            Form1.LabelBase.Text = "Standard"
            Form1.TextUCL.Text = "0"
            Form1.TextCL.Text = "0"
            Form1.TextLCL.Text = "0"
            Form1.TextSiguma.Text = "0"
            Form1.TextRUCL.Text = "0"
            Form1.TextRCL.Text = "0"
            Form1.TextRSiguma.Text = "0"
            Form1.LabelQC.Visible = False
        ElseIf FormMiddle.Visible = True Then
            FormMiddle.TextItem.Text = str
            FormMiddle.PictureBox1.Image = Nothing
            FormMiddle.PictureBox2.Image = Nothing
            FormMiddle.PictureBox4.Image = Nothing
            FormMiddle.PictureBox9.Image = Nothing
            FormMiddle.LabUnit.Text = "Unit"
            FormMiddle.LabUpCpk.Text = "0"
            FormMiddle.LabLoCpk.Text = "0"
            FormMiddle.LabelBase.Text = "Standard"
            FormMiddle.TextUCL.Text = "0"
            FormMiddle.TextCL.Text = "0"
            FormMiddle.TextLCL.Text = "0"
            FormMiddle.TextSiguma.Text = "0"
            FormMiddle.TextRUCL.Text = "0"
            FormMiddle.TextRCL.Text = "0"
            FormMiddle.TextRSiguma.Text = "0"
            FormMiddle.LabelQC.Visible = False
        ElseIf FormSmall.Visible = True Then
            FormSmall.TextItem.Text = str
            FormSmall.PictureBox1.Image = Nothing
            FormSmall.PictureBox2.Image = Nothing
            FormSmall.PictureBox4.Image = Nothing
            FormSmall.PictureBox9.Image = Nothing
            FormSmall.LabUnit.Text = "Unit"
            FormSmall.LabUpCpk.Text = "0"
            FormSmall.LabLoCpk.Text = "0"
            FormSmall.LabelBase.Text = "Standard"
            FormSmall.TextUCL.Text = "0"
            FormSmall.TextCL.Text = "0"
            FormSmall.TextLCL.Text = "0"
            FormSmall.TextSiguma.Text = "0"
            FormSmall.TextRUCL.Text = "0"
            FormSmall.TextRCL.Text = "0"
            FormSmall.TextRSiguma.Text = "0"
            FormSmall.LabelQC.Visible = False
        End If

    End Sub


    'SPC????????
    Public Function getSPCMaster() As String()
        Dim _Master(vv) As String
        Graphsmallcount = 1
        SPCDataNum = 0

        '??????????
        Dim SpcTable As New DataTable
        SpcTable = getData()

        If SpcTable Is Nothing Then
            txtclear()
            Return Nothing
        End If


        Dim c As Integer
        Dim strvalue As String

        For i = 0 To 31
            xpnbuf_X(i) = 0
            ypnbuf_X(i) = 0
            xpnbuf_R(i) = 0
            ypnbuf_R(i) = 0
        Next
        For i = 0 To vv
            _Master(i) = ""

        Next

        c = 0
        Dim s As Integer = 0
        If SpcTable.Rows.Count > vv Then
            s = SpcTable.Rows.Count - vv + 1
        End If
        Dim syosu As Integer
        Dim temp() As String

        temp = Split(PropertyTable.Rows(PropertyTable.Rows.Count - 1)("cXcl"), ".")
        If temp.Length = 1 Then
            syosu = 1
        Else
            syosu = Len(temp(1)) + 1
        End If

        Dim mBuf(_cate) As String


        For i = s To SpcTable.Rows.Count - 1
            If Not PropertyTable.Rows(PropertyTable.Rows.Count - 1)("dStartDate") <= SpcTable.Rows(i)("dWorkDate") Then
                Continue For
            End If

            For j As Integer = 0 To UBound(mBuf, 1)
                mBuf(j) = ""
            Next

            mBuf(_id) = SpcTable.Rows(i)("iID")
            mBuf(_wDate) = SpcTable.Rows(i)("dWorkDate")

            If (Not IsDBNull(SpcTable.Rows(i)("cValue1"))) Then
                If IsNumeric(SpcTable.Rows(i)("cValue1")) = True Then
                    temp = Split(SpcTable.Rows(i)("cValue1"), ".")
                    If temp.Length = 1 Then
                        syosu = 0
                    Else
                        syosu = Len(temp(1))
                    End If
                End If
            End If
            mBuf(_X) = Math.Round(CDbl(SpcTable.Rows(i)("cXbar")), syosu, MidpointRounding.AwayFromZero)
            If (Not IsDBNull(SpcTable.Rows(i)("cR"))) Then
                mBuf(_R) = SpcTable.Rows(i)("cR")  'R
            End If
            If (Not IsDBNull(SpcTable.Rows(i)("cLotNo"))) Then
                mBuf(_lot) = SpcTable.Rows(i)("cLotNo") '???No
            End If
            If (Not IsDBNull(SpcTable.Rows(i)("cInCharge"))) Then
                mBuf(_opName) = SpcTable.Rows(i)("cInCharge") '???
            End If
            If (Not IsDBNull(SpcTable.Rows(i)("cCategory"))) Then
                mBuf(_cate) = SpcTable.Rows(i)("cCategory") '????
            End If

            If c > 0 Then
                mBuf(_MR) = Math.Round(Math.Abs(CDbl(mBuf(_X)) - CDbl(readMaster(_Master(c - 1), _X))), syosu, MidpointRounding.AwayFromZero) '???
            ElseIf c = 0 Then
                mBuf(_MR) = 0
            End If

            strvalue = ""
            For j = 1 To 100
                If IsDBNull(SpcTable.Rows(i)("cValue" & j)) = True Then
                    Continue For
                End If
                If Not strvalue = "" Then
                    strvalue &= ","
                End If

                strvalue &= SpcTable.Rows(i)("cValue" & j)


            Next
            MesureValueBuf(c) = strvalue

            Dim str As String = ""
            For j As Integer = 0 To UBound(mBuf, 1)
                If Not j = 0 Then
                    str &= ","
                End If
                str &= mBuf(j)
            Next

            _Master(c) = str

            c += 1

        Next

        SPCDataNum = c '???????
        If c >= 30 * Graphsmallcount Then
            DispStartPosition = SPCDataNum - (30 * Graphsmallcount)
        Else
            DispStartPosition = 0
        End If

        SpcTable.Dispose()
        Return _Master

        'Catch ex As System.Exception
        '    Adapter.Dispose()
        '    Cn.Dispose()
        '    SpcTable.Dispose()

        '    StrErrMes = "SPC????????" + ", " + ex.Message & ex.StackTrace
        '    Call SaveLog(Now(), StrErrMes)
        '    Exit Sub
        'End Try

    End Function

    'SPC????????
    Public Function getAlarmMaster() As String()()


        Dim _Master(vv)() As String
        For i As Integer = 0 To UBound(_Master, 1)
            ReDim _Master(i)(3 - 1)
            For j As Integer = 0 To UBound(_Master(i), 1)
                _Master(i)(j) = "0,00000000" '??????,?????? j=0:X,1:R,2:MR
            Next
        Next
        getAlarmMaster = _Master

        Dim AlarmTable As New DataTable
        AlarmTable = getAlarm()

        If AlarmTable Is Nothing Then
            Return Nothing
        End If
        If AlarmTable.Rows.Count = 0 Then
            Return _Master
        End If

        For i As Integer = 0 To AlarmTable.Rows.Count - 1

            If IsDBNull(AlarmTable.Rows(i)("iID")) = True Then Continue For
            Dim a_id As String = AlarmTable.Rows(i)("iID")

            If IsDBNull(AlarmTable.Rows(i)("cGraphFormat")) = True Then Continue For

            Dim p As Integer = 0
            If AlarmTable.Rows(i)("cGraphFormat") = "X" Then
                p = 0
            ElseIf AlarmTable.Rows(i)("cGraphFormat") = "R" Then
                p = 1
            ElseIf AlarmTable.Rows(i)("cGraphFormat") = "MR" Then
                p = 2
            Else
                Continue For
            End If

            Dim come As Integer = 1 '??????????
            Dim CommentFlag As Boolean = False
            If (Not IsDBNull(AlarmTable.Rows(i)("cSurveyIncharge"))) Then '??????????????
                If AlarmTable.Rows(i)("cSurveyIncharge") <> "" Then
                    come = 2 '???????
                End If
            End If
            If (Not IsDBNull(AlarmTable.Rows(i)("cSurveyResult"))) Then '??????????????
                If AlarmTable.Rows(i)("cSurveyResult") <> "" Then
                    come = 2 '???????
                End If
            End If
            If (Not IsDBNull(AlarmTable.Rows(i)("cTreatIncharge"))) Then '??????????????
                If AlarmTable.Rows(i)("cTreatIncharge") <> "" Then
                    come = 2 '???????
                End If
            End If
            If (Not IsDBNull(AlarmTable.Rows(i)("cTreatResult"))) Then '??????????????
                If AlarmTable.Rows(i)("cTreatResult") <> "" Then
                    come = 2 '???????
                End If
            End If
            If (Not IsDBNull(AlarmTable.Rows(i)("cTreatEffect"))) Then '??????????????
                If AlarmTable.Rows(i)("cTreatEffect") <> "" Then
                    come = 2 '???????
                End If
            End If
            If (Not IsDBNull(AlarmTable.Rows(i)("cMaintenanceID"))) Then '??????????????
                If AlarmTable.Rows(i)("cMaintenanceID") <> "" Then
                    come = 2 '???????
                End If
            End If
            If (Not IsDBNull(AlarmTable.Rows(i)("cApproverName"))) Then '??????????
                If AlarmTable.Rows(i)("cApproverName") <> "" Then
                    come = 3 'QC????
                End If
            End If

            Dim naiyou As String = ""
            For j As Integer = 0 To UBound(M_Data, 1)
                Dim strId As String = readMaster(M_Data(j), _id)

                If Not a_id = strId Then Continue For

                For k As Integer = 1 To 8
                    Dim _s As String = ""
                    If (Not IsDBNull(AlarmTable.Rows(i)("cSpcrule" & k))) Then
                        If CBool(AlarmTable.Rows(i)("cSpcrule" & k)) Then
                            _s = "1"
                        Else
                            _s = "0"
                        End If

                    End If
                    naiyou &= _s
                Next

                _Master(j)(p) = come & "," & naiyou

                Exit For
            Next

        Next

        Return _Master

    End Function


    'SPC??????????
    Public Sub GetAlarmData_kai()


        Dim aDataBuf(,) As String

        '????????
        Dim AlarmResult(8 - 1) As Boolean


        aDataBuf = getAlarmbuf()

        For j = 0 To SPCDataNum - 1

            For i As Integer = 0 To UBound(AlarmResult, 1)
                AlarmResult(i) = False
            Next


            'If PropertyTable.Rows(PropertyTable.Rows.Count - 1)("aStartDate") <= CDate(aDataBuf(j, p_dWorkDate)) Then 'strAlarmStartDate???????????


            If CBool(aDataBuf(j, p_cSpcRule0)) Then AlarmResult(0) = checkSpcRule0(aDataBuf, j, _X) '(SPC???1) 1??3???????
            If CBool(aDataBuf(j, p_cSpcRule1)) Then AlarmResult(1) = checkSpcRule1(aDataBuf, j, _X) '(SPC???2) 8?????????
            If CBool(aDataBuf(j, p_cSpcRule2)) Then AlarmResult(2) = checkSpcRule2(aDataBuf, j, _X) '(SPC???3) 3????2??2???????
            If CBool(aDataBuf(j, p_cSpcRule3)) Then AlarmResult(3) = checkSpcRule3(aDataBuf, j, _X) '(SPC???4) 5????4??1???????
            If CBool(aDataBuf(j, p_cSpcRule4)) Then AlarmResult(4) = checkSpcRule4(aDataBuf, j, _X) '(SPC???5) 15????1???????
            If CBool(aDataBuf(j, p_cSpcRule5)) Then AlarmResult(5) = checkSpcRule5(aDataBuf, j, _X) '(SPC???6) 8????1???????
            If CBool(aDataBuf(j, p_cSpcRule6)) Then AlarmResult(6) = checkSpcRule6(aDataBuf, j, _X) '(SPC???7) 7?????or??
            If CBool(aDataBuf(j, p_cSpcRule7)) Then AlarmResult(7) = checkSpcRule7(aDataBuf, j, _X) '(SPC???8) 14???????????



            For i As Integer = 0 To UBound(AlarmResult, 1)
                If AlarmResult(i) = False Then Continue For

                Dim strId As String = readMaster(M_Data(j), _id)
                If GetAlarmInfo_Server_kai(strId, "X") = True Then Continue For
                INSERT_AlarmInfo_kai(strId, "X", AlarmResult)
                Exit For
            Next

            'SPCAlarmBuf(j) = CStr(AlarmResult1) & "," & CStr(AlarmResult2) & "," & CStr(AlarmResult3) & "," & CStr(AlarmResult4) & "," & CStr(AlarmResult5) & "," & CStr(AlarmResult6) & "," & CStr(AlarmResult7) & "," & CStr(AlarmResult8)



            'R?????3?????
            For i As Integer = 0 To UBound(AlarmResult, 1)
                AlarmResult(i) = False
            Next

            AlarmResult(0) = checkSpcRule0(aDataBuf, j, _R) '(SPC???1) 1??3???????


            For i As Integer = 0 To UBound(AlarmResult, 1)
                If AlarmResult(i) = False Then Continue For

                Dim strId As String = readMaster(M_Data(j), _id)
                If GetAlarmInfo_Server_kai(strId, "R") = True Then Continue For
                INSERT_AlarmInfo_kai(strId, "R", AlarmResult)
                Exit For
            Next


            'SPCRAlarmBuf(j) = CStr(AlarmResult1) & "," & CStr(AlarmResult2) & "," & CStr(AlarmResult3) & "," & CStr(AlarmResult4) & "," & CStr(AlarmResult5) & "," & CStr(AlarmResult6) & "," & CStr(AlarmResult7) & "," & CStr(AlarmResult8)


        Next


    End Sub
    Private Function getAlarmbuf() As String(,)
        getAlarmbuf = Nothing

        Dim _aDataBuf(SPCDataNum - 1, 20 - 1) As String

        For j = 0 To SPCDataNum - 1
            Dim p As Integer
            p = 0
            For k = 0 To PropertyTable.Rows.Count - 1
                If IsDBNull(PropertyTable.Rows(k)("cApprovalDate")) = False Then
                    If readMaster(M_Data(j), _wDate) > PropertyTable.Rows(k)("cApprovalDate") Then '1?????????????Propaty?????????????????
                        p = k
                    End If
                End If
            Next

            _aDataBuf(j, p_iID) = p
            _aDataBuf(j, p_dWorkDate) = readMaster(M_Data(j), _wDate)
            _aDataBuf(j, p_cUsl) = PropertyTable.Rows(p)("cUsl")
            _aDataBuf(j, p_cLsl) = PropertyTable.Rows(p)("cLsl")
            _aDataBuf(j, p_cXcl) = PropertyTable.Rows(p)("cXcl")
            _aDataBuf(j, p_cXucl) = PropertyTable.Rows(p)("cXucl")
            _aDataBuf(j, p_cXlcl) = PropertyTable.Rows(p)("cXlcl")
            _aDataBuf(j, p_cXdev) = PropertyTable.Rows(p)("cXdev")
            If (Not IsDBNull(PropertyTable.Rows(p)("cRucl"))) Then
                _aDataBuf(j, p_cRucl) = PropertyTable.Rows(p)("cRucl")
            Else
                _aDataBuf(j, p_cRucl) = ""
            End If
            'If (Not IsDBNull(PropertyTable.Rows(p)("cRlcl"))) Then
            '    _aDataBuf(j, p_cRlcl) = PropertyTable.Rows(p)("cRlcl")
            'Else
            '    _aDataBuf(j, p_cRlcl) = ""
            'End If
            _aDataBuf(j, p_cRlcl) = "0"
            If (Not IsDBNull(PropertyTable.Rows(p)("cMRucl"))) Then
                _aDataBuf(j, p_cMRucl) = PropertyTable.Rows(p)("cMRucl")
            Else
                _aDataBuf(j, p_cMRucl) = ""
            End If
            _aDataBuf(j, p_cMR) = PropertyTable.Rows(p)("cMR")
            _aDataBuf(j, p_cSpcRule0) = PropertyTable.Rows(p)("cSpcRule1")
            _aDataBuf(j, p_cSpcRule1) = PropertyTable.Rows(p)("cSpcRule2")
            _aDataBuf(j, p_cSpcRule2) = PropertyTable.Rows(p)("cSpcRule3")
            _aDataBuf(j, p_cSpcRule3) = PropertyTable.Rows(p)("cSpcRule4")
            _aDataBuf(j, p_cSpcRule4) = PropertyTable.Rows(p)("cSpcRule5")
            _aDataBuf(j, p_cSpcRule5) = PropertyTable.Rows(p)("cSpcRule6")
            _aDataBuf(j, p_cSpcRule6) = PropertyTable.Rows(p)("cSpcRule7")
            _aDataBuf(j, p_cSpcRule7) = PropertyTable.Rows(p)("cSpcRule8")
        Next

        Return _aDataBuf

    End Function
    Private Function checkSpcRule0(ByVal aDataBuf(,) As String, ByVal j As Integer, ByVal mode As Integer) As Boolean
        checkSpcRule0 = False
        Dim atai As Single = CSng(readMaster(M_Data(j), mode))
        Dim _lcl As Integer = 0
        Dim _ucl As Integer = 0

        If mode = _X Then
            _lcl = p_cXlcl
            _ucl = p_cXucl
        ElseIf mode = _R Then
            _lcl = p_cRlcl
            _ucl = p_cRucl
        End If

        If atai < CSng(aDataBuf(j, _lcl)) Or CSng(aDataBuf(j, _ucl)) < atai Then
            Return True
        End If

    End Function
    Private Function checkData(ByVal aDataBuf(,) As String, ByVal j As Integer, ByVal p As Integer) As Boolean

        checkData = True

        If j + 1 < p Then '????????(numP-1)??????
            Return False
        End If

        For k As Integer = 1 To p - 1 'numP???cl,ucl,lcl,dev????(???numP?????)
            If Not (aDataBuf(j, p_cXcl) = aDataBuf(j - k, p_cXcl) And aDataBuf(j, p_cXucl) = aDataBuf(j - k, p_cXucl) And aDataBuf(j, p_cXlcl) = aDataBuf(j - k, p_cXlcl) And aDataBuf(j, p_cXdev) = aDataBuf(j - k, p_cXdev)) Then
                Return False '?????????
            End If
        Next

    End Function

    Private Function checkSpcRule1(ByVal aDataBuf(,) As String, ByVal j As Integer, ByVal mode As Integer) As Boolean
        checkSpcRule1 = False

        Dim numP As Integer = 8 '????

        '???? ??????????
        If checkData(aDataBuf, j, numP) = False Then Return False

        '???????????????????????????8?
        Dim cl As Single = CSng(aDataBuf(j, p_cXcl))
        Dim atai(numP - 1) As Single
        For i As Integer = 0 To UBound(atai, 1)
            atai(i) = CSng(readMaster(M_Data(j - i), mode)) - cl
            If atai(i) = 0 Then
                Return False '?????????F
            End If
        Next

        Dim c_m As Integer = 0 '??????
        Dim c_p As Integer = 0 '?????
        For i As Integer = 0 To UBound(atai, 1)
            If atai(i) < 0 Then
                c_m += 1
            Else
                c_p += 1
            End If
        Next
        If c_m = numP Or c_p = numP Then '?????????????
            Return True
        End If

    End Function
    Private Function checkSpcRule2(ByVal aDataBuf(,) As String, ByVal j As Integer, ByVal mode As Integer) As Boolean
        checkSpcRule2 = False

        Dim numP As Integer = 3 '????

        '???? ??????????
        If checkData(aDataBuf, j, numP) = False Then Return False

        Dim cl As Single = CSng(aDataBuf(j, p_cXcl))
        Dim dev As Single = CSng(aDataBuf(j, p_cXdev))

        Dim atai2siguma_p(numP - 1) As Single '?-(cl+2*siguma)
        Dim atai2siguma_m(numP - 1) As Single '?-(cl-2*siguma)
        For i As Integer = 0 To UBound(atai2siguma_p, 1)
            atai2siguma_p(i) = CSng(readMaster(M_Data(j - i), mode)) - (cl + 2 * dev)
            atai2siguma_m(i) = CSng(readMaster(M_Data(j - i), mode)) - (cl - 2 * dev)
        Next

        Dim c As Integer = 0
        For i As Integer = 0 To UBound(atai2siguma_p, 1)
            If 0 < atai2siguma_p(i) Then
                c += 1 'cl+2*siguma???????
            End If
        Next
        If 2 <= c Then
            Return True
        End If
        c = 0
        For i As Integer = 0 To UBound(atai2siguma_p, 1)
            If atai2siguma_m(i) < 0 Then
                c += 1 'cl-2*siguma???????
            End If
        Next
        If 2 <= c Then
            Return True
        End If



    End Function
    Private Function checkSpcRule3(ByVal aDataBuf(,) As String, ByVal j As Integer, ByVal mode As Integer) As Boolean
        checkSpcRule3 = False

        Dim numP As Integer = 5 '????

        '???? ??????????
        If checkData(aDataBuf, j, numP) = False Then Return False

        Dim cl As Single = CSng(aDataBuf(j, p_cXcl))
        Dim dev As Single = CSng(aDataBuf(j, p_cXdev))


        Dim atai2siguma_p(numP - 1) As Single '?-(cl+siguma)
        Dim atai2siguma_m(numP - 1) As Single '?-(cl-siguma)
        For i As Integer = 0 To UBound(atai2siguma_p, 1)
            atai2siguma_p(i) = CSng(readMaster(M_Data(j - i), mode)) - (cl + dev)
            atai2siguma_m(i) = CSng(readMaster(M_Data(j - i), mode)) - (cl - dev)
        Next

        Dim c As Integer = 0
        For i As Integer = 0 To UBound(atai2siguma_p, 1)
            If 0 < atai2siguma_p(i) Then
                c += 1 'cl+siguma???????
            End If
        Next
        If 4 <= c Then
            Return True
        End If
        c = 0
        For i As Integer = 0 To UBound(atai2siguma_p, 1)
            If atai2siguma_m(i) < 0 Then
                c += 1 'cl-siguma???????
            End If
        Next
        If 4 <= c Then
            Return True
        End If


    End Function
    Private Function checkSpcRule4(ByVal aDataBuf(,) As String, ByVal j As Integer, ByVal mode As Integer) As Boolean
        checkSpcRule4 = False

        Dim numP As Integer = 15 '????

        '???? ??????????
        If checkData(aDataBuf, j, numP) = False Then Return False

        Dim cl As Single = CSng(aDataBuf(j, p_cXcl))
        Dim dev As Single = CSng(aDataBuf(j, p_cXdev))

        Dim atai(numP - 1) As Single
        For i As Integer = 0 To UBound(atai, 1)
            atai(i) = CSng(readMaster(M_Data(j - i), mode))
        Next


        'cl-?~cl+??????????????????????(False)
        For i As Integer = 0 To UBound(atai, 1)
            If atai(i) < cl - dev Then
                Return False
            End If
            If cl + dev < atai(i) Then
                Return False
            End If
        Next
        '?????????cl-?~cl+??????(???????)(True)
        Return True


    End Function
    Private Function checkSpcRule5(ByVal aDataBuf(,) As String, ByVal j As Integer, ByVal mode As Integer) As Boolean
        checkSpcRule5 = False

        Dim numP As Integer = 8 '????

        '???? ??????????
        If checkData(aDataBuf, j, numP) = False Then Return False

        Dim cl As Single = CSng(aDataBuf(j, p_cXcl))
        Dim dev As Single = CSng(aDataBuf(j, p_cXdev))

        Dim atai2siguma_p(numP - 1) As Single '?-(cl+siguma)
        Dim atai2siguma_m(numP - 1) As Single '?-(cl-siguma)
        For i As Integer = 0 To UBound(atai2siguma_p, 1)
            atai2siguma_p(i) = CSng(readMaster(M_Data(j - i), mode)) - (cl + dev)
            atai2siguma_m(i) = CSng(readMaster(M_Data(j - i), mode)) - (cl - dev)
        Next

        Dim c_p As Integer = 0
        Dim c_m As Integer = 0
        For i As Integer = 0 To UBound(atai2siguma_p, 1)
            If 0 < atai2siguma_p(i) Then
                c_p += 1 'cl+siguma???????
            End If
        Next

        For i As Integer = 0 To UBound(atai2siguma_p, 1)
            If atai2siguma_m(i) < 0 Then
                c_m += 1 'cl-siguma???????
            End If
        Next

        '??cl-?~cl+?????????????????(True)
        If c_p = 0 Then Return False
        If c_m = 0 Then Return False

        If c_p + c_m = numP Then
            Return True
        End If


    End Function
    Private Function checkSpcRule6(ByVal aDataBuf(,) As String, ByVal j As Integer, ByVal mode As Integer) As Boolean
        checkSpcRule6 = False

        Dim numP As Integer = 7 '????

        '???? ??????????
        If checkData(aDataBuf, j, numP) = False Then Return False

        Dim cl As Single = CSng(aDataBuf(j, p_cXcl))
        Dim dev As Single = CSng(aDataBuf(j, p_cXdev))

        Dim atai(numP - 1) As Single
        For i As Integer = 0 To UBound(atai, 1)
            atai(i) = CSng(readMaster(M_Data(j - i), mode))
        Next

        Dim atai_sa(UBound(atai, 1) - 1) As Single
        For i As Integer = 0 To UBound(atai, 1) - 1
            atai_sa(i) = atai(i) - atai(i + 1)
            If atai_sa(i) = 0 Then
                Return False '???????F
            End If
        Next


        Dim c As Integer = 0
        For i As Integer = 0 To UBound(atai_sa, 1)
            If 0 < atai_sa(i) Then
                c += 1 '????
            ElseIf atai_sa(i) < 0 Then
                c -= 1 '????
            End If
        Next

        If Math.Abs(c) = numP - 1 Then '??????????????????
            Return True
        End If


    End Function
    Private Function checkSpcRule7(ByVal aDataBuf(,) As String, ByVal j As Integer, ByVal mode As Integer) As Boolean
        checkSpcRule7 = False

        Dim numP As Integer = 14 '????

        '???? ??????????
        If checkData(aDataBuf, j, numP) = False Then Return False

        Dim cl As Single = CSng(aDataBuf(j, p_cXcl))
        Dim dev As Single = CSng(aDataBuf(j, p_cXdev))

        Dim atai(numP - 1) As Single
        For i As Integer = 0 To UBound(atai, 1)
            atai(i) = CSng(readMaster(M_Data(j - i), mode))
        Next

        Dim atai_sa(UBound(atai, 1) - 1) As Single
        For i As Integer = 0 To UBound(atai, 1) - 1
            atai_sa(i) = atai(i) - atai(i + 1)
            If atai_sa(i) = 0 Then
                Return False '???????F
            End If
        Next


        Dim c As Integer = 0
        Dim f_sa As Integer = 0
        For i As Integer = 0 To UBound(atai_sa, 1)
            If 0 < atai_sa(i) Then
                c += 1 '????
            ElseIf atai_sa(i) < 0 Then
                c -= 1 '????
            End If

            If i = 0 Then
                f_sa = c
            End If

            '???????????c ? (1 0 1 0 1 0) ? (-1 0 -1 0 -1 0)   ??????????????????(false)
            If i Mod 2 = 0 Then
                If Not c = f_sa Then Return False
            ElseIf i Mod 2 = 1 Then
                If Not c = 0 Then Return False
            End If

        Next

        Return True


    End Function
    '??????????????????
    Public Function GetAlarmInfo_Server_kai(ByVal _ID As String, ByVal _Mode As String) As Boolean
        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim Adapter As New SqlDataAdapter
        Dim strSQL As String = ""
        Dim table As New DataTable
        Try
            GetAlarmInfo_Server_kai = False

            Cn.ConnectionString = StrServerConnection
            table.Clear()

            strSQL = "SELECT iID"
            strSQL &= " FROM SPC_Alarm"
            strSQL &= " WHERE iID = '" & _ID & "'"

            For i As Integer = 0 To UBound(TreeName, 1)
                strSQL &= " AND"
                strSQL &= " cTreeName" & i + 1 & " = '" & TreeName(i) & "'"
            Next

            strSQL &= " AND [cGraphFormat] = '" & _Mode & "'"
            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(table)

            Adapter.Dispose()
            Cn.Dispose()
            table.Dispose()

            If table.Rows.Count > 0 Then
                Return True
            End If

        Catch ex As System.Exception
            Adapter.Dispose()
            Cn.Dispose()
            table.Dispose()

            StrErrMes = "????????????" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Function
        End Try
    End Function
    '??????????????????
    Public Function getAlarm() As DataTable
        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim Adapter As New SqlDataAdapter
        Dim strSQL As String = ""
        Dim _atable As New DataTable
        getAlarm = Nothing
        Try


            Cn.ConnectionString = StrServerConnection
            _atable.Clear()

            strSQL = "SELECT *"
            strSQL &= " FROM SPC_Alarm"

            For i As Integer = 0 To UBound(TreeName, 1)
                If i = 0 Then
                    strSQL &= " WHERE"
                Else
                    strSQL &= " AND"
                End If

                strSQL &= " cTreeName" & i + 1 & " = '" & TreeName(i) & "'"
            Next
            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(_atable)

            Adapter.Dispose()
            Cn.Dispose()
            _atable.Dispose()

            Return _atable


        Catch ex As System.Exception
            Adapter.Dispose()
            Cn.Dispose()
            _atable.Dispose()

            StrErrMes = "????????????" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Function
        End Try
    End Function
    '*******************************************************************
    'SPC????????????????????
    '*******************************************************************
    Public Sub INSERT_AlarmInfo_kai(ByVal _ID As String, ByVal _Mode As String, ByVal _Alarm() As Boolean)

        Dim Cn As New SqlConnection
        Dim strSQL As String
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        Dim trans As SqlTransaction '??????????

        Try

            Cn.ConnectionString = StrServerConnection

            Cn.Open()
            trans = Cn.BeginTransaction
            SQLCm.Transaction = trans

            strSQL = ""
            strSQL = "INSERT INTO SPC_Alarm VALUES ("
            strSQL &= "'" & _ID & "'," 'iID
            strSQL &= "'" & _Mode & "'," 'cGraphFormat

            For i As Integer = 0 To UBound(TreeName, 1)
                strSQL &= "'" & TreeName(i) & "'," 'cTreeName1
            Next

            For i As Integer = 0 To UBound(_Alarm, 1)
                strSQL &= "'" & CStr(_Alarm(i)) & "'," 'cSpcrule1-8

            Next

            'strSQL &= "'" & CStr(_Alarm1) & "'," 'cSpcrule1
            'strSQL &= "'" & CStr(_Alarm2) & "'," 'cSpcrule2
            'strSQL &= "'" & CStr(_Alarm3) & "'," 'cSpcrule3
            'strSQL &= "'" & CStr(_Alarm4) & "'," 'cSpcrule4
            'strSQL &= "'" & CStr(_Alarm5) & "'," 'cSpcrule5
            'strSQL &= "'" & CStr(_Alarm6) & "'," 'cSpcrule6
            'strSQL &= "'" & CStr(_Alarm7) & "'," 'cSpcrule7
            'strSQL &= "'" & CStr(_Alarm8) & "'," 'cSpcrule8

            strSQL &= "''," 'cSurveyIncharge
            strSQL &= "''," 'cSurveyResult
            strSQL &= "''," 'cTreatIncharge
            strSQL &= "''," 'cTreatResult
            strSQL &= "''," 'cTreatEffect
            strSQL &= "''," 'cApprovalDate
            strSQL &= "''," 'cApproverName
            strSQL &= "'')" 'cMaintenanceID


            SQLCm.CommandText = strSQL
            SQLCm.ExecuteNonQuery()

            trans.Commit()
            Cn.Close()

        Catch ex As Exception
            If IsNothing(trans) = False Then
                trans.Rollback()
            End If
            StrErrMes = "SPC????????????" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Sub
        End Try

    End Sub
    '????????
    Public Sub GraphDisp()

        'X????????????
        GraphDisp1(StrResolution)

        'R????????????
        GraphDisp2(StrResolution, MRFlag)
        '?????????????
        GraphDisp4(StrResolution)
        '???????????
        GraphDisp7(StrResolution)


    End Sub

    Public Sub GraphDisp1(ByVal Size As String)


        Dim xpn, ypn, xp, n, yp, yh, i, j, pno, jk, yps, ypf, ypa, k, p As Integer
        Dim Bairitu, yp0 As Double
        Dim f As New Font("MS P????", 10)
        Dim dbl1, dbl2, dblLow, Data1, DataMAX, DataMIN As Double
        Dim strData As String
        Dim colbuf(5000) As Integer
        Dim xp_old, yp_old, null_bit, end_bit As Integer
        Dim ul, ll As Integer
        Dim strUcl, strStep As String
        Dim strLcl As String
        Dim strCl As String

        Dim Hsum As Double '???
        Dim Bnum As Double '??
        Dim Siguma As Double '????
        Dim sum, ave As Double '?
        Dim Jsum As Double '????
        Dim c As Integer
        Dim UpperCpk, LowerCpk As Double

        '????????=============================================
        Dim g As Graphics
        If Size = "MAX" Then
            With Form1.PictureBox1
                .Image = New Bitmap(1355, 746)
                g = Graphics.FromImage(.Image)
            End With
        ElseIf Size = "Middle" Then
            With FormMiddle.PictureBox1
                .Image = New Bitmap(1355, 746)
                g = Graphics.FromImage(.Image)
            End With
        ElseIf Size = "MIN" Then
            With FormSmall.PictureBox1
                .Image = New Bitmap(1355, 746)
                g = Graphics.FromImage(.Image)
            End With
        End If


        Dim APen As New Pen(Color.Green, 2)
        APen.DashStyle = Drawing2D.DashStyle.Dot
        Dim BPen As New Pen(Color.Black, 1)
        BPen.DashStyle = Drawing2D.DashStyle.Dot
        Dim CPen As New Pen(Color.Green, 2)
        CPen.DashStyle = Drawing2D.DashStyle.Solid

        '????????
        Dim DPen As New Pen(Color.Blue, 2)
        DPen.DashStyle = Drawing2D.DashStyle.Solid
        Dim EPen As New Pen(Color.Green, 1)
        EPen.DashStyle = Drawing2D.DashStyle.Solid
        Dim FPen As New Pen(Color.Red, 2)
        FPen.DashStyle = Drawing2D.DashStyle.Dash
        Dim GPen As New Pen(Color.Red, 2)
        GPen.DashStyle = Drawing2D.DashStyle.Solid
        Dim HPen As New Pen(Color.Red, 3)
        HPen.DashStyle = Drawing2D.DashStyle.Solid
        Dim IPen As New Pen(Color.Blue, 1)
        IPen.DashStyle = Drawing2D.DashStyle.Solid
        Dim JPen As New Pen(Color.Blue, 1)
        JPen.DashStyle = Drawing2D.DashStyle.Dot
        Dim KPen As New Pen(Color.Purple, 2)
        KPen.DashStyle = Drawing2D.DashStyle.Dash

        Dim c1 As New SolidBrush(Color.FromArgb(255, 255, 38, 38))  'Red
        Dim c2 As New SolidBrush(Color.FromArgb(255, 238, 228, 255))  '
        Dim c3 As New SolidBrush(Color.FromArgb(255, 255, 250, 55))  '3
        Dim c4 As New SolidBrush(Color.Blue)
        Dim A1Pen As New Pen(Color.Red, 1)
        A1Pen.DashStyle = Drawing2D.DashStyle.Solid
        Dim A2Pen As New Pen(Color.Green, 1)
        A2Pen.DashStyle = Drawing2D.DashStyle.Solid
        Dim B1Pen As New Pen(Color.Black, 2)
        B1Pen.DashStyle = Drawing2D.DashStyle.Solid

        '================================================================


        If Size = "MAX" Then

            Form1.LabUpCpk.Text = ""
            Form1.LabLoCpk.Text = ""
            Form1.TextUCL.Text = ""
            Form1.TextCL.Text = ""
            Form1.TextLCL.Text = ""
            Form1.TextSiguma.Text = ""
            Form1.LabelBase.Text = ""

        ElseIf Size = "Middle" Then

            FormMiddle.LabUpCpk.Text = ""
            FormMiddle.LabLoCpk.Text = ""
            FormMiddle.TextUCL.Text = ""
            FormMiddle.TextCL.Text = ""
            FormMiddle.TextLCL.Text = ""
            FormMiddle.TextSiguma.Text = ""
            FormMiddle.LabelBase.Text = ""

        ElseIf Size = "MIN" Then

            FormSmall.LabUpCpk.Text = ""
            FormSmall.LabLoCpk.Text = ""
            FormSmall.TextUCL.Text = ""
            FormSmall.TextCL.Text = ""
            FormSmall.TextLCL.Text = ""
            FormSmall.TextSiguma.Text = ""
            FormSmall.LabelBase.Text = ""

        End If

        xp = 0
        yp0 = 0
        yp = 0

        If Size = "MAX" Then
            yh = 428 '???????????????
        ElseIf Size = "Middle" Then
            yh = 362  '???????????????
        ElseIf Size = "MIN" Then
            yh = 260 '???????????????
        End If


        jk = DispStartPosition  '?????????????????
        PropertyNo = PropertyTable.Rows.Count - 1

        '???????????========================
        dbl1 = PropertyTable.Rows(PropertyNo)("cScl")
        dbl2 = PropertyTable.Rows(PropertyNo)("cTolerance") / 5
        If IsDBNull(PropertyTable.Rows(PropertyNo)("cUnit")) = False Then
            strUnit = PropertyTable.Rows(PropertyNo)("cUnit")
        End If

        Dim strStanUR As String = ""
        If PropertyTable.Rows(PropertyNo)("cLimitType") = "UpperLower" Then
            strStanUR = PropertyTable.Rows(PropertyNo)("cLsl") & " - " & PropertyTable.Rows(PropertyNo)("cUsl") & " " & strUnit
        ElseIf PropertyTable.Rows(PropertyNo)("cLimitType") = "Upper" Then
            strStanUR = "<= " & PropertyTable.Rows(PropertyNo)("cUsl") & " " & strUnit
        ElseIf PropertyTable.Rows(PropertyNo)("cLimitType") = "Lower" Then
            strStanUR = ">= " & PropertyTable.Rows(PropertyNo)("cLsl") & " " & strUnit 'Ver2.11 change cUsl cLsl
        End If


        If Size = "MAX" Then
            '?????
            Form1.LabelBase.Text = strStanUR
        ElseIf Size = "Middle" Then
            '?????
            FormMiddle.LabelBase.Text = strStanUR
        ElseIf Size = "MIN" Then
            '?????
            FormSmall.LabelBase.Text = strStanUR
        End If

        strStep = CStr(dbl2)      'STEP

        If Size = "MAX" Then
            Bairitu = 40 / dbl2       '?40Pix
        ElseIf Size = "Middle" Then
            Bairitu = 35 / dbl2       '?35Pix
        ElseIf Size = "MIN" Then
            Bairitu = 25 / dbl2       '?25Pix
        End If

        dblLow = dbl1 - dbl2 * 5

        null_bit = 1
        sum = 0
        c = 0
        Jsum = 0

        Dim x00 As Integer
        Dim x01 As Integer
        Dim x02 As Integer
        Dim x03 As Integer

        For j = 0 To 29 * Graphsmallcount
            Try
                If jk < 0 OrElse jk > UBound(M_Data) Then
                    jk += 1
                    Continue For
                End If
                If String.IsNullOrEmpty(M_Data(jk)) Then
                    jk += 1
                    Continue For
                End If
                strData = readMaster(M_Data(jk), _X)
                If strData = "" Then
                    jk += 1
                    Continue For
                End If

                p = 0

                For k = 0 To PropertyNo
                    If IsDBNull(PropertyTable.Rows(k)("cApprovalDate")) = False Then
                        If readMaster(M_Data(jk), _wDate) > PropertyTable.Rows(k)("cApprovalDate") Then
                            p = k
                        End If
                    End If
                    'jk += 1

                Next

                X_USL = PropertyTable.Rows(p)("cUsl")
                X_LSL = PropertyTable.Rows(p)("cLsl")
                X_CL = PropertyTable.Rows(p)("cXcl")
                X_UCL = PropertyTable.Rows(p)("cXucl")
                X_LCL = PropertyTable.Rows(p)("cXlcl")
                X_Shiguma = PropertyTable.Rows(p)("cXdev")
                X_SCL = PropertyTable.Rows(p)("cScl")
                X_kousa = PropertyTable.Rows(p)("cTolerance")
                X_gType = PropertyTable.Rows(p)("cLimitType")
                If Size = "MAX" Then
                    'xpn = xp + j * (30 / Graphsmallcount) + 15 + 120
                    xpn = xp + j * (30 / Graphsmallcount) + 15
                    x00 = 15
                ElseIf Size = "Middle" Then
                    xpn = xp + j * (25 / Graphsmallcount) + 10
                    x00 = 13
                ElseIf Size = "MIN" Then
                    xpn = xp + j * (20 / Graphsmallcount) + 10
                    x00 = 10
                End If

                If X_gType = "UpperLower" Then
                    yp0 = ((X_USL + X_LSL) / 2 - X_SCL)
                    yp = yp0 * Bairitu
                End If

                'Upper:?????Lower:????
                '??????????
                If X_gType <> "Lower" Then

                    'If Size = "MAX" Then
                    '    ypn = 28
                    'ElseIf Size = "Middle" Then
                    '    ypn = 12
                    'ElseIf Size = "MIN" Then
                    '    ypn = 10
                    'End If
                    ypn = yp + yh - (X_USL - dblLow) * Bairitu
                    g.DrawLine(HPen, xpn - x00, ypn, xpn + x00, ypn)
                End If
                If X_gType <> "Upper" Then
                    '??????????

                    'If Size = "MAX" Then
                    '    ypn = 428
                    'ElseIf Size = "Middle" Then
                    '    ypn = 362
                    'ElseIf Size = "MIN" Then
                    '    ypn = 42
                    'End If
                    ypn = yp + yh - (X_LSL - dblLow) * Bairitu
                    g.DrawLine(HPen, xpn - x00, ypn, xpn + x00, ypn)
                End If

                '------------   CL,UCL,LCL????   -------------

                'CL?????????
                Data1 = (X_CL - dblLow) * Bairitu
                ypn = yp + yh - Data1

                g.DrawLine(APen, xpn - x00, ypn, xpn + x00, ypn)
                ycl = ypn
                'UCL?????????
                Data1 = (X_UCL - dblLow) * Bairitu
                ypn = yp + yh - Data1

                g.DrawLine(FPen, xpn - x00, ypn, xpn + x00, ypn)
                yucl = ypn
                'LCL?????????
                Data1 = (X_LCL - dblLow) * Bairitu
                ypn = yp + yh - Data1
                g.DrawLine(FPen, xpn - x00, ypn, xpn + x00, ypn)
                ylcl = ypn
                '??????????Data1
                Data1 = (CDbl(strData) - dblLow) * Bairitu


                'Cpk?????????????????
                sum += CDbl(strData)
                Jsum += CDbl(strData) * CDbl(strData)

                c += 1


                ypn = yp + yh - Data1


                '??????????0
                colbuf(j) = 0
                '???????1
                Dim ala As Integer = readMaster(M_Alarm(jk)(0), 0)


                If ala = 1 Then
                    colbuf(j) = 1
                End If

                If ala = 2 Or ala = 3 Then
                    colbuf(j) = 1
                    ypf = ypn

                    ypa = ypf - 50
                    If ypa < 0 Then
                        ypa = 50
                        ypf = ypa + 50
                    End If
                    If ypa > 400 Then
                        ypa = 400
                        ypf = ypa + 50
                    End If

                    g.DrawLine(B1Pen, xpn, ypf - 50, xpn, ypa + 23) '????

                    If Size = "MAX" Then
                        x01 = 15
                    ElseIf Size = "Middle" Then
                        x01 = 15
                    ElseIf Size = "MIN" Then
                        x01 = 15
                    End If

                    k = 0
                    For ii = 0 To 7 '?(??)
                        If ala = 3 Then
                            g.DrawLine(A2Pen, xpn + k, ypf - 50 + ii, xpn + k, ypa + x01 - ii)
                        ElseIf ala = 2 Then
                            g.DrawLine(A1Pen, xpn + k, ypf - 50 + ii, xpn + k, ypa + x01 - ii)
                        End If


                        k += 1
                    Next
                End If




                xpnbuf_X(j) = xpn
                ypnbuf_X(j) = ypn
                If j > 0 And null_bit = 0 Then

                    g.DrawLine(DPen, xp_old, yp_old, xpn, ypn) '?????

                End If

                null_bit = 0
                xp_old = xpn
                yp_old = ypn
                end_bit = 1

                jk += 1

            Catch ex As Exception
                jk += 1
                Continue For
            End Try

        Next

        Hsum = Jsum - (sum * sum) / c
        Bnum = Hsum / (c - 1)
        Siguma = Math.Sqrt(Bnum)
        ave = sum / c


        If X_gType = "Lower" Then
            UpperCpk = 0
            LowerCpk = (ave - X_LSL) / (3 * Siguma)
        ElseIf X_gType = "Upper" Then
            UpperCpk = (X_USL - ave) / (3 * Siguma)
            LowerCpk = 0
        Else
            UpperCpk = (X_USL - ave) / (3 * Siguma)
            LowerCpk = (ave - X_LSL) / (3 * Siguma)
        End If

        If Size = "MAX" Then
            Form1.LabUpCpk.Text = Mid(UpperCpk, 1, 4)
            Form1.LabLoCpk.Text = Mid(LowerCpk, 1, 4)
        ElseIf Size = "Middle" Then
            FormMiddle.LabUpCpk.Text = Mid(UpperCpk, 1, 4)
            FormMiddle.LabLoCpk.Text = Mid(LowerCpk, 1, 4)
        ElseIf Size = "MIN" Then
            FormSmall.LabUpCpk.Text = Mid(UpperCpk, 1, 4)
            FormSmall.LabLoCpk.Text = Mid(LowerCpk, 1, 4)
        End If


        '???????????===============================
        If Size = "MAX" Then

            strUcl = CStr(X_UCL)
            Form1.labUCL.Top = yucl - 7 + Form1.PictureBox1.Top
            Form1.TextUCL.Text = strUcl

            strCl = CStr(X_CL)
            Form1.TextCL.Text = strCl
            Form1.labCL.Top = ycl - 7 + Form1.PictureBox1.Top

            strLcl = CStr(X_LCL)
            Form1.labLCL.Top = ylcl - 7 + Form1.PictureBox1.Top
            Form1.TextLCL.Text = strLcl

            Form1.TextSiguma.Text = CStr(X_Shiguma)

        ElseIf Size = "Middle" Then

            strUcl = CStr(X_UCL)
            FormMiddle.labUCL.Top = yucl - 7 + FormMiddle.PictureBox1.Top
            FormMiddle.TextUCL.Text = strUcl

            strCl = CStr(X_CL)
            FormMiddle.TextCL.Text = strCl
            FormMiddle.labCL.Top = ycl - 7 + FormMiddle.PictureBox1.Top

            strLcl = CStr(X_LCL)
            FormMiddle.labLCL.Top = ylcl - 7 + FormMiddle.PictureBox1.Top
            FormMiddle.TextLCL.Text = strLcl

            FormMiddle.TextSiguma.Text = CStr(X_Shiguma)

        ElseIf Size = "MIN" Then

            strUcl = CStr(X_UCL)
            FormSmall.labUCL.Top = yucl - 7 + FormSmall.PictureBox1.Top
            FormSmall.TextUCL.Text = strUcl

            strCl = CStr(X_CL)
            FormSmall.TextCL.Text = strCl
            FormSmall.labCL.Top = ycl - 7 + FormSmall.PictureBox1.Top

            strLcl = CStr(X_LCL)
            FormSmall.labLCL.Top = ylcl - 7 + FormSmall.PictureBox1.Top
            FormSmall.TextLCL.Text = strLcl

            FormSmall.TextSiguma.Text = CStr(X_Shiguma)

        End If
        '=====================================================

        '??????????===============================================
        n = j '????????????
        jk = DispStartPosition
        For j = 0 To n - 1
            xpn = xpnbuf_X(j)
            ypn = ypnbuf_X(j)
            If xpn <> "0" Then
                If readMaster(M_Data(jk), _cate) <> "R" Then
                    If colbuf(j) = 1 Then 'SPC?????????????????
                        g.FillEllipse(c1, xpn - 4, ypn - 4, 7, 7)
                        g.DrawEllipse(EPen, xpn - 4, ypn - 4, 7, 7)
                    Else 'SPC?????????????????
                        g.FillEllipse(c2, xpn - 4, ypn - 4, 7, 7)
                        g.DrawEllipse(EPen, xpn - 4, ypn - 4, 7, 7)
                    End If
                Else '?????????????
                    g.FillEllipse(c4, xpn - 4, ypn - 4, 7, 7)
                    g.DrawEllipse(EPen, xpn - 4, ypn - 4, 7, 7)
                End If
            End If


            jk += 1
        Next
        '====================================================================

        '????????????============================
        dbl1 = X_SCL       '?????
        dbl2 = X_kousa / 5       'STEP
        dbl1 = dbl1 + dbl2 * 5 + yp0
        For i = 0 To 10
            If Size = "MAX" Then

                If dbl1 > 999 Then
                    Form1.LabXBar(i).Text = Format(dbl1, "0")
                Else
                    Form1.LabXBar(i).Text = Format(dbl1, "0.00")
                End If

            ElseIf Size = "Middle" Then

                If dbl1 > 999 Then
                    FormMiddle.LabXBar_Middle(i).Text = Format(dbl1, "0")
                Else
                    FormMiddle.LabXBar_Middle(i).Text = Format(dbl1, "0.00")
                End If

            ElseIf Size = "MIN" Then

                If dbl1 > 999 Then
                    FormSmall.LabXBar_Small(i).Text = Format(dbl1, "0")
                Else
                    FormSmall.LabXBar_Small(i).Text = Format(dbl1, "0.00")
                End If

            End If

            dbl1 -= dbl2
        Next
        '====================================================
        '??????????==================================================
        '??

        If Size = "MAX" Then
            x02 = 30
            x03 = 40
        ElseIf Size = "Middle" Then
            x02 = 25
            x03 = 35
        ElseIf Size = "MIN" Then
            x02 = 20
            x03 = 25
        End If

        For j = 1 To 35
            g.DrawLine(BPen, xp + j * x02, 0, xp + j * x02, 500) '??
        Next
        '??
        If X_Shiguma <> 0 Then '????????????1????
            yps = yp + yh - (CDbl(X_CL) - dblLow) * Bairitu
            For j = 0 To 200
                g.DrawLine(BPen, xp, CInt(yps + j * X_Shiguma / strStep * x03), xp + 36 * 30, CInt(yps + j * X_Shiguma / strStep * x03))
                g.DrawLine(BPen, xp, CInt(yps - j * X_Shiguma / strStep * x03), xp + 36 * 30, CInt(yps - j * X_Shiguma / strStep * x03))
            Next

        Else '????????????40?????????
            For j = 1 To 15
                g.DrawLine(BPen, xp, yp + j * x03, xp + 36 * 30, yp + j * x03)
            Next
        End If
        '======================================================================

        c1.Dispose()
        c2.Dispose()
        c3.Dispose()
        f.Dispose()
        g.Dispose()

        Dim str As String = ""
        For l As Integer = 0 To UBound(TreeName, 1)
            If TreeName(l) = "" Then Exit For
            If Not str = "" Then
                str &= "  "
            End If
            str &= TreeName(l)
        Next


        If Size = "MAX" Then
            Form1.TextItem.Text = str
            Form1.labTitle.Text = PropertyTable.Rows(PropertyNo)("cMachineNo") & "  " & PropertyTable.Rows(PropertyNo)("cControlItem") & " " & "Control chart"
        ElseIf Size = "Middle" Then
            FormMiddle.TextItem.Text = str
            FormMiddle.labTitle.Text = PropertyTable.Rows(PropertyNo)("cMachineNo") & "  " & PropertyTable.Rows(PropertyNo)("cControlItem") & " " & "Control chart"
        ElseIf Size = "MIN" Then
            FormSmall.TextItem.Text = str
            FormSmall.labTitle.Text = PropertyTable.Rows(PropertyNo)("cMachineNo") & "  " & PropertyTable.Rows(PropertyNo)("cControlItem") & " " & "Control chart"
        End If

    End Sub
    'R??????
    Public Sub GraphDisp2(ByVal Size As String, ByVal MR As Boolean)

        Dim xpn, ypn, ypf, ypa, xp, yp, yh, i, p As Integer
        Dim Bairitu As Double
        Dim f As New Font("MS P????", 10)
        Dim dbl1, dbl2, dblLow, Data1 As Double
        Dim strData As String
        Dim colbuf(5000) As Integer
        Dim xp_old, yp_old, null_bit, end_bit As Integer
        Dim UclChangeFlag As Boolean
        Dim UclChangeFlag2 As Boolean
        Dim dblSiguma, dblCl As Double
        Dim yuclR, yclR, k As Integer
        '????????=============================================
        Dim g As Graphics

        If Size = "MAX" Then
            With Form1.PictureBox2
                .Image = New Bitmap(1050, 275)
                g = Graphics.FromImage(.Image)
            End With
        ElseIf Size = "Middle" Then
            With FormMiddle.PictureBox2
                .Image = New Bitmap(1050, 275)
                g = Graphics.FromImage(.Image)
            End With
        ElseIf Size = "MIN" Then
            With FormSmall.PictureBox2
                .Image = New Bitmap(1050, 275)
                g = Graphics.FromImage(.Image)
            End With
        End If



        Dim APen As New Pen(Color.Green, 2)
        APen.DashStyle = Drawing2D.DashStyle.Dot
        Dim BPen As New Pen(Color.Black, 1)
        BPen.DashStyle = Drawing2D.DashStyle.Dot
        Dim CPen As New Pen(Color.Green, 2)
        CPen.DashStyle = Drawing2D.DashStyle.Solid
        Dim DPen As New Pen(Color.DarkOliveGreen, 2)
        DPen.DashStyle = Drawing2D.DashStyle.Solid
        Dim EPen As New Pen(Color.Red, 1.5)
        EPen.DashStyle = Drawing2D.DashStyle.Solid
        Dim D2Pen As New Pen(Color.Cyan, 2)
        D2Pen.DashStyle = Drawing2D.DashStyle.Solid
        Dim E2Pen As New Pen(Color.Black, 1.5)
        E2Pen.DashStyle = Drawing2D.DashStyle.Solid
        Dim FPen As New Pen(Color.Red, 2)
        FPen.DashStyle = Drawing2D.DashStyle.Dash
        Dim HPen As New Pen(Color.Red, 3)
        HPen.DashStyle = Drawing2D.DashStyle.Solid
        Dim c1 As New SolidBrush(Color.FromArgb(255, 255, 38, 38))  'Red
        Dim c2 As New SolidBrush(Color.FromArgb(255, 235, 253, 0))  'Green
        Dim c3 As New SolidBrush(Color.FromArgb(255, 50, 200, 50))  '3
        Dim A1Pen As New Pen(Color.Red, 1)
        A1Pen.DashStyle = Drawing2D.DashStyle.Solid
        Dim A2Pen As New Pen(Color.Green, 1)
        A2Pen.DashStyle = Drawing2D.DashStyle.Solid
        Dim B1Pen As New Pen(Color.Black, 2)
        B1Pen.DashStyle = Drawing2D.DashStyle.Solid
        '===============================================================            
        If Size = "MAX" Then




            If readMaster(M_Data(DispStartPosition), _R) = "" Then
                Form1.PictureBox2.BackColor = Color.Gray
                Exit Sub
            Else
                Form1.PictureBox2.BackColor = Color.White
            End If

        ElseIf Size = "Middle" Then

            If readMaster(M_Data(DispStartPosition), _R) = "" Then
                FormMiddle.PictureBox2.BackColor = Color.Gray
                Exit Sub
            Else
                FormMiddle.PictureBox2.BackColor = Color.White
            End If

        ElseIf Size = "MIN" Then

            If readMaster(M_Data(DispStartPosition), _R) = "" Then
                FormSmall.PictureBox2.BackColor = Color.Gray
                Exit Sub
            Else
                FormSmall.PictureBox2.BackColor = Color.White
            End If

        End If
        UclChangeFlag = False
        UclChangeFlag2 = False


        If Size = "MAX" Then
            Form1.TextRUCL.Text = ""
            Form1.TextRCL.Text = ""
            Form1.TextRSiguma.Text = ""
            yh = 270
        ElseIf Size = "Middle" Then
            FormMiddle.TextRUCL.Text = ""
            FormMiddle.TextRCL.Text = ""
            FormMiddle.TextRSiguma.Text = ""
            yh = 250
        ElseIf Size = "MIN" Then
            FormSmall.TextRUCL.Text = ""
            FormSmall.TextRCL.Text = ""
            FormSmall.TextRSiguma.Text = ""
            yh = 145
        End If



        xp = 0
        yp = 0




        Dim x00 As Integer
        Dim x01 As Integer
        Dim x02 As Integer
        Dim x03 As Integer
        Dim x04 As Integer
        Dim x05 As Integer
        Dim x06 As Integer

        If Size = "MAX" Then
            x00 = 30
            x01 = 30
            x02 = 30
            x03 = 30
            yh = 270
        ElseIf Size = "Middle" Then
            x00 = 25
            x01 = 30
            x02 = 20
            x03 = 30
            yh = 250
        ElseIf Size = "MIN" Then
            x00 = 20
            x01 = 20
            x02 = 20
            x03 = 20
            yh = 145
        End If
        For j = 1 To 35
            g.DrawLine(BPen, xp + j * x00, yp, xp + j * x00, yp + 270)
        Next

        For j = 1 To 10
            g.DrawLine(BPen, xp, yp + j * x01, xp + 35 * x00, yp + j * x01)
        Next

        If MR Then 'MR???
            If PropertyTable.Rows(PropertyNo)("cMRucl") = 0 Then 'MR????????????????R???????????
                dbl2 = (PropertyTable.Rows(PropertyNo)("cRucl") * 30) / 100   'STEP
            Else
                dbl2 = (PropertyTable.Rows(PropertyNo)("cMRucl") * 30) / 100   'STEP
            End If
        Else 'R???
            dbl2 = (PropertyTable.Rows(PropertyNo)("cRucl") * x02) / 100   'STEP

        End If

        Bairitu = x03 / dbl2       '?40Pix

        null_bit = 1


        i = DispStartPosition '????????????

        If Size = "MAX" Then
            'x04 = 15 + 120
            x04 = 15
            x05 = 15
        ElseIf Size = "Middle" Then
            x04 = 10
            x05 = 13
        ElseIf Size = "MIN" Then
            x04 = 10
            x05 = 10
        End If

        For j = 0 To 29 * Graphsmallcount
            If i < 0 OrElse i > UBound(M_Data) Then
                Exit For
            End If
            If M_Data(i) Is Nothing OrElse M_Data(i) = "" Then
                If i < UBound(M_Data) Then i += 1
                Continue For
            End If

            If MR Then
                strData = readMaster(M_Data(i), _MR)
            Else
                strData = readMaster(M_Data(i), _R)
            End If

            If strData <> "" Then

                xpn = xp + j * (x00 / Graphsmallcount) + x04
                '------------   CL,UCL,LCL????   -------------

                p = 0

                For k = 0 To PropertyNo
                    If IsDBNull(PropertyTable.Rows(k)("cApprovalDate")) = False Then
                        If readMaster(M_Data(i), _wDate) > PropertyTable.Rows(k)("cApprovalDate") Then
                            p = k
                        End If
                    End If
                Next

                R_UCL = PropertyTable.Rows(p)("cRucl")
                R_CL = PropertyTable.Rows(p)("cRcl")
                R_Shiguma = PropertyTable.Rows(p)("cRdev")

                MR_UCL = PropertyTable.Rows(p)("cMRucl")
                MR_CL = PropertyTable.Rows(p)("cMRcl")
                MR_Shiguma = PropertyTable.Rows(p)("cMRdev")


                If MR Then
                    dblSiguma = MR_UCL
                    dblCl = MR_CL
                    R_Shiguma = MR_Shiguma
                Else
                    dblSiguma = R_UCL
                    dblCl = R_CL
                End If



                'CL?????
                Data1 = (CDbl(dblCl)) * Bairitu
                ypn = yp + yh - Data1                '
                If ypn > yp + yh Then ypn = yp + yh
                g.DrawLine(APen, xpn - x05, ypn, xpn + x05, ypn)
                yclR = ypn

                'UCL?????
                Data1 = (CDbl(dblSiguma)) * Bairitu
                ypn = yp + yh - Data1                '
                If ypn > yp + yh Then ypn = yp + yh
                g.DrawLine(FPen, xpn - x05, ypn, xpn + x05, ypn)
                yuclR = ypn

                Data1 = (CDbl(strData) - dblLow) * Bairitu
                ypn = yp + yh - Data1                '


                Dim ala As Integer = readMaster(M_Alarm(i)(0), 0)

                If MR Then
                    ala = readMaster(M_Alarm(i)(2), 0)
                Else
                    ala = readMaster(M_Alarm(i)(1), 0)
                End If

                '??????????0
                colbuf(j) = 0
                '???????1
                If ala = 1 Then
                    colbuf(j) = 1
                End If

                '??????????????????
                If ala = 2 Or ala = 3 Then
                    colbuf(j) = 1
                    ypf = ypn

                    ypa = ypf - 50
                    If ypa < 0 Then
                        ypa = 50
                        ypf = ypa + 50
                    End If
                    If ypa > 400 Then
                        ypa = 400
                        ypf = ypa + 50
                    End If

                    g.DrawLine(B1Pen, xpn, ypf - 50, xpn, ypa + 23)
                    k = 0
                    For ii = 0 To 7
                        If ala = 3 Then 'QC?????????????
                            g.DrawLine(A2Pen, xpn + k, ypf - 50 + ii, xpn + k, ypa + 15 - ii)
                        ElseIf ala = 2 Then 'QC???????????????
                            g.DrawLine(A1Pen, xpn + k, ypf - 50 + ii, xpn + k, ypa + 15 - ii)
                        End If

                        k += 1
                    Next
                End If

                xpnbuf_R(j) = xpn
                ypnbuf_R(j) = ypn
                'BufNoR(j) = i
                If j > 0 And null_bit = 0 Then
                    g.DrawLine(DPen, xp_old, yp_old, xpn, ypn) '?????
                End If
                null_bit = 0
                xp_old = xpn
                yp_old = ypn
                end_bit = 1
                i += 1
            End If
        Next


        If Size = "MAX" Then
            Form1.labRUCL.Top = yuclR - 5 + Form1.PictureBox2.Top
            Form1.TextRUCL.Text = dblSiguma
            Form1.labRCL.Top = yclR - 5 + Form1.PictureBox2.Top
            Form1.TextRCL.Text = dblCl
            Form1.TextRSiguma.Text = R_Shiguma
        ElseIf Size = "Middle" Then
            FormMiddle.labRUCL.Top = yuclR - 5 + FormMiddle.PictureBox2.Top
            FormMiddle.TextRUCL.Text = dblSiguma
            FormMiddle.labRCL.Top = yclR - 5 + FormMiddle.PictureBox2.Top
            FormMiddle.TextRCL.Text = dblCl
            FormMiddle.TextRSiguma.Text = R_Shiguma
        ElseIf Size = "MIN" Then
            FormSmall.labRUCL.Top = yuclR - 5 + FormSmall.PictureBox2.Top
            FormSmall.TextRUCL.Text = dblSiguma
            FormSmall.labRCL.Top = yclR - 5 + FormSmall.PictureBox2.Top
            FormSmall.TextRCL.Text = dblCl
            FormSmall.TextRSiguma.Text = R_Shiguma
        End If


        '???????????==========================
        For j = 0 To 30 * Graphsmallcount
            xpn = xpnbuf_R(j)
            ypn = ypnbuf_R(j)
            If xpn <> "0" Then
                If colbuf(j) = 1 Then '????????????????
                    g.FillEllipse(c1, xpn - 4, ypn - 4, 7, 7)
                    g.DrawEllipse(EPen, xpn - 4, ypn - 4, 7, 7)
                Else '????????????????
                    g.FillEllipse(c2, xpn - 4, ypn - 4, 7, 7)
                    g.DrawEllipse(EPen, xpn - 4, ypn - 4, 7, 7)
                End If
            End If
        Next
        '================================================

        '??????????==========================
        If Size = "MAX" Then
            x06 = 9
        ElseIf Size = "Middle" Then
            x06 = 8
        ElseIf Size = "MIN" Then
            x06 = 7
        End If
        dbl1 = dbl2 * x06
        For i = 0 To x06

            If Size = "MAX" Then
                If dbl1 > 999 Then
                    Form1.LabR(i).Text = Format(dbl1, "0")
                Else
                    Form1.LabR(i).Text = Format(dbl1, "0.00")
                End If
            ElseIf Size = "Middle" Then
                If dbl1 > 99 Then
                    FormMiddle.LabR_Middle(i).Text = Format(dbl1, "0")
                Else
                    FormMiddle.LabR_Middle(i).Text = Format(dbl1, "0.00")
                End If
            ElseIf Size = "MIN" Then
                If dbl1 > 99 Then
                    FormSmall.LabR_Small(i).Text = Format(dbl1, "0")
                Else
                    FormSmall.LabR_Small(i).Text = Format(dbl1, "0.00")
                End If
            End If


            dbl1 -= dbl2
        Next
        '==============================================

        c1.Dispose()
        c2.Dispose()
        c3.Dispose()
        f.Dispose()
        g.Dispose()



    End Sub

    '?????????????????
    Public Sub GraphDisp4(ByVal Size As String)

        Dim j, k, xp, yp As Integer
        Dim f As New Font("Segoe UI", 8, FontStyle.Regular)
        Dim f2 As New Font("Segoe UI", 7, FontStyle.Regular)
        Dim g As Graphics
        Dim rect As New RectangleF '???????
        Dim sf As New StringFormat() '???????
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center


        If Size = "MAX" Then
            With Form1.PictureBox4
                .Image = New Bitmap(1355, 150)
                g = Graphics.FromImage(.Image)
            End With
        ElseIf Size = "Middle" Then
            With FormMiddle.PictureBox4
                .Image = New Bitmap(1355, 150)
                g = Graphics.FromImage(.Image)
            End With
        ElseIf Size = "MIN" Then
            With FormSmall.PictureBox4
                .Image = New Bitmap(1355, 150)
                g = Graphics.FromImage(.Image)
            End With
        End If
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias

        Dim Stf As New StringFormat(StringFormatFlags.DirectionVertical)
        Dim str1 As String
        Dim dt1 As Date
        Dim APen As New Pen(Color.Black, 1)
        APen.DashStyle = Drawing2D.DashStyle.Solid

        Dim c1 As New SolidBrush(Color.FromArgb(255, 208, 241, 255))  '??

        xp = 0
        yp = 0
        Dim x00 As Integer
        Dim x01 As Integer
        Dim x02 As Integer
        Dim x03 As Integer
        Dim x04 As Integer
        Dim x05 As Integer
        If Size = "MAX" Then
            x00 = 30
            x01 = 26
            x02 = 26
            x03 = 26
            x04 = 6
            'x05 = 0
            x05 = 4
        ElseIf Size = "Middle" Then
            x00 = 25
            x01 = 25
            x02 = 25
            x03 = 25
            x04 = 3
            x05 = 1
        ElseIf Size = "MIN" Then
            x00 = 20
            x01 = 25
            x02 = 25
            x03 = 25
            x04 = 3
            x05 = 1
        End If
        g.FillRectangle(c1, xp, yp, x00 * 6, x01)
        g.FillRectangle(c1, xp + 6 * x00, yp, x00 * 33, x01)
        g.DrawLine(APen, xp, yp, xp + 37 * x00, yp)
        g.DrawLine(APen, xp, yp + x01, xp + 37 * x00, yp + x01)
        g.DrawLine(APen, xp, yp + x01 + x02, xp + 37 * x00, yp + x01 + x02)
        g.DrawLine(APen, xp, yp + x01 + x02 + x03, xp + 37 * x00, yp + x01 + x02 + x03)
        'g.DrawLine(APen, xp, yp + x01 + x02 + x03 + x04, xp + 37 * x00, yp + x01 + x02 + x03 + x04)

        k = DispStartPosition

        For j = 0 To 36
            If k < 0 OrElse k > UBound(M_Data) Then
                k += (1 * Graphsmallcount)
                Continue For
            End If
            If M_Data(k) Is Nothing OrElse M_Data(k) = "" Then
                k += (1 * Graphsmallcount)
                Continue For
            End If
            Dim wdate As String = readMaster(M_Data(k), _wDate)

            If j >= x04 And k < SPCDataNum And wdate <> "" And j <> 36 Then
                Dim strX As String = readMaster(M_Data(k), _X)
                Dim strR As String = readMaster(M_Data(k), _R)
                dt1 = CDate(wdate)
                str1 = dt1.Month.ToString.PadLeft(2, "0") & Environment.NewLine & dt1.Day.ToString.PadLeft(2, "0")
                '??????
                rect = New RectangleF(xp + (j - x05) * x00, yp + 2, x00, x01)
                g.DrawString(str1, f, Brushes.Black, rect, sf)
                rect = New RectangleF(xp, yp + 2, 2 * x00, x01)
                g.DrawString("Date", f, Brushes.Black, rect, sf)
                '???X???
                rect = New RectangleF(xp + (j - x05) * x00 + 1, yp + x01 + 2, x00, x02)
                g.DrawString(strX, f, Brushes.Black, rect, sf)
                rect = New RectangleF(xp, yp + x01 + 2, 2 * x00, x02)
                g.DrawString("XBar", f, Brushes.Black, rect, sf)
                '???R???
                rect = New RectangleF(xp + (j - x05) * x00 + 1, yp + x01 + x02 + 2, x00, x03)
                g.DrawString(strR, f, Brushes.Black, rect, sf)
                rect = New RectangleF(xp, yp + x01 + x02 + 2, 2 * x00, x02)
                g.DrawString("R", f, Brushes.Black, rect, sf)

                k += (1 * Graphsmallcount)
            End If
            If j <> 1 Then
                g.DrawLine(APen, xp + j * x00, yp, xp + j * x00, yp + 450)
            End If
        Next
        If Size = "MAX" Then
            Form1.LabUnit.Text = strUnit
        ElseIf Size = "Middle" Then
            FormMiddle.LabUnit.Text = strUnit
        ElseIf Size = "MIN" Then
            FormSmall.LabUnit.Text = strUnit
        End If




        c1.Dispose()
        f.Dispose()
        g.Dispose()


    End Sub
    '???????????
    Public Sub GraphDisp7(ByVal Size As String)
        Dim xp, yp, yh, i, j, jk, yps, ul, ll As Integer
        Dim Bairitu, yp0 As Double
        Dim dbl1, dbl2, dblLow, dblData As Double
        'Dim xpnbuf(31), ypnbuf(31), colbuf(31) As Integer
        Dim HistogramBuf(3000) As String
        Dim HistogramCount As Integer = 0
        Dim strData, strStep As String
        Dim Cl_s1, s1_s2, s2_s3, s3_s4, s4_s5, s5_s6, s6_s7, s7_s8, s8_s9, s9_s10 As Integer 'UCL??
        Dim Cl_ms1, ms1_ms2, ms2_ms3, ms3_ms4, ms4_ms5, ms5_ms6, ms6_ms7, ms7_ms8, ms8_ms9, ms9_ms10 As Integer 'LCL??
        Dim local_yucl, local_ylcl, local_ycl As Integer


        Dim g As Graphics

        If Size = "MAX" Then
            With Form1.PictureBox9
                .Image = New Bitmap(Form1.PictureBox9.Width, Form1.PictureBox9.Height)
                g = Graphics.FromImage(.Image)
            End With
            yh = 428
        ElseIf Size = "Middle" Then
            With FormMiddle.PictureBox9
                .Image = New Bitmap(FormMiddle.PictureBox9.Width, FormMiddle.PictureBox9.Height)
                g = Graphics.FromImage(.Image)
            End With
            yh = 362
        ElseIf Size = "MIN" Then
            With FormSmall.PictureBox9
                .Image = New Bitmap(FormSmall.PictureBox9.Width, FormSmall.PictureBox9.Height)
                g = Graphics.FromImage(.Image)
            End With
            yh = 260
        End If
        g.Clear(Color.White)
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

        Dim APen As New Pen(Color.Green, 2)
        APen.DashStyle = Drawing2D.DashStyle.Dot

        Dim BPen As New Pen(Color.Black, 1)
        BPen.DashStyle = Drawing2D.DashStyle.Dot

        Dim CPen As New Pen(Color.Orange, 1)
        CPen.DashStyle = Drawing2D.DashStyle.Solid

        Dim FPen As New Pen(Color.Red, 2)
        FPen.DashStyle = Drawing2D.DashStyle.Dash

        Dim HPen As New Pen(Color.Red, 3)
        HPen.DashStyle = Drawing2D.DashStyle.Solid

        Dim B1Pen As New Pen(Color.Black, 1)
        B1Pen.DashStyle = Drawing2D.DashStyle.Solid
        '===============================================================

        Dim x00 As Integer
        Dim x01 As Integer
        Dim x02 As Integer
        Dim x03 As Integer
        'Dim x04 As Integer
        If Size = "MAX" Then
            x00 = 40
            x01 = 25
            x02 = 2
            x03 = 10
        ElseIf Size = "Middle" Then
            x00 = 35
            x01 = 20
            x02 = 1
            x03 = 5
        ElseIf Size = "MIN" Then
            x00 = 25
            x01 = 20
            x02 = 1
            x03 = 5
        End If

        xp = 0
        yp = 0
        yp0 = 0

        Dim p As Integer = 0
        If PropertyTable IsNot Nothing AndAlso PropertyTable.Rows.Count > 0 Then
            If DispStartPosition < M_Data.Length AndAlso M_Data(DispStartPosition) IsNot Nothing Then
                Dim currentDataDate As String = readMaster(M_Data(DispStartPosition), _wDate)
                For k = 0 To PropertyTable.Rows.Count - 1
                    If IsDBNull(PropertyTable.Rows(k)("cApprovalDate")) = False Then
                        If String.Compare(currentDataDate, PropertyTable.Rows(k)("cApprovalDate").ToString()) > 0 Then
                            p = k
                        End If
                    End If
                Next
            End If
            X_SCL = PropertyTable.Rows(p)("cScl")
            X_kousa = PropertyTable.Rows(p)("cTolerance")
            X_CL = PropertyTable.Rows(p)("cXcl")
            X_USL = PropertyTable.Rows(p)("cUsl")
            X_LSL = PropertyTable.Rows(p)("cLsl")
            X_UCL = PropertyTable.Rows(p)("cXucl")
            X_LCL = PropertyTable.Rows(p)("cXlcl")
            X_Shiguma = PropertyTable.Rows(p)("cXdev")
            X_gType = PropertyTable.Rows(p)("cLimitType")
        Else
            g.Dispose()
            Exit Sub
        End If

        dbl1 = X_SCL
        dbl2 = X_kousa / 5 'STEP
        If dbl2 = 0 Then dbl2 = 1
        strStep = CStr(dbl2)      'STEP
        Bairitu = x00 / dbl2       '?40Pix
        dblLow = dbl1 - dbl2 * 5

        If X_gType = "UpperLower" Then
            yp0 = ((X_USL + X_LSL) / 2 - X_SCL)
            yp = yp0 * Bairitu
        End If


        ul = yp + yh - (X_USL - dblLow) * Bairitu
        g.DrawLine(HPen, 0, ul, 1500, ul)
        ll = yp + yh - (X_LSL - dblLow) * Bairitu
        g.DrawLine(HPen, 0, ll, 1500, ll)


        jk = DispStartPosition

        For j = 1 To 100
            g.DrawLine(BPen, xp + j * x01, 0, xp + j * x01, 1000)
        Next

        yps = yp + yh - (X_CL - dblLow) * Bairitu

        If X_Shiguma <> 0 Then
            For j = 0 To 100
                Dim y_plus As Integer = CInt(yps + j * X_Shiguma / strStep * x00)
                Dim y_minus As Integer = CInt(yps - j * X_Shiguma / strStep * x00)
                g.DrawLine(BPen, xp, y_plus, xp + 36 * 30, y_plus)
                g.DrawLine(BPen, xp, y_minus, xp + 36 * 30, y_minus)
            Next
        Else
            For j = 1 To 15
                g.DrawLine(BPen, xp, yp + j * x00, xp + 36 * 30, yp + j * x00)
            Next
        End If


        For i = DispStartPosition To DispStartPosition + 29 + (30 * (Graphsmallcount - 1))
            If i > UBound(M_Data) Then Exit For
            If M_Data(i) Is Nothing Then Continue For
            strData = readMaster(M_Data(i), _X)
            'strData = SPCXDataBuf(i)
            If strData <> "" Then
                HistogramBuf(HistogramCount) = strData
                HistogramCount += 1
                jk += 1
            End If
        Next

        If X_Shiguma = 0 And HistogramCount > 1 Then
            Dim h_sum As Double = 0
            Dim h_sumSq As Double = 0
            For i = 0 To HistogramCount - 1
                Dim val As Double = CDbl(HistogramBuf(i))
                h_sum += val
                h_sumSq += val * val
            Next
            Dim h_mean As Double = h_sum / HistogramCount
            Dim h_var As Double = (h_sumSq - (h_sum * h_sum) / HistogramCount) / (HistogramCount - 1)
            If h_var > 0 Then X_Shiguma = Math.Sqrt(h_var)
            If X_CL = 0 Then X_CL = h_mean
        End If

        If X_Shiguma <= 0.000001 Then X_Shiguma = dbl2
        yps = yp + yh - (X_CL - dblLow) * Bairitu


        local_yucl = yp + yh - (X_UCL - dblLow) * Bairitu
        local_ylcl = yp + yh - (X_LCL - dblLow) * Bairitu
        local_ycl = yp + yh - (X_CL - dblLow) * Bairitu
        '?????????=================

        g.DrawLine(FPen, 0, local_yucl, 1500, local_yucl)
        g.DrawLine(FPen, 0, local_ylcl, 1500, local_ylcl)
        g.DrawLine(APen, 0, local_ycl, 1500, local_ycl)

        Dim barWidth As Integer = x03
        If Graphsmallcount >= 2 Then barWidth = x02

        If X_Shiguma <> 0 Then
            For i = 0 To HistogramCount - 1
                dblData = HistogramBuf(i)
                If dblData > (CDbl(X_CL) - (CDbl(X_Shiguma) / 2)) And dblData <= (CDbl(X_CL) + (CDbl(X_Shiguma) / 2)) Then '-0.5?~0.5????
                    Cl_s1 += barWidth
                ElseIf dblData > (CDbl(X_CL) + (CDbl(X_Shiguma) / 2)) And dblData <= (CDbl(X_CL) + CDbl(X_Shiguma * 1.5)) Then '1?-2????
                    s1_s2 += barWidth
                ElseIf dblData > (CDbl(X_CL) + CDbl(X_Shiguma * 1.5)) And dblData <= (CDbl(X_CL) + CDbl(X_Shiguma * 2.5)) Then '2?-3????
                    s2_s3 += barWidth
                ElseIf dblData > (CDbl(X_CL) + CDbl(X_Shiguma * 2.5)) And dblData <= (CDbl(X_CL) + CDbl(X_Shiguma * 3.5)) Then '3?-4????
                    s3_s4 += barWidth
                ElseIf dblData > (CDbl(X_CL) + CDbl(X_Shiguma * 3.5)) And dblData <= (CDbl(X_CL) + CDbl(X_Shiguma * 4.5)) Then '4?-5????
                    s4_s5 += barWidth
                ElseIf dblData > (CDbl(X_CL) + CDbl(X_Shiguma * 4.5)) And dblData <= (CDbl(X_CL) + CDbl(X_Shiguma * 5.5)) Then '5?-6????
                    s5_s6 += barWidth
                ElseIf dblData > (CDbl(X_CL) + CDbl(X_Shiguma * 5.5)) And dblData <= (CDbl(X_CL) + CDbl(X_Shiguma * 6.5)) Then '6?-7????
                    s6_s7 += barWidth
                ElseIf dblData > (CDbl(X_CL) + CDbl(X_Shiguma * 6.5)) And dblData <= (CDbl(X_CL) + CDbl(X_Shiguma * 7.5)) Then '7?-8????
                    s7_s8 += barWidth
                ElseIf dblData > (CDbl(X_CL) + CDbl(X_Shiguma * 7.5)) And dblData <= (CDbl(X_CL) + CDbl(X_Shiguma * 8.5)) Then '8?-9????
                    s8_s9 += barWidth
                ElseIf dblData > (CDbl(X_CL) + CDbl(X_Shiguma * 8.5)) Then '9??????
                    s9_s10 += barWidth
                ElseIf dblData <= (CDbl(X_CL) - (CDbl(X_Shiguma) / 2)) And dblData > (CDbl(X_CL) - (CDbl(X_Shiguma) * 1.5)) Then 'CL--1????
                    Cl_ms1 += barWidth
                ElseIf dblData <= (CDbl(X_CL) - CDbl(X_Shiguma * 1.5)) And dblData > (CDbl(X_CL) - CDbl(X_Shiguma * 2.5)) Then '-1?--2????
                    ms1_ms2 += barWidth
                ElseIf dblData <= (CDbl(X_CL) - CDbl(X_Shiguma * 2.5)) And dblData > (CDbl(X_CL) - CDbl(X_Shiguma * 3.5)) Then '-2?--3????
                    ms2_ms3 += barWidth
                ElseIf dblData <= (CDbl(X_CL) - CDbl(X_Shiguma * 3.5)) And dblData > (CDbl(X_CL) - CDbl(X_Shiguma * 4.5)) Then '-3?--4????
                    ms3_ms4 += barWidth
                ElseIf dblData <= (CDbl(X_CL) - CDbl(X_Shiguma * 4.5)) And dblData > (CDbl(X_CL) - CDbl(X_Shiguma * 5.5)) Then '-4?--5????
                    ms4_ms5 += barWidth
                ElseIf dblData <= (CDbl(X_CL) - CDbl(X_Shiguma * 5.5)) And dblData > (CDbl(X_CL) - CDbl(X_Shiguma * 6.5)) Then '-5?--6????
                    ms5_ms6 += barWidth
                ElseIf dblData <= (CDbl(X_CL) - CDbl(X_Shiguma * 6.5)) And dblData > (CDbl(X_CL) - CDbl(X_Shiguma * 7.5)) Then '-6?--7????
                    ms6_ms7 += barWidth
                ElseIf dblData <= (CDbl(X_CL) - CDbl(X_Shiguma * 7.5)) And dblData > (CDbl(X_CL) - CDbl(X_Shiguma * 8.5)) Then '-7?--8????
                    ms7_ms8 += barWidth
                ElseIf dblData <= (CDbl(X_CL) - CDbl(X_Shiguma * 8.5)) And dblData > (CDbl(X_CL) - CDbl(X_Shiguma * 9.5)) Then '-8?--9????
                    ms8_ms9 += barWidth
                ElseIf dblData <= (CDbl(X_CL) - CDbl(X_Shiguma * 9.5)) Then '-9??????
                    ms9_ms10 += barWidth
                End If
            Next
        End If



        If Size = "MAX" Then
            Form1.Label24.Text = "0"
            Form1.Label25.Text = "25"
            Form1.Label26.Text = "50"
            Form1.Label27.Text = "75"
            Form1.Label28.Text = "100"
        ElseIf Size = "Middle" Then
            'FormMiddle.Label24.Text = "0"
            'FormMiddle.Label25.Text = "25"
            'FormMiddle.Label26.Text = "50"
            'FormMiddle.Label27.Text = "75"
            'FormMiddle.Label28.Text = "100"
        ElseIf Size = "MIN" Then
            'FormSmall.Label24.Text = "0"
            'FormSmall.Label25.Text = "25"
            'FormSmall.Label26.Text = "50"
            'FormSmall.Label27.Text = "75"
            'FormSmall.Label28.Text = "100"
        End If


        '-0.5?~-1.5???????????==================================
        yps = yp + yh - (CDbl(X_CL - (X_Shiguma / 2)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)

            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, Cl_ms1, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, Cl_ms1, j + yps)
                g.DrawLine(B1Pen, Cl_ms1, yps, Cl_ms1, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, Cl_ms1, j + yps)
            End If

        Next

        '-1.5?~-2.5???????????==================================
        yps = yp + yh - (CDbl(X_CL - X_Shiguma * 1.5) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, ms1_ms2, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, ms1_ms2, j + yps)
                g.DrawLine(B1Pen, ms1_ms2, yps, ms1_ms2, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, ms1_ms2, j + yps)
            End If
        Next

        '-2.5?~-3.5???????????==================================
        yps = yp + yh - (CDbl(X_CL - (X_Shiguma * 2.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, ms2_ms3, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, ms2_ms3, j + yps)
                g.DrawLine(B1Pen, ms2_ms3, yps, ms2_ms3, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, ms2_ms3, j + yps)
            End If
        Next

        '-3.5?~-4.5???????????==================================
        yps = yp + yh - (CDbl(X_CL - (X_Shiguma * 3.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, ms3_ms4, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, ms3_ms4, j + yps)
                g.DrawLine(B1Pen, ms3_ms4, yps, ms3_ms4, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, ms3_ms4, j + yps)
            End If
        Next

        '-4.5?~-5.5???????????==================================
        yps = yp + yh - (CDbl(X_CL - (X_Shiguma * 4.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, ms4_ms5, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, ms4_ms5, j + yps)
                g.DrawLine(B1Pen, ms4_ms5, yps, ms4_ms5, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, ms4_ms5, j + yps)
            End If
        Next

        '-5.5?~-6.5???????????==================================
        yps = yp + yh - (CDbl(X_CL - (X_Shiguma * 5.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, ms5_ms6, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, ms5_ms6, j + yps)
                g.DrawLine(B1Pen, ms5_ms6, yps, ms5_ms6, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, ms5_ms6, j + yps)
            End If
        Next

        '-6.5?~-7.5???????????==================================
        yps = yp + yh - (CDbl(X_CL - (X_Shiguma * 6.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, ms6_ms7, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, ms6_ms7, j + yps)
                g.DrawLine(B1Pen, ms6_ms7, yps, ms6_ms7, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, ms6_ms7, j + yps)
            End If
        Next

        '-7.5?~-8.5???????????==================================
        yps = yp + yh - (CDbl(X_CL - (X_Shiguma * 7.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, ms7_ms8, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, ms7_ms8, j + yps)
                g.DrawLine(B1Pen, ms7_ms8, yps, ms7_ms8, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, ms7_ms8, j + yps)
            End If
        Next

        '-8.5?~-9.5???????????==================================
        yps = yp + yh - (CDbl(X_CL - (X_Shiguma * 8.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, ms8_ms9, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, ms8_ms9, j + yps)
                g.DrawLine(B1Pen, ms8_ms9, yps, ms8_ms9, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, ms8_ms9, j + yps)
            End If
        Next

        '-9.5?~-10.5???????????==================================
        yps = yp + yh - (CDbl(X_CL - (X_Shiguma * 9.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, ms9_ms10, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, ms9_ms10, j + yps)
                g.DrawLine(B1Pen, ms9_ms10, yps, ms9_ms10, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, ms9_ms10, j + yps)
            End If
        Next

        '-0.5?~0.5???????????==================================
        yps = yp + yh - ((CDbl(X_CL) + (CDbl(X_Shiguma) / 2)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, Cl_s1, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, Cl_s1, j + yps)
                g.DrawLine(B1Pen, Cl_s1, yps, Cl_s1, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, Cl_s1, j + yps)
            End If
        Next

        '0.5?~1.5???????????==================================
        yps = yp + yh - ((CDbl(X_CL) + CDbl(X_Shiguma * 1.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, s1_s2, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, s1_s2, j + yps)
                g.DrawLine(B1Pen, s1_s2, yps, s1_s2, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, s1_s2, j + yps)
            End If
        Next

        '1.5?~2.5???????????==================================
        yps = yp + yh - ((CDbl(X_CL) + CDbl(X_Shiguma * 2.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, s2_s3, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, s2_s3, j + yps)
                g.DrawLine(B1Pen, s2_s3, yps, s2_s3, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, s2_s3, j + yps)
            End If
        Next

        '2.5?~3.5???????????==================================
        yps = yp + yh - ((CDbl(X_CL) + CDbl(X_Shiguma * 3.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, s3_s4, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, s3_s4, j + yps)
                g.DrawLine(B1Pen, s3_s4, yps, s3_s4, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, s3_s4, j + yps)
            End If
        Next

        '3.5?~4.5???????????==================================
        yps = yp + yh - ((CDbl(X_CL) + CDbl(X_Shiguma * 4.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, s4_s5, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, s4_s5, j + yps)
                g.DrawLine(B1Pen, s4_s5, yps, s4_s5, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, s4_s5, j + yps)
            End If
        Next

        '4.5?~5.5???????????==================================
        yps = yp + yh - ((CDbl(X_CL) + CDbl(X_Shiguma * 5.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, s5_s6, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, s5_s6, j + yps)
                g.DrawLine(B1Pen, s5_s6, yps, s5_s6, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, s5_s6, j + yps)
            End If
        Next

        '5.5?~6.5???????????==================================
        yps = yp + yh - ((CDbl(X_CL) + CDbl(X_Shiguma * 6.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, s6_s7, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, s6_s7, j + yps)
                g.DrawLine(B1Pen, s6_s7, yps, s6_s7, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, s6_s7, j + yps)
            End If
        Next

        '6.5?~7.5???????????==================================
        yps = yp + yh - ((CDbl(X_CL) + CDbl(X_Shiguma * 7.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, s7_s8, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, s7_s8, j + yps)
                g.DrawLine(B1Pen, s7_s8, yps, s7_s8, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, s7_s8, j + yps)
            End If
        Next

        '7.5?~8.5???????????==================================
        yps = yp + yh - ((CDbl(X_CL) + CDbl(X_Shiguma * 8.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, s8_s9, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, s8_s9, j + yps)
                g.DrawLine(B1Pen, s8_s9, yps, s8_s9, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, s8_s9, j + yps)
            End If
        Next

        '8.5?~9.5???????????==================================
        yps = yp + yh - ((CDbl(X_CL) + CDbl(X_Shiguma * 9.5)) - dblLow) * Bairitu
        For j = 0 To CInt(X_Shiguma / strStep * x00)
            If j = 0 Then
                g.DrawLine(B1Pen, 0, j + yps, s9_s10, j + yps)
            ElseIf j = CInt(X_Shiguma / strStep * x00) Then
                g.DrawLine(B1Pen, 0, j + yps, s9_s10, j + yps)
                g.DrawLine(B1Pen, s9_s10, yps, s9_s10, j + yps)
            Else
                g.DrawLine(CPen, 0, j + yps, s9_s10, j + yps)
            End If
        Next


        g.Dispose()

    End Sub

    '??????????????????
    Public Function getTreeData() As String(,)

        getTreeData = Nothing
        Dim _TreeRist(,) As String


        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim Adapter As New SqlDataAdapter
        Dim strSQL As String = ""
        Dim P_Table As New DataTable
        Dim AMP_Table As New DataTable


        Try

            Cn.ConnectionString = StrServerConnection
            P_Table.Clear()

            strSQL = "SELECT DISTINCT "

            For i As Integer = 0 To 10 - 1
                strSQL &= "cTreeName" & i + 1
                If Not i = 10 - 1 Then
                    strSQL &= " + ',' + "
                End If
            Next
            strSQL &= " AS Tree"

            strSQL &= " FROM SPC_Property"

            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(P_Table)
            P_Table.Dispose()

            If P_Table.Rows.Count = 0 Then
                Return Nothing
            End If

            ReDim _TreeRist(P_Table.Rows.Count - 1, 2 - 1)

            For i As Integer = 0 To UBound(_TreeRist, 1)
                _TreeRist(i, 0) = P_Table.Rows(i)("Tree")
                _TreeRist(i, 1) = "0"
            Next


            AMP_Table.Clear()

            strSQL = "SELECT DISTINCT "

            For i As Integer = 0 To 10 - 1
                strSQL &= "SPC_Alarm.cTreeName" & i + 1
                If Not i = 10 - 1 Then
                    strSQL &= " + ',' + "
                End If
            Next
            strSQL &= " AS Tree"

            strSQL &= " FROM SPC_Alarm"
            strSQL &= " LEFT OUTER JOIN SPC_Master ON SPC_Alarm.iID = SPC_Master.iID"
            strSQL &= " LEFT OUTER JOIN SPC_Property ON SPC_Property.cTreeName1 =  SPC_Alarm.cTreeName1"
            strSQL &= " AND SPC_Property.cTreeName2 =  SPC_Alarm.cTreeName2"
            strSQL &= " AND SPC_Property.cTreeName3 =  SPC_Alarm.cTreeName3"
            strSQL &= " AND SPC_Property.cTreeName4 =  SPC_Alarm.cTreeName4"
            strSQL &= " AND SPC_Property.cTreeName5 =  SPC_Alarm.cTreeName5"
            strSQL &= " AND SPC_Property.cTreeName6 =  SPC_Alarm.cTreeName6"
            strSQL &= " AND SPC_Property.cTreeName7 =  SPC_Alarm.cTreeName7"
            strSQL &= " AND SPC_Property.cTreeName8 =  SPC_Alarm.cTreeName8"
            strSQL &= " AND SPC_Property.cTreeName9 =  SPC_Alarm.cTreeName9"
            strSQL &= " AND SPC_Property.cTreeName10 =  SPC_Alarm.cTreeName10"

            strSQL &= " WHERE SPC_Property.aStartDate < SPC_Master.dWorkDate" '?????????
            strSQL &= " AND SPC_Alarm.cApproverName = ''" 'QC??????


            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(AMP_Table)

            AMP_Table.Dispose()


            For i As Integer = 0 To AMP_Table.Rows.Count - 1
                For j As Integer = 0 To UBound(_TreeRist, 1)
                    If Not _TreeRist(j, 0) = AMP_Table.Rows(i)("Tree") Then Continue For
                    _TreeRist(j, 1) = "1"
                    Exit For
                Next
            Next


            Adapter.Dispose()
            Cn.Dispose()
            P_Table.Dispose()
            AMP_Table.Dispose()

            Return _TreeRist

        Catch ex As System.Exception
            Adapter.Dispose()
            Cn.Dispose()
            P_Table.Dispose()
            AMP_Table.Dispose()
            StrErrMes = "????????????" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Return Nothing
        End Try


    End Function

#Region "?????"
    Public mouseY As Integer
    Public i_old As Integer = 1000
    Public Sub popUp(ByVal mx1 As Integer, ByVal my1 As Integer, ByVal objName As String)
        If M_Data Is Nothing Then Exit Sub
        If FormAlarmDisp.Visible = True Then Exit Sub
        If FormAlarmInput.Visible = True Then Exit Sub

        Dim dc, i As Integer
        Dim _mode As String

        If objName = "PictureBox1" Then
            _mode = "X"
        ElseIf objName = "PictureBox2" Then
            _mode = "R"
        Else
            Exit Sub
        End If

        If _mode = "R" Then
            If MRFlag = True Then
                _mode = "MR"
            End If
        End If

        dc = 0

        For i = 0 To 31
            If _mode = "X" Then

                If Not (xpnbuf_X(i) - 10 < mx1 And xpnbuf_X(i) + 10 > mx1) Then Continue For

                If 0 < ypnbuf_X(i) Then '??????????
                    If Not (ypnbuf_X(i) - 10 < my1 And ypnbuf_X(i) + 10 > my1) Then Continue For
                End If
                dc = 1
                Exit For
            ElseIf _mode = "R" Or _mode = "MR" Then


                If Not (xpnbuf_R(i) - 10 < mx1 And xpnbuf_R(i) + 10 > mx1) Then Continue For

                If 0 < ypnbuf_R(i) Then '??????????
                    If Not (ypnbuf_R(i) - 10 < my1 And ypnbuf_R(i) + 10 > my1) Then Continue For
                End If
                dc = 1
                Exit For

            End If

        Next

        If dc = 0 Then
            i_old = 1000
            FormPopupNew.Close()
            Exit Sub
        End If



        If i_old = i Then Exit Sub '??????????


        'popUp????(X???pic2????R???pic1???)
        Dim form_top As Integer
        Dim form_left As Integer
        Dim pic1_top As Integer
        Dim pic1_left As Integer
        Dim pic4_top As Integer
        Dim pic4_left As Integer

        If Form1.Visible = True Then
            form_top = Form1.Top
            form_left = Form1.Left
            pic1_top = Form1.PictureBox1.Top
            pic1_left = Form1.PictureBox1.Left
            pic4_top = Form1.PictureBox4.Top
            pic4_left = Form1.PictureBox4.Left
        ElseIf FormMiddle.Visible = True Then
            form_top = FormMiddle.Top
            form_left = FormMiddle.Left
            pic1_top = FormMiddle.PictureBox1.Top
            pic1_left = FormMiddle.PictureBox1.Left
            pic4_top = FormMiddle.PictureBox4.Top
            pic4_left = FormMiddle.PictureBox4.Left
        ElseIf FormSmall.Visible = True Then
            form_top = FormSmall.Top
            form_left = FormSmall.Left
            pic1_top = FormSmall.PictureBox1.Top
            pic1_left = FormSmall.PictureBox1.Left
            pic4_top = FormSmall.PictureBox4.Top
            pic4_left = FormSmall.PictureBox4.Left
        End If

        Dim pic_top As Integer
        Dim pic_left As Integer
        If _mode = "X" Then
            pic_top = pic4_top
            pic_left = pic4_left
        Else
            pic_top = pic1_top
            pic_left = pic1_left
        End If

        mouseY = form_top + pic_top + 30 '?????????



        i_old = i
        Display_Popup(i, _mode)


    End Sub


    Public Sub Display_Popup(ByVal d As Integer, ByVal Mode As String)
        Dim po As Integer
        Dim temp() As String
        Dim strAlarmName As String = ""

        Dim x1, x2, x3, x4, y0, y1, y2, y3, y4, y5, y6, y7, Lot0_x, Lot0_y, Lot_x, Lot_y, Va0_x, Va0_y, Va_x, Va_y As Integer

        x1 = 10
        x2 = 3
        x3 = 10
        y0 = 10
        Lot0_x = 68
        Lot0_y = 18
        y1 = 3
        y2 = 7
        y3 = y1
        y4 = y2
        y5 = y1
        y6 = y1
        y7 = y0
        Va0_x = Lot0_x
        Va0_y = Lot0_y
        Va_x = Va0_x
        Va_y = Va0_y
        x4 = x1

        Lot_x = Va_x + x3 + Va0_x + x2 + Va_x
        Lot_y = Lot0_y

        po = DispStartPosition + d


        If readMaster(M_Data(po), _X) <> "" Then

            Try
                For i = 0 To FormPopupNew.Values0.Length - 1
                    FormPopupNew.Controls.Remove(FormPopupNew.Values0(i))
                Next
                For i = 0 To FormPopupNew.Values.Length - 1
                    FormPopupNew.Controls.Remove(FormPopupNew.Values(i))
                Next
            Catch ex As Exception

            End Try

            temp = Split(MesureValueBuf(po), ",")
            Dim leng As Integer = UBound(temp, 1) + 1
            If 20 < leng Then leng = 20
            FormPopupNew.Values0 = New System.Windows.Forms.Label(leng - 1) {}
            FormPopupNew.SuspendLayout()

            For i As Integer = 0 To UBound(FormPopupNew.Values0, 1)
                FormPopupNew.Values0(i) = New System.Windows.Forms.Label
                FormPopupNew.Values0(i).Name = "Values0" & (i).ToString()
                FormPopupNew.Values0(i).Text = "Value" & (i).ToString().PadLeft(2, "0")
                FormPopupNew.Values0(i).Font = New Font("Meiryo UI", 10)
                FormPopupNew.Values0(i).AutoSize = False
                FormPopupNew.Values0(i).TextAlign = ContentAlignment.MiddleCenter
                FormPopupNew.Values0(i).BorderStyle = BorderStyle.None
                FormPopupNew.Values0(i).Location = New Point(x1 + (i \ 5) * (Va0_x + x2 + Va_x + x3), y0 + Lot0_y + y1 + Lot0_y + y2 + (i Mod 5) * (Va0_y + y3))
                FormPopupNew.Values0(i).Size = New System.Drawing.Size(Va0_x, Va0_y)
                FormPopupNew.Values0(i).BackColor = Color.FromArgb(255, 255, 128)
            Next

            FormPopupNew.Controls.AddRange(FormPopupNew.Values0)
            FormPopupNew.ResumeLayout(False)


            FormPopupNew.Values = New System.Windows.Forms.Label(leng - 1) {}
            FormPopupNew.SuspendLayout()

            For i As Integer = 0 To UBound(FormPopupNew.Values, 1)
                FormPopupNew.Values(i) = New System.Windows.Forms.Label
                FormPopupNew.Values(i).Name = "Values" & (i).ToString()
                FormPopupNew.Values(i).Text = temp(i)
                FormPopupNew.Values(i).Font = New Font("Meiryo UI", 10)
                FormPopupNew.Values(i).AutoSize = False
                FormPopupNew.Values(i).TextAlign = ContentAlignment.MiddleCenter
                FormPopupNew.Values(i).BorderStyle = BorderStyle.Fixed3D
                FormPopupNew.Values(i).Location = New Point(x1 + (i \ 5) * (Va0_x + x2 + Va_x + x3) + Va0_x + x2, y0 + Lot0_y + y1 + Lot0_y + y2 + (i Mod 5) * (Va0_y + y3))
                FormPopupNew.Values(i).Size = New System.Drawing.Size(Va_x, Va_y)
                FormPopupNew.Values(i).BackColor = Color.FromArgb(255, 255, 255)
            Next

            FormPopupNew.Controls.AddRange(FormPopupNew.Values)
            FormPopupNew.ResumeLayout(False)



            FormPopupNew.Labels = New System.Windows.Forms.Label(12 - 1) {}
            FormPopupNew.SuspendLayout()

            For i As Integer = 0 To UBound(FormPopupNew.Labels, 1)
                FormPopupNew.Labels(i) = New System.Windows.Forms.Label
                FormPopupNew.Labels(i).Name = "Labels" & (i).ToString()

                FormPopupNew.Labels(i).Font = New Font("Meiryo UI", 10)
                FormPopupNew.Labels(i).AutoSize = False
                FormPopupNew.Labels(i).TextAlign = ContentAlignment.MiddleCenter
                FormPopupNew.Labels(i).BorderStyle = BorderStyle.FixedSingle
                FormPopupNew.Labels(i).BackColor = Color.FromArgb(215, 255, 255)
                If i Mod 2 = 1 Then
                    FormPopupNew.Labels(i).BorderStyle = BorderStyle.Fixed3D
                    FormPopupNew.Labels(i).BackColor = Color.FromArgb(255, 255, 255)
                End If
            Next

            'Data
            FormPopupNew.Labels(0).Text = "Date"
            FormPopupNew.Labels(0).Location = New Point(x1, y0)
            FormPopupNew.Labels(0).Size = New System.Drawing.Size(Lot0_x, Lot0_y)
            FormPopupNew.Labels(1).Text = readMaster(M_Data(po), _wDate)
            FormPopupNew.Labels(1).Location = New Point(x1 + Lot0_x + x2, FormPopupNew.Labels(0).Top)
            FormPopupNew.Labels(1).Size = New System.Drawing.Size(Lot_x, Lot_y)

            'LotNo
            FormPopupNew.Labels(2).Text = "LotNo"
            FormPopupNew.Labels(2).Location = New Point(FormPopupNew.Labels(0).Left, y0 + Lot0_y + y1)
            FormPopupNew.Labels(2).Size = New System.Drawing.Size(Lot0_x, Lot0_y)
            FormPopupNew.Labels(3).Text = readMaster(M_Data(po), _lot)
            FormPopupNew.Labels(3).Location = New Point(FormPopupNew.Labels(1).Left, FormPopupNew.Labels(2).Top)
            FormPopupNew.Labels(3).Size = New System.Drawing.Size(Lot_x, Lot_y)

            Dim ybuf As Integer = 0
            If UBound(FormPopupNew.Values0, 1) < 5 Then
                ybuf = (UBound(FormPopupNew.Values0, 1) + 1) * (Va0_y + y3) - y3 + y4
            Else
                ybuf = 5 * (Va0_y + y3) - y3 + y4
            End If

            'Ave
            FormPopupNew.Labels(4).Text = "Ave"
            FormPopupNew.Labels(4).Location = New Point(FormPopupNew.Labels(0).Left, FormPopupNew.Labels(2).Bottom + y2 + ybuf)
            FormPopupNew.Labels(4).Size = New System.Drawing.Size(Lot0_x, Lot0_y)
            FormPopupNew.Labels(5).Text = readMaster(M_Data(po), _X)
            FormPopupNew.Labels(5).Location = New Point(FormPopupNew.Labels(1).Left, FormPopupNew.Labels(4).Top)
            FormPopupNew.Labels(5).Size = New System.Drawing.Size(Lot0_x, Lot0_y)

            'Operater
            FormPopupNew.Labels(6).Text = "OP"
            FormPopupNew.Labels(6).Location = New Point(FormPopupNew.Labels(0).Left, FormPopupNew.Labels(4).Bottom + y5)
            FormPopupNew.Labels(6).Size = New System.Drawing.Size(Lot0_x, Lot0_y)
            FormPopupNew.Labels(7).Text = readMaster(M_Data(po), _opName)
            FormPopupNew.Labels(7).Location = New Point(FormPopupNew.Labels(1).Left, FormPopupNew.Labels(6).Top)
            FormPopupNew.Labels(7).Size = New System.Drawing.Size(Lot_x, Lot0_y)

            'Range
            FormPopupNew.Labels(8).Text = "Range"
            FormPopupNew.Labels(8).Location = New Point(FormPopupNew.Labels(5).Right + x3, FormPopupNew.Labels(4).Top)
            FormPopupNew.Labels(8).Size = New System.Drawing.Size(Lot0_x, Lot0_y)
            FormPopupNew.Labels(9).Text = readMaster(M_Data(po), _R)
            FormPopupNew.Labels(9).Location = New Point(FormPopupNew.Labels(8).Right + x2, FormPopupNew.Labels(8).Top)
            FormPopupNew.Labels(9).Size = New System.Drawing.Size(Lot0_x, Lot0_y)



            Dim SPCMes(8 - 1) As String
            If StrLanguage = "Japanese" Then
                SPCMes(0) = "?1??3???????"
                SPCMes(1) = "?8?????????"
                SPCMes(2) = "?3????2??2???????"
                SPCMes(3) = "?5????4??1???????"
                SPCMes(4) = "?15????1???????"
                SPCMes(5) = "?8????1???????"
                SPCMes(6) = "?7?????or??"
                SPCMes(7) = "?14???????????"
            ElseIf StrLanguage = "English" Then
                SPCMes(0) = "?Any single data point falls outside The 3? limit from the centerline"
                SPCMes(1) = "?Eight consecutive points fall on the same side of the centerline"
                SPCMes(2) = "?Two out of three consecutive points fall beyond the 2? limit"
                SPCMes(3) = "?Four out of five consecutive points fall beyond the 1? limit"
                SPCMes(4) = "?Fifteen consective points fall within ?1?"
                SPCMes(5) = "?Eight consective points fall beyond the 1? limit"
                SPCMes(6) = "?Seven consective points fall continuous rise or descent"
                SPCMes(7) = "?Fourteen consective points fall alternate up and down"
            End If
            Dim p As Integer = 0
            If Mode = "X" Then
                p = 0
            ElseIf Mode = "R" Then
                p = 1
            ElseIf Mode = "MR" Then
                p = 2
            End If

            Dim naiyou As String = "00000000"

            naiyou = readMaster(M_Alarm(po)(p), 1)
            If Not naiyou = "" Then
                For i As Integer = 0 To naiyou.Length - 1
                    If CBool(naiyou.Substring(i, 1)) Then
                        strAlarmName &= SPCMes(i)
                    End If
                Next
            End If





            'Alarm
            FormPopupNew.Labels(10).Text = "Alarm"
            FormPopupNew.Labels(10).Location = New Point(FormPopupNew.Labels(0).Left, FormPopupNew.Labels(6).Bottom + y6)
            FormPopupNew.Labels(10).Size = New System.Drawing.Size(Lot0_x, Lot0_y)
            FormPopupNew.Labels(11).Text = strAlarmName
            FormPopupNew.Labels(11).Location = New Point(FormPopupNew.Labels(1).Left, FormPopupNew.Labels(10).Top)
            FormPopupNew.Labels(11).Size = New System.Drawing.Size(Lot_x, Lot0_y * 4 - 15)
            FormPopupNew.Labels(11).TextAlign = ContentAlignment.TopLeft

            FormPopupNew.Controls.AddRange(FormPopupNew.Labels)
            FormPopupNew.ResumeLayout(False)


            FormPopupNew.Width = x1 + ((UBound(FormPopupNew.Values, 1) \ 5) + 1) * (Va0_x + x2 + Va_x + x3) - x3 + x4 + 6 '+6????
            FormPopupNew.Height = FormPopupNew.Labels(11).Bottom + y7 + 29


            If FormPopupNew.Width < x1 + Va0_x + x2 + Va_x + x3 + Va0_x + x2 + Va_x + x4 + 6 Then
                FormPopupNew.Width = x1 + Va0_x + x2 + Va_x + x3 + Va0_x + x2 + Va_x + x4 + 6
            End If
            'MsgBox(x1 & " " & (UBound(FormPopupNew.Values, 1) \ 5) & " " & (Va0_x + x2 + Va_x + x3) - x3 + x4)
            FormPopupNew.Show()

        End If

    End Sub


    Public Sub alarmInfo(ByVal mx1 As Integer, ByVal my1 As Integer, ByVal objName As String, ByVal btnName As String)
        If M_Data Is Nothing Then Exit Sub
        FormAlarmDisp.Close()
        FormAlarmInput.Close()
        FormPopupNew.Close()

        Dim dc, i As Integer

        Dim _mode As String

        If objName = "PictureBox1" Then
            _mode = "X"
        ElseIf objName = "PictureBox2" Then
            _mode = "R"
        Else
            Exit Sub
        End If

        If _mode = "R" Then
            If MRFlag = True Then
                _mode = "MR"
            End If
        End If


        dc = 0
        For i = 0 To 31

            If _mode = "X" Then

                If Not (xpnbuf_X(i) - 10 < mx1 And xpnbuf_X(i) + 10 > mx1) Then Continue For
                ypnbuf_X(i) = my1
                dc = 1
                Exit For

            ElseIf _mode = "R" Or _mode = "MR" Then

                If Not (xpnbuf_R(i) - 10 < mx1 And xpnbuf_R(i) + 10 > mx1) Then Continue For
                ypnbuf_R(i) = my1
                dc = 1
                Exit For

            End If
        Next


        If dc = 0 Then Exit Sub

        SerectPoint = DispStartPosition + i

        Dim p As Integer = 0
        If _mode = "X" Then
            p = 0
        ElseIf _mode = "R" Then
            p = 1
        ElseIf _mode = "MR" Then
            p = 2
        Else
            Exit Sub
        End If

        If InStr(M_Alarm(SerectPoint)(p), "1") = 0 Then Exit Sub '??????


        If btnName = "Right" Then '?????
            '???????????????????????
            Get_AlarmInfo("Write", _mode)
            FormAlarmInput.Show()
        ElseIf btnName = "Left" Then    '???????????????????????
            '??????????????
            Get_AlarmInfo("Read", _mode)
            FormAlarmDisp.Show()
        End If


    End Sub

    '??????????????????
    Public Sub Get_AlarmInfo(ByVal _RorW As String, ByVal _mode As String) '???ID???? ModeRead or Write
        Dim strID As String = readMaster(M_Data(SerectPoint), _id)

        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim Adapter As New SqlDataAdapter
        Dim strSQL As String = ""
        Dim table As New DataTable
        Dim n As Integer
        Try

            Cn.ConnectionString = StrServerConnection
            table.Clear()

            strSQL = "SELECT *"
            strSQL &= " FROM SPC_Alarm"
            strSQL &= " WHERE iID = '" & strID & "'"

            For i As Integer = 0 To UBound(TreeName, 1)
                strSQL &= " AND"
                strSQL &= " cTreeName" & i + 1 & " = '" & TreeName(i) & "'"
            Next

            strSQL &= " AND cGraphFormat = '" & _mode & "'"

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


            Display_AlarmInfo(table, _RorW, _mode)

        Catch ex As System.Exception
            Adapter.Dispose()
            Cn.Dispose()

            StrErrMes = "????????????" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Sub
        End Try
    End Sub

    '?????????????
    ' [Module2.vb]
    ' ก๊อปปี้ไปวางทับ Sub Display_AlarmInfo ของเดิมทั้งหมด
    Public Sub Display_AlarmInfo(ByVal _table As DataTable, ByVal _RorW As String, ByVal _mode As String)
        Dim strAlm As String = ""
        Try
            ' [ป้องกัน 1] เช็คก่อนว่ามีตารางและมีข้อมูลไหม
            If _table Is Nothing OrElse _table.Rows.Count = 0 Then
                ' ถ้าไม่มีข้อมูล ให้จบการทำงานเลย (กันเด้ง)
                Exit Sub
            End If

            If StrLanguage = "Japanese" Then
                If _table.Rows(0)("cSpcrule1").ToString = "True" Then
                    strAlm &= "?1??3???????"
                ElseIf _table.Rows(0)("cSpcrule2").ToString = "True" Then
                    strAlm &= "?8?????????"
                ElseIf _table.Rows(0)("cSpcrule3").ToString = "True" Then
                    strAlm &= "?3????2??2???????"
                ElseIf _table.Rows(0)("cSpcrule4").ToString = "True" Then
                    strAlm &= "?5????4??1???????"
                ElseIf _table.Rows(0)("cSpcrule5").ToString = "True" Then
                    strAlm &= "?15????1???????"
                ElseIf _table.Rows(0)("cSpcrule6").ToString = "True" Then
                    strAlm &= "?8????1???????"
                ElseIf _table.Rows(0)("cSpcrule7").ToString = "True" Then
                    strAlm &= "?7?????or??"
                ElseIf _table.Rows(0)("cSpcrule8").ToString = "True" Then
                    strAlm &= "?14???????????"
                End If
            ElseIf StrLanguage = "English" Then
                If _table.Rows(0)("cSpcrule1").ToString = "True" Then
                    strAlm &= "?Any single data point falls outside The 3? limit from the centerline"
                ElseIf _table.Rows(0)("cSpcrule2").ToString = "True" Then
                    strAlm &= "?Eight consecutive points fall on the same side of the centerline"
                ElseIf _table.Rows(0)("cSpcrule3").ToString = "True" Then
                    strAlm &= "?Two out of three consecutive points fall beyond the 2? limit"
                ElseIf _table.Rows(0)("cSpcrule4").ToString = "True" Then
                    strAlm &= "?Four out of five consecutive points fall beyond the 1? limit"
                ElseIf _table.Rows(0)("cSpcrule5").ToString = "True" Then
                    strAlm &= "?Fifteen consective points fall within ?1?"
                ElseIf _table.Rows(0)("cSpcrule6").ToString = "True" Then
                    strAlm &= "?Eight consective points fall beyond the 1? limit"
                ElseIf _table.Rows(0)("cSpcrule7").ToString = "True" Then
                    strAlm &= "?Seven consective points fall continuous rise or descent"
                ElseIf _table.Rows(0)("cSpcrule8").ToString = "True" Then
                    strAlm &= "?Fourteen consective points fall alternate up and down"
                End If
            End If

            If _RorW = "Read" Then
                ' [ป้องกัน 2] ใช้ TryCatch ย่อย หรือเช็ค Null ก่อนเสมอ
                FormAlarmDisp.LabDate.Text = readMaster(M_Data(SerectPoint), _wDate)
                FormAlarmDisp.LabSPCAlarm.Text = strAlm

                If Not IsDBNull(_table.Rows(0)("cMaintenanceID")) Then
                    FormAlarmDisp.LabAnother.Text = _table.Rows(0)("cMaintenanceID")
                Else
                    FormAlarmDisp.LabAnother.Text = ""
                End If
                If Not IsDBNull(_table.Rows(0)("cApproverName")) Then
                    FormAlarmDisp.LabQC.Text = _table.Rows(0)("cApproverName")
                Else
                    FormAlarmDisp.LabQC.Text = ""
                End If
                If Not IsDBNull(_table.Rows(0)("cTreatEffect")) Then
                    FormAlarmDisp.LabCheck.Text = _table.Rows(0)("cTreatEffect")
                Else
                    FormAlarmDisp.LabCheck.Text = ""
                End If
                If Not IsDBNull(_table.Rows(0)("cTreatResult")) Then
                    FormAlarmDisp.LabAction.Text = _table.Rows(0)("cTreatResult")
                Else
                    FormAlarmDisp.LabAction.Text = ""
                End If
                If Not IsDBNull(_table.Rows(0)("cTreatIncharge")) Then
                    FormAlarmDisp.LabPerson2.Text = _table.Rows(0)("cTreatIncharge")
                Else
                    FormAlarmDisp.LabPerson2.Text = ""
                End If
                If Not IsDBNull(_table.Rows(0)("cSurveyResult")) Then
                    FormAlarmDisp.LabResult.Text = _table.Rows(0)("cSurveyResult")
                Else
                    FormAlarmDisp.LabResult.Text = ""
                End If
                If Not IsDBNull(_table.Rows(0)("cSurveyIncharge")) Then
                    FormAlarmDisp.LabPerson.Text = _table.Rows(0)("cSurveyIncharge")
                Else
                    FormAlarmDisp.LabPerson.Text = ""
                End If

                FormAlarmDisp.TextMode.Text = _mode

            ElseIf _RorW = "Write" Then

                FormAlarmInput.TextDate.Text = readMaster(M_Data(SerectPoint), _wDate)
                FormAlarmInput.TextSPCAlarm.Text = strAlm

                If Not IsDBNull(_table.Rows(0)("cMaintenanceID")) Then
                    FormAlarmInput.TextAnother.Text = _table.Rows(0)("cMaintenanceID")
                Else
                    FormAlarmInput.TextAnother.Text = ""
                End If
                If Not IsDBNull(_table.Rows(0)("cApproverName")) Then
                    FormAlarmInput.TextQC.Text = _table.Rows(0)("cApproverName")
                Else
                    FormAlarmInput.TextQC.Text = ""
                End If
                If Not IsDBNull(_table.Rows(0)("cTreatEffect")) Then
                    FormAlarmInput.TextCheck.Text = _table.Rows(0)("cTreatEffect")
                Else
                    FormAlarmInput.TextCheck.Text = ""
                End If
                If Not IsDBNull(_table.Rows(0)("cTreatResult")) Then
                    FormAlarmInput.TextAction.Text = _table.Rows(0)("cTreatResult")
                Else
                    FormAlarmInput.TextAction.Text = ""
                End If
                If Not IsDBNull(_table.Rows(0)("cTreatIncharge")) Then
                    FormAlarmInput.TextPerson2.Text = _table.Rows(0)("cTreatIncharge")
                Else
                    FormAlarmInput.TextPerson2.Text = ""
                End If
                If Not IsDBNull(_table.Rows(0)("cSurveyResult")) Then
                    FormAlarmInput.TextResult.Text = _table.Rows(0)("cSurveyResult")
                Else
                    FormAlarmInput.TextResult.Text = ""
                End If
                If Not IsDBNull(_table.Rows(0)("cSurveyIncharge")) Then
                    FormAlarmInput.TextPerson.Text = _table.Rows(0)("cSurveyIncharge")
                Else
                    FormAlarmInput.TextPerson.Text = ""
                End If

                FormAlarmInput.TextMode.Text = _mode
            End If

        Catch ex As Exception
            ' บันทึก Error แต่ไม่เด้งปิดโปรแกรม
            StrErrMes = "Alarm Info Display Error: " & ex.Message
            Call SaveLog(Now(), StrErrMes)
        End Try
    End Sub

#End Region



End Module
