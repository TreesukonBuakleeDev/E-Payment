Imports System.Data
Imports System.Data.SqlClient

Public Class clsCValidate

    Public Function ValidateData(ByRef dtGridview As DataGridView, ByVal Progressbar1 As ProgressBar) As Boolean 'ปุ่ม Export เรียกไม่มี Report 
        Dim ret As Boolean = True

        For i = 0 To dtGridview.RowCount - 1
            If IsDBNull(dtGridview.Rows(i).Cells("Check").Value) Then
                Continue For
            End If

            Progressbar1.Value = 10

            If dtGridview.Rows(i).Cells("Check").Value = True Then
                Dim StrCheck As String
                connection = New SqlConnection(conStr)
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                sql1 = "SELECT * FROM APTCR WHERE CNTBTCH = " & dtGridview.Rows(i).Cells("BatchNo").Value
                command = New SqlCommand(sql1, connection)
                adapter = New SqlDataAdapter(command)
                dataSt = New DataSet()
                adapter.Fill(dataSt, "DataFromEntry")
                For k = 0 To dataSt.Tables("DataFromEntry").Rows.Count - 1
                    StrCheck = Nothing
                    connection = New SqlConnection(conStr)
                    If connection.State = ConnectionState.Closed Then
                        connection.Open()
                    End If
                    sql1 = "SELECT * FROM APVNR WHERE IDVEND = '" & dataSt.Tables("DataFromEntry").Rows(k)("IDVEND") & "'"
                    command = New SqlCommand(sql1, connection)
                    adapter = New SqlDataAdapter(command)
                    adapter.Fill(dataSt, "DataPayment")
                    If dataSt.Tables("DataPayment").Rows.Count < 1 Then
                        MessageBox.Show("Please Set Remit To Location")
                    ElseIf dataSt.Tables("DataPayment").Rows.Count > 0 Then
                        If dataSt.Tables("DataPayment").Rows(k)("IDVENDRMIT").ToString.Trim <> "BTMU" Then
                            StrCheck &= "No BTMU" & vbCrLf
                            ret = False
                        End If
                        If dataSt.Tables("DataPayment").Rows(k)("RMITNAME").ToString.Trim = "" Then
                            StrCheck &= ":" & "No RMITNAME" & vbCrLf
                            ret = False
                        End If
                        If dataSt.Tables("DataPayment").Rows(k)("TEXTSTRE1").ToString.Trim = "" Then
                            StrCheck &= ":" & "No Address" & vbCrLf
                            ret = False
                        End If
                        If dataSt.Tables("DataPayment").Rows(k)("NAMECITY").ToString.Trim = "" Then
                            StrCheck &= ":" & "No Bank Code" & vbCrLf  'NAMECITY
                            ret = False
                        End If
                        If dataSt.Tables("DataPayment").Rows(k)("CODESTTE").ToString.Trim = "" Then
                            StrCheck &= ":" & "No Branch Code" & vbCrLf 'CODESTTE
                            ret = False
                        End If
                        If dataSt.Tables("DataPayment").Rows(k)("CODEPSTL").ToString.Trim = "" Then
                            StrCheck &= ":" & "No Account Number" & vbCrLf 'CODEPSTL
                            ret = False
                        End If
                        'If dataSt.Tables("DataPayment").Rows(k)("EMAIL").ToString.Trim = "" Then
                        '    StrCheck &= ":" & "No EMAIL" & vbCrLf
                        '    ret = False
                        'End If
                        If dataSt.Tables("DataPayment").Rows(k)("CODECTRY").ToString.Trim = "" Then
                            StrCheck &= ":" & "No Service Type" & vbCrLf 'CODECTRY
                            ret = False
                        End If
                        'If dataSt.Tables("DataPayment").Rows(k)("TEXTPHON1").ToString.Trim = "" Then
                        '    StrCheck &= ":" & "No Telephone" & vbCrLf 'TEXTPHON1
                        '    ret = False
                        'End If
                        'If dataSt.Tables("DataPayment").Rows(k)("TEXTPHON2").ToString.Trim = "" Then
                        '    StrCheck &= ":" & "No Fax" & vbCrLf 'TEXTPHON2
                        '    ret = False
                        'End If
                    End If
                    ' If ret = False Then
                    If Not (StrCheck Is Nothing) Then
                        MessageBox.Show("BatchNo:" & dataSt.Tables("DataFromEntry").Rows(k)("CNTBTCH") & "-" & _
                                        dataSt.Tables("DataFromEntry").Rows(k)("CNTENTR") & "  " & vbCrLf & _
                                        "IDVEND :" & dataSt.Tables("DataFromEntry").Rows(k)("IDVEND") & "  " & vbCrLf & _
                                        StrCheck, "Remit To Location", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        dtGridview.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    End If
                Next

                If Progressbar1.Value = 100 Then
                    Progressbar1.Value = 0
                    Progressbar1.Value += 10
                End If
                Progressbar1.Value += 10

            End If

            Progressbar1.Value = 100

        Next
        connection.Close()
        Return ret
    End Function

    Public Function ValidateDataWithReport(ByRef dtGridview As DataGridView, ByRef dtInvalid As DataTable) As Boolean
        Dim ret As Boolean = True
        dtInvalid.TableName = "dtInvalid"
        dtInvalid.Columns.Add("IDVEND") : dtInvalid.Columns.Add("IDVENDRMIT")
        dtInvalid.Columns.Add("RMITNAME") : dtInvalid.Columns.Add("TEXTSTRE1")
        dtInvalid.Columns.Add("NAMECITY") : dtInvalid.Columns.Add("CODESTTE")
        dtInvalid.Columns.Add("CODEPSTL") : dtInvalid.Columns.Add("EMAIL")
        dtInvalid.Columns.Add("CODECTRY") : dtInvalid.Columns.Add("TEXTPHON1")
        dtInvalid.Columns.Add("TEXTPHON2")
        Dim IDVEND, IDVENDRMIT, RMITNAME, TEXTSTRE1, NAMECITY, CODESTTE, CODEPSTL, EMAIL, CODECTRY, TEXTPHON1, TEXTPHON2 As String

        For i = 0 To dtGridview.RowCount - 1
            If IsDBNull(dtGridview.Rows(i).Cells("Check").Value) Then
                Continue For
            End If
            If dtGridview.Rows(i).Cells("Check").Value = True Then
                Dim StrCheck As String
                connection = New SqlConnection(conStr)
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                sql1 = "SELECT * FROM APTCR WHERE CNTBTCH = " & dtGridview.Rows(i).Cells("BatchNo").Value
                command = New SqlCommand(sql1, connection)
                adapter = New SqlDataAdapter(command)
                dataSt = New DataSet()
                adapter.Fill(dataSt, "DataFromEntry")
                For k = 0 To dataSt.Tables("DataFromEntry").Rows.Count - 1
                    IDVEND = Nothing : IDVENDRMIT = Nothing : RMITNAME = Nothing : TEXTSTRE1 = Nothing : NAMECITY = Nothing
                    CODESTTE = Nothing : CODEPSTL = Nothing : EMAIL = Nothing : CODECTRY = Nothing : TEXTPHON1 = Nothing : TEXTPHON2 = Nothing
                    StrCheck = Nothing
                    connection = New SqlConnection(conStr)
                    If connection.State = ConnectionState.Closed Then
                        connection.Open()
                    End If
                    sql1 = "SELECT * FROM APVNR WHERE IDVEND = '" & dataSt.Tables("DataFromEntry").Rows(k)("IDVEND") & "'"
                    command = New SqlCommand(sql1, connection)
                    adapter = New SqlDataAdapter(command)
                    adapter.Fill(dataSt, "DataPayment")
                    If dataSt.Tables("DataPayment").Rows.Count < 1 Then
                        MessageBox.Show("Please Set Remit To Location")
                    ElseIf dataSt.Tables("DataPayment").Rows.Count > 0 Then
                        IDVEND = dataSt.Tables("DataPayment").Rows(k)("IDVEND")
                        If dataSt.Tables("DataPayment").Rows(k)("IDVENDRMIT").ToString.Trim <> "BTMU" Then
                            StrCheck &= "No BTMU" & vbCrLf
                            'IDVENDRMIT = "None BTMU"
                            IDVENDRMIT = "Empty"
                            ret = False
                        End If
                        If dataSt.Tables("DataPayment").Rows(k)("RMITNAME").ToString.Trim = "" Then
                            StrCheck &= ":" & "None RMITNAME" & vbCrLf
                            'RMITNAME = "None RMITNAME"
                            RMITNAME = "Empty"
                            ret = False
                        End If
                        If dataSt.Tables("DataPayment").Rows(k)("TEXTSTRE1").ToString.Trim = "" Then
                            StrCheck &= ":" & "None Address" & vbCrLf
                            'TEXTSTRE1 = "None Address"
                            TEXTSTRE1 = "Empty"
                            ret = False
                        End If
                        If dataSt.Tables("DataPayment").Rows(k)("NAMECITY").ToString.Trim = "" Then
                            StrCheck &= ":" & "None BankCode" & vbCrLf  'NAMECITY
                            'NAMECITY = "None Bankcode"
                            NAMECITY = "Empty"
                            ret = False
                        End If
                        If dataSt.Tables("DataPayment").Rows(k)("CODESTTE").ToString.Trim = "" Then
                            StrCheck &= ":" & "None BranchCode" & vbCrLf 'CODESTTE
                            'CODESTTE = "None BranchCode"
                            CODESTTE = "Empty"
                            ret = False
                        End If
                        If dataSt.Tables("DataPayment").Rows(k)("CODEPSTL").ToString.Trim = "" Then
                            StrCheck &= ":" & "None AccountNumber" & vbCrLf 'CODEPSTL
                            'CODEPSTL = "None Account Number"
                            CODEPSTL = "Empty"
                            ret = False
                        End If
                        'If dataSt.Tables("DataPayment").Rows(k)("EMAIL").ToString.Trim = "" Then
                        '    StrCheck &= ":" & "None EMAIL" & vbCrLf
                        '    'EMAIL = "None Email"
                        '    EMAIL = "Empty"
                        '    ret = False
                        'End If
                        If dataSt.Tables("DataPayment").Rows(k)("CODECTRY").ToString.Trim = "" Then
                            StrCheck &= ":" & "None Service Type" & vbCrLf 'CODECTRY
                            'CODECTRY = "None Service Type"
                            CODECTRY = "Empty"
                            ret = False
                        End If
                        'If dataSt.Tables("DataPayment").Rows(k)("TEXTPHON1").ToString.Trim = "" Then
                        '    StrCheck &= ":" & "None Telephone" & vbCrLf 'TEXTPHON1
                        '    'TEXTPHON1 = "None Telephone"
                        '    TEXTPHON1 = "Empty"
                        '    ret = False
                        'End If
                        'If dataSt.Tables("DataPayment").Rows(k)("TEXTPHON2").ToString.Trim = "" Then
                        '    StrCheck &= ":" & "None Fax" & vbCrLf 'TEXTPHON2
                        '    'TEXTPHON2 = "None Fax"
                        '    TEXTPHON2 = "Empty"
                        '    ret = False
                        'End If
                    End If
                    'If ret = False Then
                    If Not (StrCheck Is Nothing) Then
                        dtInvalid.Rows.Add(IDVEND, IDVENDRMIT, RMITNAME, TEXTSTRE1, NAMECITY, CODESTTE, CODEPSTL, _
                                           EMAIL, CODECTRY, TEXTPHON1, TEXTPHON2)
                        dtGridview.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    End If
                Next
            End If
        Next
        If Not ret Then
            Dim frm As New clsVCrystalReportPreview
            frm = New clsVCrystalReportPreview(dtInvalid)
            frm.ShowDialog()
        End If
        connection.Close()
        Return ret
    End Function

    Public Function ValidateReference(ByVal Datagridview As DataGridView) As Boolean
        Dim ret As Boolean = True
        For i = 0 To Datagridview.Rows.Count - 1
            If IsDBNull(Datagridview.Rows(i).Cells("Check").Value) Then
                Continue For
            End If
            If Datagridview.Rows(i).Cells("Check").Value = True Then
                connection = New SqlConnection(conStr)
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                sql1 = "SELECT CNTBTCH,CNTRMIT,TEXTREF,CNTRMIT FROM APTCU WHERE CNTBTCH = " & Datagridview.Rows(i).Cells("BatchNo").Value
                sql1 &= " AND (substring(IDDISTCODE,1,2) = 'T3' Or Substring(IDDISTCODE,1,3) = 'T53') "
                command = New SqlCommand(sql1, connection)
                adapter = New SqlDataAdapter(command)
                dataSt = New DataSet()
                adapter.Fill(dataSt, "DataReference")
                For k = 0 To dataSt.Tables("DataReference").Rows.Count - 1
                    Dim REFArr() As String = dataSt.Tables("DataReference").Rows(k)("TEXTREF").ToString.Split(",")
                    Dim IDVEND_APVEN As String = ""

                    If REFArr(0) = "" Then
                        sql1 = "SELECT * " & vbCrLf
                        sql1 &= "   FROM APVEN " & vbCrLf
                        sql1 &= "WHERE VENDORID = " & vbCrLf
                        sql1 &= "               (SELECT IDVEND " & vbCrLf
                        sql1 &= "                FROM APTCR " & vbCrLf
                        sql1 &= "                WHERE CNTBTCH = " & dataSt.Tables("DataReference").Rows(k)("CNTBTCH") & " " & vbCrLf
                        sql1 &= "                      AND CNTENTR = " & dataSt.Tables("DataReference").Rows(k)("CNTRMIT") & ")"
                    ElseIf REFArr(0) <> "" Then
                        sql1 = "SELECT * FROM APVEN WHERE VENDORID = '" & REFArr(0) & "'"
                    End If


                    command = New SqlCommand(sql1, connection)
                    Dim Result As String = command.ExecuteScalar()
                    If Result Is Nothing Then
                        Datagridview.Rows(i).DefaultCellStyle.BackColor = Color.Green
                        MessageBox.Show("BatchNo :" & dataSt.Tables("DataReference").Rows(k)("CNTBTCH") _
                                        & "-" & dataSt.Tables("DataReference").Rows(k)("CNTRMIT") & vbCrLf _
                                        & "IDVEND : " & REFArr(0) & " Invalid" _
                                        , "Data WHT Invalid ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ret = False
                    End If
                Next
            End If
        Next
        connection.Close()
        Return ret
    End Function

    Public Function ValidateReferenceWithReport(ByVal Datagridview As DataGridView) As Boolean
        Dim ret As Boolean = True
        Dim dtReferenceInvalid As New DataTable
        dtReferenceInvalid.TableName = "dtReferenceInvalid"
        dtReferenceInvalid.Columns.Add("BatchNo") : dtReferenceInvalid.Columns.Add("Entry")
        dtReferenceInvalid.Columns.Add("Line") : dtReferenceInvalid.Columns.Add("Reference")

        For i = 0 To Datagridview.Rows.Count - 1
            If IsDBNull(Datagridview.Rows(i).Cells("Check").Value) Then
                Continue For
            End If
            If Datagridview.Rows(i).Cells("Check").Value = True Then
                connection = New SqlConnection(conStr)
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                sql1 = "SELECT CNTBTCH,CNTRMIT,TEXTREF,CNTSEQ FROM APTCU WHERE CNTBTCH = " & Datagridview.Rows(i).Cells("BatchNo").Value
                sql1 &= " AND (substring(IDDISTCODE,1,2) = 'T3' Or Substring(IDDISTCODE,1,3) = 'T53') "
                command = New SqlCommand(sql1, connection)
                adapter = New SqlDataAdapter(command)
                dataSt = New DataSet()
                adapter.Fill(dataSt, "DataReference")
                For k = 0 To dataSt.Tables("DataReference").Rows.Count - 1
                    Dim REFArr() As String = dataSt.Tables("DataReference").Rows(k)("TEXTREF").ToString.Split(",")
                    If REFArr(0) = "" Then
                        sql1 = "SELECT * " & vbCrLf
                        sql1 &= "   FROM APVEN " & vbCrLf
                        sql1 &= "WHERE VENDORID = " & vbCrLf
                        sql1 &= "               (SELECT IDVEND " & vbCrLf
                        sql1 &= "                FROM APTCR " & vbCrLf
                        sql1 &= "                WHERE CNTBTCH = " & dataSt.Tables("DataReference").Rows(k)("CNTBTCH") & " " & vbCrLf
                        sql1 &= "                      AND CNTENTR = " & dataSt.Tables("DataReference").Rows(k)("CNTRMIT") & ")"
                    ElseIf REFArr(0) <> "" Then
                        sql1 = "SELECT * FROM APVEN WHERE VENDORID = '" & REFArr(0) & "'"
                    End If
                    command = New SqlCommand(sql1, connection)
                    Dim Result As String = command.ExecuteScalar()
                    If Result Is Nothing Then
                        Datagridview.Rows(i).DefaultCellStyle.BackColor = Color.Green
                        dtReferenceInvalid.Rows.Add(dataSt.Tables("DataReference").Rows(k)("CNTBTCH"), _
                                                    dataSt.Tables("DataReference").Rows(k)("CNTRMIT"), _
                                                    dataSt.Tables("DataReference").Rows(k)("CNTSEQ"), _
                                                    dataSt.Tables("DataReference").Rows(k)("TEXTREF"))
                        ret = False
                    End If
                Next
            End If
        Next
        If Not ret Then
            Dim frm As New clsVCrystalReportPreview
            frm = New clsVCrystalReportPreview(dtReferenceInvalid)
            frm.ShowDialog()
        End If
        connection.Close()

        Return ret
    End Function

End Class





'Public Function ValidateData(ByRef dtGridview As DataGridView, Optional ByRef dtInvalid As DataTable = Nothing) As Boolean
'    Dim ret As Boolean = True

'    dtInvalid.Columns.Add("IDVEND") : dtInvalid.Columns.Add("IDVENDRMIT")
'    dtInvalid.Columns.Add("RMITNAME") : dtInvalid.Columns.Add("TEXTSTRE1")
'    dtInvalid.Columns.Add("NAMECITY") : dtInvalid.Columns.Add("CODESTTE")
'    dtInvalid.Columns.Add("CODEPSTL") : dtInvalid.Columns.Add("EMAIL")
'    dtInvalid.Columns.Add("CODECTRY") : dtInvalid.Columns.Add("TEXTPHON1")
'    dtInvalid.Columns.Add("TEXTPHON2")
'    Dim IDVEND, IDVENDRMIT, RMITNAME, TEXTSTRE1, NAMECITY, CODESTTE, CODEPSTL, EMAIL, CODECTRY, TEXTPHON1, TEXTPHON2 As String

'    For i = 0 To dtGridview.RowCount - 1
'        If dtGridview.Rows(i).Cells("Check").Value = True Then
'            Dim StrCheck As String = ""
'            connection = New SqlConnection(conStr)
'            If connection.State = ConnectionState.Closed Then
'                connection.Open()
'            End If
'            sql1 = "SELECT * FROM APTCR WHERE CNTBTCH = " & dtGridview.Rows(i).Cells("BatchNo").Value
'            command = New SqlCommand(sql1, connection)
'            adapter = New SqlDataAdapter(command)
'            dataSt = New DataSet()
'            adapter.Fill(dataSt, "DataFromEntry")
'            For k = 0 To dataSt.Tables("DataFromEntry").Rows.Count - 1
'                connection = New SqlConnection(conStr)
'                If connection.State = ConnectionState.Closed Then
'                    connection.Open()
'                End If
'                sql1 = "SELECT * FROM APVNR WHERE IDVEND = '" & dataSt.Tables("DataFromEntry").Rows(k)("IDVEND") & "'"
'                command = New SqlCommand(sql1, connection)
'                adapter = New SqlDataAdapter(command)
'                adapter.Fill(dataSt, "DataPayment")
'                If dataSt.Tables("DataPayment").Rows.Count < 1 Then
'                    MessageBox.Show("Please Set Remit To Location")
'                ElseIf dataSt.Tables("DataPayment").Rows.Count > 0 Then
'                    IDVEND = dataSt.Tables("DataPayment").Rows(0)("IDVEND")

'                    If dataSt.Tables("DataPayment").Rows(0)("IDVENDRMIT").ToString.Trim <> "BTMU" Then
'                        StrCheck &= "No BTMU" & vbCrLf
'                        IDVENDRMIT = "No BTMU"
'                        ret = False
'                    End If
'                    If dataSt.Tables("DataPayment").Rows(0)("RMITNAME").ToString.Trim = "" Then
'                        StrCheck &= ":" & "No RMITNAME" & vbCrLf
'                        RMITNAME = "No RMITNAME"
'                        ret = False
'                    End If
'                    If dataSt.Tables("DataPayment").Rows(0)("TEXTSTRE1").ToString.Trim = "" Then
'                        StrCheck &= ":" & "No Address" & vbCrLf
'                        TEXTSTRE1 = "No Address"
'                        ret = False
'                    End If
'                    If dataSt.Tables("DataPayment").Rows(0)("NAMECITY").ToString.Trim = "" Then
'                        StrCheck &= ":" & "No Bank Code" & vbCrLf  'NAMECITY
'                        NAMECITY = "No Bankcode"
'                        ret = False
'                    End If
'                    If dataSt.Tables("DataPayment").Rows(0)("CODESTTE").ToString.Trim = "" Then
'                        StrCheck &= ":" & "No Branch Code" & vbCrLf 'CODESTTE
'                        CODESTTE = "No BranchCode"
'                        ret = False
'                    End If
'                    If dataSt.Tables("DataPayment").Rows(0)("CODEPSTL").ToString.Trim = "" Then
'                        StrCheck &= ":" & "No Account Number" & vbCrLf 'CODEPSTL
'                        CODEPSTL = "No Account Number"
'                        ret = False
'                    End If
'                    If dataSt.Tables("DataPayment").Rows(0)("EMAIL").ToString.Trim = "" Then
'                        StrCheck &= ":" & "No EMAIL" & vbCrLf
'                        EMAIL = "No Email"
'                        ret = False
'                    End If
'                    If dataSt.Tables("DataPayment").Rows(0)("CODECTRY").ToString.Trim = "" Then
'                        StrCheck &= ":" & "No Service Type" & vbCrLf 'CODECTRY
'                        CODECTRY = "No Service Type"
'                        ret = False
'                    End If
'                    If dataSt.Tables("DataPayment").Rows(0)("TEXTPHON1").ToString.Trim = "" Then
'                        StrCheck &= ":" & "No Telephone" & vbCrLf 'TEXTPHON1
'                        TEXTPHON1 = "No Telephone"
'                        ret = False
'                    End If
'                    If dataSt.Tables("DataPayment").Rows(0)("TEXTPHON2").ToString.Trim = "" Then
'                        StrCheck &= ":" & "No Fax" & vbCrLf 'TEXTPHON2
'                        TEXTPHON2 = "NoFax"
'                        ret = False
'                    End If
'                End If
'                If ret = False Then
'                    MessageBox.Show("BatchNo:" & dataSt.Tables("DataFromEntry").Rows(k)("CNTBTCH") & "-" & _
'                                    dataSt.Tables("DataFromEntry").Rows(k)("CNTENTR") & "  " & vbCrLf & _
'                                    "IDVEND :" & dataSt.Tables("DataFromEntry").Rows(k)("IDVEND") & "  " & vbCrLf & _
'                                    StrCheck, "Remit To Location", MessageBoxButtons.OK, MessageBoxIcon.Error)

'                    dtInvalid.Rows.Add(IDVEND, IDVENDRMIT, RMITNAME, TEXTSTRE1, NAMECITY, CODESTTE, CODEPSTL, _
'                                       EMAIL, CODECTRY, TEXTPHON1, TEXTPHON2)

'                    dtGridview.Rows(i).DefaultCellStyle.BackColor = Color.Red
'                End If
'            Next
'        End If
'    Next

'    connection.Close()

'    Return ret
'End Function
