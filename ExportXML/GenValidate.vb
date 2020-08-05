Imports System.Data.SqlClient
Module GenValidate
    Public dtchk As DataTable = New DataTable()
    Public dtErr As DataTable = New DataTable()
    Public dtMC As DataTable = New DataTable()
    Public dtFT As DataTable = New DataTable()

#Region "Validate"
    Sub ValidateRmitto(ByVal DGVLIST As DataGridView)
        cleardt()
        Try
            dtchk = CType(DGVLIST.DataSource, DataTable).Copy
            dtErr = CType(DGVLIST.DataSource, DataTable).Copy

            For i = 0 To dtchk.Rows.Count - 1
                Dim Miscode As String = ""
                If IsDBNull(dtchk.Rows(i).Item("Check").ToString) Then
                    Continue For
                End If
                Dim CheckState As String = dtchk.Rows(i).Item("Check").ToString
                If CheckState = "True" Then
                    Miscode = dtchk.Rows(i).Item("VendorCode").ToString
                    Dim chkMiscode = chkRmitto(Miscode)
                    If chkMiscode.Rows.Count = 0 Then
                        dtchk.Rows(i).Delete()
                        '>>log null Rmitto Code
                    Else
                        dtErr.Rows(i).Delete()
                    End If
                Else
                    dtchk.Rows(i).Delete()
                    dtErr.Rows(i).Delete()
                End If
            Next
            dtchk.AcceptChanges()
            dtErr.AcceptChanges()
            DisplayValidate(dtErr)

            ValidateTypeMC(dtchk)
            ValidateTypeFT(dtchk)

        Catch ex As Exception
            MessageBox.Show("Error 43: " & ex.Message)
        End Try
    End Sub
    Sub ValidateTypeMC(ByVal dtchk As DataTable)
        dtMC = dtchk.Copy
        dtErr = dtchk.Copy

        Try


            For i = 0 To dtMC.Rows.Count - 1
                Dim batchno As String = ""
                Dim entryno As String = ""
                Dim BeneficiaryName As String = ""
                Dim TOTALAMOUNT As Integer
                Dim Acct, TaxId, RemittoAddress, OptService As String
                Dim Taxbase As String
                batchno = dtMC.Rows(i).Item("BatchNo").ToString
                entryno = dtMC.Rows(i).Item("Entry").ToString
                BeneficiaryName = dtMC.Rows(i).Item("VendorName").ToString
                TOTALAMOUNT = CInt(dtMC.Rows(i).Item("Amount").ToString)
                DataGETP = clsCGenerateFile.GetDataP(batchno)

                For p = 0 To DataGETP.Rows.Count - 1
                    Acct = DataGETP.Rows(0).Item("AccountNo").ToString.TrimEnd
                    TaxId = DataGETP.Rows(0).Item("TAXID").ToString.TrimEnd
                    RemittoAddress = DataGETP.Rows(0).Item("ADDR").ToString.TrimEnd
                    OptService = DataGETP.Rows(0).Item("COUNTRY").ToString.TrimEnd
                Next

                DataGETW = clsCGenerateFile.GetDataW(batchno, entryno)
                For w = 0 To DataGETW.Rows.Count - 1
                    Taxbase = DataGETW.Rows(0).Item("TAXBASE").ToString.TrimEnd
                Next
                Dim strErr As String = ""
                Dim checkMC = chkTypeMC(batchno, entryno)
                Dim checkMCBeneficiaryName = hasSpecialChar(BeneficiaryName)
                Dim chkOptService = hasDefineOptService(OptService)
                For j = 0 To checkMC.Rows.Count - 1

                    If checkMC.Rows(j).Item("BankCode").ToString.TrimEnd = "039" Then
                        dtMC.Rows(i).Delete()
                        Exit For
                    End If

                    If checkMCBeneficiaryName = True Or IsDBNull(BeneficiaryName) Or BeneficiaryName = "" Then
                        dtMC.Rows(i).Delete()
                        strErr = "1"
                        '>>log Rmitto Name  = "" or Funny character
                    End If

                    If TOTALAMOUNT = 0 Or IsDBNull(TOTALAMOUNT) Or TOTALAMOUNT.ToString = "" Then
                        dtMC.Rows(i).Delete()
                        strErr = strErr & "2"
                        'log TOTALAMT = 0 or  TOTALAMT = ""
                    End If

                    If Taxbase = 0 Or IsDBNull(Taxbase) Or Taxbase.ToString = "" Then
                        dtMC.Rows(i).Delete()
                        strErr = strErr & "3"
                        'log Taxbase = 0
                    End If

                    If Acct = "" Or IsDBNull(Acct) Then
                        dtMC.Rows(i).Delete()
                        strErr = strErr & "4"
                        'log Account Number = ""
                    Else
                        If Acct.Length > 12 Then
                            dtMC.Rows(i).Delete()
                            strErr = strErr & "5"
                            'log Account number > 12 digit
                        End If
                    End If

                    If IsNumeric(TaxId) = True Then
                        Select Case TaxId.Length
                            Case 10
                            Case 13
                            Case Else
                                dtMC.Rows(i).Delete()
                                strErr = strErr & "6"
                        End Select
                        'log Tax ID <> 10 or Tax ID <> 13
                    Else
                        dtMC.Rows(i).Delete()
                        strErr = strErr & "7"
                        'log TaxID not number

                    End If

                    If RemittoAddress = "" Or IsDBNull(RemittoAddress) Then
                        dtMC.Rows(i).Delete()
                        strErr = strErr & "8"
                        'log RemittoAddress = ""
                    End If

                    If chkOptService = False Then
                        dtMC.Rows(i).Delete()
                        strErr = strErr & "9"
                    End If

                Next
                If strErr <> "" Then

                Else
                    dtErr.Rows(i).Delete()
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Error 153: " & ex.Message)
        End Try
        dtMC.AcceptChanges()
        dtErr.AcceptChanges()
        DisplayValidate(dtErr)
    End Sub

    Sub ValidateTypeFT(ByVal dtchk As DataTable)
        dtFT = dtchk.Copy
        dtErr = dtchk.Copy

        Try
            For i = 0 To dtFT.Rows.Count - 1
                Dim batchno As String = ""
                Dim entryno As String = ""
                Dim BeneficiaryName As String = ""
                Dim TOTALAMOUNT, Taxbase As Integer
                Dim Acct, TaxId, RemittoAddress, OptService As String

                batchno = dtFT.Rows(i).Item("BatchNo").ToString
                entryno = dtFT.Rows(i).Item("Entry").ToString
                BeneficiaryName = dtFT.Rows(i).Item("VendorName").ToString
                TOTALAMOUNT = CInt(dtFT.Rows(i).Item("Amount").ToString)
                DataGETP = clsCGenerateFile.GetDataP(batchno)
                For p = 0 To DataGETP.Rows.Count - 1
                    Acct = DataGETP.Rows(0).Item("AccountNo").ToString.TrimEnd
                    TaxId = DataGETP.Rows(0).Item("TAXID").ToString.TrimEnd
                    RemittoAddress = DataGETP.Rows(0).Item("ADDR").ToString.TrimEnd
                    OptService = DataGETP.Rows(0).Item("COUNTRY").ToString.TrimEnd
                Next

                DataGETW = clsCGenerateFile.GetDataW(batchno, entryno)
                For w = 0 To DataGETW.Rows.Count - 1
                    Taxbase = DataGETW.Rows(0).Item("TAXBASE").ToString.TrimEnd
                Next

                Dim strErr As String = ""
                Dim checkMC = chkTypeMC(batchno, entryno)
                Dim checkMCBeneficiaryName = hasSpecialChar(BeneficiaryName)
                Dim chkOptService = hasDefineOptService(OptService)

                For j = 0 To checkMC.Rows.Count - 1
                    If checkMC.Rows(j).Item("BankCode").ToString.TrimEnd <> "039" Then
                        dtFT.Rows(i).Delete()
                        Exit For
                    End If

                    If checkMCBeneficiaryName = True Or IsDBNull(BeneficiaryName) Or BeneficiaryName = "" Then
                        dtFT.Rows(i).Delete()
                        strErr = strErr & "1"
                        '>>log Rmitto Name  = "" or Funny character
                    End If

                    If TOTALAMOUNT = 0 Or IsDBNull(TOTALAMOUNT) Or TOTALAMOUNT.ToString = "" Then
                        dtFT.Rows(i).Delete()
                        strErr = strErr & "2"
                        'log TOTALAMT = 0 or  TOTALAMT = ""
                    End If

                    If Taxbase = 0 Or IsDBNull(Taxbase) Or Taxbase.ToString = "" Then
                        dtFT.Rows(i).Delete()
                        strErr = strErr & "3"
                        'log Taxbase = 0 
                    End If

                    If Acct = "" Or IsDBNull(Acct) Then
                        dtFT.Rows(i).Delete()
                        strErr = strErr & "4"
                        'log Account Number = ""
                    Else
                        If Acct.Length > 12 Then
                            dtFT.Rows(i).Delete()
                            strErr = strErr & "5"
                            'log Account number > 12 digit
                        End If
                    End If

                    If IsNumeric(TaxId) = True Then
                        Select Case TaxId.Length
                            Case 10
                            Case 13
                            Case Else
                                dtFT.Rows(i).Delete()
                                strErr = strErr & "6"
                        End Select
                        'log Tax ID <> 10 or Tax ID <> 13
                    Else
                        dtFT.Rows(i).Delete()
                        strErr = strErr & "7"
                        'log TaxID not number

                    End If

                    If RemittoAddress = "" Or IsDBNull(RemittoAddress) Then
                        dtFT.Rows(i).Delete()
                        strErr = strErr & "8"
                        'log RemittoAddress = ""
                    End If

                    If chkOptService = False Then
                        dtFT.Rows(i).Delete()
                        strErr = strErr & "9"
                    End If

                Next
                If strErr <> "" Then

                Else
                    dtErr.Rows(i).Delete()
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Error 265: " & ex.Message)
        End Try
        dtFT.AcceptChanges()
        dtErr.AcceptChanges()
        DisplayValidate(dtErr)
    End Sub

#End Region

#Region "ProcessData"

    Public Function chkRmitto(ByVal Miscode As String) As DataTable
        Dim DataCheck As String = ""
        dataSt = New DataSet()
        connection = New SqlConnection(conStr)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If
        Dim str As String
        str = "SELECT IDVEND FROM APVNR WHERE IDVEND = '" & Miscode & "'"
        command = New SqlCommand(str, connection)
        adapter = New SqlDataAdapter(command)
        dataSt = New DataSet()
        adapter.Fill(dataSt, "DataCheck")
        connection.Close()

        Return dataSt.Tables("DataCheck")
    End Function

    Public Function chkTypeMC(ByVal batchno As String, ByVal entryno As String) As DataTable
        Dim DataCheck As String = ""
        dataSt = New DataSet()
        connection = New SqlConnection(conStr)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If

        sql1 = "SELECT R.NAMECITY AS BankCode "
        sql1 &= "FROM CBBTHD E " & vbCrLf
        sql1 &= "   LEFT JOIN APVNR R ON E.MISCCODE = R.IDVEND " & vbCrLf
        sql1 &= "   LEFT JOIN APVEN V ON E.MISCCODE = V.VENDORID " & vbCrLf
        sql1 &= "WHERE LTRIM(RTRIM(E.BATCHID))  = '" & batchno & "'" & vbCrLf
        sql1 &= "   AND LTRIM(RTRIM(E.ENTRYNO))  = '" & entryno & "' " & vbCrLf
        sql1 &= "   AND E.ENTRYTYPE = 1 "
        command = New SqlCommand(sql1, connection)
        adapter = New SqlDataAdapter(command)
        dataSt = New DataSet()
        adapter.Fill(dataSt, "DataCheck")
        connection.Close()

        Return dataSt.Tables("DataCheck")
    End Function

    Public Function hasSpecialChar(ByVal input As String) As Boolean
        Dim specialChar As String = "^~!#$%^&+*=|;<>?\*'"
        For Each item In specialChar
            If input.Contains(item) Then Return True
        Next
        Return False
    End Function

    Public Sub chkAddr(ByRef addr As String)
        If addr.Length > 140 Then
            addr = addr.Remove(139, addr.Length - 139)
        End If
    End Sub

    Public Function hasDefineOptService(ByVal input As String) As Boolean
        Dim DefineOptService As String = "023459"
        Dim arrstr(3) As String
        If input = "" Or IsDBNull(input) Then
            Return False
        Else
            arrstr = input.Split(",")
            If arrstr.Count >= 4 Then
                If arrstr(3).ToString <> "" Or IsNumeric(arrstr(3)) = True Then
                    input = arrstr(3).ToString
                Else
                    input = "0"
                End If
            Else
                input = "0"
            End If
            For Each item In DefineOptService
                If input.Contains(item) Then Return True
            Next
        End If
    End Function

#End Region
    Sub cleardt()
        dtchk.Rows.Clear()
        dtchk.Columns.Clear()
        dtchk.Clear()

        dtErr.Rows.Clear()
        dtErr.Columns.Clear()
        dtErr.Clear()

        DataGETP.Rows.Clear()
        DataGETP.Columns.Clear()
        DataGETP.Clear()

        DataGETI.Rows.Clear()
        DataGETI.Columns.Clear()
        DataGETI.Clear()

        DataGETC.Rows.Clear()
        DataGETC.Columns.Clear()
        DataGETC.Clear()

    End Sub

    Sub DisplayValidate(ByVal dt As DataTable)
        Try
            For j = 0 To clsVExport.DataGridView1.Rows.Count - 1
                If clsVExport.DataGridView1.Rows(j).Cells("Check").Value.ToString = "True" Then

                    Dim batchDGV = "", entryDGV = ""
                    batchDGV = clsVExport.DataGridView1.Rows(j).Cells("Batchno").Value.ToString
                    entryDGV = clsVExport.DataGridView1.Rows(j).Cells("Entry").Value.ToString
                    For i = 0 To dt.Rows.Count - 1
                        Dim batchno = "", entryno = ""
                        batchno = dt.Rows(i).Item("Batchno").ToString
                        entryno = dt.Rows(i).Item("Entry").ToString
                        If batchDGV = batchno And entryDGV = entryno Then
                            clsVExport.DataGridView1.Rows(j).DefaultCellStyle.BackColor = Color.Salmon
                        Else

                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Error 413: " & ex.Message)
        End Try
    End Sub
End Module
