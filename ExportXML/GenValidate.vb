Imports System.Data.SqlClient
Imports System.IO

Module GenValidate

    Public dtchk As DataTable = New DataTable()
    Public dtErr As DataTable = New DataTable()
    Public dtMC As DataTable = New DataTable()
    Public dtFT As DataTable = New DataTable()
    Public strLog As String = ""

#Region "Validate"
    Sub ValidateRmitto(ByVal DGVLIST As DataGridView)
        cleardt()
        strLog = ""
        Try
            dtchk = CType(DGVLIST.DataSource, DataTable).Copy
            dtErr = CType(DGVLIST.DataSource, DataTable).Copy
            'Call clsVExport.SelectCBX_ALL(dtchk)
            'Call clsVExport.SelectCBX_ALL(dtErr)

            For i = 0 To dtchk.Rows.Count - 1
                Dim Miscode As String = ""
                If IsDBNull(dtchk.Rows(i).Item("Check").ToString) Then
                    Continue For
                End If
                Dim CheckState As String = dtchk.Rows(i).Item("Check").ToString
                If CheckState = "True" Then
                    Dim strErr, strErrValue, batchno, entryno As String
                    Miscode = dtchk.Rows(i).Item("VendorCode").ToString
                    batchno = dtchk.Rows(i).Item("Batchno").ToString
                    entryno = dtchk.Rows(i).Item("Entry").ToString
                    Dim chkMiscode = chkRmitto(Miscode)
                    If chkMiscode.Rows.Count = 0 Then
                        dtchk.Rows(i).Delete()
                        strErr = "00"
                        strErrValue = "NO Remit-To Location" & vbCrLf & "-" & vbCrLf & "-" & vbCrLf & "-" & vbCrLf & "-" & vbCrLf & "-" & vbCrLf & "-" & vbCrLf & "-" & vbCrLf & "-"
                        InsertTempReport(strErr, strErrValue, batchno, entryno)
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

            GenReportValidate.ShowReport(dtchk)

        Catch ex As Exception
            'MessageBox.Show("Error 57 Validate : " & ex.Message)
            strLog &= "Error 57 Validate : " & ex.Message
        End Try
        'write log validate
        If strLog <> "" Then
            WRITELOG(strLog)
        End If
    End Sub
    Sub ValidateTypeMC(ByVal dtchk As DataTable)
        dtMC = dtchk.Copy
        dtErr = dtchk.Copy
        Try
            For i = 0 To dtMC.Rows.Count - 1
                DataGETP.Rows.Clear()
                DataGETW.Rows.Clear()
                Dim batchno As String = ""
                Dim entryno As String = ""
                Dim BeneficiaryName As String = ""
                Dim TOTALAMOUNT As Integer
                Dim Acct, TaxId, RemittoAddress, OptService, DateINV As String
                Dim Taxbase As String = Nothing

                batchno = dtMC.Rows(i).Item("BatchNo").ToString.TrimEnd
                entryno = dtMC.Rows(i).Item("Entry").ToString.TrimEnd
                BeneficiaryName = dtMC.Rows(i).Item("VendorName").ToString.TrimEnd
                TOTALAMOUNT = CInt(dtMC.Rows(i).Item("Amount").ToString)
                DataGETP = clsCGenerateFile.GetDataP(batchno, entryno)

                For p = 0 To DataGETP.Rows.Count - 1
                    Acct = DataGETP.Rows(0).Item("AccountNo").ToString.TrimEnd
                    TaxId = DataGETP.Rows(0).Item("TAXID").ToString.TrimEnd
                    RemittoAddress = DataGETP.Rows(0).Item("ADDR").ToString.TrimEnd
                    OptService = DataGETP.Rows(0).Item("COUNTRY").ToString.TrimEnd
                    DateINV = DataGETP.Rows(0).Item("DATE").ToString.TrimEnd
                Next

                DataGETW = clsCGenerateFile.GetDataW(batchno, entryno)
                For w = 0 To DataGETW.Rows.Count - 1
                    Taxbase = DataGETW.Rows(0).Item("WHTAMOUNT").ToString.TrimEnd
                Next
                Dim strErr As String = ""
                Dim strErrValue As String = ""
                Dim checkMC = chkTypeMC(batchno, entryno)
                Dim checkMCBeneficiaryName = hasSpecialChar(BeneficiaryName)
                Dim chkOptService = hasDefineOptService(OptService)
                Dim chkDateINV = hasCheckNowDate(DateINV)
                Dim chkTaxId = hasCheckNumeric(TaxId)
                For j = 0 To checkMC.Rows.Count - 1
                    Dim bankCode = checkMC.Rows(j).Item("BankCode").ToString.TrimEnd
                    If bankCode = "039" Then
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

                    'If Taxbase = 0 Or IsDBNull(Taxbase) Or Taxbase = Nothing Then
                    If Taxbase <> Nothing Then
                        If Taxbase = 0 Then
                            dtMC.Rows(i).Delete()
                            strErr = strErr & "3"
                            'log Taxbase = 0
                        Else
                            Taxbase = 0
                        End If
                    Else
                        Taxbase = 0
                    End If

                    If chkDateINV = False Then
                        dtMC.Rows(i).Delete()
                        strErr = strErr & "4"
                        'log DateINV > DateNow
                    End If

                    If Acct = "" Or IsDBNull(Acct) Then
                        dtMC.Rows(i).Delete()
                        strErr = strErr & "5"
                        'log Account Number = ""
                    Else

                        If Acct.Length >= 13 Then
                            dtMC.Rows(i).Delete()
                            strErr = strErr & "5"
                            'log Account number > 12 digit
                        Else
                            If Acct.Length = 12 Then
                                If IsNumeric(Acct) = False Then
                                    If Acct.Contains("H") Or Acct.Contains("F") Then
                                    Else
                                        dtMC.Rows(i).Delete()
                                        strErr = strErr & "5"
                                    End If
                                Else
                                    dtMC.Rows(i).Delete()
                                    strErr = strErr & "5"
                                End If
                                'ElseIf Acct.Length <> 11 Then
                                '    dtMC.Rows(i).Delete()
                                '    strErr = strErr & "5"
                            End If
                        End If

                    End If

                    If chkTaxId = True And TaxId <> "" And IsDBNull(TaxId) = False Then
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
                        strErr = strErr & "6"
                        'log TaxID not number

                    End If

                    'Check Bankcode 
                    If bankCode <> "" Then
                    Else
                        dtMC.Rows(i).Delete()
                        strErr = strErr & "7"
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
                    strErrValue = IIf(BeneficiaryName <> "", BeneficiaryName, "EMPTY") & vbCrLf & IIf(TOTALAMOUNT.ToString <> "", TOTALAMOUNT, "EMPTY") & vbCrLf & IIf(Taxbase <> Nothing, Taxbase, "EMPTY") & vbCrLf & IIf(Acct <> "", Acct, "EMPTY") & vbCrLf & IIf(TaxId <> "", TaxId, "EMPTY") & vbCrLf & IIf(RemittoAddress <> "", RemittoAddress, "EMPTY") & vbCrLf & IIf(OptService <> "", OptService, "EMPTY") & vbCrLf & IIf(DateINV <> "", DateINV, "EMPTY") & vbCrLf & IIf(bankCode <> "", bankCode, "EMPTY")
                Next
                If strErr <> "" Then
                    GenValidateReport(strErr, strErrValue, batchno, entryno)
                Else
                    dtErr.Rows(i).Delete()
                End If
            Next
        Catch ex As Exception
            'MessageBox.Show("Error 208 Validate : " & ex.Message)
            strLog &= "Error 208 Validate : " & ex.Message
        End Try
        dtMC.AcceptChanges()
        dtErr.AcceptChanges()
        DisplayValidate(dtErr)

    End Sub

    Sub ValidateTypeFT(ByVal dtchk As DataTable)
        dtFT = dtchk.Copy
        dtErr = dtchk.Copy
        Dim batchno As String
        Dim entryno As String
        Try
            For i = 0 To dtFT.Rows.Count - 1
                batchno = ""
                entryno = ""
                Dim BeneficiaryName As String = ""
                Dim TOTALAMOUNT As Integer
                Dim Taxbase As Integer = Nothing
                Dim Acct, TaxId, RemittoAddress, OptService, DateINV As String

                batchno = dtFT.Rows(i).Item("BatchNo").ToString.TrimEnd
                entryno = dtFT.Rows(i).Item("Entry").ToString.TrimEnd
                BeneficiaryName = dtFT.Rows(i).Item("VendorName").ToString.TrimEnd
                TOTALAMOUNT = CInt(dtFT.Rows(i).Item("Amount").ToString)
                DataGETP = clsCGenerateFile.GetDataP(batchno, entryno)
                For p = 0 To DataGETP.Rows.Count - 1
                    Acct = DataGETP.Rows(0).Item("AccountNo").ToString.TrimEnd
                    TaxId = DataGETP.Rows(0).Item("TAXID").ToString.TrimEnd
                    RemittoAddress = DataGETP.Rows(0).Item("ADDR").ToString.TrimEnd
                    OptService = DataGETP.Rows(0).Item("COUNTRY").ToString.TrimEnd
                    DateINV = DataGETP.Rows(0).Item("DATE").ToString.TrimEnd
                Next

                DataGETW = clsCGenerateFile.GetDataW(batchno, entryno)
                For w = 0 To DataGETW.Rows.Count - 1
                    Taxbase = DataGETW.Rows(0).Item("WHTAMOUNT").ToString.TrimEnd
                Next

                Dim strErr As String = ""
                Dim strErrValue As String = ""
                Dim checkMC = chkTypeMC(batchno, entryno)
                Dim checkMCBeneficiaryName = hasSpecialChar(BeneficiaryName)
                Dim chkOptService = hasDefineOptService(OptService)
                Dim chkDateINV = hasCheckNowDate(DateINV)
                Dim chkTaxId = hasCheckNumeric(TaxId)

                For j = 0 To checkMC.Rows.Count - 1
                    Dim bankCode = checkMC.Rows(j).Item("BankCode").ToString.TrimEnd
                    If bankCode <> "039" Then
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

                    'If Taxbase = 0 Or IsDBNull(Taxbase) Or Taxbase = Nothing Then
                    If Taxbase <> Nothing Then
                        If Taxbase = 0 Then
                            dtMC.Rows(i).Delete()
                            strErr = strErr & "3"
                            'log Taxbase = 0
                        Else
                            Taxbase = 0
                        End If
                    Else
                        Taxbase = 0
                    End If

                    If chkDateINV = False Then
                        dtFT.Rows(i).Delete()
                        strErr = strErr & "4"
                        'log DateINV > DateNow
                    End If

                    If Acct = "" Or IsDBNull(Acct) Then
                        dtFT.Rows(i).Delete()
                        strErr = strErr & "5"
                        'log Account Number = ""
                    Else

                        If Acct.Length >= 13 Then
                            dtMC.Rows(i).Delete()
                            strErr = strErr & "5"
                            'log Account number > 12 digit
                        Else
                            If Acct.Length = 12 Then
                                If IsNumeric(Acct) = False Then
                                    If Acct.Contains("H") Or Acct.Contains("F") Then
                                    Else
                                        dtMC.Rows(i).Delete()
                                        strErr = strErr & "5"
                                    End If
                                Else
                                    dtMC.Rows(i).Delete()
                                    strErr = strErr & "5"
                                End If
                                'ElseIf Acct.Length <> 11 Then
                                '    dtMC.Rows(i).Delete()
                                '    strErr = strErr & "5"
                            End If
                        End If

                    End If

                    If chkTaxId = True And TaxId <> "" And IsDBNull(TaxId) = False Then
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
                        strErr = strErr & "6"
                        'log TaxID not number

                    End If

                    'Check Bankcode 
                    If bankCode <> "" Then
                    Else
                        dtMC.Rows(i).Delete()
                        strErr = strErr & "7"
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
                    strErrValue = IIf(BeneficiaryName <> "", BeneficiaryName, "EMPTY") & vbCrLf & IIf(TOTALAMOUNT.ToString <> "", TOTALAMOUNT, "EMPTY") & vbCrLf & IIf(Taxbase.ToString <> "", Taxbase, "EMPTY") & vbCrLf & IIf(Acct <> "", Acct, "EMPTY") & vbCrLf & IIf(TaxId <> "", TaxId, "EMPTY") & vbCrLf & IIf(RemittoAddress <> "", RemittoAddress, "EMPTY") & vbCrLf & IIf(OptService <> "", OptService, "EMPTY") & vbCrLf & IIf(DateINV <> "", DateINV, "EMPTY") & vbCrLf & IIf(bankCode <> "", bankCode, "EMPTY")
                Next
                If strErr <> "" Then
                    GenValidateReport(strErr, strErrValue, batchno, entryno)
                Else
                    dtErr.Rows(i).Delete()
                End If
            Next
        Catch ex As Exception
            'MessageBox.Show("Error 376 Validate : " & ex.Message)
            strLog &= "Error 376 Validate : " & ex.Message & "Batch" & batchno & "Entry" & entryno
        End Try
        dtFT.AcceptChanges()
        dtErr.AcceptChanges()
        DisplayValidate(dtErr)
    End Sub

#End Region

#Region "ProcessData"

    Public Function chkRmitto(ByVal Miscode As String) As DataTable
        Try
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
        Catch ex As Exception
            'MessageBox.Show("Error 303 Validate : " & ex.Message)
            strLog &= "Error 397 Validate : " & ex.Message
        End Try
        Return dataSt.Tables("DataCheck")
    End Function

    Public Function chkTypeMC(ByVal batchno As String, ByVal entryno As String) As DataTable
        Try
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
        Catch ex As Exception
            'MessageBox.Show("Error 424 Validate : " & ex.Message)
            strLog &= "Error 424 Validate : " & ex.Message
        End Try
        Return dataSt.Tables("DataCheck")
    End Function

    Public Function hasSpecialChar(ByVal input As String) As Boolean
        Try
            Dim specialChar As String = "^~!#$%^&+*=|;<>?\*'"
            For Each item In specialChar
                If input.Contains(item) Then Return True
            Next
        Catch ex As Exception
            'MessageBox.Show("Error 344 Validate : " & ex.Message)
            strLog &= "Error 438 Validate : " & ex.Message
        End Try
        Return False
    End Function

    Public Sub chkAddr(ByRef addr As String)
        Try
            If addr.Length > 140 Then
                addr = addr.Remove(139, addr.Length - 139)
            End If
        Catch ex As Exception
            'MessageBox.Show("Error 353 Validate : " & ex.Message)
            strLog &= "Error 450 Validate : " & ex.Message
        End Try
    End Sub

    Public Function hasDefineOptService(ByVal input As String) As Boolean
        Try
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
        Catch ex As Exception
            'MessageBox.Show("Error 390 Validate : " & ex.Message)
            strLog &= "Error 390 Validate : " & ex.Message
        End Try
    End Function
    Public Function hasCheckNowDate(ByVal DateINV As String) As Boolean
        Try
            Dim DateNow = Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00")

            If DateINV = "" Or IsDBNull(DateINV) Then
                Return False
            Else
                If DateINV > DateNow Then
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            'MessageBox.Show("Error 410 Validate: " & ex.Message)
            strLog &= "Error 495 Validate: " & ex.Message
        End Try
    End Function

    Public Function hasCheckNumeric(ByVal input As String) As Boolean
        Try
            If IsDBNull(input) = True Or input = "" Then
                Return False
            Else
                Dim specialChar As String = "0123456789"
                For Each item In specialChar
                    If input.Contains(item) Then Return True
                Next
            End If
        Catch ex As Exception
            'MessageBox.Show("Error 420 Validate : " & ex.Message)
            strLog &= "Error 511 Validate : " & ex.Message
        End Try

    End Function

#End Region

#Region "EVENT"
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
            'MessageBox.Show("Error 431 Validate : " & ex.Message)
            strLog &= "Error 565 Validate : " & ex.Message
        End Try
    End Sub

    Sub GenValidateReport(ByVal strErr As String, ByVal strErrValue As String, ByVal batchno As String, ByVal entryno As String)
        InsertTempReport(strErr, strErrValue, batchno, entryno)
    End Sub
    Sub CreateTempTable()
        Try
            Dim dtchkTemp As DataTable = New DataTable()
            connection = New SqlConnection(conStr)
            If connection.State = ConnectionState.Closed Then
                connection.Open()
            End If
            Dim STREXIST As String
            STREXIST = "SELECT name FROM sys.tables WHERE name = 'FMSEPAYTEMP'"
            Dim cmdquery As SqlDataAdapter = New SqlDataAdapter(STREXIST, connection)
            cmdquery.Fill(dtchkTemp)

            If dtchkTemp.Rows.Count = 0 Then
                Dim str As String
                str = "CREATE TABLE FMSEPAYTEMP " & vbCrLf
                str &= "(ERRCODE NVARCHAR(20)," & vbCrLf
                str &= " BATCH NVARCHAR(100)," & vbCrLf
                str &= " ENTRYNO NVARCHAR(100)," & vbCrLf
                str &= " BeneficiaryName NVARCHAR(100)," & vbCrLf
                str &= " TOTALAMOUNT NVARCHAR(100)," & vbCrLf
                str &= " Taxbase NVARCHAR(100)," & vbCrLf
                str &= " Acct NVARCHAR(100)," & vbCrLf
                str &= " TaxId NVARCHAR(100)," & vbCrLf
                str &= " RemittoAddress NVARCHAR(100)," & vbCrLf
                str &= " OptService NVARCHAR(100)," & vbCrLf
                str &= " DATEINV NVARCHAR(20)," & vbCrLf
                str &= " BankCode NVARCHAR(20))"
                Dim cmd As SqlCommand = New SqlCommand(str, connection)
                cmd.ExecuteNonQuery()
            End If
        Catch ex As Exception
            'MessageBox.Show("Error 468 Validate : " & ex.Message)
            strLog &= "Error 604 Validate : " & ex.Message
        End Try
    End Sub

    Sub InsertTempReport(ByVal strErr As String, ByVal strErrValue As String, ByVal batchno As String, ByVal entryno As String)
        Try
            CreateTempTable()
            connection = New SqlConnection(conStr)
            If connection.State = ConnectionState.Closed Then
                connection.Open()
            End If

            Dim arrStrErr(8) As String
            arrStrErr = strErrValue.Split(New String() {Environment.NewLine},
              StringSplitOptions.RemoveEmptyEntries)

            Dim str As String = ""
            str = "INSERT INTO FMSEPAYTEMP" & vbCrLf
            str &= "VALUES('" & strErr & "' " & vbCrLf
            str &= ",'" & batchno & "'" & vbCrLf
            str &= ",'" & entryno & "'" & vbCrLf
            str &= ",'" & arrStrErr(0).ToString & "'" & vbCrLf
            str &= ",'" & arrStrErr(1).ToString & "'" & vbCrLf
            str &= ",'" & arrStrErr(2).ToString & "'" & vbCrLf
            str &= ",'" & arrStrErr(3).ToString & "'" & vbCrLf
            str &= ",'" & arrStrErr(4).ToString & "'" & vbCrLf
            str &= ",'" & arrStrErr(5).ToString & "'" & vbCrLf
            str &= ",'" & arrStrErr(6).ToString & "'" & vbCrLf
            str &= ",'" & arrStrErr(7).ToString & "'" & vbCrLf
            str &= ",'" & arrStrErr(8).ToString & "')" & vbCrLf

            Dim cmd As SqlCommand = New SqlCommand(str, connection)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            'MessageBox.Show("Error 560 Validate : " & ex.Message)
            strLog &= "Error 639 Validate : " & ex.Message & "Batch" & batchno & "Entry" & entryno
        End Try
    End Sub

    Sub WRITELOG(ByVal strlog As String)
        Try

            If (Not System.IO.Directory.Exists(My.Application.Info.DirectoryPath & "\Log")) Then
                System.IO.Directory.CreateDirectory(My.Application.Info.DirectoryPath & "\Log")
            End If
            Dim FILE_text1 As String
            Dim getnow As String
            getnow = Date.Now.Day & Date.Now.Month & Date.Now.Year & "_" & Date.Now.Hour & Date.Now.Minute
            strlog &= Date.Now & " " & strlog
            FILE_text1 = My.Application.Info.DirectoryPath & "\Log\log" & getnow & ".txt"

            Dim objWriter As New System.IO.StreamWriter(FILE_text1)
            objWriter.Write(strlog)
            objWriter.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

#End Region
End Module
