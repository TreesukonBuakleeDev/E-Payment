Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text

Public Class clsCGenerateFile
    Public Shared clsW As New clsCWriteLogFile

#Region "WriteTextFile"

    Public Function WriteTextFile(ByVal DataGridView1 As DataGridView, ByVal TypeExport As String, ByVal Progressbar1 As ProgressBar) As Boolean
        Dim status As Boolean
        Dim DataForExport As New DataView(DataGridView1.DataSource, "Check = TRUE", "BatchNo", DataViewRowState.CurrentRows)
        Dim dtList As DataTable = DataForExport.ToTable()
        If TypeExport = 6 Then
            status = WriteExportManyFile(dtList, Progressbar1) 'แยกbatch 
        ElseIf TypeExport = 7 Then
            status = WriteExportOneFile(dtList) 'รวมbatch
        End If
        Return status
    End Function

    Public Function WriteExportOneFile(ByVal dtList As DataTable) As Boolean
        Dim ret As Boolean = True
        CheckExistDirectory()
        Dim fileForWrite As StreamWriter
        Dim DateFileName As String = GetDateForFileName(Now)
        fileForWrite = My.Computer.FileSystem.OpenTextFileWriter(BaseValiabled.FileEx & "\Payment" & DateFileName & ".txt", True, Encoding.ASCII)
        fileForWrite.WriteLine("F2")
        fileForWrite.WriteLine("")
        fileForWrite.WriteLine("")
        fileForWrite.WriteLine("")
        Try
            For i = 0 To dtList.Rows.Count - 1
                Dim BatchNo As String = dtList.Rows(i)("BatchNo")
                Dim DataP As DataTable = GetDataP(BatchNo) '
                For l = 0 To DataP.Rows.Count - 1
                    Dim Entry As String = DataP.Rows(l)("CNTENTR")
                    Dim FaxNo As String = DataP.Rows(l)("FAXNo").ToString.Replace(" ", "")
                    Dim Amount As Decimal = DataP.Rows(l)("Amount")
                    fileForWrite.WriteLine(Chr(34) & DataP.Rows(l)("RecordType").ToString.Trim & Chr(34) _
                                           & "," & Chr(34) & DataP.Rows(l)("BeneficiaryName").ToString.Trim & Chr(34) _
                                           & "," & Chr(34) & DataP.Rows(l)("AccountNo").ToString.Trim & Chr(34) _
                                           & "," & Chr(34) & DataP.Rows(l)("BankCode").ToString.Trim & Chr(34) _
                                           & "," & Chr(34) & DataP.Rows(l)("BranchCode").ToString.Trim & Chr(34) _
                                           & "," & Chr(34) & DataP.Rows(l)("BankCharge").ToString.Trim & Chr(34) _
                                           & "," & Chr(34) & DataP.Rows(l)("ServiceType").ToString.Trim & Chr(34) _
                                           & "," & Chr(34) & Amount.ToString("0.00") & Chr(34) _
                                           & "," & Chr(34) & FaxNo & Chr(34) _
                                           & "," & Chr(34) & DataP.Rows(l)("Email").ToString.Trim _
                                           & "," & Chr(34) & Chr(34) _
                                           & "," & Chr(34) & DataP.Rows(l)("CNTBTCH") & "/" & DataP.Rows(l)("CNTENTR") & Chr(34) _
                                           & "," & Chr(34) & Chr(34))

                    Dim DataW As DataTable = GetDataW(BatchNo, Entry)
                    If Not IsDBNull(DataW) AndAlso DataW.Rows.Count > 0 Then
                        For k = 0 To DataW.Rows.Count - 1
                            Dim WTHAmount As Decimal = DataW.Rows(k)("WHTAmount")
                            Dim AmountPaid As Decimal = DataW.Rows(k)("AmountPaid")
                            fileForWrite.WriteLine(Chr(34) & "W" & Chr(34) _
                                         & "," & Chr(34) & DataW.Rows(k)("WHTaxNo").ToString.Trim & Chr(34) _
                                         & "," & Chr(34) & DataW.Rows(k)("VENDNAME").ToString.Trim & Chr(34) _
                                         & "," & Chr(34) & DataW.Rows(k)("ADDRESS").ToString.Trim & Chr(34) _
                                         & "," & Chr(34) & DataW.Rows(k)("TAXID").ToString.Trim & Chr(34) _
                                         & "," & Chr(34) & Chr(34) _
                                         & "," & Chr(34) & DataW.Rows(k)("TAXForm").ToString.Trim & Chr(34) _
                                         & "," & Chr(34) & Chr(34) _
                                         & "," & Chr(34) & Chr(34) _
                                         & "," & Chr(34) & DataW.Rows(k)("TypeOfIncome").ToString.Trim & Chr(34) _
                                         & "," & Chr(34) & DataW.Rows(k)("TypeOfIncomeDiscription").ToString.Trim & Chr(34) _
                                         & "," & Chr(34) & DataW.Rows(k)("Date").ToString.Substring(2, 6) & Chr(34) _
                                         & "," & Chr(34) & AmountPaid.ToString("0.00") & Chr(34) _
                                         & "," & Chr(34) & DataW.Rows(k)("TaxRate").ToString.Trim & Chr(34) _
                                         & "," & Chr(34) & WTHAmount.ToString("0.00") & Chr(34) _
                                         & "," & Chr(34) & Chr(34) _
                                         & "," & Chr(34) & Chr(34) _
                                         & "," & Chr(34) & Chr(34) _
                                         & "," & Chr(34) & Chr(34) _
                                         & "," & Chr(34) & Chr(34) _
                                         & "," & Chr(34) & Chr(34) _
                                         & "," & Chr(34) & Chr(34) _
                                         & "," & Chr(34) & Chr(34) _
                                         & "," & Chr(34) & Chr(34) _
                                         & "," & Chr(34) & Chr(34) _
                                         & "," & Chr(34) & Chr(34) _
                                         & "," & Chr(34) & Chr(34))
                        Next
                    End If

                    Dim dtDataI As DataTable = GetDataI(BatchNo, Entry)
                    If Not IsDBNull(dtDataI) Then
                        For k = 0 To dtDataI.Rows.Count - 1
                            fileForWrite.WriteLine(Chr(34) & "I" & Chr(34) _
                                                    & "," & Chr(34) & dtDataI.Rows(k)("IDINVC").ToString.Trim & Chr(34) _
                                                    & "," & Chr(34) & dtDataI.Rows(k)("DATEINVC").ToString.Substring(2, 6) & Chr(34) _
                                                    & "," & Chr(34) & Chr(34) _
                                                    & "," & Chr(34) & Chr(34) _
                                                    & "," & Chr(34) & Chr(34) _
                                                    & "," & Chr(34) & CDbl(dtDataI.Rows(k)("AMTGROSTOT").ToString.Trim) & Chr(34))
                        Next
                    End If
                Next
            Next
            fileForWrite.Close()
        Catch ex As Exception
            clsW.WriteLogError("<WriteTextFile>/Export : " & ex.Message.ToString())
            fileForWrite.Dispose()
            ret = False
        End Try
        Return ret
    End Function

    Public Function WriteExportManyFile(ByVal dtList As DataTable, ByVal Progressbar1 As ProgressBar) As Boolean
        Dim ret As Boolean = True
        Try
            Dim fileForWrite As StreamWriter
            Dim datefilename As String = GetDateForFileName(Now)
            For i = 0 To dtList.Rows.Count - 1
                fileForWrite = My.Computer.FileSystem.OpenTextFileWriter(BaseValiabled.FileEx &
                                                                         "\Payment" & datefilename & "-" & dtList.Rows(i)("BatchNo") & ".txt", False, Encoding.Unicode)

                Progressbar1.Value = 20

                fileForWrite.WriteLine("F2")
                fileForWrite.WriteLine("")
                fileForWrite.WriteLine("")
                fileForWrite.WriteLine("")



                Dim BatchNo As String = dtList.Rows(i)("BatchNo")
                Dim DataP As DataTable = GetDataP(BatchNo)
                For l = 0 To DataP.Rows.Count - 1
                    Dim Entry As String = DataP.Rows(l)("CNTENTR")
                    Dim FaxNo As String = DataP.Rows(l)("FAXNo").ToString.Replace(" ", "")
                    Dim Amount As Decimal = DataP.Rows(l)("Amount")
                    fileForWrite.WriteLine(Chr(34) & DataP.Rows(l)("RecordType").ToString.Trim & Chr(34) _
                                           & "," & Chr(34) & DataP.Rows(l)("BeneficiaryName").ToString.Trim & Chr(34) _
                                           & "," & Chr(34) & DataP.Rows(l)("AccountNo").ToString.Trim & Chr(34) _
                                           & "," & Chr(34) & DataP.Rows(l)("BankCode").ToString.Trim & Chr(34) _
                                           & "," & Chr(34) & DataP.Rows(l)("BranchCode").ToString.Trim & Chr(34) _
                                           & "," & Chr(34) & DataP.Rows(l)("BankCharge").ToString.Trim & Chr(34) _
                                           & "," & Chr(34) & DataP.Rows(l)("ServiceType").ToString.Trim & Chr(34) _
                                           & "," & Chr(34) & Amount.ToString("0.00") & Chr(34) _
                                           & "," & Chr(34) & FaxNo & Chr(34) _
                                           & "," & Chr(34) & DataP.Rows(l)("Email").ToString.Trim & Chr(34) _
                                           & "," & Chr(34) & Chr(34) _
                                           & "," & Chr(34) & DataP.Rows(l)("CNTBTCH") & "/" & DataP.Rows(l)("CNTENTR") & Chr(34) _
                                           & "," & Chr(34) & Chr(34))


                    Dim DataW As DataTable = GetDataW(BatchNo, Entry)
                    If Not IsDBNull(DataW) AndAlso DataW.Rows.Count > 0 Then
                        Dim TypeOfIncome(2) As String
                        Dim TypeOfIncomeDescription(2) As String
                        Dim DateWHT(2) As String
                        Dim AmountPaid(2) As Decimal
                        Dim Taxrate(2) As Integer
                        Dim WHTAmount(2) As Decimal
                        For k = 0 To DataW.Rows.Count - 1
                            Select Case k
                                Case 0
                                    TypeOfIncome(k) = DataW.Rows(k)("TypeOfIncome")
                                    TypeOfIncomeDescription(k) = DataW.Rows(k)("TypeOfIncomeDiscription")
                                    DateWHT(k) = DataW.Rows(k)("Date")
                                    AmountPaid(k) = DataW.Rows(k)("AmountPaid")
                                    Taxrate(k) = DataW.Rows(k)("TaxRate")
                                    WHTAmount(k) = DataW.Rows(k)("WHTAmount")
                                Case 1
                                    TypeOfIncome(k) = DataW.Rows(k)("TypeOfIncome")
                                    TypeOfIncomeDescription(k) = DataW.Rows(k)("TypeOfIncomeDiscription")
                                    DateWHT(k) = DataW.Rows(k)("Date")
                                    AmountPaid(k) = DataW.Rows(k)("AmountPaid")
                                    Taxrate(k) = DataW.Rows(k)("TaxRate")
                                    WHTAmount(k) = DataW.Rows(k)("WHTAmount")
                                Case 2
                                    TypeOfIncome(k) = DataW.Rows(k)("TypeOfIncome")
                                    TypeOfIncomeDescription(k) = DataW.Rows(k)("TypeOfIncomeDiscription")
                                    DateWHT(k) = DataW.Rows(k)("Date")
                                    AmountPaid(k) = DataW.Rows(k)("AmountPaid")
                                    Taxrate(k) = DataW.Rows(k)("TaxRate")
                                    WHTAmount(k) = DataW.Rows(k)("WHTAmount")
                            End Select
                        Next
                        fileForWrite.WriteLine(Chr(34) & "W" & Chr(34) _
                                     & "," & Chr(34) & DataW.Rows(0)("WHTaxNo").ToString.Trim & Chr(34) _
                                     & "," & Chr(34) & DataW.Rows(0)("VENDNAME").ToString.Trim & Chr(34) _
                                     & "," & Chr(34) & DataW.Rows(0)("ADDRESS").ToString.Trim & Chr(34) _
                                     & "," & Chr(34) & DataW.Rows(0)("TAXID").ToString.Trim & Chr(34) _
                                     & "," & Chr(34) & Chr(34) _
                                     & "," & Chr(34) & DataW.Rows(0)("TAXForm").ToString.Trim & Chr(34) _
                                     & "," & Chr(34) & Chr(34) _
                                     & "," & Chr(34) & Chr(34) _
                                     & "," & Chr(34) & TypeOfIncome(0) & Chr(34) _
                                     & "," & Chr(34) & TypeOfIncomeDescription(0) & Chr(34) _
                                     & "," & Chr(34) & DateWHT(0) & Chr(34) _
                                     & "," & Chr(34) & IIf(AmountPaid(0) = 0, "", AmountPaid(0)) & Chr(34) _
                                     & "," & Chr(34) & IIf(Taxrate(0) = 0, "", Taxrate(0)) & Chr(34) _
                                     & "," & Chr(34) & IIf(WHTAmount(0) = 0, "", WHTAmount(0)) & Chr(34) _
                                     & "," & Chr(34) & TypeOfIncome(1) & Chr(34) _
                                     & "," & Chr(34) & TypeOfIncomeDescription(1) & Chr(34) _
                                     & "," & Chr(34) & DateWHT(1) & Chr(34) _
                                     & "," & Chr(34) & IIf(AmountPaid(1) = 0, "", AmountPaid(1)) & Chr(34) _
                                     & "," & Chr(34) & IIf(Taxrate(1) = 0, "", Taxrate(1)) & Chr(34) _
                                     & "," & Chr(34) & IIf(WHTAmount(1) = 0, "", WHTAmount(1)) & Chr(34) _
                                     & "," & Chr(34) & TypeOfIncome(2) & Chr(34) _
                                     & "," & Chr(34) & TypeOfIncomeDescription(2) & Chr(34) _
                                     & "," & Chr(34) & DateWHT(2) & Chr(34) _
                                     & "," & Chr(34) & IIf(AmountPaid(2) = 0, "", AmountPaid(2)) & Chr(34) _
                                     & "," & Chr(34) & IIf(Taxrate(2) = 0, "", Taxrate(2)) & Chr(34) _
                                     & "," & Chr(34) & IIf(WHTAmount(2) = 0, "", WHTAmount(2)) & Chr(34))
                    End If

                    Dim dtDataI As DataTable = GetDataI(BatchNo, Entry)
                    If Not IsDBNull(dtDataI) Then
                        For k = 0 To dtDataI.Rows.Count - 1
                            fileForWrite.WriteLine(Chr(34) & "I" & Chr(34) _
                                                    & "," & Chr(34) & dtDataI.Rows(k)("IDINVC").ToString.Trim & Chr(34) _
                                                    & "," & Chr(34) & dtDataI.Rows(k)("DATEINVC").ToString.Substring(2, 6) & Chr(34) _
                                                    & "," & Chr(34) & Chr(34) _
                                                    & "," & Chr(34) & Chr(34) _
                                                    & "," & Chr(34) & Chr(34) _
                                                    & "," & Chr(34) & CDbl(dtDataI.Rows(k)("AMTGROSTOT").ToString.Trim) & Chr(34))
                        Next

                        Progressbar1.Value = 100

                    End If
                Next
                fileForWrite.Close()

            Next
        Catch ex As Exception
            clsW.WriteLogError("<WriteTextFile>/Export : " & ex.Message.ToString())
            ret = False
        End Try
        Return ret
    End Function
#End Region

#Region "Get Data"

    Public Shared Function GetDataP(ByVal BatchNo As String) As DataTable
        Dim DataP As String = ""
        connection = New SqlConnection(conStr)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If

        sql1 = "SELECT 'P' AS RecordType   " & vbCrLf
        sql1 &= "   ,R.RMITNAME AS BeneficiaryName " & vbCrLf
        sql1 &= "   ,R.CODEPSTL AS AccountNo " & vbCrLf
        sql1 &= "   ,R.NAMECITY AS BankCode " & vbCrLf
        sql1 &= "   ,R.CODESTTE AS BranchCode" & vbCrLf
        sql1 &= "   ,'BEN' AS BankCharge " & vbCrLf
        sql1 &= "   ,'04' AS ServiceType " & vbCrLf
        sql1 &= "   ,ABS(E.TOTAMOUNT) AS Amount  " & vbCrLf
        sql1 &= "   ,R.TEXTPHON2 AS FaxNo " & vbCrLf
        sql1 &= "   ,R.EMAIL AS Email" & vbCrLf
        sql1 &= "   ,E.BATCHID AS CNTBTCH " & vbCrLf
        sql1 &= "   ,E.ENTRYNO AS CNTENTR " & vbCrLf
        sql1 &= "   ,CONCAT(RTRIM(R.TEXTSTRE1),RTRIM(R.TEXTSTRE2)) AS ADDR" & vbCrLf
        sql1 &= "   ,RTRIM(R.TEXTSTRE3) AS CITY" & vbCrLf
        sql1 &= "   ,RTRIM(R.TEXTSTRE4) AS POSTCODE " & vbCrLf
        sql1 &= "   ,V.BRN AS TAXID " & vbCrLf
        sql1 &= "   ,R.CODECTRY AS COUNTRY" & vbCrLf
        sql1 &= "FROM CBBTHD E " & vbCrLf
        sql1 &= "   LEFT JOIN APVNR R ON E.MISCCODE = R.IDVEND " & vbCrLf
        sql1 &= "   LEFT JOIN APVEN V ON E.MISCCODE = V.VENDORID " & vbCrLf
        sql1 &= "WHERE LTRIM(RTRIM(E.BATCHID))  = '" & BatchNo & "'" & vbCrLf
        sql1 &= "   AND E.ENTRYTYPE = 1 "

        command = New SqlCommand(sql1, connection)
        adapter = New SqlDataAdapter(command)
        dataSt = New DataSet()
        adapter.Fill(dataSt, "DataP")
        connection.Close()

        Return dataSt.Tables("DataP")
    End Function

    Public Shared Function GetDataW(ByVal BatchNo As String, ByVal Entry As String) As DataTable
        Try
            Dim DataW As String = ""
            connection = New SqlConnection(conStr)
            If connection.State = ConnectionState.Closed Then
                connection.Open()
            End If

            sql1 = "SELECT ROW_NUMBER() OVER(ORDER BY D.ENTRYNO ASC) AS Row " & vbCrLf
            sql1 &= "   ,D.COMMENTS AS TEXTREF " & vbCrLf
            sql1 &= "   ,SUBSTRING(D.SRCECODE,LEN(D.SRCECODE),1) AS TAXRATE " & vbCrLf
            sql1 &= "   ,D.SRCECODE AS IDDISTCODE " & vbCrLf
            sql1 &= "   ,H.DATE AS Date" & vbCrLf
            'sql1 &= "   ,0 AS AmountPaid" & vbCrLf
            sql1 &= "   ,H.MISCCODE AS IDVEND " & vbCrLf
            sql1 &= "   ,D.QUANTITY AS TAXBASE " & vbCrLf
            sql1 &= "   ,D.DTLAMOUNT AS WHTAMOUNT " & vbCrLf
            sql1 &= "   ,ABS(H.TOTAMOUNT) AS TOTAMOUNT" & vbCrLf
            sql1 &= "FROM CBBTDT D" & vbCrLf
            sql1 &= "   LEFT JOIN CBBTHD H ON D.BATCHID = H.BATCHID AND  D.ENTRYNO = H.ENTRYNO  " & vbCrLf
            sql1 &= "WHERE D.BATCHID = " & BatchNo & " AND D.ENTRYNO = " & Entry & vbCrLf
            sql1 &= "   AND (SUBSTRING(D.SRCECODE,1,2) = 'T3' OR SUBSTRING(D.SRCECODE,1,3) = 'T53')" & vbCrLf
            sql1 &= "   AND H.ENTRYTYPE = 1 " & vbCrLf

            command = New SqlCommand(sql1, connection)
            adapter = New SqlDataAdapter(command)
            dataSt = New DataSet()
            adapter.Fill(dataSt, "DataW")

            dataSt.Tables("DataW").Columns.Add("AmountPaid", Type.GetType("System.String"))
            dataSt.Tables("DataW").Columns.Add("WHTAXNo")
            dataSt.Tables("DataW").Columns.Add("VENDNAME")
            dataSt.Tables("DataW").Columns.Add("ADDRESS")
            dataSt.Tables("DataW").Columns.Add("TAXID")
            dataSt.Tables("DataW").Columns.Add("TAXForm")
            dataSt.Tables("DataW").Columns.Add("TypeOfIncome")
            dataSt.Tables("DataW").Columns.Add("TypeOfIncomeDiscription")
            dataSt.Tables("DataW").Columns.Add("WHTAmount")

            For i = 0 To dataSt.Tables("DataW").Rows.Count - 1
                Dim REF As String = dataSt.Tables("DataW").Rows(i)("TEXTREF")
                Dim REFArr() As String = REF.Split(",")

                Dim PayVendor As String = ""

                If REFArr(0) = "" Then
                    PayVendor = dataSt.Tables("DataW").Rows(i)("IDVEND")
                Else
                    PayVendor = REFArr(0)
                End If

                sql1 = "SELECT R.RMITNAME AS VENDNAME" & vbCrLf
                sql1 &= "   ,CONCAT(RTRIM(R.TEXTSTRE1),RTRIM(R.TEXTSTRE2),RTRIM(R.TEXTSTRE3),RTRIM(R.TEXTSTRE4)) AS [ADDRESS]" & vbCrLf
                sql1 &= "   ,O.VALUE AS TAXID     " & vbCrLf
                sql1 &= "FROM APVNR R" & vbCrLf
                sql1 &= "    LEFT JOIN APVEN V ON R.IDVEND = V.VENDORID" & vbCrLf
                sql1 &= "    LEFT JOIN APVENO O ON V.VENDORID = O.VENDORID AND OPTFIELD = 'TAXID'" & vbCrLf
                sql1 &= "WHERE V.VENDORID = '" & PayVendor & "'"


                command = New SqlCommand(sql1, connection)
                adapter = New SqlDataAdapter(command)
                adapter.Fill(dataSt, "DataW_VENDOR")
                'dataSt.Tables("DataW").Rows(i)("WHTAXNo") = IIf(IS null(REFArr(4)), 0, REFArr(4))
                If REFArr.Length < 5 Then
                    dataSt.Tables("DataW").Rows(i)("WHTAXNo") = 0
                Else
                    dataSt.Tables("DataW").Rows(i)("WHTAXNo") = REFArr(4)
                End If
                dataSt.Tables("DataW").Rows(i)("VENDNAME") = dataSt.Tables("DataW_VENDOR").Rows(i)("VENDNAME")
                dataSt.Tables("DataW").Rows(i)("ADDRESS") = dataSt.Tables("DataW_VENDOR").Rows(i)("ADDRESS")
                dataSt.Tables("DataW").Rows(i)("TAXID") = dataSt.Tables("DataW_VENDOR").Rows(i)("TAXID")
                dataSt.Tables("DataW").Rows(i)("TAXForm") = IIf(Len(dataSt.Tables("DataW").Rows(i)("IDDISTCODE").ToString.Trim) = 3, "4", "7")

                If REFArr(1) = 5 And dataSt.Tables("DataW").Rows(i)("Row") = 1 Then
                    dataSt.Tables("DataW").Rows(i)("TypeOfIncome") = 51
                ElseIf REFArr(1) = 5 And dataSt.Tables("DataW").Rows(i)("Row") = 2 Then
                    dataSt.Tables("DataW").Rows(i)("TypeOfIncome") = 52
                ElseIf REFArr(1) = 6 And (dataSt.Tables("DataW").Rows(i)("Row") = 1 Or dataSt.Tables("DataW").Rows(i)("Row") = 3) Then
                    dataSt.Tables("DataW").Rows(i)("TypeOfIncome") = 61
                ElseIf REFArr(1) = 6 And (dataSt.Tables("DataW").Rows(i)("Row") = 2 Or dataSt.Tables("DataW").Rows(i)("Row") = 4) Then
                    dataSt.Tables("DataW").Rows(i)("TypeOfIncome") = 62
                Else
                    dataSt.Tables("DataW").Rows(i)("TypeOfIncome") = REFArr(1)
                End If

                dataSt.Tables("DataW").Rows(i)("TypeOfIncomeDiscription") = REFArr(2)
                dataSt.Tables("DataW").Rows(i)("WHTAmount") = dataSt.Tables("DataW").Rows(i)("WHTAMOUNT")
                dataSt.Tables("DataW").Rows(i)("AmountPaid") = dataSt.Tables("DataW").Rows(i)("TAXBASE")
            Next
            connection.Close()

        Catch ex As Exception
            clsW.WriteLogError("GETDataW\ : " & ex.Message.ToString())
            ' ret = False
        End Try

        Return dataSt.Tables("DataW")
    End Function

    Public Shared Function GetDataI(ByVal BatchNo As String, ByVal Entry As String) As DataTable
        connection = New SqlConnection(conStr)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If

        sql1 = "SELECT D.DOCNUMBER AS IDINVC " & vbCrLf
        sql1 &= "   ,D.DATEDOC AS DATEINVC " & vbCrLf
        sql1 &= "   ,H.TEXTDESC AS INVCDESC " & vbCrLf
        sql1 &= "   ,D.APPLAMOUNT AS AMTGROSTOT " & vbCrLf
        sql1 &= "   ,'' AS AMTTAXDIST " & vbCrLf
        sql1 &= "   ,'' AS AMTINVCTOT " & vbCrLf
        sql1 &= "FROM CBBTSD D " & vbCrLf
        sql1 &= "   LEFT JOIN CBBTHD H ON D.BATCHID = H.BATCHID AND D.ENTRYNO = H.ENTRYNO" & vbCrLf
        sql1 &= "WHERE D.BATCHID = " & BatchNo & vbCrLf
        sql1 &= "   AND D.ENTRYNO = " & Entry & vbCrLf

        command = New SqlCommand(sql1, connection)
        adapter = New SqlDataAdapter(command)
        dataSt = New DataSet()
        adapter.Fill(dataSt, "DataI")
        connection.Close()
        Return dataSt.Tables("DataI")
    End Function

    Public Shared Function GetDataCB(ByVal BatchNo As String, ByVal Entry As String) As DataTable
        Try
            Dim DataCB As String = ""
            sql1 = ""
            connection = New SqlConnection(conStr)
            If connection.State = ConnectionState.Closed Then
                connection.Open()
            End If
            sql1 &= "SELECT CBBTHD.[DATE]" & vbCrLf
            sql1 &= ",  REFERENCE" & vbCrLf
            sql1 &= ",  MISCCODE" & vbCrLf
            sql1 &= "FROM CBBTHD  " & vbCrLf
            sql1 &= "WHERE BATCHID = '" & BatchNo & "'" & vbCrLf
            sql1 &= "AND ENTRYNO = '" & Entry & "'"

            command = New SqlCommand(sql1, connection)
            adapter = New SqlDataAdapter(command)
            dataSt = New DataSet()
            adapter.Fill(dataSt, "DataCB")
            connection.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return dataSt.Tables("DataCB")
    End Function

#End Region

    Private Sub CheckExistDirectory()
        If Not (IO.Directory.Exists(BaseValiabled.FileEx)) Then
            IO.Directory.CreateDirectory(BaseValiabled.FileEx)
        End If
    End Sub

    Private Function GetDateForFileName(ByVal Datenow As String) As String
        Datenow = Datenow.Replace("/", "-")
        Datenow = Datenow.Replace(" ", "_")
        Datenow = Datenow.Replace(":", ";")
        Return Datenow
    End Function
    Public Shared Sub STOREDATA(ByVal dt As DataTable, ByVal MsgId_Value As String, ByVal getnow As String)
        sql1 = ""
        connection = New SqlConnection(conStr)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If

        For i = 0 To dt.Rows.Count - 1
            Dim batchno = dt.Rows(i).Item("BatchNo").ToString
            Dim Entryno = dt.Rows(i).Item("Entry").ToString
            sql1 = "INSERT INTO FMSEpayEx VALUES('" & MsgId_Value & "','" & batchno & "','" & Entryno & "','" & getnow & "')"
            Dim cmd As SqlCommand = New SqlCommand(sql1, connection)
            cmd.ExecuteNonQuery()

        Next

    End Sub

    Public Shared Sub Running(ByRef RunNo As String)
        Dim Data As String = ""
        connection = New SqlConnection(conStr)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If
        sql1 = "SELECT MAX(ID) as ID FROM FMSEPayEx"
        command = New SqlCommand(sql1, connection)
        adapter = New SqlDataAdapter(command)
        Dim dataSTRN = New DataTable()
        adapter.Fill(dataSTRN)

        connection.Close()

        For i = 0 To dataSTRN.Rows.Count - 1
            RunNo = dataSTRN.Rows(0).Item("ID").ToString
            Exit For
        Next

        If RunNo = "" Or IsDBNull(RunNo) Then
            RunNo = Format(1, "0000")
        Else
            RunNo = Format(RunNo + 1, "0000")
        End If

    End Sub
End Class

