Imports System.Xml
Imports System.Xml.Linq

Module GenerateXML
    Public DataGETP As DataTable = New DataTable()
    Public DataGETI As DataTable = New DataTable()
#Region "GETDATA"
    Sub GETALLDATA(ByRef dt As DataTable)
        Dim DGVLIST As DataGridView = clsVExport.DataGridView1
        Dim Batchno, EntryNo As String
        For i = 0 To DGVLIST.Rows.Count - 1
            Batchno = DGVLIST.Rows(i).Cells("BatchNo").Value.ToString
            EntryNo = DGVLIST.Rows(i).Cells("Entry").Value.ToString
            DataGETP = clsCGenerateFile.GetDataP(Batchno)

        Next

    End Sub

#End Region
    Sub GENXML(ByVal DGVLIST As DataGridView)

        GETALLDATA(DataGETP)

        Dim strSql As String = ""
        Dim settingPath As XmlWriterSettings = New XmlWriterSettings()
        settingPath.Indent = True
        Dim pathaddr As String
        'pathaddr = clsVConfiguration.TxtFileEx.Text + "\MCXML.xml"
        pathaddr = "D:\TTET_File_BK\MCXML.xml"

        Dim MsgId_Value, CreDtTm_Value, NbOfTxs_Value, Nm_Value, Id_Value As String

        MsgId_Value = "1911040001"
        CreDtTm_Value = "2019-10-29"
        NbOfTxs_Value = "6"
        For i = 0 To DataGETP.Rows.Count - 1
            Nm_Value = DataGETP.Rows(0).Item("BeneficiaryName").ToString
        Next
        Id_Value = "02010094"

        Dim Document = New XElement("Document")
        Dim CstmrCdtTrfInitn = New XElement("CstmrCdtTrfInitn")

        Dim GrpHdr = New XElement("GrpHdr",
                                  New XElement("MsgId", MsgId_Value),
                                  New XElement("CreDtTm", CreDtTm_Value),
                                  New XElement("NbOfTxs", NbOfTxs_Value))
        Dim InitgPty = New XElement("InitgPty",
                                    New XElement("Nm", Nm_Value))
        Dim Id = New XElement("Id")

        '#HEADER
        Dim OrgId = New XElement("OrgId")
        H_Othr(OrgId)

        '#DETAIL
        'NODE: RmtInf
        Dim RmtInf = New XElement("RmtInf")
        D_RmtInf(RmtInf)

        'NODE: PmtId
        Dim PmtId = New XElement("PmtId")
        D_PmtId(PmtId)

        'NODE: PmtTpInf
        Dim PmtTpInf = New XElement("PmtTpInf")
        D_PmtTpInf(PmtTpInf)

        'NODE: Amt
        Dim Amt = New XElement("Amt")
        D_Amt(Amt)

        'NODE: CdtrAgt
        '>> CdtrAgt value
        Dim CdtrAgt = New XElement("CdtrAgt")
        D_CdtrAgt(CdtrAgt)
        'NODE: Cdtr
        Dim Cdtr = New XElement("Cdtr")
        D_Cdtr(Cdtr)

        'NODE: CdtrAcct
        Dim CdtrAcct = New XElement("CdtrAcct")
        D_CdtrAcct(CdtrAcct)

        'NODE: UltmtCdtr
        Dim UltmtCdtr = New XElement("UltmtCdtr")

        'NODE: InstrForDbtrAgt
        Dim InstrForDbtrAgt = New XElement("InstrForDbtrAgt")

        'NODE: Tax
        Dim Tax = New XElement("Tax")

        'Document.Add(CstmrCdtTrfInitn, GrpHdr, MsgId, NbOfTxs, InitgPty, Nm, Id, OrgId, Othr)

        '#Add Header
        CstmrCdtTrfInitn.Add(GrpHdr)
        GrpHdr.Add(InitgPty)
        InitgPty.Add(Id)
        Id.Add(OrgId)

        '#Add detail
        CstmrCdtTrfInitn.Add(PmtId)
        CstmrCdtTrfInitn.Add(PmtTpInf)
        CstmrCdtTrfInitn.Add(Amt)
        CstmrCdtTrfInitn.Add(CdtrAgt)
        CstmrCdtTrfInitn.Add(Cdtr)
        CstmrCdtTrfInitn.Add(CdtrAcct)
        CstmrCdtTrfInitn.Add(RmtInf)

        Document.Add(CstmrCdtTrfInitn)


        Dim reader = Document.CreateReader()
        reader.ReadInnerXml()
        reader.MoveToContent()

        Using writer As XmlWriter = XmlWriter.Create(pathaddr, settingPath)
            writer.WriteStartDocument()
            writer.WriteRaw(reader.ReadOuterXml)

        End Using

    End Sub

#Region "ProcessDATA"

#End Region

#Region "ProcessXML"

#Region "HEADDER"
    Sub H_Othr(ByRef OrgId As XElement)
        Dim i As Integer
        i = 1
        Do While i <= 3 'Repeat othr 3 time
            Select Case i
                Case 1
                    Dim othr1 = New XElement("Othr",
                                             New XElement("Id", ">>MGEB Code1"))
                    Dim SchmeNm = New XElement("SchmeNm",
                                             New XElement("Prtry", "Mizuho Global e - Banking"))
                    othr1.Add(SchmeNm)
                    OrgId.Add(othr1)

                Case 2
                    Dim othr2 = New XElement("Othr",
                                             New XElement("Id", "Version1"))
                    Dim SchmeNm = New XElement("SchmeNm",
                                             New XElement("Prtry", ">>Program Name"))
                    othr2.Add(SchmeNm)
                    OrgId.Add(othr2)
                Case Else
                    Dim othr3 = New XElement("Othr",
                                             New XElement("Id", "4.2.7"))
                    Dim SchmeNm = New XElement("SchmeNm",
                                             New XElement("Prtry", "Mizuho XML Mapping"))
                    othr3.Add(SchmeNm)
                    OrgId.Add(othr3)
            End Select
            i = i + 1
        Loop

    End Sub
#End Region

#Region "DETAIL"

#Region "RmtInf"
    Sub D_RmtInf(ByRef RmtInf As XElement)
        Try

            For i = 0 To clsVExport.DataGridView1.Rows.Count - 1

                If IsDBNull(clsVExport.DataGridView1.Rows(i).Cells("Check").Value) Then
                    Continue For
                End If
                If clsVExport.DataGridView1.Rows(i).Cells("Check").Value = True Then
                    Dim batchno, entryno, TOTALAMOUNT, BankCode As String
                    batchno = clsVExport.DataGridView1.Rows(i).Cells("BatchNo").Value.ToString
                    entryno = clsVExport.DataGridView1.Rows(i).Cells("Entry").Value.ToString
                    'TOTALAMOUNT = clsVExport.DataGridView1.Rows(i).Cells("Amount").Value.ToString
                    'BankCode = clsVExport.DataGridView1.Rows(i).Cells("BankCode").Value.ToString

                    Dim IDINVC As String
                    Dim APPLAMOUNT As String
                    Dim INVCDESC As String
                    DataGETI = clsCGenerateFile.GetDataI(batchno, entryno)

                    Dim strRfrdDocInf As String
                    Dim strRfrdDocAmt As String
                    Dim strCdtrRefInf As String

                    For j = 0 To DataGETI.Rows.Count - 1
                        strRfrdDocInf = ""
                        strRfrdDocAmt = ""
                        strCdtrRefInf = ""

                        'ELEMENT : RfrdDocInf 
                        IDINVC = DataGETI.Rows(j).Item("IDINVC").ToString()
                        D_RmtInf_GenRfrdDocInf(IDINVC, strRfrdDocInf)
                        Dim xstrRfrdDocInf As XElement = XElement.Parse(strRfrdDocInf)
                        Dim Strd = New XElement("Strd", xstrRfrdDocInf)

                        'ELEMENT: RfrdDocAmt
                        APPLAMOUNT = DataGETI.Rows(j).Item("AMTGROSTOT").ToString
                        D_RmtInf_GenRfrdDocAmt(APPLAMOUNT, strRfrdDocAmt)
                        Dim xstrRfrdDocAmt As XElement = XElement.Parse(strRfrdDocAmt)
                        Strd.Add(xstrRfrdDocAmt)

                        'ELEMENT: CdtrRefInf
                        INVCDESC = DataGETI.Rows(j).Item("INVCDESC").ToString
                        D_RmtInf_GenCdtrRefInf(INVCDESC, strCdtrRefInf)
                        Dim xstrCdtrRefInf As XElement = XElement.Parse(strCdtrRefInf)
                        Strd.Add(xstrCdtrRefInf)

                        RmtInf.Add(Strd)
                    Next
                End If

            Next
        Catch ex As Exception
            MessageBox.Show("Error 180: " & ex.Message)
        End Try
    End Sub

    Sub D_RmtInf_GenRfrdDocInf(ByVal IDINVC As String, ByRef str As String)
        Dim RfrdDocInf = New XElement("RfrdDocInf")
        Dim Tp = New XElement("Tp")
        Dim CdOrPrtry = New XElement("CdOrPrtry", New XElement("Cd", "CINV"))
        Dim nb = New XElement("Nb", IDINVC.TrimEnd)
        RfrdDocInf.Add(Tp)
        RfrdDocInf.Add(nb)
        Tp.Add(CdOrPrtry)
        str = RfrdDocInf.ToString
    End Sub

    Sub D_RmtInf_GenRfrdDocAmt(ByVal APPLAMOUNT As String, ByRef str As String)
        Dim RfrdDocAmt = New XElement("RfrdDocAmt")
        Dim DuePyblAmt = New XElement("DuePyblAmt", New XAttribute("Ccy", "THB"), APPLAMOUNT)
        Dim RmtdAmt = New XElement("RmtdAmt", New XAttribute("Ccy", "THB"), APPLAMOUNT)
        RfrdDocAmt.Add(DuePyblAmt)
        RfrdDocAmt.Add(RmtdAmt)
        str = RfrdDocAmt.ToString
    End Sub
    Sub D_RmtInf_GenCdtrRefInf(ByVal INVCDESC As String, ByRef str As String)
        Dim CdtrRefInf = New XElement("CdtrRefInf")
        Dim Tp = New XElement("Tp")
        '>>> Validate INVDESC < 35 character
        Dim CdOrPrtry = New XElement("CdOrPrtry", New XElement("Prtry", INVCDESC))
        Dim Ref = New XElement("Ref", 1)
        CdtrRefInf.Add(Tp)
        CdOrPrtry.Add(Ref)
        Tp.Add(CdOrPrtry)
        str = CdtrRefInf.ToString
    End Sub
#End Region

#Region "PmtId"
    Sub D_PmtId(ByRef PmtId As XElement)
        '>>>
        Dim InstrId = New XElement("InstrId", "")
        Dim EndToEndId = New XElement("EndToEndId", "")

        PmtId.Add(InstrId)
        PmtId.Add(EndToEndId)
    End Sub
#End Region

#Region "PmtTpInf"
    Sub D_PmtTpInf(ByRef PmtTpInf As XElement)
        Dim LclInstrm = New XElement("LclInstrm", New XElement("Prtry", "DB*MC"))
        Dim CtgyPurp = New XElement("CtgyPurp", New XElement("Cd", "SUPP"))

        PmtTpInf.Add(LclInstrm)
        PmtTpInf.Add(CtgyPurp)

    End Sub

#End Region

#Region "Amt"
    Sub D_Amt(ByRef Amt As XElement)
        '>>AMT value 
        Dim TOTALAMOUNT As String
        Try
            For i = 0 To clsVExport.DataGridView1.Rows.Count - 1

                If IsDBNull(clsVExport.DataGridView1.Rows(i).Cells("Check").Value) Then
                    Continue For
                End If
                If clsVExport.DataGridView1.Rows(i).Cells("Check").Value = True Then
                    TOTALAMOUNT = clsVExport.DataGridView1.Rows(i).Cells("Amount").Value.ToString
                    Dim InstdAmt = New XElement("InstdAmt", New XAttribute("Ccy", "THB"), TOTALAMOUNT)
                    Amt.Add(InstdAmt)
                End If
            Next

        Catch ex As Exception
            MessageBox.Show("Error 300:" & ex.Message)
        End Try
    End Sub

#End Region

#Region "ChrgBr"
    Sub D_ChrgBr(ByRef ChrgBr As XElement)
        '>>ChrgBr value
        ChrgBr = New XElement("ChrgBr", "")
    End Sub

#End Region

#Region "CdtrAgt"
    Sub D_CdtrAgt(ByRef CdtrAgt As XElement)
        Try
            Dim BankCode As String
            For i = 0 To clsVExport.DataGridView1.Rows.Count - 1
                BankCode = ""
                If IsDBNull(clsVExport.DataGridView1.Rows(i).Cells("Check").Value) Then
                    Continue For
                End If
                If clsVExport.DataGridView1.Rows(i).Cells("Check").Value = True Then
                    BankCode = clsVExport.DataGridView1.Rows(i).Cells("BankCode").Value.ToString()
                    Dim FinInstnId = New XElement("FinInstnId")
                    Dim ClrSysMmbId = New XElement("ClrSysMmbId")
                    Dim ClrSysId = New XElement("ClrSysId", New XElement("Cd", "THCBC"))
                    Dim MmbId = New XElement("MmbId", BankCode)

                    CdtrAgt.Add(FinInstnId)
                    FinInstnId.Add(ClrSysMmbId)
                    ClrSysMmbId.Add(ClrSysId)
                    ClrSysMmbId.Add(MmbId)
                End If
            Next

        Catch ex As Exception
            MessageBox.Show("Error 320: " & ex.Message)
        End Try
    End Sub

#End Region

#Region "Cdtr"
    Sub D_Cdtr(ByRef Cdtr As XElement)
        Try
            For i = 0 To clsVExport.DataGridView1.Rows.Count - 1

                If IsDBNull(clsVExport.DataGridView1.Rows(i).Cells("Check").Value) Then
                    Continue For
                End If
                If clsVExport.DataGridView1.Rows(i).Cells("Check").Value = True Then
                    Dim BatchNo, CNTBTCH, Address, VendName, Acct, TAXID As String
                    BatchNo = clsVExport.DataGridView1.Rows(i).Cells("BatchNo").Value.ToString
                    Address = ""
                    VendName = ""
                    TAXID = ""
                    For j = 0 To DataGETP.Rows.Count - 1
                        CNTBTCH = ""
                        CNTBTCH = DataGETP.Rows(j).Item("CNTBTCH").ToString
                        If BatchNo = CNTBTCH Then
                            Address = DataGETP.Rows(j).Item("ADDR").ToString
                            VendName = DataGETP.Rows(j).Item("BeneficiaryName").ToString
                            TAXID = DataGETP.Rows(j).Item("TAXID").ToString
                        End If
                        Dim Nm = New XElement("Nm", VendName)
                        Dim PstlAdr = New XElement("PstlAdr", New XElement("PstCd", ""), New XElement("TwnNm", ""), New XElement("Ctry", "TH"), New XElement("AdrLine", Address))
                        Cdtr.Add("Nm")
                        Cdtr.Add("PstlAdr")
                    Next

                End If
            Next

        Catch ex As Exception
            MessageBox.Show("Error 370: " & ex.Message)
        End Try
    End Sub

#End Region

#Region "CdtrAcct"
    Sub D_CdtrAcct(ByRef CdtrAcct As XElement)
        Try
            For i = 0 To clsVExport.DataGridView1.Rows.Count - 1

                If IsDBNull(clsVExport.DataGridView1.Rows(i).Cells("Check").Value) Then
                    Continue For
                End If
                If clsVExport.DataGridView1.Rows(i).Cells("Check").Value = True Then
                    Dim BatchNo, CNTBTCH, Acct As String
                    BatchNo = clsVExport.DataGridView1.Rows(i).Cells("BatchNo").ToString
                    Acct = ""
                    For j = 0 To DataGETP.Rows.Count - 1
                        CNTBTCH = ""
                        CNTBTCH = DataGETP.Rows(j).Item("CNTBTCH").ToString
                        If BatchNo = CNTBTCH Then
                            Acct = DataGETP.Rows(j).Item("AccountNo").ToString
                        End If
                        Dim Id = New XElement("Id", New XElement("Othr", New XElement("Id", Acct)))
                        CdtrAcct.Add(Id)
                    Next
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Error 400: " & ex.Message)
        End Try


    End Sub
#End Region


#End Region

#End Region

End Module
