Imports System.Xml
Imports System.Xml.Linq

Module GenerateXML
    Public DataGETP As DataTable = New DataTable()
    Public DataGETI As DataTable = New DataTable()
    Public DataGETW As DataTable = New DataTable()
    Public DataGETC As DataTable = New DataTable()

#Region "GETDATA"
    'Sub GENXML(ByVal DGVLIST As DataGridView)
    Sub GENXML(ByVal dtMC As DataTable)
        Try
            Dim strSql As String = ""
            Dim settingPath As XmlWriterSettings = New XmlWriterSettings()
            settingPath.Indent = True
            Dim pathaddr As String
            Dim getnow, RunNo, gettime As String
            getnow = Date.Now.ToString("yyyyMMdd")
            gettime = Date.Now.ToString("Hmm")
            'pathaddr = clsVConfiguration.TxtFileEx.Text + "\MCXML.xml"
            pathaddr = Config.GetPrivateProfileString("CONFIG", "FILEEx", "")
            pathaddr = pathaddr + "\MCXML_" & getnow & gettime & ".xml"

            Dim MsgId_Value, CreDtTm_Value, NbOfTxs_Value, Nm_Value, Id_Value As String
            clsCGenerateFile.Running(RunNo)
            MsgId_Value = "MC" & getnow & "_" & RunNo
            CreDtTm_Value = DateTime.Now.ToString("o")
            NbOfTxs_Value = clsVExport.CountSelect.ToString

            Nm_Value = "KST"
            Id_Value = "02010222"

            Dim Document = New XElement("Document")

            Dim xsdValue As XNamespace = "http://www.w3.org/2001/XMLSchema"
            Dim xsdATT As XAttribute = New XAttribute(XNamespace.Xmlns + "xsd", xsdValue)
            Document.Add(xsdATT)

            Dim xsiValue As XNamespace = "http://www.w3.org/2001/XMLSchema-instance"
            Dim xsiATT As XAttribute = New XAttribute(XNamespace.Xmlns + "xsi", xsiValue)
            Document.Add(xsiATT)

            Dim urnValue As XNamespace = "urn:std:iso:20022:tech:xsd:pain.001.001.04"
            Dim urnATT As XAttribute = New XAttribute("xmlns", urnValue)
            Document.Add(urnATT)

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

            '#Add Header
            CstmrCdtTrfInitn.Add(GrpHdr)
            GrpHdr.Add(InitgPty)
            InitgPty.Add(Id)
            Id.Add(OrgId)

            '#DETAIL
            Dim PmtInf As New XElement("PmtInf")
            'NODE: Dbtr 
            Dim PmtInfId = New XElement("PmtInfId", "001")
            PmtInf.Add(PmtInfId)

            'NODE: PmtMtd 
            Dim PmtMtd = New XElement("PmtMtd", "TRF")
            PmtInf.Add(PmtMtd)

            'NODE: ReqdExctnDt
            Dim ReqdExctnDt As XElement
            D_ReqdExctnDt(ReqdExctnDt, dtMC)
            PmtInf.Add(ReqdExctnDt)

            'NODE: Dbtr 
            Dim Dbtr = New XElement("Dbtr", New XElement("Nm", Nm_Value))
            PmtInf.Add(Dbtr)

            'NODE: DbtrAcct
            Dim DbtrAcct = New XElement("DbtrAcct")
            D_DbtrAcct(DbtrAcct)
            PmtInf.Add(DbtrAcct)

            'NODE: DbtrAgt
            Dim DbtrAgt = New XElement("DbtrAgt", New XElement("FinInstnId", New XElement("BICFI", "MHCBTHBK")))
            PmtInf.Add(DbtrAgt)

            'NODE: RmtInf (Insert Each Entry)

            Dim Batchno, EntryNo As String
            Dim cntRnd As Integer = 1
            For i = 0 To dtMC.Rows.Count - 1
                Dim RmtInf = New XElement("RmtInf")
                Dim strRmtInf As String
                strRmtInf = ""
                Batchno = ""
                EntryNo = ""

                Batchno = dtMC.Rows(i).Item("BatchNo").ToString
                EntryNo = dtMC.Rows(i).Item("Entry").ToString

                Dim Ustrd = New XElement("Ustrd")
                D_Ustrd(Ustrd, Batchno, EntryNo)
                RmtInf.Add(Ustrd)
                D_RmtInf(RmtInf, Batchno, EntryNo)
                '***>>>
                GENCdtTrfTxInf(RmtInf, strRmtInf, Batchno, EntryNo, cntRnd)
                Dim xstrRmtInf As XElement = XElement.Parse(strRmtInf)

                PmtInf.Add(xstrRmtInf)
                cntRnd = cntRnd + 1
            Next

            CstmrCdtTrfInitn.Add(PmtInf)

            '#GEN DOC
            Document.Add(CstmrCdtTrfInitn)

            Dim reader = Document.CreateReader()
            reader.ReadInnerXml()
            reader.MoveToContent()

            Using writer As XmlWriter = XmlWriter.Create(pathaddr, settingPath)
                writer.WriteStartDocument()
                writer.WriteRaw(reader.ReadOuterXml)

            End Using
            'Insert Data to FMSEPayEx
            clsCGenerateFile.STOREDATA(dtMC, MsgId_Value, getnow)
            DisplayValidate(dtErr)
        Catch ex As Exception
            MessageBox.Show("Error 140: " & ex.Message)
        End Try
    End Sub
    Sub GENCdtTrfTxInf(ByVal RmtInf As XElement, ByRef strRmtInf As String, ByVal Batchno As String, ByVal EntryNo As String, ByVal cntRnd As Integer)
        Try
            Dim CdtTrfTxInf = New XElement("CdtTrfTxInf")
            'NODE : PmtId
            Dim PmtId = New XElement("PmtId")
            D_PmtId(PmtId, Batchno, EntryNo, cntRnd)
            CdtTrfTxInf.Add(PmtId)

            'NODE :  PmtTpInf
            Dim PmtTpInf = New XElement("PmtTpInf")
            D_PmtTpInf(PmtTpInf)
            CdtTrfTxInf.Add(PmtTpInf)

            'NODE: Amt
            Dim Amt = New XElement("Amt")
            D_Amt(Amt, Batchno, EntryNo)
            CdtTrfTxInf.Add(Amt)

            'NODE: ChrgBr
            Dim ChrgBr = New XElement("ChrgBr")
            D_ChrgBr(ChrgBr)
            CdtTrfTxInf.Add(ChrgBr)

            'NODE: CdtrAgt
            '>> CdtrAgt value
            Dim CdtrAgt = New XElement("CdtrAgt")
            D_CdtrAgt(CdtrAgt, Batchno, EntryNo)
            CdtTrfTxInf.Add(CdtrAgt)
            'NODE: Cdtr
            Dim Cdtr = New XElement("Cdtr")
            D_Cdtr(Cdtr, Batchno, EntryNo)
            CdtTrfTxInf.Add(Cdtr)

            'NODE: CdtrAcct
            Dim CdtrAcct = New XElement("CdtrAcct")
            D_CdtrAcct(CdtrAcct, Batchno, EntryNo)
            CdtTrfTxInf.Add(CdtrAcct)

            'NODE: InstrForDbtrAgt
            Dim InstrForDbtrAgt As XElement = Nothing
            D_InstrForDbtrAgt(InstrForDbtrAgt, Batchno)
            CdtTrfTxInf.Add(InstrForDbtrAgt)

            'NODE: TAX
            Dim Tax As XElement = Nothing
            D_Tax(Tax, Batchno, EntryNo)
            CdtTrfTxInf.Add(Tax)

            'NODE : RmtInf
            CdtTrfTxInf.Add(RmtInf)

            strRmtInf = CdtTrfTxInf.ToString
        Catch ex As Exception
            MessageBox.Show("Error 195: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "HEADDER"
    Sub H_Othr(ByRef OrgId As XElement)
        Try
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
                                                 New XElement("Prtry", "SAGE300"))
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
        Catch ex As Exception
            MessageBox.Show("Error 233: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "DETAIL"
#Region "DbtrAcct"
    Sub D_DbtrAcct(ByRef DbtrAcct As XElement)
        Dim Id = New XElement("Id")
        Dim Othr = New XElement("Othr", New XElement("Id", "H15764665544"))
        Id.Add(Othr)
        DbtrAcct.Add(Id)

    End Sub
#End Region
#Region "Ustrd"
    Sub D_Ustrd(ByRef Ustrd As XElement, ByVal Batchno As String, ByVal EntryNo As String)
        Try
            Dim TAXIDVENDOR As String = ""
            DataGETP = clsCGenerateFile.GetDataP(Batchno)

            For n = 0 To DataGETP.Rows.Count - 1
                TAXIDVENDOR = DataGETP.Rows(n).Item("TAXID").ToString
            Next

            Dim strUstrdInv = New XElement("Ustrd", "InvoiceMC" & Batchno & EntryNo)
            Dim strUstrdTax = New XElement("Ustrd", "1|0105539110845")
            Dim strUstrdTaxVEND = New XElement("Ustrd", "1|" & TAXIDVENDOR.TrimEnd)
            Ustrd.Add(strUstrdInv)
            For i As Integer = 0 To 5
                Dim strUstrd = New XElement("Ustrd", "||")
                Ustrd.Add(strUstrd)
            Next
            Ustrd.Add(strUstrdTax)
            Ustrd.Add(strUstrdTaxVEND)
        Catch ex As Exception
            MessageBox.Show("Error 269: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "RmtInf"
    Sub D_RmtInf(ByRef RmtInf As XElement, ByVal Batchno As String, ByVal EntryNo As String)
        Try
            For i = 0 To dtMC.Rows.Count - 1

                If IsDBNull(dtMC.Rows(i).Item("Check").ToString) Then
                    Continue For
                End If

                Dim IDINVC As String
                Dim APPLAMOUNT As String
                Dim INVCDESC As String
                DataGETI = clsCGenerateFile.GetDataI(Batchno, EntryNo)
                Dim BatchnoDGV, ENTRYDGV As String
                BatchnoDGV = dtMC.Rows(i).Item("BatchNo").ToString
                ENTRYDGV = dtMC.Rows(i).Item("Entry").ToString
                Dim strRfrdDocInf As String
                Dim strRfrdDocAmt As String
                Dim strCdtrRefInf As String
                For j = 0 To DataGETI.Rows.Count - 1
                    strRfrdDocInf = ""
                    strRfrdDocAmt = ""
                    strCdtrRefInf = ""
                    If Batchno = BatchnoDGV And EntryNo = ENTRYDGV Then
                        'ELEMENT : RfrdDocInf 
                        IDINVC = DataGETI.Rows(j).Item("IDINVC").ToString()
                        D_RmtInf_GenRfrdDocInf(IDINVC.Trim, strRfrdDocInf)
                        Dim xstrRfrdDocInf As XElement = XElement.Parse(strRfrdDocInf)
                        Dim Strd = New XElement("Strd", xstrRfrdDocInf)

                        'ELEMENT: RfrdDocAmt
                        APPLAMOUNT = DataGETI.Rows(j).Item("AMTGROSTOT").ToString
                        D_RmtInf_GenRfrdDocAmt(APPLAMOUNT.TrimEnd, strRfrdDocAmt)
                        Dim xstrRfrdDocAmt As XElement = XElement.Parse(strRfrdDocAmt)
                        Strd.Add(xstrRfrdDocAmt)

                        'ELEMENT: CdtrRefInf
                        INVCDESC = DataGETI.Rows(j).Item("INVCDESC").ToString
                        D_RmtInf_GenCdtrRefInf(INVCDESC.TrimEnd, strCdtrRefInf)
                        Dim xstrCdtrRefInf As XElement = XElement.Parse(strCdtrRefInf)
                        Strd.Add(xstrCdtrRefInf)

                        'ELEMENT: AddtlRmtInf
                        Dim AddtlRmtInf = New XElement("AddtlRmtInf", "099")
                        Strd.Add(AddtlRmtInf)

                        RmtInf.Add(Strd)
                    End If
                Next
            Next
        Catch ex As Exception
            MessageBox.Show("Error 325: " & ex.Message)
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
        Dim CdOrPrtry = New XElement("CdOrPrtry", New XElement("Prtry", INVCDESC.TrimEnd))
        Dim Ref = New XElement("Ref", "1 B")
        CdtrRefInf.Add(Tp)
        CdtrRefInf.Add(Ref)
        Tp.Add(CdOrPrtry)
        str = CdtrRefInf.ToString
    End Sub
#End Region

#Region "PmtId"
    Sub D_PmtId(ByRef PmtId As XElement, ByVal Batchno As String, ByVal EntryNo As String, ByVal cntRnd As Integer)

        Try
            For i = 0 To dtMC.Rows.Count - 1
                Dim batchnoDGV = dtMC.Rows(i).Item("BatchNo").ToString
                Dim EntryNoDGV = dtMC.Rows(i).Item("Entry").ToString

                If IsDBNull(dtMC.Rows(i).Item("Check").ToString) Then
                    Continue For
                End If

                If Batchno = batchnoDGV And EntryNo = EntryNoDGV Then
                    Dim InstrId As XElement
                    D_PmtId_GenInstrId(InstrId, Batchno, EntryNo)
                    Dim EndToEndId = New XElement("EndToEndId", Format(cntRnd, "0000000000"))
                    PmtId.Add(InstrId)
                    PmtId.Add(EndToEndId)
                End If
            Next

        Catch ex As Exception
            MessageBox.Show("Error 385: " & ex.Message)
        End Try
    End Sub

    Sub D_PmtId_GenInstrId(ByRef InstrId As XElement, ByVal batchno As String, ByVal entryno As String)
        DataGETC.Rows.Clear()
        DataGETC.Columns.Clear()

        Dim RefCheque As String
        DataGETC = clsCGenerateFile.GetDataCB(batchno, entryno)

        For i = 0 To DataGETC.Rows.Count - 1
            RefCheque = DataGETC.Rows(i).Item("REFERENCE").ToString
            InstrId = New XElement("InstrId", RefCheque.TrimEnd)
        Next

    End Sub
#End Region

#Region "PmtTpInf"
    Sub D_PmtTpInf(ByRef PmtTpInf As XElement)
        Dim LclInstrm = New XElement("LclInstrm", New XElement("Prtry", "DB*MC"))
        Dim CtgyPurp = New XElement("CtgyPurp", New XElement("Cd", "SUPP"))

        PmtTpInf.Add(LclInstrm)
        PmtTpInf.Add(CtgyPurp)

    End Sub

    Sub D_ReqdExctnDt(ByRef ReqdExctnDt As XElement, ByVal dtMC As DataTable)
        Try
            Dim CBDATE As String
            Dim BatchnoDGV As String = ""
            Dim ENTRYDGV As String = ""
            DataGETC.Rows.Clear()
            DataGETC.Columns.Clear()

            For i = 0 To dtMC.Rows.Count - 1

                If IsDBNull(dtMC.Rows(i).Item("Check").ToString) Then
                    Continue For
                End If
                If dtMC.Rows(i).Item("Check").ToString = True Then
                    BatchnoDGV = dtMC.Rows(i).Item("BatchNo").ToString
                    ENTRYDGV = dtMC.Rows(i).Item("Entry").ToString
                    DataGETC = clsCGenerateFile.GetDataCB(BatchnoDGV, ENTRYDGV)
                    For j = 0 To DataGETC.Rows.Count - 1
                        CBDATE = DataGETC.Rows(j).Item("DATE").ToString.TrimEnd
                        CBDATE = CBDATE.Substring(0, 4) & "-" & CBDATE.Substring(4, 2) & "-" & CBDATE.Substring(6, 2)
                        ReqdExctnDt = New XElement("ReqdExctnDt", CBDATE)
                        Exit For
                    Next
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Error 437: " & ex.Message)
        End Try
    End Sub

#End Region

#Region "Amt"
    Sub D_Amt(ByRef Amt As XElement, ByVal Batchno As String, ByVal EntryNo As String)
        '>>AMT value 
        Dim TOTALAMOUNT As String
        Try
            For i = 0 To dtMC.Rows.Count - 1
                Dim batchnoDGV = dtMC.Rows(i).Item("BatchNo").ToString
                Dim EntryNoDGV = dtMC.Rows(i).Item("Entry").ToString

                If IsDBNull(dtMC.Rows(i).Item("Check").ToString) Then
                    Continue For
                End If
                If dtMC.Rows(i).Item("Check").ToString = True Then
                    If Batchno = batchnoDGV And EntryNo = EntryNoDGV Then
                        TOTALAMOUNT = dtMC.Rows(i).Item("Amount").ToString
                        Dim InstdAmt = New XElement("InstdAmt", New XAttribute("Ccy", "THB"), TOTALAMOUNT)
                        Amt.Add(InstdAmt)
                    End If
                End If
            Next

        Catch ex As Exception
            MessageBox.Show("Error 467:" & ex.Message)
        End Try
    End Sub

#End Region

#Region "ChrgBr"
    Sub D_ChrgBr(ByRef ChrgBr As XElement)
        '>>ChrgBr value
        ChrgBr = New XElement("ChrgBr", "CRED")
    End Sub

#End Region

#Region "CdtrAgt"
    Sub D_CdtrAgt(ByRef CdtrAgt As XElement, ByVal Batchno As String, ByVal EntryNo As String)

        Try
            Dim BankCode As String
            For i = 0 To dtMC.Rows.Count - 1
                DataGETP.Rows.Clear()
                DataGETP.Columns.Clear()
                DataGETP = clsCGenerateFile.GetDataP(Batchno)
                Dim batchnoDGV As String = dtMC.Rows(i).Item("BatchNo").ToString
                Dim EntryNoDGV As String = dtMC.Rows(i).Item("Entry").ToString
                If dtMC.Rows(i).Item("Check").ToString = True Then
                    BankCode = ""
                    For j = 0 To DataGETP.Rows.Count - 1
                        BankCode = DataGETP.Rows(0).Item("BankCode").ToString
                    Next
                    If IsDBNull(dtMC.Rows(i).Item("Check").ToString) Then
                        Continue For
                    End If
                    If Batchno = batchnoDGV And EntryNo = EntryNoDGV Then
                        Dim BANKNAME As String = GenMaster.getBankname(BankCode)
                        Dim strClrSysMmbId As String = ""
                        Dim FinInstnId = New XElement("FinInstnId")
                        D_CdtrAgt_GENClrSysMmbId(BankCode.TrimEnd, strClrSysMmbId)
                        Dim xstrClrSysMmbId As XElement = XElement.Parse(strClrSysMmbId)
                        FinInstnId.Add(xstrClrSysMmbId)
                        Dim Nm = New XElement("Nm", BANKNAME)
                        FinInstnId.Add(Nm)
                        CdtrAgt.Add(FinInstnId)
                    End If
                End If
            Next

        Catch ex As Exception
            MessageBox.Show("Error 515: " & ex.Message)
        End Try
    End Sub

    Sub D_CdtrAgt_GENClrSysMmbId(ByVal BANKCODE As String, ByRef str As String)
        Dim ClrSysMmbId = New XElement("ClrSysMmbId")
        Dim ClrSysId = New XElement("ClrSysId", New XElement("Cd", "THCBC"))
        Dim MmbId = New XElement("MmbId", BANKCODE.TrimEnd)

        ClrSysMmbId.Add(ClrSysId)
        ClrSysMmbId.Add(MmbId)

        str = ClrSysMmbId.ToString

    End Sub

#End Region

#Region "Cdtr"
    Sub D_Cdtr(ByRef Cdtr As XElement, ByVal Batchno As String, ByVal EntryNo As String)
        Try
            For i = 0 To dtMC.Rows.Count - 1

                If IsDBNull(dtMC.Rows(i).Item("Check").ToString) Then
                    Continue For
                End If
                If dtMC.Rows(i).Item("Check").ToString = True Then
                    DataGETP.Rows.Clear()
                    DataGETP.Columns.Clear()
                    DataGETP = clsCGenerateFile.GetDataP(Batchno)
                    Dim batchnoDGV, EntryNoDGV, Address, VendName, TAXID, POSTCODE, CITY As String
                    batchnoDGV = dtMC.Rows(i).Item("BatchNo").ToString
                    EntryNoDGV = dtMC.Rows(i).Item("Entry").ToString
                    Address = ""
                    VendName = ""
                    TAXID = ""
                    POSTCODE = ""
                    CITY = ""
                    Dim strPstlAdr As String = ""
                    For j = 0 To DataGETP.Rows.Count - 1
                        Address = DataGETP.Rows(j).Item("ADDR").ToString
                        POSTCODE = DataGETP.Rows(j).Item("POSTCODE").ToString
                        CITY = DataGETP.Rows(j).Item("CITY").ToString
                        VendName = DataGETP.Rows(j).Item("BeneficiaryName").ToString
                        TAXID = DataGETP.Rows(j).Item("TAXID").ToString
                    Next
                    If Batchno = batchnoDGV And EntryNo = EntryNoDGV Then
                        chkAddr(Address)
                        Dim Nm = New XElement("Nm", VendName.TrimEnd)
                        D_Cdtr_GENPstlAdr(Address.TrimEnd, strPstlAdr, POSTCODE, CITY)
                        Dim xstrPstlAdr As XElement = XElement.Parse(strPstlAdr)
                        Cdtr.Add(Nm)
                        Cdtr.Add(xstrPstlAdr)
                    End If
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Error 570: " & ex.Message)
        End Try
    End Sub
    Sub D_Cdtr_GENPstlAdr(ByVal Address As String, ByRef str As String, ByVal POSTCODE As String, ByVal CITY As String)
        Dim PstlAdr = New XElement("PstlAdr", New XElement("PstCd", POSTCODE.TrimEnd), New XElement("TwnNm", CITY.TrimEnd), New XElement("Ctry", "TH"), New XElement("AdrLine", Address.TrimEnd))
        str = PstlAdr.ToString
    End Sub
#End Region

#Region "CdtrAcct"
    Sub D_CdtrAcct(ByRef CdtrAcct As XElement, ByVal BatchNo As String, ByVal EntryNo As String)
        Try
            For i = 0 To dtMC.Rows.Count - 1

                If IsDBNull(dtMC.Rows(i).Item("Check").ToString) Then
                    Continue For
                End If
                If dtMC.Rows(i).Item("Check").ToString = True Then
                    DataGETP.Rows.Clear()
                    DataGETP.Columns.Clear()
                    DataGETP = clsCGenerateFile.GetDataP(BatchNo)
                    Dim BatchDGV, EntryNoDGV, Acct As String
                    BatchDGV = dtMC.Rows(i).Item("BatchNo").ToString
                    EntryNoDGV = dtMC.Rows(i).Item("Entry").ToString
                    Acct = ""
                    For j = 0 To DataGETP.Rows.Count - 1
                        Acct = DataGETP.Rows(j).Item("AccountNo").ToString
                    Next
                    If BatchNo = BatchDGV And EntryNo = EntryNoDGV Then
                        Dim Id = New XElement("Id", New XElement("Othr", New XElement("Id", Acct.TrimEnd)))
                        CdtrAcct.Add(Id)
                    End If
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Error 606: " & ex.Message)
        End Try

    End Sub
#End Region

#Region "InstrForDbtrAgt"
    Sub D_InstrForDbtrAgt(ByRef InstrForDbtrAgt As XElement, ByVal batchno As String)
        Dim optService As String = "0"
        Dim arrStr(3) As String
        DataGETP = clsCGenerateFile.GetDataP(batchno)
        Try
            For i = 0 To DataGETP.Rows.Count - 1
                optService = DataGETP.Rows(i).Item("COUNTRY").ToString
                arrStr = optService.Split(",")
                If arrStr.Count >= 4 Then
                    If arrStr(3).ToString <> "" Or IsNumeric(arrStr(3)) = True Then
                        optService = arrStr(3).ToString
                    Else
                        optService = "0"
                    End If
                Else
                    optService = "0"
                End If
            Next
        Catch ex As Exception
            optService = "0"
        End Try
        InstrForDbtrAgt = New XElement("InstrForDbtrAgt", "O/" & optService.TrimEnd & "/O")
    End Sub

#End Region

#Region "Tax"
    Sub D_Tax(ByRef Tax As XElement, ByVal Batchno As String, ByVal EntryNo As String)
        DataGETP.Rows.Clear()
        DataGETP.Columns.Clear()
        DataGETP.Clear()
        Tax = New XElement("Tax")
        Dim TaxId, RefNbValue As String
        RefNbValue = Batchno.Remove(0, 3) & EntryNo.Remove(0, 2)

        DataGETP = clsCGenerateFile.GetDataP(Batchno)
        For i = 0 To DataGETP.Rows.Count - 1
            TaxId = DataGETP.Rows(0).Item("TAXID").ToString.TrimEnd
            Exit For
        Next

        Dim TaxTp As String = "002"
        Dim arrStr(3) As String
        DataGETP = clsCGenerateFile.GetDataP(Batchno)
        Try
            For i = 0 To DataGETP.Rows.Count - 1
                TaxTp = DataGETP.Rows(i).Item("COUNTRY").ToString.TrimEnd
                arrStr = TaxTp.Split(",")
                If arrStr.Count >= 4 Then
                    If arrStr(2).ToString <> "" Then
                        TaxTp = arrStr(2).ToString
                    Else
                        TaxTp = "002"
                    End If
                Else
                    TaxTp = "002"
                End If
            Next
        Catch ex As Exception
            TaxTp = "002"
        End Try

        Dim Cdtr = New XElement("Cdtr", New XElement("TaxId", TaxId.TrimEnd), New XElement("TaxTp", TaxTp))
        Dim RefNb = New XElement("RefNb", "WHT" & RefNbValue)
        Dim mtd = New XElement("mtd", "B")
        Dim Rcrd As XElement
        D_Tax_GenRcrd(Rcrd, Batchno, EntryNo)

        Tax.Add(Cdtr)
        Tax.Add(RefNb)
        Tax.Add(mtd)
        Tax.Add(Rcrd)

    End Sub

    Sub D_Tax_GenRcrd(ByRef Rcrd As XElement, ByVal batchno As String, ByVal EntryNo As String)

        Dim TpValue As Integer
        Dim WHTCODE As String = ""
        Dim getWHTCODE As String = ""
        Dim getAddtlInf As String = ""
        Dim arrAddt(3) As String
        DataGETW.Rows.Clear()
        DataGETW.Columns.Clear()
        DataGETW.Clear()

        DataGETW = clsCGenerateFile.GetDataW(batchno, EntryNo)
        Rcrd = New XElement("Rcrd")

        If DataGETW.Rows.Count = 0 Then
            TpValue = 0
        Else
            TpValue = 1
        End If

        For i = 0 To DataGETW.Rows.Count - 1
            getWHTCODE = DataGETW.Rows(i).Item("IDDISTCODE").ToString.TrimEnd
            Select Case getWHTCODE
                Case "T31"
                    WHTCODE = "016"
                Case "T33"
                    WHTCODE = "007"
                Case "T35"
                    WHTCODE = "009"
                Case "T531"
                    WHTCODE = "016"
                Case "T532"
                    WHTCODE = "014"
                Case "T533"
                    WHTCODE = "007"
                Case "T535"
                    WHTCODE = "009"
                Case Else
                    WHTCODE = "007"
            End Select
            Try
                getAddtlInf = DataGETW.Rows(i).Item("TEXTREF").ToString.TrimEnd
                arrAddt = getAddtlInf.Split(",")
                If arrAddt.Count >= 4 Then
                    If arrAddt(2).ToString <> "" Then
                        getAddtlInf = arrAddt(2).ToString
                    Else
                        getAddtlInf = "etc"
                    End If
                Else
                    getAddtlInf = "etc"
                End If

            Catch ex As Exception
                getAddtlInf = "etc"
            End Try
        Next

        Dim TP = New XElement("TP", TpValue)
        Dim Ctgy = New XElement("Ctgy", 0)

        Dim CtgyDtls = New XElement("CtgyDtls", WHTCODE)
        Dim TaxAmt As XElement
        D_Tax_GenTaxAmt(TaxAmt, DataGETW)

        Rcrd.Add(TP)
        Rcrd.Add(Ctgy)
        Rcrd.Add(CtgyDtls)
        Rcrd.Add(TaxAmt)
        Dim AddtlInf = New XElement("AddtlInf", getAddtlInf)
        Rcrd.Add(AddtlInf)

    End Sub
    Sub D_Tax_GenTaxAmt(ByRef TaxAmt As XElement, ByVal DataGETW As DataTable)
        Dim TaxbaseAmt As String = ""
        Dim TOTAMT As String = ""
        Dim Rate As Double
        Dim WHTAMOUNT As String = ""

        For i = 0 To DataGETW.Rows.Count - 1
            TaxbaseAmt = DataGETW.Rows(i).Item("TAXBASE").ToString.TrimEnd
            TOTAMT = DataGETW.Rows(i).Item("TOTAMOUNT").ToString.TrimEnd
            WHTAMOUNT = DataGETW.Rows(i).Item("WHTAMOUNT").ToString.TrimEnd
        Next
        If TaxbaseAmt <> 0 Then
            Rate = (WHTAMOUNT / TaxbaseAmt) * 100
        Else
            Rate = 0
        End If

        TaxAmt = New XElement("TaxAmt")
        Dim taxRate = New XElement("Rate", Math.Round(Rate))
        Dim TaxblBaseAmt = New XElement("TaxblBaseAmt", TaxbaseAmt)
        Dim TltAmt = New XElement("TltAmt", TOTAMT)
        Dim Dtls As XElement
        D_Tax_GenDtls(Dtls, WHTAMOUNT)

        TaxAmt.Add(taxRate)
        TaxAmt.Add(TaxblBaseAmt)
        TaxAmt.Add(TltAmt)
        TaxAmt.Add(Dtls)
    End Sub

    Sub D_Tax_GenDtls(ByRef Dtls As XElement, ByVal WHTAMOUNT As String)
        Dtls = New XElement("Dtls")
        Dim AmtTax = New XElement("Amt", WHTAMOUNT)
        Dtls.Add(AmtTax)
    End Sub

#End Region

#End Region

End Module
