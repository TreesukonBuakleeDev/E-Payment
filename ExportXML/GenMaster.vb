Imports System.Data.SqlClient
Public Class GenMaster


    Public dtbank As DataTable = New DataTable()
    Public Shared Function checkExistView() As DataTable
        Dim dtchk As DataTable = New DataTable()
        connection = New SqlConnection(conStr)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If
        Dim STREXIST As String
        STREXIST = "SELECT name FROM sys.tables WHERE name = 'FMSEPAYBANKMASTER'"
        Dim cmdquery As SqlDataAdapter = New SqlDataAdapter(STREXIST, connection)
        cmdquery.Fill(dtchk)
        If dtchk.Rows.Count = 0 Then
            Dim str As String
            str = "CREATE TABLE FMSEPAYBANKMASTER " & vbCrLf
            str &= "(BANKCODE VARCHAR(3)," & vbCrLf
            str &= "BANKNAME VARCHAR(140))"

            str &= "INSERT INTO FMSEPAYBANKMASTER" & vbCrLf
            str &= "(BANKCODE,BANKNAME)" & vbCrLf
            str &= "VALUES" & vbCrLf
            str &= " ('001','BANK OF THAILAND')," & vbCrLf
            str &= "('002','BANGKOK BANK PUBLIC COMPANY LTD.')," & vbCrLf
            str &= "('004','KASIKORNBANK PUBLIC COMPANY LTD.')," & vbCrLf
            str &= "('005','THE ROYAL BANK OF SCOTLAND PLC')," & vbCrLf
            str &= "('006','KRUNG THAI BANK PUBLIC COMPANY LTD.')," & vbCrLf
            str &= "('008','JPMORGAN CHASE BANK, NATIONAL ASSOCIATION')," & vbCrLf
            str &= "('009','OVER SEA-CHINESE BANKING CORPORATION LIMITED')," & vbCrLf
            str &= "('011','TMB BANK PUBLIC COMPANY LIMITED')," & vbCrLf
            str &= "('014','SIAM COMMERCIAL BANK PUBLIC COMPANY LTD.')," & vbCrLf
            str &= "('017','CITIBANK, N.A.')," & vbCrLf
            str &= "('018','SUMITOMO MITSUI BANKING CORPORATION')," & vbCrLf
            str &= "('020','STANDARD CHARTERED BANK (THAI) PUBLIC COMPANY LIMITED')," & vbCrLf
            str &= "('022','CIMB THAI BANK Public Company Limited')," & vbCrLf
            str &= "('023','RHB BANK BERHAD')," & vbCrLf
            str &= "('024','UNITED OVERSEAS BANK (THAI) PUBLIC COMPANY LIMITED')," & vbCrLf
            str &= "('025','BANK OF AYUDHYA PUBLIC COMPANY LTD.')," & vbCrLf
            str &= "('026','MEGA INTERNATIONAL COMMERCIAL BANK PUBLIC COMPANY LIMITED')," & vbCrLf
            str &= "('027','BANK OF AMERICA, NATIONAL ASSOCIATION')," & vbCrLf
            str &= "('029','INDIAN OVERSEA BANK')," & vbCrLf
            str &= "('030','THE GOVERNMENT SAVINGS BANK')," & vbCrLf
            str &= "('031','THE HONGKONG AND SHANGHAI BANKING CORPORATION LTD.')," & vbCrLf
            str &= "('032','DEUTSCHE BANK AG.')," & vbCrLf
            str &= "('033','THE GOVERNMENT HOUSING BANK')," & vbCrLf
            str &= "('034','BANK FOR AGRICULTURE AND AGRICULTURAL COOPERATIVES')," & vbCrLf
            str &= "('035','EXPORT-IMPORT BANK OF THAILAND')," & vbCrLf
            str &= "('039','Mizuho Bank, Ltd. Bangkok Branch')," & vbCrLf
            str &= "('045','BNP PARIBAS')," & vbCrLf
            str &= "('052','BANK OF CHINA (THAI) PUBLIC COMPANY LIMITED')," & vbCrLf
            str &= "('065','THANACHART BANK PUBLIC COMPANY LTD.')," & vbCrLf
            str &= "('066','ISLAMIC BANK OF THAILAND')," & vbCrLf
            str &= "('067','TISCO BANK PUBLIC COMPANY LIMITED')," & vbCrLf
            str &= "('069','KIATNAKIN BANK PUBLIC COMPANY LIMITED')," & vbCrLf
            str &= "('070','INDUSTRIAL AND COMMERCIAL BANK OF CHINA (THAI) PUBLIC COMPANY LIMITED')," & vbCrLf
            str &= "('071','THE THAI CREDIT RETAIL BANK PUBLIC COMPANY LIMITED')," & vbCrLf
            str &= "('073','LAND AND HOUSES BANK PUBLIC COMPANY LIMITED')," & vbCrLf
            str &= "('079','ANZ BANK (THAI) PUBLIC COMPANY LIMITED')," & vbCrLf
            str &= "('080','SUMITOMO MITSUI TRUST BANK (THAI) PUBLIC COMPANY LIMITED')," & vbCrLf
            str &= "('098','SMALL AND MEDIUM ENTERPRISE DEVELOPMENT BANK OF THAILAND')"

            Dim cmd As SqlCommand = New SqlCommand(str, connection)
            cmd.ExecuteNonQuery()
        End If
        Return dtchk
    End Function

    Public Shared Function getBankname(ByVal bankcode As String) As String
        Dim dtbank As DataTable = New DataTable()
        Dim checkExist = checkExistView()
        Dim bankname As String = ""
        Dim STREXIST As String
        STREXIST = "SELECT BANKNAME FROM FMSEPAYBANKMASTER where BANKCODE = '" & bankcode.TrimEnd & "'"
        Dim cmdquery As SqlDataAdapter = New SqlDataAdapter(STREXIST, connection)
        cmdquery.Fill(dtbank)

        For i = 0 To dtbank.Rows.Count - 1
            bankname = dtbank.Rows(0).Item("BANKNAME").ToString
        Next
        Return bankname

    End Function

    Public Shared Sub CheckExistFMSEpayEX()
        Dim dtchk As DataTable = New DataTable()
        connection = New SqlConnection(conStr)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If
        Dim STREXIST As String
        STREXIST = "SELECT name FROM sys.tables WHERE name = 'FMSEPayEx'"
        Dim cmdquery As SqlDataAdapter = New SqlDataAdapter(STREXIST, connection)
        cmdquery.Fill(dtchk)
        If dtchk.Rows.Count = 0 Then
            Dim str As String
            str = "CREATE TABLE [dbo].[FMSEPayEx](" & vbCrLf
            str = "[ID] [bigint] IDENTITY(1,1) NOT NULL," & vbCrLf
            str = "[RUNNO] [nvarchar](30) NULL," & vbCrLf
            str = "[CNTBTCH] [nvarchar](30) NULL," & vbCrLf
            str = "[CNTENTY] [nvarchar](30) NULL," & vbCrLf
            str = "[EXPORTDATE] [nvarchar](30) NULL) " & vbCrLf
            Dim cmd As SqlCommand = New SqlCommand(str, connection)
            cmd.ExecuteNonQuery()
        End If

    End Sub


End Class
