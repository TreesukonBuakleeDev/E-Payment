Imports System.Data
Imports System.Data.SqlClient

Public Class clsCGetData
    Public Function Getdata(Optional ByVal BatchFrom As String = "", Optional ByVal BatchTo As String = "", Optional ByVal ReExport As Boolean = False) As DataTable
        Dim Condition As String = ""
        If BatchFrom <> "" AndAlso BatchTo <> "" Then
            Condition &= "AND B.[BATCHID] BETWEEN " & BatchFrom & " AND " & BatchTo
        End If

        If Not ReExport Then
            Condition &= "   AND B.[BATCHID] NOT IN(SELECT CNTBTCH FROM FMSEPayEx)"
        End If

        connection = New SqlConnection(conStr)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If

        sql1 = "Select DISTINCT 'False' AS 'Check' " & vbCrLf
        sql1 &= "   ,B.[BATCHID] AS BatchNo " & vbCrLf
        sql1 &= "   ,'' AS BatchDate " & vbCrLf
        sql1 &= "   ,CBBTHD.TEXTDESC AS BatchDesc " & vbCrLf
        sql1 &= "   ,CBBTHD.ENTRYNO As 'Entry' " & vbCrLf
        sql1 &= "   ,CBBTHD.MISCCODE AS VendorCode " & vbCrLf
        sql1 &= "   ,APVNR.RMITNAME AS VendorName " & vbCrLf
        sql1 &= "   ,CBBTHD.APPLAMOUNT AS Amount " & vbCrLf
        sql1 &= "   ,B.BANKCODE AS BankCode " & vbCrLf
        sql1 &= "   ,'' AS Currency " & vbCrLf
        sql1 &= "   ,CASE B.[STATUS] " & vbCrLf
        sql1 &= "       WHEN 0 THEN 'Open' " & vbCrLf
        sql1 &= "       WHEN 1 THEN 'Ready to Post' " & vbCrLf
        sql1 &= "       WHEN 2 THEN 'Deleted' " & vbCrLf
        sql1 &= "       WHEN 3 THEN 'Posted' " & vbCrLf
        sql1 &= "   END  AS BatchStatus " & vbCrLf
        sql1 &= "   ,CASE B.BATCHTYPE " & vbCrLf
        sql1 &= "       WHEN 0 THEN 'Normal' " & vbCrLf
        sql1 &= "       WHEN 1 THEN 'Transfer' " & vbCrLf
        sql1 &= "   END AS BatchType " & vbCrLf
        sql1 &= "   ,'' AS DateRate " & vbCrLf
        sql1 &= "   ,CASE " & vbCrLf
        sql1 &= "       WHEN E.CNTBTCH IS NULL THEN 'Not Export' " & vbCrLf
        sql1 &= "       WHEN E.CNTBTCH IS NOT NULL THEN 'Exported' " & vbCrLf
        sql1 &= "   END AS ExportStatus" & vbCrLf
        sql1 &= "FROM CBBCTL B " & vbCrLf
        sql1 &= "   LEFT JOIN CBBANK BANK ON B.BANKCODE = BANK.BANKCODE " & vbCrLf
        sql1 &= "   LEFT OUTER JOIN CBBTHD ON CBBTHD.BATCHID = B.BATCHID " & vbCrLf
        sql1 &= "   LEFT OUTER JOIN APVNR ON CBBTHD.MISCCODE = APVNR.IDVEND " & vbCrLf
        sql1 &= "   LEFT JOIN FMSEPayEx E ON B.BATCHID = E.CNTBTCH " & vbCrLf
        sql1 &= "WHERE 1 = 1 " & vbCrLf
        sql1 &= "   AND B.BATCHID IN (SELECT BATCHID FROM CBBTHD WHERE TOTAMOUNT < 0)" & vbCrLf
        sql1 &= "   AND B.BATCHTYPE = 0 " & vbCrLf
        'sql1 &= "   AND B.BANKCODE IN ('BTMU','SASMBC') " & vbCrLf 'Fix ไว้ 
        sql1 &= "   AND B.BANKCODE IN ('" & BaseValiabled.Bankcode & "') " & vbCrLf 'Fix ไว้ 
        sql1 &= "   AND B.[STATUS] = 0" & Condition

        command = New SqlCommand(sql1, connection)
        adapter = New SqlDataAdapter(command)
        dataSt = New DataSet()
        adapter.Fill(dataSt, "DataExport")

        connection.Close()

        Return dataSt.Tables("DataExport")
    End Function
End Class
