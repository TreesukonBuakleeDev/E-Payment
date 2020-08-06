Imports System.Data.SqlClient
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Public Class GenReportValidate
    Structure Parameter
        Dim ParameterName As String
        Dim ParameterValue As String
    End Structure
    Structure ReportCondition
        Dim ConnectionString As String
        Dim ReportFile As String
        Dim Fomula As String
    End Structure
    Public ReportAttribute As ReportCondition
    Public ReportAttributeParameter() As Parameter

    Sub ShowReport(ByRef dtchk As DataTable)
        Try
            Dim dtchkErr As String = ""
            Dim str As String
            str = "SELECT * from FMSEPAYTEMP"
            command = New SqlCommand(str, connection)
            adapter = New SqlDataAdapter(command)
            dataSt = New DataSet()
            adapter.Fill(dataSt, "dtchkErr")
            connection.Close()

            If dataSt.Tables("dtchkErr").Rows.Count = 0 Then
                Dim resultdialog As DialogResult = MessageBox.Show("ท่านต้องการ Print Reconcile ใช่หรือไม่",
                                                                              "Report Reconcile", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                If resultdialog = Windows.Forms.DialogResult.Yes Then
                    Dim frm As New clsVCrystalReportPreview(clsVExport.DataGridView1, clsVExport.TypeReportReconcile.Summary)
                    frm.ShowDialog()
                End If
                'clsVExport.DataGridView1.DataSource.Rows.Clear()
            Else
                GenReportValidate_Load()
            End If
        Catch ex As Exception
            MessageBox.Show("Error 40: " & ex.Message)
        End Try

    End Sub

    Sub GenReportValidate_Load()
        Dim strFormula As String = ""
        Dim clsRptViewer As New GenReportValidate
        Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        connection = New SqlConnection(conStr)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If

        Dim conStrREPORT As New SqlClient.SqlConnectionStringBuilder(conStr)
        rpt.Load(My.Application.Info.DirectoryPath & "\Reports\GenPrintValidate.rpt")

        clsRptViewer.ReportAttribute.Fomula = strFormula

        'ReDim clsRptViewer.ReportAttributeParameter(2)

        'clsRptViewer.ReportAttributeParameter(1).ParameterName = "batchno"
        'clsRptViewer.ReportAttributeParameter(1).ParameterValue = batchno

        'clsRptViewer.ReportAttributeParameter(2).ParameterName = "entryno"
        'clsRptViewer.ReportAttributeParameter(2).ParameterValue = entryno

        Dim crTable As CrystalDecisions.CrystalReports.Engine.Table
        Dim crTableLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
        Dim ConnInfo As New CrystalDecisions.Shared.ConnectionInfo

        ConnInfo.ServerName = conStrREPORT.DataSource
        ConnInfo.UserID = conStrREPORT.UserID
        ConnInfo.Password = conStrREPORT.Password
        ConnInfo.DatabaseName = conStrREPORT.InitialCatalog
        ConnInfo.IntegratedSecurity = False

        Try
            For Each crTable In rpt.Database.Tables
                crTableLogonInfo = crTable.LogOnInfo
                crTableLogonInfo.ConnectionInfo = ConnInfo
                crTable.ApplyLogOnInfo(crTableLogonInfo)
            Next
            'If (clsRptViewer.ReportAttributeParameter IsNot Nothing) Then
            '    For Each obj As Parameter In clsRptViewer.ReportAttributeParameter
            '        If (obj.ParameterValue <> String.Empty) OrElse (obj.ParameterName <> String.Empty) Then
            '            rpt.SetParameterValue(obj.ParameterName, obj.ParameterValue)
            '        End If
            '    Next
            'End If

            ''rpt.RecordSelectionFormula = clsRptViewer.ReportAttribute.Fomula

            'Open Crystal report viewer' 
            Using objForm As New Windows.Forms.Form
                objForm.StartPosition = FormStartPosition.CenterScreen
                'objForm.Text = Utilities.Info.Title
                objForm.WindowState = FormWindowState.Maximized

                Using rptViewer As New CrystalDecisions.Windows.Forms.CrystalReportViewer

                    'rptViewer.DisplayGroupTree = False
                    rptViewer.ShowCloseButton = False
                    rptViewer.ShowGroupTreeButton = False
                    rptViewer.ShowTextSearchButton = False
                    rptViewer.ShowZoomButton = False

                    rptViewer.Dock = DockStyle.Fill

                    objForm.Controls.Add(rptViewer)

                    rptViewer.ReportSource = rpt

                    objForm.ShowDialog()

                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Error 116: " & ex.Message)
        End Try

    End Sub


End Class