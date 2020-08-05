Imports System.Data
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine

Public Class clsVExport
    Dim ClsVali As New clsCValidate
    Dim ClsGet As New clsCGetData 'clsCGetData
    Dim clsGen As New clsCGenerateFile 'ClsCGenerateFile
    Dim clsLog As New clsCWriteLogFile
    Dim clsW As New clsCWriteLogFile


#Region "Resize"
    Dim CW As Integer = Me.Width ' Current Width
    Dim CH As Integer = Me.Height ' Current Height
    Dim IW As Integer = Me.Width ' Initial Width
    Dim IH As Integer = Me.Height ' Initial Height

    Private Sub FormExport_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        Dim RW As Double = (Me.Width - CW) / CW ' Ratio change of width
        Dim RH As Double = (Me.Height - CH) / CH ' Ratio change of height

        For Each Ctrl As Control In Controls
            Ctrl.Width += CInt(Ctrl.Width * RW)
            Ctrl.Height += CInt(Ctrl.Height * RH)
            Ctrl.Left += CInt(Ctrl.Left * RW)
            Ctrl.Top += CInt(Ctrl.Top * RH)
        Next

        CW = Me.Width
        CH = Me.Height

    End Sub
#End Region

    Enum TypeReportReconcile
        Summary = 1
        Detail = 2
    End Enum

#Region "Load Form"

    Private Sub clsVExport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadControl()
        Dim PathEx As String = Config.GetPrivateProfileString("CONFIG", "FILEEx", "")
        BaseValiabled.FileEx = Config.GetPrivateProfileString("CONFIG", "FILEEx", "")
        Dim PathErr As String = Config.GetPrivateProfileString("CONFIG", "FILEERR", "")
        BaseValiabled.FileErr = Config.GetPrivateProfileString("CONFIG", "FILEERR", "")
    End Sub

    Private Sub LoadControl()
        btnExport.Enabled = False
        btnValidate.Enabled = False
        btnPrintReconcile.Enabled = False

        Dim dtgTextbox As New DataGridViewTextBoxColumn
        Dim dtgCheckbox As New DataGridViewCheckBoxColumn

        dtgCheckbox = New DataGridViewCheckBoxColumn
        dtgCheckbox.DataPropertyName = "Check"
        dtgCheckbox.HeaderText = "Check"
        dtgTextbox.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dtgCheckbox.Name = "Check"
        dtgCheckbox.Width = 50
        DataGridView1.Columns.Add(dtgCheckbox)

        dtgTextbox = New DataGridViewTextBoxColumn
        dtgTextbox.DataPropertyName = "BatchNo"
        dtgTextbox.HeaderText = "BatchNo"
        dtgTextbox.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dtgTextbox.Name = "BatchNo"
        dtgTextbox.ReadOnly = False
        dtgTextbox.Width = 60
        DataGridView1.Columns.Add(dtgTextbox)

        dtgTextbox = New DataGridViewTextBoxColumn
        dtgTextbox.DataPropertyName = "Entry"
        dtgTextbox.HeaderText = "Entry"
        dtgTextbox.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dtgTextbox.Name = "Entry"
        dtgTextbox.ReadOnly = False
        dtgTextbox.Width = 60
        DataGridView1.Columns.Add(dtgTextbox)

        dtgTextbox = New DataGridViewTextBoxColumn
        dtgTextbox.DataPropertyName = "BatchDate"
        dtgTextbox.HeaderText = "BatchDate"
        dtgTextbox.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dtgTextbox.Name = "BatchDate"
        dtgTextbox.ReadOnly = False
        dtgTextbox.Visible = False
        dtgTextbox.Width = 80
        DataGridView1.Columns.Add(dtgTextbox)

        dtgTextbox = New DataGridViewTextBoxColumn
        dtgTextbox.DataPropertyName = "BatchDesc"
        dtgTextbox.HeaderText = "BatchDesc"
        dtgTextbox.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dtgTextbox.Name = "BatchDesc"
        dtgTextbox.ReadOnly = False
        dtgTextbox.Width = 150
        DataGridView1.Columns.Add(dtgTextbox)

        dtgTextbox = New DataGridViewTextBoxColumn
        dtgTextbox.DataPropertyName = "VendorCode"
        dtgTextbox.HeaderText = "VendorCode"
        dtgTextbox.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dtgTextbox.Name = "VendorCode"
        dtgTextbox.ReadOnly = False
        dtgTextbox.Width = 95
        dtgTextbox.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns.Add(dtgTextbox)

        dtgTextbox = New DataGridViewTextBoxColumn
        dtgTextbox.DataPropertyName = "VendorName"
        dtgTextbox.HeaderText = "VendorName"
        dtgTextbox.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dtgTextbox.Name = "VendorName"
        dtgTextbox.ReadOnly = False
        dtgTextbox.Width = 150
        dtgTextbox.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns.Add(dtgTextbox)

        dtgTextbox = New DataGridViewTextBoxColumn
        dtgTextbox.DataPropertyName = "Amount"
        dtgTextbox.HeaderText = "Amount"
        dtgTextbox.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dtgTextbox.Name = "Amount"
        dtgTextbox.ReadOnly = False
        dtgTextbox.Width = 95
        dtgTextbox.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns.Add(dtgTextbox)

        dtgTextbox = New DataGridViewTextBoxColumn
        dtgTextbox.DataPropertyName = "BankCode"
        dtgTextbox.HeaderText = "BankCode"
        dtgTextbox.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dtgTextbox.Name = "BankCode"
        dtgTextbox.ReadOnly = False
        dtgTextbox.Width = 80
        DataGridView1.Columns.Add(dtgTextbox)

        dtgTextbox = New DataGridViewTextBoxColumn
        dtgTextbox.DataPropertyName = "Currency"
        dtgTextbox.HeaderText = "Currency"
        dtgTextbox.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dtgTextbox.Name = "Currency"
        dtgTextbox.ReadOnly = False
        dtgTextbox.Visible = False
        dtgTextbox.Width = 80
        DataGridView1.Columns.Add(dtgTextbox)

        dtgTextbox = New DataGridViewTextBoxColumn
        dtgTextbox.DataPropertyName = "BatchStatus"
        dtgTextbox.HeaderText = "BatchStatus"
        dtgTextbox.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dtgTextbox.Name = "BatchStatus"
        dtgTextbox.ReadOnly = False
        dtgTextbox.Width = 80
        DataGridView1.Columns.Add(dtgTextbox)

        dtgTextbox = New DataGridViewTextBoxColumn
        dtgTextbox.DataPropertyName = "BatchType"
        dtgTextbox.HeaderText = "BatchType"
        dtgTextbox.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dtgTextbox.Name = "BatchType"
        dtgTextbox.ReadOnly = False
        dtgTextbox.Width = 80
        DataGridView1.Columns.Add(dtgTextbox)

        dtgTextbox = New DataGridViewTextBoxColumn
        dtgTextbox.DataPropertyName = "DateRate"
        dtgTextbox.HeaderText = "DateRate"
        dtgTextbox.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dtgTextbox.Name = "DateRate"
        dtgTextbox.ReadOnly = False
        dtgTextbox.Visible = False
        dtgTextbox.Width = 80
        DataGridView1.Columns.Add(dtgTextbox)

        dtgTextbox = New DataGridViewTextBoxColumn
        dtgTextbox.DataPropertyName = "ExportStatus"
        dtgTextbox.HeaderText = "ExportStatus"
        dtgTextbox.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dtgTextbox.Name = "ExportStatus"
        dtgTextbox.ReadOnly = 80
        DataGridView1.Columns.Add(dtgTextbox)

    End Sub

#End Region

#Region "Event Form"

    Private Sub btnGetList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetList.Click
        DataGridView1.DataSource = ClsGet.Getdata(txtBatchFrom.Text.Trim, txtBatchTo.Text.Trim, CheckBox1.Checked)
    End Sub

    Private Sub btnValidate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnValidate.Click
        If Not CountCheckBox() Then
            Exit Sub
        End If
        For i = 0 To DataGridView1.RowCount - 1
            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
        Next
        '----------------------------------------------

        Dim dtDataInvalid As New DataTable
        Dim status As Boolean = True
        status = ClsVali.ValidateDataWithReport(DataGridView1, dtDataInvalid) 'ตรวจสอบ data พร้อมออก Report 
 
        If status Then status = ClsVali.ValidateReferenceWithReport(DataGridView1)

    End Sub

    Private Sub btnPrintReconcile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintReconcile.Click
        If Not CountCheckBox() Then
            Exit Sub
        End If

        Dim frm As New clsVCrystalReportPreview
        Dim resultdialog As DialogResult = MessageBox.Show("ท่านต้องการรายงานแบบ Summary ใช่หรือไม่", "ชนิดรายงาน", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
        If resultdialog = Windows.Forms.DialogResult.Yes Then
            frm = New clsVCrystalReportPreview(DataGridView1, TypeReportReconcile.Summary)
        Else
            frm = New clsVCrystalReportPreview(DataGridView1, TypeReportReconcile.Detail)
        End If
        frm.ShowDialog()
    End Sub

    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        If Not CountCheckBox() Then
            Exit Sub
        End If
        For i = 0 To DataGridView1.RowCount - 1
            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
        Next
        Dim Status As Boolean
        Dim typeExport As String = 6
        If ClsVali.ValidateData(DataGridView1, ProgressBar1) Then
            If ClsVali.ValidateReference(DataGridView1) Then 'Coloumn Reference ในหน้า Adjust"
                If DataGridView1.RowCount > 1 Then
                    typeExport = 6
                End If
                Status = clsGen.WriteTextFile(DataGridView1, typeExport, ProgressBar1)
                If Status Then
                    ProgressBar1.Value = 0
                    MessageBox.Show("Complete")
                    clsLog.SaveExported(DataGridView1) 'Save รายการที่ Export แล้ว
                    Dim resultdialog As DialogResult = MessageBox.Show("ท่านต้องการ Print Reconcile ใช่หรือไม่", _
                                                                       "Report Reconcile", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                    If resultdialog = Windows.Forms.DialogResult.Yes Then
                        Dim frm As New clsVCrystalReportPreview(DataGridView1, TypeReportReconcile.Summary)
                        frm.ShowDialog()
                    End If
                    DataGridView1.DataSource.Rows.Clear()
                Else
                    MessageBox.Show("Data Incomplete")
                    MessageBox.Show("WriteTextFile")
                    clsW.WriteLogError("WriteTextFile")
                End If
            Else
                MessageBox.Show("Data Incomplete")
                MessageBox.Show("Validate Reference")
                clsW.WriteLogError("Validate Reference")
            End If
        Else
            MessageBox.Show("Data Incomplete")
            MessageBox.Show("Validate Data")
            clsW.WriteLogError("Validate Data")
        End If
    End Sub

#End Region

    Private Sub txtBatchTo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBatchTo.KeyPress
        If txtBatchFrom.Text <> "" And txtBatchTo.Text <> "" Then
            If Asc(e.KeyChar) = (13) Then
                btnGetList.PerformClick()
            End If
        End If
    End Sub

    Private Sub DataGridView1_DataSourceChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DataSourceChanged
        If DataGridView1.Rows.Count > 0 Then
            btnValidate.Enabled = True
            btnExport.Enabled = True
            btnPrintReconcile.Enabled = True
        End If
    End Sub
    Private Function CountCheckBox() As Boolean ' ฟังชั่นสำหรับ ตรวจสอบว่ามีการเลือกรายการหรือยัง
        Dim CountSelect As Integer = 0
        For i = 0 To DataGridView1.RowCount - 1
            If IsDBNull(DataGridView1.Rows(i).Cells("Check").Value) Then
                Continue For
            End If
            If DataGridView1.Rows(i).Cells("Check").Value = True Then
                CountSelect += 1
            End If
        Next
        If CountSelect < 1 Then
            MessageBox.Show("Please Select Transaction", "Transaction", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        Else
            Return True
        End If
    End Function
    Public Shared Function CountSelect() As Integer ' count number of checkedbox
        CountSelect = 0
        For i = 0 To clsVExport.DataGridView1.Rows.Count - 1
            If IsDBNull(clsVExport.DataGridView1.Rows(i).Cells("Check").Value) Then
                Continue For
            End If
            If clsVExport.DataGridView1.Rows(i).Cells("Check").Value = True Then
                CountSelect += 1
            End If
        Next
        If CountSelect < 1 Then
            MessageBox.Show("Please Select Transaction", "Transaction", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        Else
            Return CountSelect
        End If
    End Function

    Private Sub Btn_GenXML_Click(sender As Object, e As EventArgs) Handles Btn_GenXML.Click
        GenValidate.cleardt()
        ValidateRmitto(DataGridView1)
        If dtMC.Rows.Count <> 0 Then
            GENXML(dtMC)
        End If

        If dtFT.Rows.Count <> 0 Then
            GENXMLFT(dtFT)
        End If

    End Sub

    Private Sub Btn_GenValidate_Click(sender As Object, e As EventArgs) Handles Btn_GenValidate.Click
        ValidateRmitto(DataGridView1)
    End Sub

End Class