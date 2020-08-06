Imports System.Data
Imports System.Data.SqlClient

Public Class Form1
    Dim sys As New clsCSys
    Dim clsW As New clsCWriteLogFile
    Dim WithEvents tm As New Timer With {.Interval = 1000, .Enabled = True}

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ReadConfig()
            conStr = "Data Source= " & BaseValiabled.SQLServer & " ;Initial Catalog= " & BaseValiabled.SQLDatabase & _
            ";User id=" & BaseValiabled.SQLUserID & " ;password=" & BaseValiabled.SQLPsw & ""
            Dim form As Form = clsVLogin
            'clsVLogin.ShowDialog()
            connection = New SqlConnection(conStr)
            If connection.State = ConnectionState.Closed Then
                connection.Open()
                If connection.State = ConnectionState.Open Then
                End If
            End If
          
        Catch ex As Exception
            clsW.WriteLogError("<Database>/Invalid to Server : " & ex.Message.ToString())
        End Try
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        sys.ShowChildFormmodule(clsVConfiguration)
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        sys.ShowChildFormmodule(test2)
    End Sub

    Private Sub ReadConfig()
        BaseValiabled.SQLServer = BaseClass.Encription.AES_Decrypt(Config.GetPrivateProfileString("CONFIG", "DBSERVER", ""), "ABCDEF")
        BaseValiabled.SQLDatabase = BaseClass.Encription.AES_Decrypt(Config.GetPrivateProfileString("CONFIG", "DBNAME", ""), "ABCDEF")
        BaseValiabled.SQLUserID = BaseClass.Encription.AES_Decrypt(Config.GetPrivateProfileString("CONFIG", "DBUSERID", ""), "ABCDEF")
        BaseValiabled.SQLPsw = BaseClass.Encription.AES_Decrypt(Config.GetPrivateProfileString("CONFIG", "DBPASSWORD", ""), "ABCDEF")
        BaseValiabled.FileEx = Config.GetPrivateProfileString("CONFIG", "FILEEX", "")
        BaseValiabled.FileErr = Config.GetPrivateProfileString("CONFIG", "FILEERR", "")
        BaseValiabled.Bankcode = Config.GetPrivateProfileString("CONFIG", "BANKCODE", "")
        If BaseValiabled.FileEx = "" OrElse BaseValiabled.FileErr = "" Then
            Dim frm As New clsVConfiguration()
            frm.ShowDialog()
        End If
    End Sub

    Private Sub ExportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportToolStripMenuItem.Click
        sys.ShowChildFormmodule(clsVExport)
    End Sub

    Private Sub Tm_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tm.Tick
        tssTime.Text = Now.ToString("HH:mm:ss")
    End Sub
 
    Private Sub ToolStripDropDownButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton1.Click

    End Sub


End Class
