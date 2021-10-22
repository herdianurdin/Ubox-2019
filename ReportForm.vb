Imports System.Data.OleDb
Imports Microsoft.Reporting.WinForms

Public Class ReportForm
    ' Bagian Sub LoadReport
    Sub LoadReport(SetReport As String, TableName As String)
        Dim rptDS As ReportDataSource
        Dim cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" &
                                    Application.StartupPath & "\KelolaBarang.accdb")
        Dim ds As New KelolaBarangDataSet
        Dim da As New OleDbDataAdapter

        With ReportViews.LocalReport
            .ReportPath = Application.StartupPath & "\" & SetReport & ".rdlc"
            .DataSources.Clear()
        End With
        cn.Open()
        da.SelectCommand = New OleDbCommand("SELECT * FROM " & TableName, cn)
        da.Fill(ds.Tables(TableName))
        cn.Close()

        rptDS = New ReportDataSource(SetReport, ds.Tables(TableName))
        ReportViews.LocalReport.DataSources.Add(rptDS)
        ReportViews.SetDisplayMode(DisplayMode.PrintLayout)
        ReportViews.RefreshReport()
    End Sub
    Sub LoadReportFilterKode(SetReport As String, TableName As String, Kode As String)
        Dim rptDS As ReportDataSource
        Dim cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" &
                                    Application.StartupPath & "\KelolaBarang.accdb")
        Dim ds As New KelolaBarangDataSet
        Dim da As New OleDbDataAdapter

        With ReportViews.LocalReport
            .ReportPath = Application.StartupPath & "\" & SetReport & ".rdlc"
            .DataSources.Clear()
        End With
        cn.Open()
        da.SelectCommand = New OleDbCommand("SELECT * FROM " & TableName & " WHERE kode_barang='" & Kode & "'", cn)
        da.Fill(ds.Tables(TableName))
        cn.Close()

        rptDS = New ReportDataSource(SetReport, ds.Tables(TableName))
        ReportViews.LocalReport.DataSources.Add(rptDS)
        ReportViews.SetDisplayMode(DisplayMode.PrintLayout)
        ReportViews.RefreshReport()
    End Sub
    Sub LoadReportFilterTanggal(SetReport As String, TableName As String, Dates As String, DateStart As Date, DateEnd As Date)
        Dim rptDS As ReportDataSource
        Dim cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" &
                                    Application.StartupPath & "\KelolaBarang.accdb")
        Dim ds As New KelolaBarangDataSet
        Dim da As New OleDbDataAdapter

        With ReportViews.LocalReport
            .ReportPath = Application.StartupPath & "\" & SetReport & ".rdlc"
            .DataSources.Clear()
        End With
        cn.Open()
        da.SelectCommand = New OleDbCommand("SELECT * FROM " & TableName & " WHERE " & Dates &
                                            " BETWEEN #" & DateStart & "# AND #" & DateEnd & "#", cn)
        da.Fill(ds.Tables(TableName))
        cn.Close()

        rptDS = New ReportDataSource(SetReport, ds.Tables(TableName))
        ReportViews.LocalReport.DataSources.Add(rptDS)
        ReportViews.SetDisplayMode(DisplayMode.PrintLayout)
        ReportViews.RefreshReport()
    End Sub
    Sub LoadReportFilterAll(SetReport As String, TableName As String, Kode As String, Dates As String, DateStart As Date, DateEnd As Date)
        Dim rptDS As ReportDataSource
        Dim cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" &
                                    Application.StartupPath & "\KelolaBarang.accdb")
        Dim ds As New KelolaBarangDataSet
        Dim da As New OleDbDataAdapter

        With ReportViews.LocalReport
            .ReportPath = Application.StartupPath & "\" & SetReport & ".rdlc"
            .DataSources.Clear()
        End With
        cn.Open()
        da.SelectCommand = New OleDbCommand("SELECT * FROM " & TableName & " WHERE kode_barang='" & Kode & "' AND " &
                                            Dates & " BETWEEN #" & DateStart & "# AND #" & DateEnd & "#", cn)
        da.Fill(ds.Tables(TableName))
        cn.Close()

        rptDS = New ReportDataSource(SetReport, ds.Tables(TableName))
        ReportViews.LocalReport.DataSources.Add(rptDS)
        ReportViews.SetDisplayMode(DisplayMode.PrintLayout)
        ReportViews.RefreshReport()
    End Sub
    ' Bagian saat load form
    Private Sub ReportForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Bagian Header
    ' Bagian untuk menggerakan aplikasi dengan mouse
    Dim Drag As Boolean
    Dim MouseX, MouseY As Integer

    Private Sub Header_MouseDown(sender As Object, e As MouseEventArgs) Handles Header.MouseDown
        Drag = True
        MouseX = Cursor.Position.X - Me.Left
        MouseY = Cursor.Position.Y - Me.Top
    End Sub

    Private Sub Header_MouseMove(sender As Object, e As MouseEventArgs) Handles Header.MouseMove
        If Drag Then
            Me.Left = Cursor.Position.X - MouseX
            Me.Top = Cursor.Position.Y - MouseY
        End If
    End Sub

    Private Sub Header_MouseUp(sender As Object, e As MouseEventArgs) Handles Header.MouseUp
        Drag = False
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Bagian button untuk close aplikasi
    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Class