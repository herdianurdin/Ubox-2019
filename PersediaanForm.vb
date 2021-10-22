Public Class PersediaanForm
    ' Bagian saat load form
    Private Sub PersediaanForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Connects()
        ShowFillTable("persediaan_barang", TabelPersediaan)
        TxtCari.Clear()
        TabelPersediaan.Columns(0).HeaderText = "Id"
        TabelPersediaan.Columns(1).HeaderText = "Kode Barang"
        TabelPersediaan.Columns(2).HeaderText = "Nama Barang"
        TabelPersediaan.Columns(3).HeaderText = "Kuantitas"
        TabelPersediaan.Columns(4).HeaderText = "Tanggal Kadaluwarsa"
        TabelPersediaan.Columns(0).Width = 50
        TabelPersediaan.Columns(1).Width = 135
        TabelPersediaan.Columns(2).Width = 400
        TabelPersediaan.Columns(3).Width = 100
        TabelPersediaan.Columns(4).Width = 140
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

    Private Sub TabelPersediaan_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles TabelPersediaan.CellMouseClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = TabelPersediaan.Rows(e.RowIndex)
            Dim i As Integer = 0
            With row
                Do Until (ManGoods.CmbBarangKeluar.Text = .Cells(1).Value)
                    ManGoods.CmbBarangKeluar.SelectedIndex = i
                    i += 1
                Loop
                ManGoods.TxtKeluarKuantitas.Text = .Cells(3).Value
                ManGoods.DTPKeluarTanggalKadaluwarsa.Value = .Cells(4).Value
            End With
        End If
    End Sub

    Private Sub BtnCari_Click(sender As Object, e As EventArgs) Handles BtnCari.Click
        SearchFillTable("persediaan_barang", TxtCari.Text, TabelPersediaan)
    End Sub

End Class