Imports System.Data.OleDb
Public Class ManGoods
    ' Bagian saat load form
    Private Sub ManGoods_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Connects()
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

    ' Bagian button untuk meminimize aplikasi
    Private Sub BtnMinimize_Click(sender As Object, e As EventArgs) Handles BtnMinimize.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Bagian UserLogin
    ' Bagian button Login
    Private Sub BtnLogin_Click(sender As Object, e As EventArgs) Handles BtnLogin.Click
        If TxtUsername.Text = "" And TxtPassword.Text = "" Then
            MsgBox("Silakan Isi Username dan Password Anda!", vbExclamation, "Peringatan!")
            TxtUsername.Focus()
        ElseIf TxtUsername.Text = "" Then
            MsgBox("Silakan Isi Username Anda!", vbExclamation, "Peringatan")
            TxtUsername.Focus()
        ElseIf TxtPassword.Text = "" Then
            MsgBox("Silakan Isi Password Anda!", vbExclamation, "Peringatan")
        Else
            Commands = New OleDbCommand("SELECT * FROM pengguna WHERE username='" &
                                        TxtUsername.Text & "' AND password='" & TxtPassword.Text & "'", Connections)
            Readers = Commands.ExecuteReader
            If Readers.Read Then
                MsgBox("Login Berhasil!", vbInformation, "Pemberitahuan")
                LoginForm.Visible = False
                LabSUser.Text = TxtUsername.Text
                TxtUsername.Clear()
                TxtPassword.Clear()
            Else
                MsgBox("Username Atau Password Anda Salah!", vbExclamation, "Peringatan")
                TxtUsername.Focus()
            End If
        End If
    End Sub

    Private Sub BtnKeluar_Click(sender As Object, e As EventArgs) Handles BtnKeluar.Click
        If MsgBox("Yakin Tutup Aplikasi ini?", vbYesNo + vbInformation, "Tutup Aplikasi") = MsgBoxResult.Yes Then
            End
        Else
            TxtUsername.Focus()
        End If
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Sub Program Khusus
    Sub GridTableBarang(TableName, Tanggal)
        TableName.Columns(0).HeaderText = "Id"
        TableName.Columns(1).HeaderText = "Kode Barang"
        TableName.Columns(2).HeaderText = "Nama Barang"
        TableName.Columns(3).HeaderText = "Kuantitas"
        TableName.Columns(4).HeaderText = Tanggal
        TableName.Columns(5).HeaderText = "Tanggal Kadaluwarsa"
        TableName.Columns(0).Width = 50
        TableName.Columns(1).Width = 135
        TableName.Columns(2).Width = 300
        TableName.Columns(3).Width = 100
        TableName.Columns(4).Width = 120
        TableName.Columns(5).Width = 120
        TableName.Columns(4).DefaultCellStyle.Format = "dd/MM/yyyy"
        TableName.Columns(5).DefaultCellStyle.Format = "dd/MM/yyyy"
    End Sub
    Sub GridTablePersediaan(TableName)
        TableName.Columns(0).HeaderText = "Id"
        TableName.Columns(1).HeaderText = "Kode Barang"
        TableName.Columns(2).HeaderText = "Nama Barang"
        TableName.Columns(3).HeaderText = "Kuantitas"
        TableName.Columns(4).HeaderText = "Tanggal Kadaluwarsa"
        TableName.Columns(0).Width = 50
        TableName.Columns(1).Width = 135
        TableName.Columns(2).Width = 400
        TableName.Columns(3).Width = 100
        TableName.Columns(4).Width = 140
        TableName.Columns(4).DefaultCellStyle.Format = "dd/MM/yyyy"
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Bagian Barang Masuk
    ' Sub Barang Masuk
    Sub BarangMasukClear()
        CmbBarangMasuk.SelectedIndex = -1
        TxtMasukKuantitas.Clear()
        TxtMasukNamaBarang.Clear()
        DTPTanggalMasuk.Value = TanggalSekarang
        DTPMasukTanggalKadaluwarsa.Value = TanggalSekarang.AddDays(90)
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub BtnBarangMasuk_Click(sender As Object, e As EventArgs) Handles BtnBarangMasuk.Click
        FillComboBox(CmbBarangMasuk)
        BtnBarangMasuk.BackColor = Color.FromArgb(73, 190, 184)
        BtnBarangKeluar.BackColor = Color.FromArgb(43, 122, 120)
        BtnPersediaan.BackColor = Color.FromArgb(43, 122, 120)
        BtnDataBarang.BackColor = Color.FromArgb(43, 122, 120)
        BtnLaporan.BackColor = Color.FromArgb(43, 122, 120)
        BarangMasukClear()
        TxtMasukCari.Clear()
        ShowFillTable("barang_masuk", TabelMasuk)
        GridTableBarang(TabelMasuk, "Tanggal Masuk")
        BarangMasukPane.Visible = True
        BarangKeluarPane.Visible = False
        PersediaanPane.Visible = False
        DataBarangPane.Visible = False
        LaporanPane.Visible = False
    End Sub

    Private Sub CmbBarangMasuk_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbBarangMasuk.SelectedIndexChanged
        FillNamaBarang(CmbBarangMasuk, TxtMasukNamaBarang)
    End Sub

    Private Sub BtnMasukSimpan_Click(sender As Object, e As EventArgs) Handles BtnMasukSimpan.Click
        KodeBarang = CmbBarangMasuk.Text
        NamaBarang = TxtMasukNamaBarang.Text
        Tanggal = DTPTanggalMasuk.Value
        TanggalKadaluwarsa = DTPMasukTanggalKadaluwarsa.Value

        If CmbBarangMasuk.SelectedIndex = -1 And TxtMasukKuantitas.Text = "" Then
            MsgDataEmpty()
        ElseIf CmbBarangMasuk.SelectedIndex = -1 Then
            MsgChoiceKode()
        ElseIf TxtMasukKuantitas.Text = "" Then
            MsgFillKuantitas()
        ElseIf DateTime.Compare(Tanggal, TanggalSekarang) > 0 Then
            MsgDateCheck("Masuknya!")
            DTPTanggalMasuk.Value = TanggalSekarang
        ElseIf DateTime.Compare(Tanggal, TanggalSekarang.AddDays(-2)) < 0 Then
            MsgDateCheck("Masuknya!")
            DTPTanggalMasuk.Value = TanggalSekarang
        ElseIf DateTime.Compare(TanggalKadaluwarsa, TanggalSekarang.AddDays(90)) < 0 Then
            MsgDateCheck("Kadaluwarsanya!")
            DTPMasukTanggalKadaluwarsa.Value = TanggalSekarang.AddDays(90)
        Else
            If IsNumeric(TxtMasukKuantitas.Text) Then
                Kuantitas = CInt(TxtMasukKuantitas.Text)
                If Kuantitas <= 0 Then
                    MsgBox("Silakan Isi Kuantitasnya Minimal 1!", vbExclamation, "Perhatian")
                Else
                    Commands = New OleDbCommand("SELECT * FROM barang_masuk WHERE kode_barang='" &
                                                KodeBarang & "' AND tanggal_masuk=@tanggal_masuk AND tanggal_kadaluwarsa=" &
                                                "@tanggal_kadaluwarsa", Connections)
                    Commands.Parameters.Add("tanggal_masuk", OleDbType.Date).Value = Tanggal
                    Commands.Parameters.Add("tanggal_kadaluwarsa", OleDbType.Date).Value = TanggalKadaluwarsa
                    Readers = Commands.ExecuteReader
                    If Readers.Read Then
                        UpdateBarang("barang_masuk", "tanggal_masuk", Plus, KodeBarang, Kuantitas, Tanggal, TanggalKadaluwarsa)
                    Else
                        InsertBarang("barang_masuk", "tanggal_masuk", KodeBarang, NamaBarang, Kuantitas, Tanggal, TanggalKadaluwarsa)
                    End If
                    Readers.Close()
                    Commands = New OleDbCommand("SELECT * FROM persediaan_barang WHERE kode_barang='" &
                                                KodeBarang & "' AND tanggal_kadaluwarsa=" &
                                                "@tanggal_kadaluwarsa", Connections)
                    Commands.Parameters.Add("tanggal_kadaluwarsa", OleDbType.Date).Value = TanggalKadaluwarsa
                    Readers = Commands.ExecuteReader
                    If Readers.Read Then
                        UpdatePersediaan(Kuantitas, Plus, KodeBarang, TanggalKadaluwarsa)
                    Else
                        InsertPersediaan(KodeBarang, NamaBarang, Kuantitas, TanggalKadaluwarsa)
                    End If
                    Readers.Close()
                    MsgDataSaved()
                    ShowFillTable("barang_masuk", TabelMasuk)
                    BarangMasukClear()
                End If
            Else
                MsgBox("Silakan Isi Kuantitas Dengan Angka!", vbExclamation, "Perhatian")
            End If
        End If
    End Sub

    Private Sub BtnMasukHapus_Click(sender As Object, e As EventArgs) Handles BtnMasukHapus.Click
        KodeBarang = CmbBarangMasuk.Text
        Tanggal = DTPTanggalMasuk.Value
        TanggalKadaluwarsa = DTPMasukTanggalKadaluwarsa.Value

        If CmbBarangMasuk.SelectedIndex = -1 Then
            MsgChoiceKode()
        Else
            Commands = New OleDbCommand("SELECT * FROM barang_masuk WHERE kode_barang='" &
                                        KodeBarang & "' AND tanggal_masuk=@tanggal_masuk AND tanggal_kadaluwarsa=" &
                                        "@tanggal_kadaluwarsa", Connections)
            Commands.Parameters.Add("tanggal_masuk", OleDbType.Date).Value = Tanggal
            Commands.Parameters.Add("tanggal_kadaluwarsa", OleDbType.Date).Value = TanggalKadaluwarsa
            Readers = Commands.ExecuteReader
            If Readers.Read Then
                Kuantitas = CInt(Readers.GetValue(3))
                If MsgBox("Yakin Hapus Data Ini?", vbYesNo + vbInformation, "Hapus") = MsgBoxResult.Yes Then
                    Commands = New OleDbCommand("SELECT * FROM persediaan_barang WHERE kode_barang='" &
                                                KodeBarang & "' AND tanggal_kadaluwarsa=@tanggal_kadaluwarsa", Connections)
                    Commands.Parameters.Add("tanggal_kadaluwarsa", OleDbType.Date).Value = TanggalKadaluwarsa
                    Readerz = Commands.ExecuteReader
                    Readerz.Read()
                    KuantitasTemp = CInt(Readerz.GetValue(3))
                    Readerz.Close()
                    If Kuantitas > KuantitasTemp Then
                        MsgBox("Data Telah Dipindahkan Ke Barang Keluar", vbExclamation, "Perhatian")
                    ElseIf Kuantitas < KuantitasTemp Then
                        DeleteBarang("barang_masuk", "tanggal_masuk", KodeBarang, Tanggal, TanggalKadaluwarsa)
                        UpdatePersediaan(Kuantitas, Minus, KodeBarang, TanggalKadaluwarsa)
                        MsgDataDeleted()
                        BarangMasukClear()
                    Else
                        DeleteBarang("barang_masuk", "tanggal_masuk", KodeBarang, Tanggal, TanggalKadaluwarsa)
                        DeletePersediaan(KodeBarang, TanggalKadaluwarsa)
                        MsgDataDeleted()
                        BarangMasukClear()
                    End If
                    ShowFillTable("barang_masuk", TabelMasuk)
                End If
            Else
                MsgDataWantDelete()
            End If
            Readers.Close()
        End If
    End Sub
    Private Sub BtnMasukBatal_Click(sender As Object, e As EventArgs) Handles BtnMasukBatal.Click
        BarangMasukClear()
    End Sub

    Private Sub BtnMasukCari_Click(sender As Object, e As EventArgs) Handles BtnMasukCari.Click
        SearchFillTable("barang_masuk", TxtMasukCari.Text, TabelMasuk)
    End Sub

    Private Sub TabelMasuk_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles TabelMasuk.CellMouseClick
        TableGridBarangFill(e, TabelMasuk, CmbBarangMasuk, TxtMasukKuantitas, DTPTanggalMasuk, DTPMasukTanggalKadaluwarsa)
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Bagian Barang Keluar
    ' Sub Barang Keluar
    Sub BarangKeluarClear()
        CmbBarangKeluar.SelectedIndex = -1
        TxtKeluarNamaBarang.Clear()
        TxtKeluarKuantitas.Clear()
        DTPTanggalKeluar.Value = TanggalSekarang
        DTPKeluarTanggalKadaluwarsa.Value = TanggalSekarang
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub BtnBarangKeluar_Click(sender As Object, e As EventArgs) Handles BtnBarangKeluar.Click
        FillComboBox(CmbBarangKeluar)
        BtnBarangMasuk.BackColor = Color.FromArgb(43, 122, 120)
        BtnBarangKeluar.BackColor = Color.FromArgb(73, 190, 184)
        BtnPersediaan.BackColor = Color.FromArgb(43, 122, 120)
        BtnDataBarang.BackColor = Color.FromArgb(43, 122, 120)
        BtnLaporan.BackColor = Color.FromArgb(43, 122, 120)
        BarangKeluarClear()
        TxtKeluarCari.Clear()
        ShowFillTable("barang_keluar", TabelKeluar)
        GridTableBarang(TabelKeluar, "Tanggal Keluar")
        BarangMasukPane.Visible = False
        BarangKeluarPane.Visible = True
        PersediaanPane.Visible = False
        DataBarangPane.Visible = False
        LaporanPane.Visible = False
    End Sub

    Private Sub CmbBarangKeluar_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbBarangKeluar.SelectedIndexChanged
        FillNamaBarang(CmbBarangKeluar, TxtKeluarNamaBarang)
    End Sub

    Private Sub BtnKeluarSimpan_Click(sender As Object, e As EventArgs) Handles BtnKeluarSimpan.Click
        KodeBarang = CmbBarangKeluar.Text
        NamaBarang = TxtKeluarNamaBarang.Text
        Tanggal = DTPTanggalKeluar.Value
        TanggalKadaluwarsa = DTPKeluarTanggalKadaluwarsa.Value

        If CmbBarangKeluar.SelectedIndex = -1 And TxtKeluarKuantitas.Text = "" Then
            MsgDataEmpty()
        ElseIf CmbBarangKeluar.SelectedIndex = -1 Then
            MsgChoiceKode()
        ElseIf TxtKeluarKuantitas.Text = "" Then
            MsgFillKuantitas()
        ElseIf DateTime.Compare(Tanggal, TanggalSekarang) > 0 Then
            MsgDateCheck("Keluarnya!")
            DTPTanggalKeluar.Value = TanggalSekarang
        ElseIf DateTime.Compare(Tanggal, TanggalSekarang.AddDays(-2)) < 0 Then
            MsgDateCheck("Keluarnya!")
            DTPTanggalKeluar.Value = TanggalSekarang
        Else
            If IsNumeric(TxtKeluarKuantitas.Text) Then
                Kuantitas = CInt(TxtKeluarKuantitas.Text)
                If Kuantitas <= 0 Then
                    MsgBox("Silakan Isi Kuantitasnya Minimal 1!", vbExclamation, "Perhatian")
                Else
                    Commands = New OleDbCommand("SELECT * FROM persediaan_barang WHERE kode_barang='" &
                                                KodeBarang & "' AND tanggal_kadaluwarsa=" &
                                                "@tanggal_kadaluwarsa", Connections)
                    Commands.Parameters.Add("tanggal_kadaluwarsa", OleDbType.Date).Value = TanggalKadaluwarsa
                    Readers = Commands.ExecuteReader
                    If Readers.Read Then
                        KuantitasTemp = CInt(Readers.GetValue(3))
                        If Kuantitas < KuantitasTemp Then
                            InsertBarang("barang_keluar", "tanggal_keluar", KodeBarang, NamaBarang, Kuantitas, Tanggal, TanggalKadaluwarsa)
                            UpdatePersediaan(Kuantitas, Minus, KodeBarang, TanggalKadaluwarsa)
                            MsgDataSaved()
                            BarangKeluarClear()
                        ElseIf Kuantitas = KuantitasTemp Then
                            InsertBarang("barang_keluar", "tanggal_keluar", KodeBarang, NamaBarang, Kuantitas, Tanggal, TanggalKadaluwarsa)
                            DeletePersediaan(KodeBarang, TanggalKadaluwarsa)
                            MsgDataSaved()
                            BarangKeluarClear()
                        Else
                            MsgBox("Kuantitas Barang Tidak Memenuhi!", vbInformation, "Pemberitahuan")
                        End If
                        ShowFillTable("barang_keluar", TabelKeluar)
                    Else
                        MsgBox("Data Tidak Ada Di Persediaan Barang!", vbInformation, "Pemberitahuan")
                    End If
                    Readers.Close()
                End If
            Else
                MsgBox("Silakan Isi Kuantitas Dengan Angka!", vbExclamation, "Perhatian")
            End If
        End If
    End Sub

    Private Sub BtnKeluarHapus_Click(sender As Object, e As EventArgs) Handles BtnKeluarHapus.Click
        KodeBarang = CmbBarangKeluar.Text
        NamaBarang = TxtKeluarNamaBarang.Text
        Tanggal = DTPTanggalKeluar.Value
        TanggalKadaluwarsa = DTPKeluarTanggalKadaluwarsa.Value

        If CmbBarangKeluar.SelectedIndex = -1 Then
            MsgChoiceKode()
        Else
            Commands = New OleDbCommand("SELECT * FROM barang_keluar WHERE kode_barang='" &
                                        KodeBarang & "' AND tanggal_keluar=@tanggal_keluar AND tanggal_kadaluwarsa=" &
                                        "@tanggal_kadaluwarsa", Connections)
            Commands.Parameters.Add("tanggal_keluar", OleDbType.Date).Value = Tanggal
            Commands.Parameters.Add("tanggal_kadaluwarsa", OleDbType.Date).Value = TanggalKadaluwarsa
            Readers = Commands.ExecuteReader
            If Readers.Read Then
                Kuantitas = CInt(Readers.GetValue(3))
                If MsgBox("Yakin Hapus Data Ini?", vbYesNo + vbInformation, "Hapus") = MsgBoxResult.Yes Then
                    DeleteBarang("barang_keluar", "tanggal_keluar", KodeBarang, Tanggal, TanggalKadaluwarsa)
                    Commands = New OleDbCommand("SELECT * FROM persediaan_barang WHERE kode_barang='" &
                                                KodeBarang & "' AND tanggal_kadaluwarsa=@tanggal_kadaluwarsa", Connections)
                    Commands.Parameters.Add("tanggal_kadaluwarsa", OleDbType.Date).Value = TanggalKadaluwarsa
                    Readerz = Commands.ExecuteReader
                    If Readerz.Read Then
                        UpdatePersediaan(Kuantitas, Plus, KodeBarang, TanggalKadaluwarsa)
                    Else
                        InsertPersediaan(KodeBarang, NamaBarang, Kuantitas, TanggalKadaluwarsa)
                    End If
                    MsgDataDeleted()
                    ShowFillTable("barang_keluar", TabelKeluar)
                    BarangKeluarClear()
                End If
            Else
                MsgDataWantDelete()
            End If
            Readers.Close()
        End If
    End Sub

    Private Sub BtnKeluarBatal_Click(sender As Object, e As EventArgs) Handles BtnKeluarBatal.Click
        BarangKeluarClear()
    End Sub
    Sub TableGridBarangFill(e, TableGrid, CmBox, Kuantitas, Tanggal, TanggalKadaluwarsa)
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = TableGrid.Rows(e.RowIndex)
            Dim i As Integer = 0
            With row
                Do Until (CmBox.Text = .Cells(1).Value)
                    CmBox.SelectedIndex = i
                    i += 1
                Loop
                Kuantitas.Text = .Cells(3).Value
                Tanggal.Value = .Cells(4).Value
                TanggalKadaluwarsa.Value = .Cells(5).Value
            End With
        End If
    End Sub

    Private Sub BtnTampilPersediaan_Click(sender As Object, e As EventArgs) Handles BtnTampilPersediaan.Click
        PersediaanForm.ShowDialog()
    End Sub

    Private Sub BtnKeluarCari_Click(sender As Object, e As EventArgs) Handles BtnKeluarCari.Click
        SearchFillTable("barang_keluar", TxtKeluarCari.Text, TabelKeluar)
    End Sub

    Private Sub TabelKeluar_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles TabelKeluar.CellMouseClick
        TableGridBarangFill(e, TabelKeluar, CmbBarangKeluar, TxtKeluarKuantitas, DTPTanggalKeluar, DTPKeluarTanggalKadaluwarsa)
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Bagian Persediaan
    Private Sub BtnPersediaan_Click(sender As Object, e As EventArgs) Handles BtnPersediaan.Click
        TxtPersediaanCari.Clear()
        BtnBarangMasuk.BackColor = Color.FromArgb(43, 122, 120)
        BtnBarangKeluar.BackColor = Color.FromArgb(43, 122, 120)
        BtnPersediaan.BackColor = Color.FromArgb(73, 190, 184)
        BtnDataBarang.BackColor = Color.FromArgb(43, 122, 120)
        BtnLaporan.BackColor = Color.FromArgb(43, 122, 120)
        ShowFillTable("persediaan_barang", TabelPersediaan)
        GridTablePersediaan(TabelPersediaan)
        BarangMasukPane.Visible = False
        BarangKeluarPane.Visible = False
        PersediaanPane.Visible = True
        DataBarangPane.Visible = False
        LaporanPane.Visible = False
    End Sub

    Private Sub BtnPersediaanCari_Click(sender As Object, e As EventArgs) Handles BtnPersediaanCari.Click
        SearchFillTable("persediaan_barang", TxtPersediaanCari.Text, TabelPersediaan)
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Bagian Data Barang
    ' Sub Data Barang
    Sub DataBarangClear()
        TxtDataBarangKode.Clear()
        TxtDataBarangNama.Clear()
        TxtDataBarangKode.Focus()
    End Sub

    Sub MsgDataBarangKode()
        MsgBox("Silakan Isi Kode Barang Dulu!", vbExclamation, "Perhatian")
        TxtDataBarangKode.Focus()
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub BtnDataBarang_Click(sender As Object, e As EventArgs) Handles BtnDataBarang.Click
        DataBarangClear()
        TxtDataBarangCari.Clear()
        BtnBarangMasuk.BackColor = Color.FromArgb(43, 122, 120)
        BtnBarangKeluar.BackColor = Color.FromArgb(43, 122, 120)
        BtnPersediaan.BackColor = Color.FromArgb(43, 122, 120)
        BtnDataBarang.BackColor = Color.FromArgb(73, 190, 184)
        BtnLaporan.BackColor = Color.FromArgb(43, 122, 120)
        ShowFillTable("data_barang", TabelDataBarang)
        TabelDataBarang.Columns(0).HeaderText = "Kode Barang"
        TabelDataBarang.Columns(1).HeaderText = "Nama Barang"
        TabelDataBarang.Columns(0).Width = 165
        TabelDataBarang.Columns(1).Width = 660
        BarangMasukPane.Visible = False
        BarangKeluarPane.Visible = False
        PersediaanPane.Visible = False
        DataBarangPane.Visible = True
        LaporanPane.Visible = False
    End Sub

    Private Sub BtnDataBarangSimpan_Click(sender As Object, e As EventArgs) Handles BtnDataBarangSimpan.Click
        NamaBarang = TxtDataBarangNama.Text
        KodeBarang = TxtDataBarangKode.Text

        If TxtDataBarangKode.Text = "" And TxtDataBarangNama.Text = "" Then
            MsgDataEmpty()
            TxtDataBarangKode.Focus()
        ElseIf TxtDataBarangKode.Text = "" Then
            MsgDataBarangKode()
        ElseIf TxtDataBarangNama.Text = "" Then
            MsgBox("Silakan Isi Nama Barang Dulu!", vbExclamation, "Perhatian")
            TxtDataBarangNama.Focus()
        Else
            Commands = New OleDbCommand("SELECT * FROM data_barang WHERE kode_barang='" & KodeBarang &
                                        "'", Connections)
            Readers = Commands.ExecuteReader
            If Readers.Read Then
                MsgBox("Terjadi Duplikasi Data!", vbExclamation, "Peringatan")
            Else
                Commands = New OleDbCommand("INSERT INTO data_barang (kode_barang, nama_barang) VALUES " &
                                            "('" & KodeBarang & "', '" & NamaBarang & "')", Connections)
                Commands.ExecuteNonQuery()
                MsgDataSaved()
                DataBarangClear()
                ShowFillTable("data_barang", TabelDataBarang)
            End If
        End If
    End Sub

    Private Sub BtnDataBarangHapus_Click(sender As Object, e As EventArgs) Handles BtnDataBarangHapus.Click
        KodeBarang = TxtDataBarangKode.Text
        Commands = New OleDbCommand("SELECT * FROM barang_masuk WHERE kode_barang='" &
                                    KodeBarang & "'", Connections)
        Readers = Commands.ExecuteReader
        If TxtDataBarangKode.Text = "" Then
            MsgDataBarangKode()
        Else
            If Readers.Read Then
                MsgBox("Kode Barang Telah Digunakan!", vbExclamation, "Perhatian")
            ElseIf MsgBox("Yakin Hapus Data Ini?", vbYesNo + vbInformation, "Hapus") = MsgBoxResult.Yes Then
                Commands = New OleDbCommand("DELETE * FROM data_barang WHERE kode_barang='" &
                                            KodeBarang & "'", Connections)
                Commands.ExecuteNonQuery()
                DataBarangClear()
                MsgDataDeleted()
                ShowFillTable("data_barang", TabelDataBarang)
            Else
                TxtDataBarangKode.Focus()
            End If
        End If
        Readers.Close()
    End Sub

    Private Sub BtnDataBarangEdit_Click(sender As Object, e As EventArgs) Handles BtnDataBarangEdit.Click
        If TxtDataBarangKode.Text = "" And TxtDataBarangNama.Text = "" Then
            MsgDataEmpty()
            TxtDataBarangKode.Focus()
        ElseIf TxtDataBarangKode.Text = "" Then
            MsgDataBarangKode()
        Else
            KodeBarang = TxtDataBarangKode.Text
            NamaBarang = TxtDataBarangNama.Text
            UpdateNamaBarang("data_barang")
            UpdateNamaBarang("barang_masuk")
            UpdateNamaBarang("barang_keluar")
            UpdateNamaBarang("persediaan_barang")
            MsgBox("Data Berhasil Diubah!", vbInformation, "Pemberitahuan")
            DataBarangClear()
            ShowFillTable("data_barang", TabelDataBarang)
        End If
    End Sub

    Private Sub BtnDataBarangBatal_Click(sender As Object, e As EventArgs) Handles BtnDataBarangBatal.Click
        DataBarangClear()
    End Sub

    Private Sub BtnDataBarangCari_Click(sender As Object, e As EventArgs) Handles BtnDataBarangCari.Click
        SearchFillTable("data_barang", TxtDataBarangCari.Text, TabelDataBarang)
    End Sub

    Private Sub TabelDataBarang_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles TabelDataBarang.CellMouseClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = TabelDataBarang.Rows(e.RowIndex)
            With row
                TxtDataBarangKode.Text = .Cells(0).Value
                TxtDataBarangNama.Text = .Cells(1).Value
            End With
        End If
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Bagian Laporan
    Dim SetName, SetTable, SetDate As String

    Private Sub BtnLaporan_Click(sender As Object, e As EventArgs) Handles BtnLaporan.Click
        FillComboBox(CmbLaporanKode)
        CmbLaporanPilih.SelectedIndex = -1
        CheckBoxLaporan.Checked = False
        CheckBoxLaporan.Enabled = False
        DateStarts.Value = TanggalSekarang
        DateEnds.Value = TanggalSekarang
        CheckBoxTanggal.Text = "Filter Tanggal"
        TabelLaporan.Columns.Clear()
        BtnBarangMasuk.BackColor = Color.FromArgb(43, 122, 120)
        BtnBarangKeluar.BackColor = Color.FromArgb(43, 122, 120)
        BtnPersediaan.BackColor = Color.FromArgb(43, 122, 120)
        BtnDataBarang.BackColor = Color.FromArgb(43, 122, 120)
        BtnLaporan.BackColor = Color.FromArgb(73, 190, 184)
        BarangMasukPane.Visible = False
        BarangKeluarPane.Visible = False
        PersediaanPane.Visible = False
        DataBarangPane.Visible = False
        LaporanPane.Visible = True
    End Sub
    Private Sub CmbLaporanPilih_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbLaporanPilih.SelectedIndexChanged
        TabelLaporan.Columns.Clear()
        Select Case CmbLaporanPilih.Text
            Case "Barang Masuk"
                SetName = "SetMasuk"
                SetTable = "barang_masuk"
                SetDate = "tanggal_masuk"
                ShowFillTable("barang_masuk", TabelLaporan)
                GridTableBarang(TabelLaporan, "Tanggal Masuk")
                CheckBoxTanggal.Text = "Filter Tanggal Masuk"
            Case "Barang Keluar"
                SetName = "SetKeluar"
                SetTable = "barang_keluar"
                SetDate = "tanggal_keluar"
                ShowFillTable("barang_keluar", TabelLaporan)
                GridTableBarang(TabelLaporan, "Tanggal Keluar")
                CheckBoxTanggal.Text = "Filter Tanggal Keluar"
            Case "Persediaan"
                SetName = "SetPersediaan"
                SetTable = "persediaan_barang"
                SetDate = "tanggal_kadaluwarsa"
                ShowFillTable("persediaan_barang", TabelLaporan)
                CheckBoxTanggal.Text = "Filter Tanggal Kadaluwarsa"
                GridTablePersediaan(TabelLaporan)
        End Select
        CheckBoxLaporan.Enabled = True
    End Sub

    Private Sub CheckBoxLaporan_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxLaporan.CheckedChanged
        If CheckBoxLaporan.Checked Then
            CheckBoxKode.Enabled = True
            CheckBoxTanggal.Enabled = True
            BtnCetakFilter.Enabled = True
            BtnCheck.Enabled = True
            BtnCetak.Enabled = False
        Else
            CheckBoxKode.Enabled = False
            CheckBoxTanggal.Enabled = False
            CheckBoxKode.Checked = False
            CheckBoxTanggal.Checked = False
            BtnCetakFilter.Enabled = False
            BtnCheck.Enabled = False
            BtnCetak.Enabled = True
        End If
    End Sub

    Private Sub CheckBoxKode_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxKode.CheckedChanged
        If CheckBoxKode.Checked Then
            CmbLaporanKode.Enabled = True
        Else
            CmbLaporanKode.SelectedIndex = -1
            CmbLaporanKode.Enabled = False
        End If
    End Sub

    Private Sub CheckBoxTanggal_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxTanggal.CheckedChanged
        If CheckBoxTanggal.Checked Then
            DateStarts.Enabled = True
            DateEnds.Enabled = True
        Else
            DateStarts.Enabled = False
            DateEnds.Enabled = False
            DateStarts.Value = TanggalSekarang
            DateEnds.Value = TanggalSekarang
        End If
    End Sub
    Private Sub BtnCetak_Click(sender As Object, e As EventArgs) Handles BtnCetak.Click
        If CmbLaporanPilih.SelectedIndex = -1 Then
            MsgBox("Silakan Pilih Laporannya Dulu!", vbExclamation, "Perhatian")
        Else
            With ReportForm
                .LoadReport(SetName, SetTable)
                .ShowDialog()
            End With
        End If
    End Sub
    Private Sub BtnCheck_Click(sender As Object, e As EventArgs) Handles BtnCheck.Click
        If CheckBoxKode.Checked And CheckBoxTanggal.Checked Then
            If CmbLaporanKode.SelectedIndex = -1 Then
                MsgChoiceKode()
            Else
                FilterTableAll(SetTable, TabelLaporan, CmbLaporanKode.Text, SetDate, DateStarts.Value, DateEnds.Value)
            End If
        ElseIf CheckBoxKode.Checked Then
            If CmbLaporanKode.SelectedIndex = -1 Then
                MsgChoiceKode()
            Else
                FilterTableKode(SetTable, CmbLaporanKode.Text, TabelLaporan)
            End If
        ElseIf CheckBoxTanggal.Checked Then
            FilterTableTanggal(SetTable, TabelLaporan, SetDate, DateStarts.Value, DateEnds.Value)
        Else
            MsgFilteringFail()
        End If
    End Sub
    Private Sub BtnCetakFilter_Click(sender As Object, e As EventArgs) Handles BtnCetakFilter.Click
        If CheckBoxKode.Checked And CheckBoxTanggal.Checked Then
            If CmbLaporanKode.SelectedIndex = -1 Then
                MsgChoiceKode()
            Else
                With ReportForm
                    .LoadReportFilterAll(SetName, SetTable, CmbLaporanKode.Text, SetDate, DateStarts.Value, DateEnds.Value)
                    .ShowDialog()
                End With
            End If
        ElseIf CheckBoxKode.Checked Then
            If CmbLaporanKode.SelectedIndex = -1 Then
                MsgChoiceKode()
            Else
                With ReportForm
                    .LoadReportFilterKode(SetName, SetTable, CmbLaporanKode.Text)
                    .ShowDialog()
                End With
            End If
        ElseIf CheckBoxTanggal.Checked Then
            With ReportForm
                .LoadReportFilterTanggal(SetName, SetTable, SetDate, DateStarts.Value, DateEnds.Value)
                .ShowDialog()
            End With
        Else
            MsgFilteringFail()
        End If
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Bagian Logout
    Private Sub BtnLogout_Click(sender As Object, e As EventArgs) Handles BtnLogout.Click
        BtnLogout.BackColor = Color.FromArgb(73, 190, 184)
        If MsgBox("Logout Dari Aplikasi ini?", vbYesNo + vbInformation, "Logout") = MsgBoxResult.Yes Then
            BarangMasukPane.Visible = False
            BarangKeluarPane.Visible = False
            PersediaanPane.Visible = False
            DataBarangPane.Visible = False
            LaporanPane.Visible = False
            LoginForm.Visible = True
            BtnBarangMasuk.BackColor = Color.FromArgb(43, 122, 120)
            BtnBarangKeluar.BackColor = Color.FromArgb(43, 122, 120)
            BtnPersediaan.BackColor = Color.FromArgb(43, 122, 120)
            BtnDataBarang.BackColor = Color.FromArgb(43, 122, 120)
            BtnLaporan.BackColor = Color.FromArgb(43, 122, 120)
            BtnLogout.BackColor = Color.FromArgb(43, 122, 120)
        Else
            BtnLogout.BackColor = Color.FromArgb(43, 122, 120)
        End If
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Class