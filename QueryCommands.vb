Imports System.Data.OleDb
Module QueryCommands
    Public NamaBarang, KodeBarang As String
    Public Kuantitas, KuantitasTemp As Integer
    Public Tanggal, TanggalKadaluwarsa As Date
    Public TanggalSekarang As Date = DateTime.Now.Date
    Public Plus As String = "+"
    Public Minus As String = "-"

    Public Sub ShowFillTable(TableName, TableGrid)
        DataAdapters = New OleDbDataAdapter("SELECT * FROM " & TableName, Connections)
        DataSets = New DataSet
        DataSets.Clear()
        DataAdapters.Fill(DataSets, TableName)
        TableGrid.DataSource = DataSets.Tables(TableName)
        TableGrid.Refresh()
    End Sub

    Public Sub SearchFillTable(TableName, SomeText, TableGrid)
        Commands = New OleDbCommand("SELECT * FROM " & TableName & " WHERE nama_barang LIKE '%" &
                                    SomeText & "%'", Connections)
        Readers = Commands.ExecuteReader
        If Readers.Read Then
            MsgBox("Data Barang Ditemukan!", vbInformation, "Pemberitahuan")
            DataAdapters = New OleDbDataAdapter("SELECT * FROM " & TableName & " WHERE nama_barang LIKE '%" &
                                                SomeText & "%'", Connections)
            DataSets = New DataSet
            DataAdapters.Fill(DataSets, "GOT")
            TableGrid.DataSource = DataSets.Tables("GOT")
        Else
            MsgBox("Data Barang Tidak Ditemukan!", vbInformation, "Pemberitahuan")
        End If
    End Sub
    Public Sub FilterTableKode(TableName, SomeText, TableGrid)
        Commands = New OleDbCommand("SELECT * FROM " & TableName & " WHERE kode_barang='" &
                            SomeText & "'", Connections)
        Readers = Commands.ExecuteReader
        If Readers.Read Then
            DataAdapters = New OleDbDataAdapter("SELECT * FROM " & TableName & " WHERE kode_barang='" &
                                                SomeText & "'", Connections)
            DataSets = New DataSet
            DataAdapters.Fill(DataSets, "GOT")
            TableGrid.DataSource = DataSets.Tables("GOT")
        Else
            MsgDataNotFound()
        End If
    End Sub
    Public Sub FilterTableTanggal(TableName, TableGrid, Dates, DateStart, DateEnd)
        Commands = New OleDbCommand("SELECT * FROM " & TableName & " WHERE " & Dates & " BETWEEN #" &
                                    DateStart & "# AND #" & DateEnd & "#", Connections)
        Readers = Commands.ExecuteReader
        If Readers.Read Then
            DataAdapters = New OleDbDataAdapter("SELECT * FROM " & TableName & " WHERE " & Dates & " BETWEEN #" &
                                                DateStart & "# AND #" & DateEnd & "#", Connections)
            DataSets = New DataSet
            DataAdapters.Fill(DataSets, "GOT")
            TableGrid.DataSource = DataSets.Tables("GOT")
        Else
            MsgDataNotFound()
        End If
    End Sub
    Public Sub FilterTableAll(TableName, TableGrid, SomeText, Dates, DateStart, DateEnd)
        Commands = New OleDbCommand("SELECT * FROM " & TableName & " WHERE kode_barang='" & SomeText & "' AND " &
                                    Dates & " BETWEEN #" & DateStart & "# AND #" & DateEnd & "#", Connections)
        Readers = Commands.ExecuteReader
        If Readers.Read Then
            DataAdapters = New OleDbDataAdapter("SELECT * FROM " & TableName & " WHERE kode_barang='" & SomeText & "' AND " &
                                                Dates & " BETWEEN #" & DateStart & "# AND #" & DateEnd & "#", Connections)
            DataSets = New DataSet
            DataAdapters.Fill(DataSets, "GOT")
            TableGrid.DataSource = DataSets.Tables("GOT")
        Else
            MsgDataNotFound()
        End If
    End Sub
    Public Sub FillComboBox(CmBox)
        CmBox.Items.Clear()
        Commands = New OleDbCommand("SELECT * FROM data_barang ORDER BY kode_barang", Connections)
        Readers = Commands.ExecuteReader
        While Readers.Read
            CmBox.Items.Add(Readers.GetValue(0))
        End While
        Readers.Close()
    End Sub

    Public Sub FillNamaBarang(CmBox, TxtSomething)
        If CmBox.SelectedIndex >= 0 Then
            KodeBarang = CmBox.Text
            Commands = New OleDbCommand("SELECT * FROM data_barang WHERE kode_barang='" &
                                        KodeBarang & "'", Connections)
            Readers = Commands.ExecuteReader
            Readers.Read()
            TxtSomething.Text = Readers.GetValue(1)
            Readers.Close()
        Else
            TxtSomething.Text = ""
        End If
    End Sub

    Sub UpdateNamaBarang(TableName)
        Commands = New OleDbCommand("UPDATE " & TableName & " SET nama_barang='" &
                                    NamaBarang & "' WHERE kode_barang='" & KodeBarang & "'",
                                    Connections)
        Commands.ExecuteNonQuery()
    End Sub
    Public Sub InsertBarang(TableName, TanggalTable, KodeBarang, NamaBarang, Kuantitas, Tanggal, TanggalKadaluwarsa)
        Commands = New OleDbCommand("INSERT INTO " & TableName & " (kode_barang, nama_barang, kuantitas, " &
                                    TanggalTable & ", tanggal_kadaluwarsa) VALUES ('" & KodeBarang & "', '" &
                                    NamaBarang & "', '" & Kuantitas & "', '" & Tanggal & "', '" &
                                    TanggalKadaluwarsa & "')", Connections)
        Commands.ExecuteNonQuery()
    End Sub
    Public Sub UpdateBarang(TableName, TanggalTable, Operators, KodeBarang, Kuantitas, Tanggal, TanggalKadaluwarsa)
        Commands = New OleDbCommand("UPDATE " & TableName & " SET kuantitas=(kuantitas " & Operators & "'" &
                                    Kuantitas & "') WHERE kode_barang='" & KodeBarang & "' AND " &
                                    TanggalTable & "=@" & TanggalTable & " And tanggal_kadaluwarsa=" &
                                    "@tanggal_kadaluwarsa", Connections)
        Commands.Parameters.Add(TanggalTable, OleDbType.Date).Value = Tanggal
        Commands.Parameters.Add("tanggal_kadaluwarsa", OleDbType.Date).Value = TanggalKadaluwarsa
        Commands.ExecuteNonQuery()
    End Sub
    Public Sub DeleteBarang(TableName, TanggalTable, KodeBarang, Tanggal, TanggalKadaluwarsa)
        Commands = New OleDbCommand("DELETE * FROM " & TableName & " WHERE kode_barang='" & KodeBarang &
                                    "' AND " & TanggalTable & "=@" & TanggalTable & " AND tanggal_kadaluwarsa=" &
                                    "@tanggal_kadaluwarsa", Connections)
        Commands.Parameters.Add(TanggalTable, OleDbType.Date).Value = Tanggal
        Commands.Parameters.Add("tanggal_kadaluwarsa", OleDbType.Date).Value = TanggalKadaluwarsa
        Commands.ExecuteNonQuery()
    End Sub

    Public Sub InsertPersediaan(KodeBarang, NamaBarang, Kuantitas, TanggalKadaluwarsa)
        Commands = New OleDbCommand("INSERT INTO persediaan_barang (kode_barang, nama_barang, kuantitas, " &
                                    "tanggal_kadaluwarsa) VALUES ('" & KodeBarang & "', '" &
                                    NamaBarang & "', '" & Kuantitas & "', '" & TanggalKadaluwarsa & "')",
                                    Connections)
        Commands.ExecuteNonQuery()
    End Sub
    Public Sub UpdatePersediaan(Kuantitas, Operators, KodeBarang, TanggalKadaluwarsa)
        Commands = New OleDbCommand("UPDATE persediaan_barang SET kuantitas=(kuantitas " &
            Operators & "'" & Kuantitas & "') WHERE kode_barang='" &
            KodeBarang & "' And tanggal_kadaluwarsa=@tanggal_kadaluwarsa",
            Connections)
        Commands.Parameters.Add("tanggal_kadaluwarsa", OleDbType.Date).Value = TanggalKadaluwarsa
        Commands.ExecuteNonQuery()
    End Sub
    Public Sub DeletePersediaan(KodeBarang, TanggalKadaluwarsa)
        Commands = New OleDbCommand("DELETE * FROM persediaan_barang WHERE kode_barang='" & KodeBarang &
                                    "' AND tanggal_kadaluwarsa=@tanggal_kadaluwarsa", Connections)
        Commands.Parameters.Add("tanggal_kadaluwarsa", OleDbType.Date).Value = TanggalKadaluwarsa
        Commands.ExecuteNonQuery()
    End Sub
End Module
