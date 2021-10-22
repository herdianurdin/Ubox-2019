Module MsgGlobal
    Public Sub MsgDataEmpty()
        MsgBox("Silakan Isi Datanya Dulu!", vbExclamation, "Perhatian")
    End Sub
    Public Sub MsgDataSaved()
        MsgBox("Data Berhasil Disimpan!", vbInformation, "Pemberitahuan")
    End Sub
    Public Sub MsgDataDeleted()
        MsgBox("Data Berhasil Dihapus!", vbInformation, "Pemberitahuan")
    End Sub
    Public Sub MsgDataWantDelete()
        MsgBox("Data Yang Ingin Dihapus Tidak Ada!", vbExclamation, "Perhatian")
    End Sub

    Public Sub MsgChoiceKode()
        MsgBox("Silakan Pilih Kode Barang Dulu!", vbExclamation, "Perhatian")
    End Sub
    Public Sub MsgFillKuantitas()
        MsgBox("Silakan Isi Kuantitasnya Dulu!", vbExclamation, "Perhatian")
    End Sub
    Public Sub MsgDateCheck(Captions)
        MsgBox("Periksa Kembali Tanggal " + Captions, vbExclamation, "Perhatian")
    End Sub
    Public Sub MsgDataNotFound()
        MsgBox("Data Tidak Ditemukan!", vbInformation, "Pemberitahuan")
    End Sub
    Public Sub MsgFilteringFail()
        MsgBox("Silakan Tandai Filternya!", vbExclamation, "Perhatian")
    End Sub
End Module
