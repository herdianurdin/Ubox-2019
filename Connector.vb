Imports System.Data.OleDb
Module Connector
    Public Connections As OleDbConnection
    Public DataAdapters As OleDbDataAdapter
    Public DataSets As DataSet
    Public Commands As OleDbCommand
    Public Readers As OleDbDataReader
    Public Readerz As OleDbDataReader
    Public Sub Connects()
        Connections = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" &
                                          Application.StartupPath & "\KelolaBarang.accdb")
        If Connections.State = ConnectionState.Closed Then
            Connections.Open()
        End If
    End Sub
End Module
