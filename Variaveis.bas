Attribute VB_Name = "Variaveis"
Public BancoDeDados As Database
Public TbContatos As Recordset
Public BuscaContato As String
Public Sub AbreArquivo()
    Set BancoDeDados = OpenDatabase(App.Path & "\Agenda.mdb")
    Set TbContatos = BancoDeDados.OpenRecordset("contatos", dbOpenTable)
End Sub
Public Sub FechaArquivo()
    TbContatos.Close
    BancoDeDados.Close
End Sub
