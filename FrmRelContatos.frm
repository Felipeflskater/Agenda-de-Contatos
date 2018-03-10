VERSION 5.00
Begin VB.Form FrmRelContatos 
   Caption         =   "Relatorio de Contatos"
   ClientHeight    =   3525
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   5295
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Ordem de Impressão"
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5055
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton OptNome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Relatorio"
      Height          =   2295
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
      Begin VB.OptionButton OptParcial 
         Caption         =   "Parcial"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton OptCompleto 
         Caption         =   "Completo"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.CommandButton BtnImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton BtnCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   2880
      Width           =   2415
   End
End
Attribute VB_Name = "FrmRelContatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnCancelar_Click()
    Unload Me
End Sub
Private Sub BtnImprimir_Click()
    If MsgBox("Deseja imprimir este Contato?", vbYesNo, "impressao") = vbNo Then
        Exit Sub
    End If
    If OptParcial.Value = True Then
        RelParcial
    Else
        RelCompleto
    End If
End Sub
Private Sub RelParcial()
 Dim Titulo As String
 Dim Linhas As Byte
    If OptCodigo.Value = True Then
        TbContatos.Index = "indcodigo"
        Titulo = "Relatorio de Contatos - Ordem de Codigo"
    Else
        TbContatos.Index = "indnome"
        Titulo = "Relatorio de Contatos - Ordem de Nome"
    End If
        Printer.Print "Codigo"; Tab(10); "Nome"; Tab(73); "Fone"
        Printer.Print String(80, "-")
        Linhas = 0
        TbContatos.MoveFirst
    Do While TbContatos.EOF = False
        Printer.Print TbContatos("Codigo"); Tab(10); TbContatos("Nome"); Tab(73); TbContatos("Fone")
        TbContatos.MoveNext
        Linhas = Linhas + 1
    If Linhas > 38 Then
        Printer.NewPage
        Printer.Print "Codigo"; Tab(10); "Nome"; Tab(73); "Fone"
        Printer.Print String(80, "-")
        Linhas = 0
    End If
    Loop
        Printer.EndDoc
End Sub
Private Sub RelCompleto()
 Dim Titulo As String
 Dim Linhas As Byte
    If OptCodigo.Value = True Then
        TbContatos.Index = "indcodigo"
        Titulo = "Relatorio de Contatos - Ordem de Codigo"
    Else
        TbContatos.Index = "indnome"
        Titulo = "Relatorio de Contatos - Ordem de Nome"
    End If
        Printer.Orientation = 2
        Cabecalho = (Titulo)
        Printer.FontSize = 8
        Printer.Print "Codigo"; Tab(10); "Nome"; Tab(73); "Endereco"; Tab(100); "Bairro"; Tab(130); "Cidade"; Tab(150); "Cep"; Tab(180); "Fone"; "Fax"; Tab(200); "Celular"; Tab(230); "Email"; Tab(250); "Data de Cadastro"; Tab(300);
        Printer.Print String(80, "-")
        Linhas = 0
        TbContatos.MoveFirst
    Do While TbContatos.EOF = False
        Printer.Print TbContatos("Codigo"); Tab(4); TbContatos("Cidade"); Tab(150); TbContatos("Cep"); Tab(50); TbContatos("Fone")
        TbContatos.MoveNext
        Linhas = Linhas + 1
    If Linhas > 38 Then
        Printer.NewPage
        Cabecalho = (Titulo)
        Printer.Print "Codigo"; Tab(10); "Nome"; Tab(73); "Endereco"; Tab(100); "Bairro"; Tab(130); "Cidade"; Tab(150); "Cep"; Tab(180); "Fone"
        Printer.Print String(80, "-")
        Linhas = 0
    End If
    Loop
        Printer.EndDoc
End Sub

