VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmCadastro 
   Caption         =   "Cadastros"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9330
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4410
   ScaleWidth      =   9330
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Menu de Botões:"
      Height          =   4335
      Left            =   6480
      TabIndex        =   23
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton BtnSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   3600
         Width           =   2535
      End
      Begin VB.CommandButton BtnImprimir 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   3120
         Width           =   2535
      End
      Begin VB.CommandButton BtnAnterior 
         Caption         =   "Anterior"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CommandButton BtnProximo 
         Caption         =   "Próximo"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton BtnLocalizar 
         Caption         =   "Localizar"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton BtnExcluir 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton BtnAlterar 
         Caption         =   "Alterar"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton BtnInserir 
         Caption         =   "Inserir"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin MSMask.MaskEdBox MskData 
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   3840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskEmail 
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   3480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskCelular 
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         Top             =   3120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskFax 
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   2760
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskFone 
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskCep 
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   2040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtCidade 
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox TxtBairro 
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox TxtEndereco 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox TxtNome 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   600
         Width           =   3735
      End
      Begin MSMask.MaskEdBox MskCodigo 
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         Caption         =   "Data de Cadastro:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "E-Mail:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Celular:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Fax:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Fone:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Cep:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Cidade:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Bairro:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Endereço:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AtualizaFormulario()
If TbContatos.RecordCount > 0 Then
MskCodigo.Text = TbContatos("Codigo")
TxtNome.Text = TbContatos("Nome")
TxtEndereco.Text = TbContatos("Endereco")
TxtBairro.Text = TbContatos("Bairro")
TxtCidade.Text = TbContatos("Cidade")
MskCep.Text = TbContatos("CEP")
MskFone.Text = TbContatos("Fone")
MskFax.Text = TbContatos("Fax")
MskCelular.Text = TbContatos("Celular")
MskEmail.Text = TbContatos("E-Mail")
MskData.Text = TbContatos("Data de Cadastro")
Else
LimpaFormulario
End If
End Sub
Private Sub AtualizaCampos()
TbContatos("Codigo") = MskCodigo.Text
TbContatos("Nome") = TxtNome.Text
TbContatos("Endereco") = TxtEndereco.Text
TbContatos("Bairro") = TxtBairro.Text
TbContatos("Cidade") = TxtCidade.Text
TbContatos("CEP") = MskCep.Text
TbContatos("Fone") = MskFone.Text
TbContatos("Fax") = MskFax.Text
TbContatos("Celular") = MskCelular.Text
TbContatos("E-Mail") = MskEmail.Text
TbContatos("Data de Cadastro") = MskData.Text
End Sub
Private Sub LimpaFormulario()
MskCodigo.Text = " "
TxtNome.Text = ""
TxtEndereco.Text = ""
TxtBairro.Text = ""
TxtCidade.Text = ""
MskCep.Text = ""
MskFone.Text = ""
MskFax.Text = ""
MskCelular.Text = ""
MskEmail.Text = ""
MskData.Text = "  /  /    "
End Sub
Private Sub HabilitaControles()
MskCodigo.Enabled = True
TxtNome.Enabled = True
TxtEndereco.Enabled = True
TxtBairro.Enabled = True
TxtCidade.Enabled = True
MskCep.Enabled = True
MskFone.Enabled = True
MskFax.Enabled = True
MskCelular.Enabled = True
MskEmail.Enabled = True
MskData.Enabled = True
End Sub
Private Sub DesabilitaControles()
MskCodigo.Enabled = False
TxtNome.Enabled = False
TxtEndereco.Enabled = False
TxtBairro.Enabled = False
TxtCidade.Enabled = False
MskCep.Enabled = False
MskFone.Enabled = False
MskFax.Enabled = False
MskCelular.Enabled = False
MskEmail.Enabled = False
MskData.Enabled = False
End Sub
Private Sub DesativaBotoes()
BtnInserir.Enabled = False
BtnAlterar.Enabled = False
BtnExcluir.Enabled = False
BtnImprimir.Enabled = False
BtnProximo.Enabled = False
BtnAnterior.Enabled = False
BtnLocalizar.Enabled = False
End Sub
Private Sub AtivaBotoes()
If TbContatos.RecordCount > 0 Then
BtnAlterar.Enabled = True
BtnExcluir.Enabled = True
BtnImprimir.Enabled = True
BtnProximo.Enabled = True
BtnAnterior.Enabled = True
BtnLocalizar.Enabled = True
Else
DesativaBotoes
End If
BtnInserir.Enabled = True
BtnSair.Enabled = True
End Sub
Private Sub form_load()
TbContatos.Index = "IndNome"
If TbContatos.RecordCount > 0 Then
AtualizaFormulario
End If
DesabilitaControles
AtivaBotoes
End Sub
Private Sub BtnAlterar_Click()
If BtnAlterar.Caption = "Alterar" Then
HabilitaControles
DesativaBotoes
BtnAlterar.Caption = "Confirmar"
BtnAlterar.Enabled = True
BtnSair.Caption = "Cancelar"
MskCodigo.Enabled = False
TxtNome.SetFocus
Else
If MsgBox("Confirma Alteração?", vbYesNo, "Alteração") = vbYes Then
TbContatos.Edit
AtualizaCampos
TbContatos.Update
End If
BtnAlterar.Caption = "Alterar"
BtnSair.Caption = "Sair"
AtivaBotoes
AtualizaFormulario
DesabilitaControles
End If
End Sub
Private Sub BtnAnterior_Click()
If TbContatos.BOF = False Then
TbContatos.MovePrevious
End If
If TbContatos.BOF = True Then
TbContatos.MoveLast
End If
AtualizaFormulario
End Sub
Private Sub BtnExcluir_click()
If MsgBox("Deseja excluir este Contato?", vbYesNo + vbDefaultButton2, "Exclusão") = vbYes Then
TbContatos.Delete
BtnAnterior_Click
AtivaBotoes
End If
End Sub
Private Sub BtnImprimir_Click()
Dim Titulo As String
If MsgBox("Deseja Imprimir este Contato?", vbYesNo, "Impressão") = vbNo Then
Exit Sub
End If
Titulo = "Ficha Individual de Contatos"
Printer.FontSize = 14
Printer.Print
Printer.Print "Codigo:"; MskCodigo.Text
Printer.Print
Printer.Print "Nome:"; TxtNome.Text
Printer.Print
Printer.Print "Bairro:"; TxtBairro.Text
Printer.Print
Printer.Print "Cidade:"; TxtCidade.Text
Printer.Print
Printer.Print "CEP:"; MskCep.Text
Printer.Print
Printer.Print "Telefone:"; MskFone.Text
Printer.Print
Printer.Print "Fax:"; MskFax.Text
Printer.Print
Printer.Print "Data de Cadastro:"; MskData.Text
Printer.EndDoc
End Sub
Private Sub BtnInserir_Click()
If BtnInserir.Caption = "Inserir" Then
LimpaFormulario
HabilitaControles
DesativaBotoes
MskData.Text = Date
BtnInserir.Enabled = True
BtnInserir.Caption = "Confirmar"
BtnSair.Caption = "Cancelar"
MskCodigo.SetFocus
Else
If MskCodigo.Text = "  " Then
MsgBox "Voce nao digitou o codigo", vbCritical, "Cadastro de Contatos"
MskCodigo.SetFocus
Exit Sub
Else
If MsgBox("Deseja gravar este Contato?", vbYesNo, "cadastro de Contatos") = vbYes Then
TbContatos.AddNew
AtualizaCampos
TbContatos.Update
End If
End If
BtnInserir.Caption = "Inserir"
BtnSair.Caption = "Sair"
AtualizaFormulario
DesabilitaControles
AtivaBotoes
End If
End Sub
Private Sub BtnLocalizar_Click()
FrmConConsulta.Show 1
If BuscaContato <> "" Then
TbContatos.Seek "=", BuscaContato
Else
TbContatos.Seek "=", TxtNome.Text
End If
AtualizaFormulario
End Sub
Private Sub BtnProximo_Click()
If TbContatos.EOF = False Then
TbContatos.MoveNext
End If
If TbContatos.EOF = True Then
TbContatos.MoveFirst
End If
AtualizaFormulario
End Sub
Private Sub BtnSair_Click()
If BtnSair.Caption = "Sair" Then
Unload Me
Else
AtualizaFormulario
DesabilitaControles
AtivaBotoes
BtnInserir.Caption = "Inserir"
BtnAlterar.Caption = "Alterar"
BtnSair.Caption = "Sair"
End If
End Sub
Private Sub MskCodigo_LostFocus()
MskCodigo.Text = Format(MskCodigo.Text, "00000")
TbContatos.Index = "IndCodigo"
TbContatos.Seek "=", MskCodigo.Text
If TbContatos.NoMatch = False Then
MsgBox "ja existe um Contato com este codigo!", vbInformation, "inclusao"
MskCodigo.Text = "  "
MskCodigo.SetFocus
End If
TbContatos.Index = "IndNome"
End Sub
Private Sub MskData_LostFocus()
If IsDate(MskData.Text) = False Then
MsgBox "Data Incorreta!", vbInformation, "Inclusão"
MskData.SetFocus
End If
End Sub
