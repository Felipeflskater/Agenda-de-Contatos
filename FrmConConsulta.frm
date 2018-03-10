VERSION 5.00
Begin VB.Form FrmConConsulta 
   Caption         =   "Localizar de Contatos"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton BtnLocalizar 
      Caption         =   "Localizar"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton BtnCancelar 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.ListBox LstContatos 
      Height          =   2400
      ItemData        =   "FrmConConsulta.frx":0000
      Left            =   0
      List            =   "FrmConConsulta.frx":0002
      TabIndex        =   3
      Top             =   840
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Digite o nome a localizar:"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.TextBox TxtLocNome 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmConConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnCancelar_Click()
BuscaContatoe = ""
Unload Me
End Sub
Private Sub BtnLocalizar_Click()
TbContatos.MoveFirst
LstContatos.Clear
Do While TbContatos.EOF = False
If InStr(LCase(TbContatos("nome")), LCase(TxtLocNome.Text)) = 1 Then
LstContatos.AddItem TbContatos("nome")
End If
TbContatos.MoveNext
Loop
End Sub
Private Sub BtnOK_Click()
BuscaContato = LstContatos.Text
Unload Me
End Sub
Private Sub form_load()
BtnLocalizar.Enabled = False
BtnOK.Enabled = False
BuscaContato = ""
CarregaContatos
End Sub
Private Sub LstContatos_Click()
BtnOK.Enabled = True
BtnOK.Default = True
End Sub
Private Sub LstContatos_DblClick()
BtnOK_Click
End Sub
Private Sub TxtLocNome_change()
If TxtLocNome.Text = "" Then
CarregaContatos
BtnLocalizar.Enabled = False
BtnOK.Enabled = False
Else
BtnLocalizar.Enabled = True
BtnOK.Enabled = True
BtnLocalizar.Default = True
BtnLocalizar_Click
End If
End Sub
Private Sub CarregaContatos()
LstContatos.Clear
TbContatos.MoveFirst
Do While TbContatos.EOF = False
LstContatos.AddItem TbContatos("nome")
TbContatos.MoveNext
Loop
End Sub
