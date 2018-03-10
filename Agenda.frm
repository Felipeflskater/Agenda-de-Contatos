VERSION 5.00
Begin VB.MDIForm MDIAgenda 
   BackColor       =   &H8000000C&
   Caption         =   "Agenda"
   ClientHeight    =   -255
   ClientLeft      =   225
   ClientTop       =   1170
   ClientWidth     =   3360
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MnuContatos 
      Caption         =   "Contatos"
      Begin VB.Menu MnuCadastros 
         Caption         =   "Cadastros"
      End
   End
   Begin VB.Menu MnuRelatorio 
      Caption         =   "Relatorio"
      Begin VB.Menu SMnuContatos 
         Caption         =   "Contatos"
      End
   End
   Begin VB.Menu MnuSobre 
      Caption         =   "Sobre"
      Begin VB.Menu MnuSistema 
         Caption         =   "Sistema"
      End
   End
   Begin VB.Menu MnuAjuda 
      Caption         =   "Ajuda"
      Begin VB.Menu SMnuSistema 
         Caption         =   "Sistema"
      End
   End
End
Attribute VB_Name = "MDIAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
AbreArquivo
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
FechaArquivo
End Sub
Private Sub MnuCadastros_Click()
FrmCadastro.Show
End Sub
Private Sub MnuSistema_Click()
FrmSobre.Show
End Sub
Private Sub SMnuContatos_Click()
FrmRelContatos.Show
End Sub
Private Sub SMnuSistema_Click()
FrmAjuda.Show
End Sub
