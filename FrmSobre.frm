VERSION 5.00
Begin VB.Form FrmSobre 
   Caption         =   "Sobre"
   ClientHeight    =   2790
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   5070
   Begin VB.CommandButton BtnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Software"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.Label Label10 
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "Felipe Prestes"
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "1.2"
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "30 de Julho de 2013"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Data de criação:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Versão:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Desenvolvedor:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmSobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnOk_Click()
Unload Me
End Sub

Private Sub Label5_Click()

End Sub
