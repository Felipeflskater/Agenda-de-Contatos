VERSION 5.00
Begin VB.Form FrmAjuda 
   Caption         =   "Ajuda"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   8940
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Relatorios de Contatos"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Instruções:"
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.Frame Frame2 
         Caption         =   "Links de Ajuda"
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   4320
         Width           =   8655
         Begin VB.CommandButton BtnSair 
            Caption         =   "Sair"
            Height          =   375
            Left            =   5640
            TabIndex        =   10
            Top             =   3240
            Width           =   2775
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Botões de Comando"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   1680
            Width           =   2775
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Entendendo os Campos"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   2775
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Sobre"
            Height          =   375
            Left            =   5640
            TabIndex        =   7
            Top             =   2760
            Width           =   2775
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Entendendo o Programa"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2775
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Localizando Contatos"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   2160
            Width           =   2775
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Cadastrando Contatos "
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   720
            Width           =   2775
         End
      End
      Begin VB.Label LblManual 
         BorderStyle     =   1  'Fixed Single
         Height          =   3495
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   8655
      End
   End
End
Attribute VB_Name = "FrmAjuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command7_Click()
FrmSobre.Show
'Sobre'
End Sub
Private Sub BtnSair_Click()
Unload Me
End Sub
Private Sub Command1_Click()
LblManual.Caption = "Este software foi desnvolvido para voce armazenar em banco de dados seus contatos. o mesmo dispoe de quatro menus sendo dois funcionais(Contatos e relatorio) e dois informativos(Sobre e ajuda) cujas respectivas funções são: Menu contatos e atraves dele que obtemos acesso ao submenu cadastros no qual armazenamos os contatos no banco de dados,temos tambem o menu relatorio que nos da acesso ao submenu contatos onde são criados relatorios constando informações sobre os contatos,o menu sobre esta ao lado de relatorios esse nos da acesso ao menu sistema trazendo imformações sobre o desenvolvedor e o software e temos tmbem o menu ajuda que nos explica desde s funções mais basicas até as avnaçadas do software."
'Entendendo o software'
End Sub
Private Sub Command2_Click()
LblManual.Caption = "Para cadastrar contatos em seu Banco de Dados,primeiramente Clicamos no menu contatos para ter acesso ao submenu cadastros,clicamos para o carregamento do mesmo,surgira uma nova tela onde existirao varios campos e comandos,para adicionar um novo contato se aperta o botao inserir, entao voce preenche os campos e aperta confirma para crialo no Banco de Dados ou cancelar se desejar desistir do cadastro"
'Cadastrando Contatos'
End Sub
Private Sub Command8_Click()
LblManual.Caption = "A tela de Cadastros contem 11 campos para armazenamento de informaçoes sendo que o campo data de cadastro e automaticamente preenchido pelo programa,o campo codigo so aceita caracteres numericos e nao pode ser o mesmo de outros contatos e nao pode ser nulo(obs: e muito bom procurar manter seu banco de dados organizado cadastrando os contatos com o campo codigo em sequencia). os campos Nome,Endereço,Bairro,Cidade e Email aceitam caracteres do tipo alfanumerico, ja os campos Cep,Fone,Fax e Celular somente suportao caracteres numericos."
'Campos'
End Sub
Private Sub Command9_Click()
LblManual.Caption = "O Subenu Cadastros contem 8 Botoes de Comando sao eles Inserir,Alterar,Excluir,Localizar,Proximo,Anterior,Imprimir e Sair. O botão inserir tem a funçao de criar um novo contao,ao clicarmos ele libera os campos para o cadastro. O botao Alterar e usado em cadastros ja criados com a finalidade de trocar ou adicionar imformaçoes do contato. O botao excluir apaga do banco de dados o contato atual mostrado na tela de cadastros. O botao localizar abre uma janela ondem podem ser localizado qualquer contato com mais facilidade.Os botoes proximo e anterior trocam o contato a ser visualizado na nos campos, o botao proximo vai para o contato seguinte de acordo com a oredem do campo codigo e o botao anterior para o contato anterior.O botao Imprimir faz a impressao de todas informaçoes do contato atual. O botao Sair sai da tela de Cadastros para que voce possa acessar outras telas."
'Botoes de Comando'
End Sub
Private Sub Command6_Click()
LblManual.Caption = "Na tela do submenu cadastros contem o botao de comando Localizar ele serve para localizarmos os contatos mais facilmente,ao clicarmos nele ele abre uma janela de busca onde temos um espaço onde digitamos o nome e ele encontra ao ser encontrado clicasse no contato  e clicase ok para que possam ser visualizadas suas informaçoes na tela cadastros mas e claro se quiser sair ou cancelar a busca do contato aperte em cancelar."
'Localizando Contatos'
End Sub
Private Sub Command3_Click()
LblManual.Caption = "Para criar um relatorio com todos os contatos cadastrados,clicasse no menu Relatorio e depois no submenu contatos,os relatorios podem ser impressos por nome(ordem alfabetica) ou podem ser impressos por codigo(ordem de numerica do codigo de cadstro dos contatos).Alem disso existem dois tipos de impressao para o relatorio o parcial(Campos Codigo,Nome e Telefone) e o completo(Todos os campos) "
'Relatorios'
End Sub
