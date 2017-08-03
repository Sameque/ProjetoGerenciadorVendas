VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmPessoaFuncionario 
   Caption         =   "Cadastro de Pessoa"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2580
   ScaleWidth      =   4290
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   2407
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Excluir o Item Selecionado"
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3375
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Sair da tela"
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   855
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Incluir novo Item"
      Top             =   1680
      Width           =   855
   End
   Begin Threed.SSFrame fraDados 
      Height          =   1455
      Left            =   45
      TabIndex        =   0
      Top             =   120
      Width           =   4170
      _Version        =   65536
      _ExtentX        =   7355
      _ExtentY        =   2566
      _StockProps     =   14
      Caption         =   "Dados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin Transportes.SuperText txtSenha 
         Height          =   510
         Left            =   2070
         TabIndex        =   3
         Top             =   825
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   900
         PasswordChar    =   "*"
         MensagemValidacao=   "a Senha"
         Label           =   "Senha"
      End
      Begin Transportes.SuperText txtUsuario 
         Height          =   510
         Left            =   75
         TabIndex        =   2
         Top             =   825
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   900
         MensagemValidacao=   "o Usuário"
         Label           =   "Usuário"
      End
      Begin Transportes.SuperDBCombo cboFuncao 
         Height          =   510
         Left            =   75
         TabIndex        =   1
         Top             =   255
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   900
         MensagemValidacao=   "a Função"
         Label           =   "Função"
      End
   End
End
Attribute VB_Name = "frmPessoaFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public id_PessoaFuncionario As Long
Public id_Pessoa As Long
Public FormChamador As Form

Dim cPessoa As clsPessoa

Private Sub cmdExcluir_Click()
On Error GoTo err_cmdExcluir_Click

    If Mensagem("Confirma exclusão?", Pergunta) = vbNo Then
        Exit Sub
    End If

    Call AbreTransacao
    If Not cPessoa.ExcluirFuncionario Then
        Call VoltaTransacao
        Mensagem "Ocorreu um erro na exclusão.", ErroCritico
        Exit Sub
    End If
    
    Call FechaTransacao
    
    Call LimparControles(Me)
'    FormChamador.id_PessoaFuncionario = 0
'    id_PessoaFuncionario = 0
    Mensagem "Exclusão efetuada.", Informacao

    Exit Sub
err_cmdExcluir_Click:
    ShowError
    Call VoltaTransacao
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPessoaFuncionario = Nothing
End Sub

Private Sub Form_Activate()
    frmMDI.Arrange vbCascade
End Sub

Private Sub Form_Load()
    Call CarregarTela
    Call CenterForm(Me)
End Sub

Private Function CarregarTela()
    Set cPessoa = New clsPessoa
    
    Call cboFuncao.Formatar("id_PessoaFuncao, ds_PessoaFuncao, qt_NivelPermissao", "0,1000,500", "false,true,true", "tbdPessoaFuncao", "", "ds_PessoaFuncao")
    
    If cPessoa.CarregarDados(id_Pessoa) Then
        txtUsuario.Text = cPessoa.ds_Usuario
        txtSenha = cPessoa.ds_Senha
        id_Pessoa = cPessoa.id_Pessoa
        
        Call cboFuncao.PesquisarCombo(True, cPessoa.id_PessoaFuncao, "", True)
    End If
    
    cboFuncao.CampoObrigatorio = True
    txtUsuario.CampoObrigatorio = True
    txtSenha.CampoObrigatorio = True
    
End Function
Private Sub cmdSair_Click()
    'FormChamador.id_PessoaFuncionario = id_PessoaFuncionario
    Unload Me
End Sub

Private Sub cmdGravar_Click()
On erro GoTo err_cmdGravar_Click

    If Not ValidarControles(Me) Then
        Exit Sub
    End If
    
    If Mensagem("Confirma Gravação?", Pergunta) = vbNo Then
        Exit Sub
    End If

    If Not CarregarPropriedades(cPessoa) Then
        Exit Sub
    End If

    Call AbreTransacao
    If Not cPessoa.GravarFuncionario Then
        Call VoltaTransacao
        Mensagem "Ocorreu um erro no processamento.", ErroCritico
        Exit Sub
    End If
    Call FechaTransacao

'    id_PessoaFuncionario = cPessoa.id_PessoaFuncionario
    Call FormChamador.AtualizarDados(id_Pessoa, 0)
    'FormChamador.id_PessoaFuncionario = id_PessoaFuncionario
    
    Call Mensagem("Gravação Efetuada", Informacao)
    cmdSair.SetFocus
    
    Exit Sub
err_cmdGravar_Click:
    ShowError
End Sub
Private Function CarregarPropriedades(cPessoa As clsPessoa) As Boolean
On Error GoTo err_CarregarPropriedades

    CarregarPropriedades = False
    
    cPessoa.id_Pessoa = id_Pessoa
    cPessoa.id_PessoaFuncionario = id_PessoaFuncionario
    cPessoa.id_PessoaFuncao = cboFuncao.ItemData2
    cPessoa.ds_Usuario = txtUsuario.Text
    cPessoa.ds_Senha = txtSenha.Text
    
    CarregarPropriedades = True
    
    Exit Function
err_CarregarPropriedades:
    ShowError
End Function
