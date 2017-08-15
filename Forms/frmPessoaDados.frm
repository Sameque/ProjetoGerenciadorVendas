VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmPessoaDados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Pessoa"
   ClientHeight    =   5310
   ClientLeft      =   3030
   ClientTop       =   4905
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8430
   Begin Transportes.SuperSpreadNovo sprContato 
      Height          =   1695
      Left            =   60
      TabIndex        =   20
      Top             =   2550
      Width           =   8340
      _extentx        =   14711
      _extenty        =   2990
      label           =   "Contatos"
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Height          =   750
      HelpContextID   =   23
      Left            =   6675
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Novo lançamento "
      Top             =   4425
      Width           =   810
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   750
      Left            =   5805
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Gravar os Dados"
      Top             =   4425
      Width           =   810
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   750
      Left            =   7545
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Sair da tela"
      Top             =   4425
      Width           =   810
   End
   Begin Threed.SSFrame fraDados 
      Height          =   2415
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   8280
      _Version        =   65536
      _ExtentX        =   14605
      _ExtentY        =   4260
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
      Begin Transportes.SuperControlNovo mskCEP 
         Height          =   315
         Left            =   2380
         TabIndex        =   4
         Top             =   1065
         Width           =   1740
         _extentx        =   2540
         _extenty        =   556
         tooltip         =   ""
         mascara         =   7
         mensagemvalidacao=   "o CEP"
      End
      Begin Transportes.SuperText txtCNPJCPF 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1065
         Width           =   2220
         _extentx        =   3916
         _extenty        =   503
         skinesl         =   -1  'True
         backcolor       =   16119285
         mensagemvalidacao=   "CNPJ ou CPF"
         campoobrigatorio=   -1  'True
      End
      Begin Transportes.SuperText txtBairro 
         Height          =   315
         Left            =   4155
         TabIndex        =   5
         Top             =   1050
         Width           =   4000
         _extentx        =   7064
         _extenty        =   556
         mensagemvalidacao=   "o Bairro"
      End
      Begin Transportes.SuperText txtEndereco 
         Height          =   315
         Left            =   135
         TabIndex        =   6
         Top             =   1620
         Width           =   4005
         _extentx        =   7064
         _extenty        =   556
         mensagemvalidacao=   "o Endereço"
      End
      Begin Transportes.SuperText txtPessoa 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   525
         Width           =   4000
         _extentx        =   7064
         _extenty        =   556
         mensagemvalidacao=   "o Nome"
         campoobrigatorio=   -1  'True
      End
      Begin Transportes.SuperText txtRazaoSocial 
         Height          =   315
         Left            =   4155
         TabIndex        =   2
         Top             =   525
         Width           =   4005
         _extentx        =   7064
         _extenty        =   556
         mensagemvalidacao=   "a Razão Social"
         campoobrigatorio=   -1  'True
      End
      Begin Transportes.SuperDBCombo cboCidade 
         Height          =   510
         Left            =   4155
         TabIndex        =   7
         Top             =   1425
         Width           =   4000
         _extentx        =   7064
         _extenty        =   900
         mensagemvalidacao=   "a Cidade"
         label           =   "Cidade"
      End
      Begin Threed.SSCheck chkFuncionario 
         Height          =   150
         Left            =   2385
         TabIndex        =   10
         Top             =   2040
         Width           =   1350
         _Version        =   65536
         _ExtentX        =   2381
         _ExtentY        =   265
         _StockProps     =   78
         Caption         =   "Funcionário"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCheck chkFornecedor 
         Height          =   150
         Left            =   1170
         TabIndex        =   9
         Top             =   2040
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   265
         _StockProps     =   78
         Caption         =   "Fornecedor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCheck chkCliente 
         Height          =   150
         Left            =   180
         TabIndex        =   8
         Top             =   2040
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   265
         _StockProps     =   78
         Caption         =   "Cliente"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCEP 
         Caption         =   "CEP"
         Height          =   195
         Left            =   2385
         TabIndex        =   14
         Top             =   885
         Width           =   1740
      End
      Begin VB.Label lblcnpjcpf 
         Caption         =   "CNPJ/CPF"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   870
         Width           =   2220
      End
      Begin VB.Label lblBairro 
         Caption         =   "Bairro"
         Height          =   195
         Left            =   4155
         TabIndex        =   16
         Top             =   855
         Width           =   4000
      End
      Begin VB.Label lblEndereco 
         Caption         =   "Endereco"
         Height          =   195
         Left            =   135
         TabIndex        =   17
         Top             =   1395
         Width           =   4000
      End
      Begin VB.Label lblPessoa 
         Caption         =   "Nome"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   3975
      End
      Begin VB.Label lblRazaoSocial 
         Caption         =   "Razão Social"
         Height          =   195
         Left            =   4155
         TabIndex        =   19
         Top             =   300
         Width           =   4000
      End
   End
End
Attribute VB_Name = "frmPessoaDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public id_Pessoa As Long
Public FormChamador As Form
Public cPessoa As clsPessoa

Private Sub Form_Load()
On Error GoTo err_Form_Load
    
    Call FormatarComponentes
    Call CarregarCampos
    Call CarregarSpread
    Call CenterForm(Me)
    
    Exit Sub
err_Form_Load:
    ShowError
End Sub
Private Function FormatarComponentes()
On Error GoTo err_FormatarComponentes
    Dim cPessoaServico As New clsPessoaServico
            
    Call cboCidade.Formatar("a.id_Cidade, a.ds_Cidade, b.ds_Estado", "0,3000,1000", "false,true,true", "tbdCidade a left join tbdEstado b on a.id_Estado = b.id_Estado", "", "ds_Cidade")
    Call sprContato.FormatarPorClasse(cPessoaServico.FormatarSpreadPessoaContato)
    
    Exit Function
err_FormatarComponentes:
    ShowError
End Function

Private Function CarregarCampos()
On Error GoTo err_CarregarCampos
    Dim cPessoaServico As New clsPessoaServico
    
    Call sprContato.CarregarPorClasse("id_Pessoa = " & cPessoa.id_Pessoa)
    Set cPessoa = cPessoaServico.CarregarPorID(id_Pessoa)
            
    With cPessoa
        If .id_Pessoa > 0 Then
            mskCEP.Text = .cd_CEP
            txtCNPJCPF.Text = .cd_cnpjcpf
            txtBairro.Text = .ds_Bairro
            txtEndereco.Text = .ds_Endereco
            txtPessoa.Text = .ds_Pessoa
            txtRazaoSocial.Text = .ds_RazaoSocial
            chkCliente = IIf(.tp_Cliente = "S", True, False)
            chkFornecedor = IIf(.tp_Fornecedor = "S", True, False)
            chkFuncionario = IIf(.tp_Funcionario = "S", True, False)
            
            Call cboCidade.PesquisarCombo(True, .id_Cidade, "", True)
        End If
    End With
    
    Exit Function
err_CarregarCampos:
    ShowError
End Function

Private Sub cmdGravar_Click()
On Error GoTo err_cmdGravar_Click

    If Not ValidarControles(Me) Then
        Exit Sub
    End If
    
    If Mensagem("Confirma Gravação?", Pergunta) = vbNo Then
        Exit Sub
    End If
    
    Call GravarDados
    
    Call Mensagem("Gravação Efetuada", Informacao)
    
    cmdNovo.SetFocus

    Exit Sub
err_cmdGravar_Click:
    ShowError
    
End Sub
Private Function GravarDados()
On Error GoTo err_GravarDados

    If Not CarregarPropriedades(cPessoa) Then
        Exit Function
    End If

    Call AbreTransacao
    If Not cPessoa.Gravar Then
        Call VoltaTransacao
        Mensagem "Ocorreu um erro no processamento.", ErroCritico
        Exit Function
    End If
    Call FechaTransacao

    id_Pessoa = cPessoa.id_Pessoa
    
    Call FormChamador.AtualizarDados(id_Pessoa)
    
    'Call cPessoa.CarregarDados(id_Pessoa)
    
    Exit Function
err_GravarDados:
    ShowError
    Call FechaTransacao
End Function
Private Function CarregarPropriedades(cPessoa As clsPessoa) As Boolean
On Error GoTo err_CarregarPropriedades

    CarregarPropriedades = False
    
    cPessoa.id_Pessoa = id_Pessoa
    cPessoa.cd_CEP = mskCEP.ClipText
    cPessoa.cd_cnpjcpf = txtCNPJCPF.Text
    cPessoa.ds_Bairro = txtBairro.Text
    cPessoa.ds_Endereco = txtEndereco.Text
    cPessoa.ds_Pessoa = txtPessoa.Text
    cPessoa.ds_RazaoSocial = txtRazaoSocial.Text
    cPessoa.id_Cidade = cboCidade.ItemData2
    cPessoa.tp_Cliente = IIf(chkCliente.Value, "S", "N")
    cPessoa.tp_Fornecedor = IIf(chkFornecedor.Value, "S", "N")
    cPessoa.tp_Funcionario = IIf(chkFuncionario.Value, "S", "N")

    CarregarPropriedades = True
    
    Exit Function
err_CarregarPropriedades:
    ShowError
End Function

Private Sub cmdNovo_Click()
    Call LimparControles(Me)
    Set cPessoa = New clsPessoa
    id_Pessoa = 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub
Public Sub AtualizarDados(id_Pessoa As Long, id_PessoaFuncionario As Long)
    'id_Pessoa = id_Pessoa
    'id_PessoaFuncionario = id_PessoaFuncionario
    cPessoa.CarregarDados (id_Pessoa)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cPessoa = Nothing
    Set Me = Nothing
End Sub

