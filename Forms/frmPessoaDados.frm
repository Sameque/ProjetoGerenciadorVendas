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
      TabIndex        =   21
      Top             =   2550
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   2990
      Label           =   "Contatos"
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Height          =   750
      HelpContextID   =   23
      Left            =   6675
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Novo lançamento "
      Top             =   4425
      Width           =   810
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   750
      Left            =   5805
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Gravar os Dados"
      Top             =   4425
      Width           =   810
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   750
      Left            =   7545
      Style           =   1  'Graphical
      TabIndex        =   14
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
      Begin VB.CommandButton cmdFuncionario 
         Caption         =   "Funcionário"
         Height          =   345
         Left            =   3675
         TabIndex        =   11
         Top             =   1995
         Width           =   1560
      End
      Begin Transportes.SuperControlNovo mskCEP 
         Height          =   315
         Left            =   2380
         TabIndex        =   4
         Top             =   1065
         Width           =   1740
         _ExtentX        =   2540
         _ExtentY        =   556
         ToolTip         =   ""
         MensagemValidacao=   "o CEP"
      End
      Begin Transportes.SuperText txtCNPJCPF 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1065
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   556
         MensagemValidacao=   "CNPJ ou CPF"
      End
      Begin Transportes.SuperText txtBairro 
         Height          =   315
         Left            =   4155
         TabIndex        =   5
         Top             =   1050
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   556
         MensagemValidacao=   "o Bairro"
      End
      Begin Transportes.SuperText txtEndereco 
         Height          =   315
         Left            =   135
         TabIndex        =   6
         Top             =   1620
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         MensagemValidacao=   "o Endereço"
      End
      Begin Transportes.SuperText txtPessoa 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   525
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   556
         MensagemValidacao=   "o Nome"
      End
      Begin Transportes.SuperText txtRazaoSocial 
         Height          =   315
         Left            =   4155
         TabIndex        =   2
         Top             =   525
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         MensagemValidacao=   "a Razão Social"
      End
      Begin Transportes.SuperDBCombo cboCidade 
         Height          =   510
         Left            =   4155
         TabIndex        =   7
         Top             =   1425
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   900
         MensagemValidacao=   "a Cidade"
         Label           =   "Cidade"
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
         TabIndex        =   15
         Top             =   885
         Width           =   1740
      End
      Begin VB.Label lblcnpjcpf 
         Caption         =   "CNPJ/CPF"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   870
         Width           =   2220
      End
      Begin VB.Label lblBairro 
         Caption         =   "Bairro"
         Height          =   195
         Left            =   4155
         TabIndex        =   17
         Top             =   855
         Width           =   4000
      End
      Begin VB.Label lblEndereco 
         Caption         =   "Endereco"
         Height          =   195
         Left            =   135
         TabIndex        =   18
         Top             =   1395
         Width           =   4000
      End
      Begin VB.Label lblPessoa 
         Caption         =   "Nome"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   3975
      End
      Begin VB.Label lblRazaoSocial 
         Caption         =   "Razão Social"
         Height          =   195
         Left            =   4155
         TabIndex        =   20
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
Public id_PessoaFuncionario As Long

Public FormChamador As Form

Dim cPessoa As clsPessoa

Private Sub chkFuncionario_LostFocus()
    If HabilitarCmdFuncionario Then
        cmdFuncionario.SetFocus
    End If
End Sub

Private Sub cmdFuncionario_Click()

    If id_Pessoa <= 0 Then
        If Mensagem("Antes de gravar os dados do funcionário é preciso gravar os dados do cadastro." _
                  & "Confirma Gravação?", Pergunta) = vbNo Then
            Exit Sub
        End If
        
        Call GravarDados
        
        Call Mensagem("Gravação Efetuada", Informacao)
        'id_PessoaFuncionario = cPessoa.id_PessoaFuncionario
    End If
    
    Set frmPessoaFuncionario.FormChamador = Me
'    frmPessoaFuncionario.id_PessoaFuncionario = cPessoa.id_PessoaFuncionario
    frmPessoaFuncionario.id_Pessoa = id_Pessoa
    frmPessoaFuncionario.Show vbModal
    
    Call CarregarTela
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cPessoa = Nothing
    Set frmPessoaDados = Nothing
End Sub

Private Sub Form_Load()
On Error GoTo err_Form_Load
    
    Call CarregarTela
    
    Call CenterForm(Me)
    
    Exit Sub
err_Form_Load:
    ShowError
End Sub
Private Function CarregarTela()
    Set cPessoa = New clsPessoa
    
    Call cboCidade.Formatar("a.id_Cidade, a.ds_Cidade, b.ds_Estado", "0,3000,1000", "false,true,true", "tbdCidade a left join tbdEstado b on a.id_Estado = b.id_Estado", "", "ds_Cidade")

    If id_Pessoa > 0 Then
        If cPessoa.CarregarDados(id_Pessoa) Then
            mskCEP.Text = cPessoa.cd_CEP
            txtCNPJCPF.Text = cPessoa.cd_cnpjcpf
            txtBairro.Text = cPessoa.ds_Bairro
            txtEndereco.Text = cPessoa.ds_Endereco
            txtPessoa.Text = cPessoa.ds_Pessoa
            txtRazaoSocial.Text = cPessoa.ds_RazaoSocial
            chkCliente = IIf(cPessoa.tp_Cliente = "S", True, False)
            chkFornecedor = IIf(cPessoa.tp_Fornecedor = "S", True, False)
            chkFuncionario = IIf(cPessoa.tp_Funcionario = "S", True, False)
            id_PessoaFuncionario = cPessoa.id_PessoaFuncionario
            
            Call HabilitarCmdFuncionario
            Call cboCidade.PesquisarCombo(True, cPessoa.id_Cidade, "", True)

        End If
        
    End If
    
    
    Call CarregarSpread
    
    mskCEP.Mascara = CEP
    cmdFuncionario.Enabled = False
    txtPessoa.CampoObrigatorio = True
    txtRazaoSocial.CampoObrigatorio = True
    txtCNPJCPF.CampoObrigatorio = True


End Function

Private Function CarregarSpread()
On Error GoTo err_CarregarSpread
    
    Dim sTabelas  As String
    Dim sCampos As String
    
    sCampos = "ds_Nome, cd_Fone, cd_Email"
    sTabelas = "tbdPessoaContato"

    Call sprContato.NovaColunaSpread(eslTexto, False, False, "ds_Nome", "Nome", 50, 255, , , , False)
    Call sprContato.NovaColunaSpread(eslTexto, False, False, "cd_Fone", "Fone", 30, 30, , , , False)
    Call sprContato.NovaColunaSpread(eslTexto, False, False, "cd_Email", "Email", 30, 255, , , , False)

    Call sprContato.FormatarNovo(21)
     

    
    Call sprContato.Carregar(Select_Table(False, sTabelas, sCampos, "id_Pessoa = " & id_Pessoa, "ds_Nome"))
    
err_CarregarSpread:
    ShowError
End Function
Private Function HabilitarCmdFuncionario() As Boolean
    'cmdFuncionario.Enabled = chkFuncionario
    'HabilitarCmdFuncionario = chkFuncionario
    cmdFuncionario.Enabled = False
    HabilitarCmdFuncionario = False
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

