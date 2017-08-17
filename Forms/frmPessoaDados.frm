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
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   750
      Left            =   4920
      TabIndex        =   21
      Top             =   4425
      Width           =   810
   End
   Begin Transportes.SuperSpreadNovo sprContato 
      Height          =   1695
      Left            =   60
      TabIndex        =   20
      Top             =   2550
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   2990
      BackColorCellAtiva=   14733514
      GrayAreaBackColor=   14670555
      SkinESL         =   -1  'True
      Label           =   "Contatos"
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
         Height          =   285
         Left            =   2385
         TabIndex        =   4
         Top             =   1065
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   503
         ToolTip         =   ""
         BackColor       =   16119285
         Mascara         =   7
         MensagemValidacao=   "o CEP"
         SkinESL         =   -1  'True
      End
      Begin Transportes.SuperText txtCNPJCPF 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1065
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   503
         SkinESL         =   -1  'True
         BackColor       =   16119285
         MensagemValidacao=   "CNPJ ou CPF"
         CampoObrigatorio=   -1  'True
      End
      Begin Transportes.SuperText txtBairro 
         Height          =   285
         Left            =   4155
         TabIndex        =   5
         Top             =   1050
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         SkinESL         =   -1  'True
         BackColor       =   16119285
         MensagemValidacao=   "o Bairro"
      End
      Begin Transportes.SuperText txtEndereco 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1620
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         SkinESL         =   -1  'True
         BackColor       =   16119285
         MensagemValidacao=   "o Endereço"
      End
      Begin Transportes.SuperText txtPessoa 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   525
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         SkinESL         =   -1  'True
         BackColor       =   16119285
         MensagemValidacao=   "o Nome"
         CampoObrigatorio=   -1  'True
      End
      Begin Transportes.SuperText txtRazaoSocial 
         Height          =   285
         Left            =   4155
         TabIndex        =   2
         Top             =   525
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         SkinESL         =   -1  'True
         BackColor       =   16119285
         MensagemValidacao=   "a Razão Social"
         CampoObrigatorio=   -1  'True
      End
      Begin Transportes.SuperDBCombo cboCidade 
         Height          =   480
         Left            =   4155
         TabIndex        =   7
         Top             =   1425
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   847
         SkinESL         =   -1  'True
         BackColor       =   16119285
         BackColorControl=   16119285
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

Public mlngPessoa As Long
Public FormChamador As Form
Public mcPessoa As New clsPessoa
Private mstrMensagemRetorno As String

Private Sub cboCidade_Change()

End Sub

Private Sub cmdExcluir_Click()
    Call Excluir
End Sub

Private Sub Form_Load()
On Error GoTo err_Form_Load
    
    Call FormatarComponentes
    Call CarregarComponentes(mlngPessoa)
    Call CenterForm(Me)
    
    Exit Sub
err_Form_Load:
    ShowError
End Sub

Private Sub cmdGravar_Click()
On Error GoTo err_cmdGravar_Click

    If Not ValidarControlesNovo(Me) Then
        Exit Sub
    End If
    
    If Mensagem("Confirma Gravação?", Pergunta) = vbNo Then
        Exit Sub
    End If
    
    If Not CarregarPropriedadesPessoa() Then
        Mensagem mstrMensagemRetorno, erro
        Exit Sub
    End If
    
    If Not CarregarPropriedadesContatos() Then
        Mensagem mstrMensagemRetorno, erro
        Exit Sub
    End If
    
    If Not GravarDados Then
        Mensagem mstrMensagemRetorno, erro
        Exit Sub
    End If
    
    Call sprContato.AtualizarStatus(mcPessoa.GetListaContatos)
    
    Call Mensagem("Gravação Efetuada", Informacao)
    
    cmdNovo.SetFocus

    Exit Sub
err_cmdGravar_Click:
    ShowError
    
End Sub

Private Sub cmdNovo_Click()
    Call LimparControles(Me)
    Set mcPessoa = Nothing
    mlngPessoa = 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub txtCNPJCPF_LostFocus()
    Call CarregarPorCNPJ
End Sub

Private Function CarregarComponentes(ByVal lngPesquisa As Long, Optional ByVal strCNPJ As String = "")
On Error GoTo err_CarregarCampos
    Dim cServicoPessoa As New clsServicoPessoa
        
    If strCNPJ <> "" Then
        Set mcPessoa = cServicoPessoa.CarregarPorCNPJ(strCNPJ, True)
    Else
        Set mcPessoa = cServicoPessoa.CarregarPorID(lngPesquisa, True)
    End If
    
    If mcPessoa.id_Pessoa <= 0 Then
        Exit Function
    End If

    Call LimparControles(Me)
    
    With mcPessoa
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
    
    Call sprContato.CarregarPorClasse("id_Pessoa = " & mcPessoa.id_Pessoa)
    
    Exit Function
err_CarregarCampos:
    ShowError
End Function

Private Function CarregarPropriedadesContatos() As Boolean
On Error GoTo err_CarregarContatos
    Dim i As Long
    
    CarregarPropriedadesContatos = False
    Call mcPessoa.ZerarListas

    For i = 1 To sprContato.MaxRows - 1
        sprContato.Row = i
                
        If sprContato.StatusGravacao(i) <> const_Inicial Then
            Call mcPessoa.AdicionarContato( _
                    sprContato.SpreadEventoName("id_PessoaContato") _
                    , sprContato.SpreadEventoName("ds_Nome") _
                    , sprContato.SpreadEventoName("cd_Fone") _
                    , sprContato.SpreadEventoName("cd_Email") _
                    , sprContato.StatusGravacao(i))
        End If
    Next i
    
    CarregarPropriedadesContatos = True
    
    Exit Function
err_CarregarContatos:
    mstrMensagemRetorno = "Erro ao carregar contato."
End Function
Private Function GravarDados() As Boolean
On Error GoTo err_GravarDados
    Dim cServicoPessoa As New clsServicoPessoa
    
    GravarDados = False
    
    Call AbreTransacao
    If Not cServicoPessoa.Salvar(mcPessoa) Then
        Call VoltaTransacao
        Mensagem cServicoPessoa.mstrMensagemRetorno, erro
        Exit Function
    End If
    Call FechaTransacao
    
    mlngPessoa = mcPessoa.id_Pessoa
    
    GravarDados = True
    Set cServicoPessoa = Nothing
    
    Exit Function
err_GravarDados:
    ShowError
    Call FechaTransacao
End Function
Private Function CarregarPropriedadesPessoa() As Boolean
On Error GoTo err_CarregarPropriedades

    CarregarPropriedadesPessoa = False
    If True Then
        With mcPessoa
            .id_Pessoa = mlngPessoa
            .cd_CEP = mskCEP.Text
            .cd_cnpjcpf = txtCNPJCPF.Text
            .ds_Bairro = txtBairro.Text
            .ds_Endereco = txtEndereco.Text
            .ds_Pessoa = txtPessoa.Text
            .ds_RazaoSocial = txtRazaoSocial.Text
            .tp_Cliente = IIf(chkCliente, "S", "N")
            .tp_Fornecedor = IIf(chkFornecedor, "S", "N")
            .tp_Funcionario = IIf(chkFuncionario, "S", "N")
            .id_Cidade = cboCidade.ItemData2
            .menumStatusGravacao = EnumStatusGravacao.IncluirOuAlterar
        End With
    End If
    CarregarPropriedadesPessoa = True
    
    Exit Function
err_CarregarPropriedades:
    mstrMensagemRetorno = "Erro ao carregar propriedades."
End Function
Private Function CarregarPorCNPJ()
On Error GoTo err_CarregarPorCNPJ

    If txtCNPJCPF.Text <> "" Then
        Call CarregarComponentes(mlngPessoa, txtCNPJCPF.Text)
    End If
    
    Exit Function
err_CarregarPorCNPJ:
    ShowError
End Function
Private Function Excluir() As Boolean
On Error GoTo err_Excluir
    Dim cServicoPessoa As New clsServicoPessoa
    
    Excluir = False
    
    Call AbreTransacao
    If Not cServicoPessoa.Excluir(mcPessoa) Then
        Call VoltaTransacao
        Mensagem cServicoPessoa.mstrMensagemRetorno, erro
        Exit Function
    End If
    Call FechaTransacao
    
    Call LimparControles(Me)
    Excluir = True
    Set cServicoPessoa = Nothing
    Set mcPessoa = Nothing
    
    Exit Function
err_Excluir:
    ShowError
End Function
Private Function FormatarComponentes()
On Error GoTo err_FormatarComponentes
    Dim cServicoPessoa As New clsServicoPessoa
            
    Call cboCidade.Formatar("a.id_Cidade, a.ds_Cidade, b.ds_Estado", "0,3000,1000", "false,true,true", "tbdCidade a left join tbdEstado b on a.id_Estado = b.id_Estado", "", "ds_Cidade")
    Call sprContato.FormatarPorClasse(cServicoPessoa.FormatarSpreadPessoaContato)
    
    Exit Function
err_FormatarComponentes:
    ShowError
End Function
Private Sub Form_Unload(Cancel As Integer)
    Set mcPessoa = Nothing
End Sub
