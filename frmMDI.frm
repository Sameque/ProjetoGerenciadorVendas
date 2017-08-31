VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "Gerenciador de Estoque"
   ClientHeight    =   7740
   ClientLeft      =   2115
   ClientTop       =   2235
   ClientWidth     =   10155
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10155
      _Version        =   65536
      _ExtentX        =   17912
      _ExtentY        =   714
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox picSetaBaixo 
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   3615
         Picture         =   "frmMDI.frx":000C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   2
         Top             =   45
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox picSetaCima 
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   4260
         Picture         =   "frmMDI.frx":0457
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   1
         Top             =   30
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label lblUsuario 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   90
         Width           =   540
      End
      Begin VB.Label lblSenha 
         Caption         =   "Senha"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   60
         Width           =   780
      End
      Begin VB.Label lblIDCliente 
         Caption         =   "IdCliente"
         Height          =   255
         Left            =   2250
         TabIndex        =   3
         Top             =   75
         Width           =   780
      End
   End
   Begin ComctlLib.StatusBar StatusBarMDI 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   6
      Top             =   7515
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   397
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastro"
      Begin VB.Menu mnuCadastroGeraCodigo 
         Caption         =   "&Gera Código"
      End
      Begin VB.Menu mnuCadastroFinanceiro 
         Caption         =   "&Financeiro"
         Begin VB.Menu mnuCadastroTipoBaixa 
            Caption         =   "&Baixa"
         End
         Begin VB.Menu mnuCadastroNatureza 
            Caption         =   "&Natureza"
         End
      End
      Begin VB.Menu mnuCadastroOperacional 
         Caption         =   "&Operacional"
         Begin VB.Menu mnuCadastroNaturezaProduto 
            Caption         =   "&Natureza de Produtos"
         End
         Begin VB.Menu mnuCadastroProduto 
            Caption         =   "&Produto"
         End
      End
      Begin VB.Menu mnuCadastroPessoal 
         Caption         =   "&Pessoa"
         Begin VB.Menu mnuCadastroFuncao 
            Caption         =   "&Funcao"
         End
         Begin VB.Menu mnuCadastroPessoa 
            Caption         =   "&Pessoa"
         End
      End
   End
   Begin VB.Menu mnuOperacional 
      Caption         =   "&Operacional"
      Begin VB.Menu mnuFrete 
         Caption         =   "&Frete"
      End
      Begin VB.Menu mnuNomenclaturaCampos 
         Caption         =   "&Nomenclatura Campos"
      End
      Begin VB.Menu mnuPedidoCompra 
         Caption         =   "&Pedido de Compra"
      End
      Begin VB.Menu mnuVenda 
         Caption         =   "&Venda de Produto"
      End
   End
   Begin VB.Menu mnuTeste 
      Caption         =   "&Teste"
      Begin VB.Menu menuTesteClacce 
         Caption         =   "&Classe"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bLoad As Boolean
Public ds_Senha As String
Public ds_Usuario As String

Private Sub MDIForm_Activate()
On Error GoTo err_MDIForm_Activate
    
    If Not bLoad Then
        If Trim(lblUsuario.Caption) = "" Then
            frmLogin.Show vbModal
        End If
        
        frmMDI.Caption = "Gerenciador de Vendas"
    End If
    
    bLoad = True
    
    Exit Sub
err_MDIForm_Activate:
    ShowError
End Sub


Private Sub MDIForm_Load()
    sPathReport = "C:\Projeto\Relatórios\"
    sIniFile = "isl.ini"
    If Not LoginPrincipal(Me, False) Then
        End
    End If
End Sub

Private Sub MDIForm_Terminate()
    cnAdo.Close
End Sub

Private Sub mnuSair_Click()
    cnAdo.Close
    End
End Sub

Private Sub mnuFuncao_Click()
    frmPessoaFuncao.Show
End Sub

Private Sub mnuGeraCodigo_Click()
    frmGerarCodigoVB6.Show
End Sub

Private Sub mnuNatureza_Click()
    frmFinanceiroNatureza.Show
End Sub

Private Sub mnuNaturezaProduto_Click()
    frmProdutoNatureza.Show
End Sub

Private Sub mnuPessoa_Click()
    frmPessoaConsulta.Show
End Sub

Private Sub mnuProduto_Click()
    frmProdutoConsulta.Show
End Sub

Private Sub mnuTipoBaixa_Click()
    frmFinanceiroTipoBaixa.Show
End Sub

Private Sub menuTesteClacce_Click()
    frmTesteClasse.Show
End Sub

Private Sub mnuCadastroFuncao_Click()
    frmPessoaFuncao.Show
End Sub

Private Sub mnuCadastroGeraCodigo_Click()
    frmGerarCodigoVB6.Show
End Sub

Private Sub mnuCadastroNatureza_Click()
    frmProdutoNatureza.Show
End Sub

Private Sub mnuCadastroNaturezaProduto_Click()
    frmFinanceiroTipoBaixa.Show
End Sub

Private Sub mnuCadastroPessoa_Click()
    frmPessoaConsulta.Show
End Sub

Private Sub mnuCadastroProduto_Click()
    frmProdutoConsulta.Show
End Sub

Private Sub mnuCadastroTipoBaixa_Click()
    frmFinanceiroNatureza.Show
End Sub

Private Sub mnuFrete_Click()
    frmFrete.Show
End Sub

Private Sub mnuNomenclaturaCampos_Click()
    frmNomenclaturaCampos.Show
End Sub

Private Sub mnuPedidoCompra_Click()
    frmPedidoCompraConsulta.Show
End Sub

Private Sub mnuVenda_Click()
    frmVendaConsulta.Show
End Sub
