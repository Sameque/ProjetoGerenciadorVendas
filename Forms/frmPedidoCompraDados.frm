VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmPedidoCompraDados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Pedido de Compra"
   ClientHeight    =   6810
   ClientLeft      =   3030
   ClientTop       =   4905
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   7875
   Begin Transportes.SuperSpreadNovo sprItens 
      Height          =   3405
      Left            =   90
      TabIndex        =   16
      Top             =   2460
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   6006
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Height          =   750
      HelpContextID   =   23
      Left            =   6075
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Novo lançamento "
      Top             =   5940
      Width           =   810
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   750
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Gravar os Dados"
      Top             =   5955
      Width           =   810
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   750
      Left            =   6990
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Sair da tela"
      Top             =   5940
      Width           =   810
   End
   Begin Threed.SSFrame fraDados 
      Height          =   1335
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   7755
      _Version        =   65536
      _ExtentX        =   13679
      _ExtentY        =   2355
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
      Begin Transportes.SuperTextMultiline txtObservacao 
         Height          =   510
         Left            =   4785
         TabIndex        =   7
         Top             =   735
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   900
         Label           =   "Observação"
      End
      Begin Transportes.SuperControlNovo mskValorCompra 
         Height          =   510
         Left            =   3225
         TabIndex        =   6
         Top             =   735
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   900
         ToolTip         =   ""
         Mascara         =   5
         Label           =   "Valor Total"
      End
      Begin Transportes.SuperControlNovo mskValorAdicional 
         Height          =   510
         Left            =   1680
         TabIndex        =   5
         Top             =   735
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   900
         ToolTip         =   ""
         Mascara         =   5
         Label           =   "Valor Adicional"
      End
      Begin Transportes.SuperControlNovo mskValorFrete 
         Height          =   510
         Left            =   135
         TabIndex        =   4
         Top             =   735
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   900
         ToolTip         =   ""
         Mascara         =   5
         Label           =   "Valor do Frete"
      End
      Begin Transportes.SuperControlNovo mskPrevisaoEntrega 
         Height          =   510
         Left            =   6270
         TabIndex        =   3
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   900
         ToolTip         =   ""
         Mascara         =   4
         Label           =   "Prev. Entrega"
      End
      Begin Transportes.SuperDBCombo cboFornecedor 
         Height          =   510
         Left            =   150
         TabIndex        =   1
         Top             =   210
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   900
         Label           =   "Fornecedor"
      End
      Begin Transportes.SuperControlNovo mskDataCompra 
         Height          =   510
         Left            =   4830
         TabIndex        =   2
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   900
         ToolTip         =   ""
         Mascara         =   4
         Label           =   "Dt. Compra"
      End
   End
   Begin Threed.SSFrame fraProduto 
      Height          =   960
      Index           =   1
      Left            =   75
      TabIndex        =   15
      Top             =   1410
      Width           =   7755
      _Version        =   65536
      _ExtentX        =   13679
      _ExtentY        =   1693
      _StockProps     =   14
      Caption         =   "Adicionar Produto "
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
      Begin Transportes.SuperControlNovo mskValorProduto 
         Height          =   510
         Left            =   5520
         TabIndex        =   10
         Top             =   255
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   900
         ToolTip         =   ""
         Mascara         =   5
         MensagemValidacao=   "o Valor"
         Label           =   "Valor"
      End
      Begin Transportes.SuperControlNovo mskQuantidadeProduto 
         Height          =   510
         Left            =   4290
         TabIndex        =   9
         Top             =   255
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   900
         ToolTip         =   ""
         Mascara         =   10
         LenDecimal      =   0
         MensagemValidacao=   "a Quantidade"
         Label           =   "Quantidade"
      End
      Begin VB.CommandButton cmdAdicionarProduto 
         Caption         =   "Adicionar Produto"
         Height          =   600
         Left            =   6795
         TabIndex        =   11
         Top             =   255
         Width           =   810
      End
      Begin Transportes.SuperDBCombo cboProduto 
         Height          =   510
         Left            =   165
         TabIndex        =   8
         Top             =   255
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   900
         MensagemValidacao=   "o Produto"
         Label           =   "Produto"
      End
   End
End
Attribute VB_Name = "frmPedidoCompraDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public id_Compra As Long
Public formChamador As Form

Dim cCompra As clsCompra

Private Sub cmdAdicionarProduto_Click()
On Error GoTo err_cmdAdicionarProduto_Click

    If Not ValidarControles(Me, "fraProduto") Then
        Exit Sub
    End If
    
    Call CarregarItemSpread
           
    Call LimparControles(Me, , "fraProduto")
    
    Call cboProduto.SetFocus
    
    Exit Sub
err_cmdAdicionarProduto_Click:
    ShowError
End Sub
Private Function CarregarItemSpread()
On Error GoTo err_CarregarItemSpread
    Dim cProduto As New clsProduto
    
    sprItens.Row = sprItens.MaxRows
    sprItens.Col = -1
    sprItens.Locked = False
    
    Call cProduto.CarregarDados(cboProduto.ItemData2)
    
    sprItens.TextCol("id_CompraItem") = 0
    sprItens.TextCol("id_Compra") = id_Compra
    sprItens.TextCol("id_Produto") = cProduto.id_Produto
    sprItens.TextCol("cd_Produto") = cProduto.cd_Produto
    sprItens.TextCol("ds_Produto") = cProduto.ds_Produto
    sprItens.TextCol("qt_Produto") = Val(mskQuantidadeProduto.Text)
    sprItens.TextCol("vl_Produto") = CDbl1(mskValorProduto.Text)
    
    sprItens.TextCol("kg_CompraItem") = cProduto.kg_Real
    
    
    sprItens.Col = sprItens.MaxCols
    sprItens.Text = const_Insert
    sprItens.Col = 2
    sprItens.Action = 0

    sprItens.MaxRows = sprItens.MaxRows + 1
    sprItens.Row = sprItens.MaxRows
    sprItens.Col = -1
    sprItens.Locked = True
   
    sprItens.Row = sprItens.MaxRows - 1
    
    Exit Function
err_CarregarItemSpread:
    ShowError
End Function
Private Sub Form_Load()
On Error GoTo err_Form_Load

    Call FormatarComponentes
    Call CarregarComponentes
    Call CenterForm(Me)
    Exit Sub
err_Form_Load:
    ShowError
End Sub
Private Function FormatarComponentes()

    Call cboFornecedor.Formatar("id_Pessoa,ds_Pessoa,cd_CNPJCPF", "0,1500,2000", "false,true,true", "tbdPessoa", "tp_Fornecedor = 'S' ", "ds_Pessoa", 2, 1500)
    Call cboProduto.Formatar("id_Produto,ds_Produto,cd_Produto", "0,1500,2000", "False,True,True", "tbdProduto", "", "ds_Produto", 2, 1500)
    
    Call sprItens.NovaColunaSpread(eslNumero, True, False, "id_CompraItem", , 0, 10)
    Call sprItens.NovaColunaSpread(eslNumero, True, False, "id_Compra", , 0, 10)
    Call sprItens.NovaColunaSpread(eslNumero, True, False, "id_Produto", , 0, 10)
    Call sprItens.NovaColunaSpread(eslTexto, True, False, "cd_Produto", "Código", 15, 20)
    Call sprItens.NovaColunaSpread(eslTexto, True, False, "ds_Produto", "Produto", 32, 255)
    Call sprItens.NovaColunaSpread(eslNumero, True, False, "qt_Produto", "Quantidade", 10, 10)
    Call sprItens.NovaColunaSpread(eslValor, True, False, "vl_Produto", "Valor", 10, 10)
    Call sprItens.NovaColunaSpread(eslValor, True, False, "kg_CompraItem", "Kg.", 10, 10)
    
    Call sprItens.FormatarNovo(21)
    
End Function

Private Function CarregarComponentes()
    
    Dim strCamposItem As String
    Dim strTabelas As String
    Set cCompra = New clsCompra
    
    If id_Compra > 0 Then
        
        If cCompra.CarregarDados(id_Compra) Then
            
            Call cboFornecedor.PesquisarCombo(True, cCompra.id_Fornecedor, "", True)
            mskDataCompra.Text = cCompra.dt_Compra
            mskPrevisaoEntrega.Text = cCompra.dt_PrevisaoEntrega
            mskValorFrete.Text = cCompra.vl_Frete
            mskValorAdicional.Text = cCompra.vl_Adicional
            mskValorCompra.Text = cCompra.vl_Compra
            txtObservacao.Text = cCompra.ds_Observacao
                        
            strCamposItem = "a.id_CompraItem, a.id_Compra, a.id_Produto, b.cd_Produto, b.ds_Produto, a.qt_produto, a.vl_Produto, a.kg_CompraItem"
            strTabelas = "tbdCompraItem a LEFT JOIN tbdProduto b ON a.id_Produto = b.id_Produto"
            Call sprItens.Carregar(Select_Table(False, strTabelas, strCamposItem, "a.id_Compra = " & id_Compra))
            
        End If
        
    End If

End Function
Private Sub Form_Unload(Cancel As Integer)
    Set cCompra = Nothing
    Set frmPedidoCompraDados = Nothing
End Sub

Private Sub cmdGravar_Click()
On Error GoTo err_cmdGravar_Click

    If Not ValidarControles(Me, , Array("fraProduto")) Then
        Exit Sub
    End If
    
    If Mensagem("Confirma Gravação?", Pergunta) = vbNo Then
        Exit Sub
    End If

    If Not CarregarPropriedades(cCompra) Then
        Exit Sub
    End If

    Call AbreTransacao
    If Not cCompra.Gravar Then
        Call VoltaTransacao
        Mensagem "Ocorreu um erro no processamento.", ErroCritico
        Exit Sub
    End If
    Call FechaTransacao

    id_Compra = cCompra.id_Compra
    
    Call formChamador.AtualizarDados(id_Compra)
    
    Call AbreTransacao
    If Not GravarItens Then
        Call VoltaTransacao
        Mensagem "Erro ao gravar o(s) produto(s)", Informacao
    End If
    Call FechaTransacao
    
    Call Mensagem("Gravação Efetuada", Informacao)
    
    cmdNovo.SetFocus
    
    Exit Sub
err_cmdGravar_Click:
Resume
    ShowError
    Call FechaTransacao
End Sub
Private Function GravarItens() As Boolean
    On Error GoTo err_Gravar
    Dim i As Long
    Dim cCompraItema As New clsCompraItem
    
    GravarItens = False
    
    For i = 1 To sprItens.MaxRows - 1
    
        sprItens.Row = i
        Call CarregarClasse(cCompraItema)

        If sprItens.StatusGravacao(i) = const_Insert Or sprItens.StatusGravacao(i) = const_Update Then
            If Not cCompraItema.Gravar Then
                Exit Function
            End If
            sprItens.TextCol("id_CompraItem") = cCompraItema.id_CompraItem
        ElseIf sprItens.StatusGravacao(i) = const_Delete Then
            If Not cCompraItema.Excluir Then
                Exit Function
            End If
        End If
        
        Call sprItens.Atualizar(i)
    Next i
    
    Set cCompraItema = Nothing
    
    GravarItens = True
    
    Exit Function
err_Gravar:
    ShowError
End Function

Private Function CarregarClasse(cCompraItema As clsCompraItem)

    cCompraItema.id_CompraItem = sprItens.SpreadEventoName("id_CompraItem")
    cCompraItema.id_Compra = id_Compra
    cCompraItema.id_Produto = sprItens.SpreadEventoName("id_Produto")
    cCompraItema.qt_Produto = sprItens.SpreadEventoName("qt_Produto")
    cCompraItema.vl_Produto = sprItens.SpreadEventoName("vl_Produto")
    cCompraItema.kg_CompraItem = sprItens.SpreadEventoName("kg_CompraItem")
 
End Function
Private Function CarregarPropriedades(cCompra As clsCompra) As Boolean
On Error GoTo err_CarregarPropriedades

    CarregarPropriedades = False
        
    cCompra.id_Compra = id_Compra
    cCompra.id_Fornecedor = cboFornecedor.ItemData2
    cCompra.dt_Compra = CDateEspecial(mskDataCompra.Text)
    cCompra.dt_PrevisaoEntrega = CDateEspecial(mskPrevisaoEntrega.Text)
    cCompra.vl_Compra = CDbl1(mskValorCompra.Text)
    cCompra.vl_Frete = CDbl1(mskValorFrete.Text)
    cCompra.vl_Adicional = CDbl1(mskValorAdicional.Text)
    cCompra.ds_Observacao = txtObservacao.Text
    
    CarregarPropriedades = True
    
    Exit Function
err_CarregarPropriedades:
    ShowError
End Function

Private Sub cmdNovo_Click()
    Call LimparControles(Me)
    Set cCompra = New clsCompra
    id_Compra = 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub
