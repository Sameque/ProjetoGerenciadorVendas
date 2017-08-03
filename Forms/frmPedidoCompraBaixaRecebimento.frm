VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmPedidoCompraBaixaRecebimento 
   Caption         =   "Cadastro de Pessoa"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   4290
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   2385
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Excluir o Item Selecionado"
      Top             =   1845
      Width           =   855
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Sair da tela"
      Top             =   1845
      Width           =   855
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   855
      Left            =   1425
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Incluir novo Item"
      Top             =   1845
      Width           =   855
   End
   Begin Threed.SSFrame fraDados 
      Height          =   1650
      Left            =   45
      TabIndex        =   0
      Top             =   120
      Width           =   4170
      _Version        =   65536
      _ExtentX        =   7355
      _ExtentY        =   2910
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
      Begin Transportes.SuperControlNovo mskDataBaixaRecebimento 
         Height          =   510
         Left            =   150
         TabIndex        =   5
         Top             =   300
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   900
         AutoTab         =   0   'False
         ToolTip         =   ""
         Mascara         =   4
         MensagemValidacao=   "a Data"
         Label           =   "Baixa Recebimento"
      End
      Begin Transportes.SuperTextMultiline txtBaixaRecebimento 
         Height          =   645
         Left            =   135
         TabIndex        =   4
         Top             =   870
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   1138
         Label           =   "Comentário"
      End
   End
End
Attribute VB_Name = "frmPedidoCompraBaixaRecebimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public id_Compra As Long
Public formChamador As Form

Private cCompra As clsCompra

Private Sub cmdExcluir_Click()
On Error GoTo err_cmdExcluir_Click

    If Mensagem("Confirma exclusão?", Pergunta) = vbNo Then
        Exit Sub
    End If
    
    Call AbreTransacao
    
    If Not EstornoEstoque Then
        Exit Sub
    End If
    
    cCompra.dt_Entrega = CDateEspecial("")
    cCompra.ds_BaixaRecebimento = ""
    
    If Not cCompra.Gravar Then
        Call VoltaTransacao
        Mensagem "Ocorreu um erro na exclusão.", ErroCritico
        Exit Sub
    End If
    
    Call FechaTransacao
    
    Call formChamador.AtualizarDados(id_Compra)

    Call LimparControles(Me)
    
    cmdGravar.Enabled = IIf(mskDataBaixaRecebimento.ClipText = "", True, False)
    
    Mensagem "Exclusão efetuada.", Informacao

    Exit Sub
err_cmdExcluir_Click:
    ShowError
    Call VoltaTransacao
End Sub

Private Function EstornoEstoque() As Boolean
On Error GoTo err_EstornoEstoque
    Dim cCompraItem As clsCompraItem
    Dim cProdutoLote As clsProdutoLote
    Dim id_ProdutoLote As Long
    
    EstornoEstoque = False
    
    If cCompra.ValidarMovimentoSaida Then
        Mensagem "Esse produto possui registro de venda", Informacao
        Exit Function
    End If
    
    For Each cCompraItem In cCompra.colCompraItem
        Set cProdutoLote = New clsProdutoLote
        id_ProdutoLote = cCompraItem.id_ProdutoLote
                
        Call cProdutoLote.CarregarDados(id_ProdutoLote)
        
        cCompraItem.id_ProdutoLote = 0
        cCompraItem.Gravar

        If Not cProdutoLote.EstornoEntrada Then
            Call VoltaTransacao
            Set cProdutoLote = Nothing
            Exit Function
        End If
        
        Set cProdutoLote = Nothing
        
    Next
    
    EstornoEstoque = True
            
    Exit Function
err_EstornoEstoque:
    ShowError
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frmPedidoCompraBaixaRecebimento = Nothing
End Sub

Private Sub Form_Activate()
    frmMDI.Arrange vbCascade
End Sub

Private Sub Form_Load()
    Call CarregarTela
    Call CenterForm(Me)
End Sub

Private Function CarregarTela()
    Set cCompra = New clsCompra
    
    If cCompra.CarregarDados(id_Compra) Then
        mskDataBaixaRecebimento.Text = CDateEspecial(cCompra.dt_Entrega)
        txtBaixaRecebimento.Text = cCompra.ds_BaixaRecebimento
        cmdGravar.Enabled = IIf(mskDataBaixaRecebimento.ClipText = "", True, False)
    End If
    
End Function
Private Sub cmdSair_Click()
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

    If Not CarregarPropriedades(cCompra) Then
        Exit Sub
    End If

    Call AbreTransacao
    If Not cCompra.Gravar Then
        Call VoltaTransacao
        Mensagem "Ocorreu um erro no processamento.", ErroCritico
        Exit Sub
    End If
        
    If Not EntradaEstoque Then
        Mensagem "Erro ao gravar o produto no estoque!", ErroCritico
    End If
    Call FechaTransacao
    
    cmdGravar.Enabled = IIf(mskDataBaixaRecebimento.ClipText = "", True, False)
    Call Mensagem("Gravação Efetuada", Informacao)
    Call formChamador.AtualizarDados(id_Compra)
    cmdSair.SetFocus
    
    Exit Sub
err_cmdGravar_Click:
    ShowError
End Sub

Private Function EntradaEstoque() As Boolean
On Error GoTo err_EntradaEstoque
    Dim cCompraItem As clsCompraItem
    Dim cProdutoLote As clsProdutoLote
    
    EntradaEstoque = False
    
    For Each cCompraItem In cCompra.colCompraItem
        If cCompra.colCompraItem.Count > 0 Then
            Set cProdutoLote = New clsProdutoLote
            
            cProdutoLote.id_Produto = cCompraItem.id_Produto
            cProdutoLote.qt_Produto = cCompraItem.qt_Produto
            cProdutoLote.vl_Compra = (cCompraItem.vl_Produto / cCompraItem.qt_Produto)
            cProdutoLote.dt_Vencimento = CDateEspecial("")
            cProdutoLote.dt_EntradaEstoque = CDateEspecial(mskDataBaixaRecebimento.Text)
            
            If Not cProdutoLote.Movimentacao(Enum_Entrada, cCompraItem.qt_Produto) Then
                Exit Function
            End If
            
            If Not cProdutoLote.Gravar Then
                Exit Function
            End If
                    
            cCompraItem.id_ProdutoLote = cProdutoLote.id_ProdutoLote
            cCompraItem.Gravar
            Set cProdutoLote = Nothing
        End If
    Next
    
    Set cCompraItem = Nothing
    
    EntradaEstoque = True
        
    Exit Function
err_EntradaEstoque:
    ShowError
End Function

Private Function CarregarPropriedades(cCompra As clsCompra) As Boolean
On Error GoTo err_CarregarPropriedades

    CarregarPropriedades = False
    
    cCompra.id_Compra = id_Compra
    cCompra.ds_BaixaRecebimento = txtBaixaRecebimento
    cCompra.dt_Entrega = CDateEspecial(mskDataBaixaRecebimento.Text)
        
    CarregarPropriedades = True
    
    Exit Function
err_CarregarPropriedades:
    ShowError
End Function
