VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmVendaDados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Vendas"
   ClientHeight    =   6810
   ClientLeft      =   3030
   ClientTop       =   4905
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   7875
   Begin Transportes.SuperSpreadNovo sprItens 
      Height          =   2625
      Left            =   45
      TabIndex        =   15
      Top             =   3225
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   4630
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Height          =   750
      HelpContextID   =   23
      Left            =   6075
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Novo lançamento "
      Top             =   5940
      Width           =   810
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   750
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Gravar os Dados"
      Top             =   5955
      Width           =   810
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   750
      Left            =   6990
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Sair da tela"
      Top             =   5940
      Width           =   810
   End
   Begin Threed.SSFrame fraDados 
      Height          =   2010
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   7755
      _Version        =   65536
      _ExtentX        =   13679
      _ExtentY        =   3545
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
      Begin Transportes.SuperControlNovo mskValorVenda 
         Height          =   510
         Left            =   150
         TabIndex        =   3
         Top             =   780
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   900
         ToolTip         =   ""
         Label           =   "Valor da Vebda"
      End
      Begin Transportes.SuperControlNovo mskValorTotal 
         Height          =   510
         Left            =   5625
         TabIndex        =   6
         Top             =   780
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   900
         ToolTip         =   ""
         Label           =   "Total"
      End
      Begin Transportes.SuperControlNovo mskValorDesconto 
         Height          =   510
         Left            =   3790
         TabIndex        =   5
         Top             =   780
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   900
         AutoTab         =   0   'False
         ToolTip         =   ""
         Mascara         =   5
         Label           =   "Desconto"
      End
      Begin Transportes.SuperTextMultiline txtObservacao 
         Height          =   630
         Left            =   120
         TabIndex        =   7
         Top             =   1290
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   1111
         Label           =   "Observação"
      End
      Begin Transportes.SuperControlNovo mskValorFrete 
         Height          =   510
         Left            =   1955
         TabIndex        =   4
         Top             =   780
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   900
         ToolTip         =   ""
         Mascara         =   5
         Label           =   "Frete"
      End
      Begin Transportes.SuperControlNovo mskPrevisaoEntrega 
         Height          =   510
         Left            =   4845
         TabIndex        =   2
         Top             =   225
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   900
         ToolTip         =   ""
         Mascara         =   4
         Label           =   "Prev. Entrega"
      End
      Begin Transportes.SuperDBCombo cboCliente 
         Height          =   510
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   900
         Label           =   "Cliente"
      End
   End
   Begin Threed.SSFrame fraProduto 
      Height          =   960
      Index           =   1
      Left            =   60
      TabIndex        =   14
      Top             =   2115
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
      Begin Transportes.SuperControlNovo mskQuantidadeProduto 
         Height          =   510
         Left            =   5535
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
         TabIndex        =   10
         Top             =   255
         Width           =   810
      End
      Begin Transportes.SuperDBCombo cboProduto 
         Height          =   510
         Left            =   165
         TabIndex        =   8
         Top             =   255
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   900
         MensagemValidacao=   "o Produto"
         Label           =   "Produto"
      End
   End
End
Attribute VB_Name = "frmVendaDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public id_Venda As Long
Public FormChamador As Form

Private cVenda As clsVenda


Private Sub cmdAdicionarProduto_Click()
On Error GoTo err_cmdAdicionarProduto_Click

    If Not ValidarControles(Me, "fraProduto") Then
        Exit Sub
    End If
    
    Call CarregarItemSpread
    Call CarregarItens
    Call CalcularTotalVenda
    Call AtualizarValores
    Call LimparControles(Me, , "fraProduto")
    Call cboProduto.SetFocus
    
    Exit Sub
err_cmdAdicionarProduto_Click:
    ShowError
End Sub

Private Function AtualizarValores()
    On Error GoTo err_AtualizarValores

    mskValorVenda.Text = Format(cVenda.vl_Venda, "###,##0.00")
    mskValorTotal.Text = Format(cVenda.vl_Total, "###,##0.00")
    mskValorDesconto.Text = Format(cVenda.vl_Desconto, "###,##0.00")
    mskValorFrete.Text = Format(cVenda.vl_Frete, "###,##0.00")
    
    Exit Function
err_AtualizarValores:
    ShowError
End Function
Private Function CarregarItemSpread() As Boolean
On Error GoTo err_CarregarItemSpread

    Dim cProdutoLote As New clsProdutoLote
    Dim i As Long

    CarregarItemSpread = False
    
    Call cProdutoLote.CarregarDados(cboProduto.ItemData2)
    
    If cProdutoLote.qt_ProdutoSaldo < Val(mskQuantidadeProduto.Text) Then
        Mensagem "Quantidade insuficiente: " & cProdutoLote.qt_ProdutoSaldo, Informacao
        Exit Function
    End If
    
    For i = 1 To sprItens.MaxRows - 1
        sprItens.Row = i

        If cProdutoLote.id_ProdutoLote = sprItens.TextCol("id_ProdutoLote") Then
            Call AlteraLinhaSpread(i, cProdutoLote)
            Set cProdutoLote = Nothing
            Exit Function
            
        End If
    Next i
    
    
    Call InserirNovaLinhaSpread(cProdutoLote)
    Set cProdutoLote = Nothing

    CarregarItemSpread = True
    
    Exit Function
err_CarregarItemSpread:
    ShowError
    Resume
End Function

Private Function AlteraLinhaSpread(lngLinha As Long, cProdutoLote As clsProdutoLote)
    sprItens.Row = lngLinha
    sprItens.Col = -1
    sprItens.Locked = False

    sprItens.TextCol("qt_Produto") = sprItens.TextCol("qt_Produto") + Val(mskQuantidadeProduto.Text)
    sprItens.TextCol("vl_Produto") = sprItens.TextCol("vl_Produto") + CDbl1(cProdutoLote.vl_Venda)
    sprItens.TextCol("vl_TotalItem") = sprItens.TextCol("vl_TotalItem") + (CDbl1(cProdutoLote.vl_Venda) * Val(mskQuantidadeProduto.Text))
    sprItens.TextCol("kg_CompraItem") = sprItens.TextCol("kg_CompraItem") + CDbl1(Val(mskQuantidadeProduto.Text) * cProdutoLote.cProduto.kg_Real)
    
    sprItens.Col = sprItens.MaxCols
    sprItens.Text = const_Update
    sprItens.Col = 2
    sprItens.Action = 0
   
    sprItens.Row = sprItens.MaxRows - 1
End Function
Private Function InserirNovaLinhaSpread(cProdutoLote As clsProdutoLote)
On Error GoTo err_InserirNovaLinhaSpread
    
    sprItens.Row = sprItens.MaxRows
    sprItens.Col = -1
    sprItens.Locked = False

    sprItens.TextCol("id_VendaItem") = 0
    sprItens.TextCol("id_Venda") = id_Venda
    sprItens.TextCol("id_ProdutoLote") = cProdutoLote.id_ProdutoLote
    sprItens.TextCol("cd_Produto") = cProdutoLote.cProduto.cd_Produto
    sprItens.TextCol("ds_Produto") = cProdutoLote.cProduto.ds_Produto
    sprItens.TextCol("qt_Produto") = Val(mskQuantidadeProduto.Text)
    sprItens.TextCol("vl_Produto") = CDbl1(cProdutoLote.vl_Venda)
    sprItens.TextCol("vl_TotalItem") = (CDbl1(cProdutoLote.vl_Venda) * Val(mskQuantidadeProduto.Text))
    sprItens.TextCol("kg_CompraItem") = CDbl1(Val(mskQuantidadeProduto.Text) * cProdutoLote.cProduto.kg_Real)
    
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
err_InserirNovaLinhaSpread:
    ShowError
End Function
Private Function CarregarItens() As Boolean
On Error GoTo err_CarregarItem
    Dim i As Long
    Dim cVendaItem As clsVendaItem
    
    CarregarItens = False
    Set cVenda.colVendaItem = New Collection
    
    For i = 1 To sprItens.MaxRows - 1
    
        Set cVendaItem = New clsVendaItem
        sprItens.Row = i
        
        Call CarregarClasse(sprItens.SpreadEventoName("id_VendaItem"), id_Venda, sprItens.SpreadEventoName("id_ProdutoLote"), sprItens.SpreadEventoName("qt_Produto"), sprItens.SpreadEventoName("vl_Produto"), sprItens.SpreadEventoName("kg_CompraItem"))
        Call cVenda.AdicionarItem(sprItens.SpreadEventoName("id_VendaItem"), id_Venda, sprItens.SpreadEventoName("qt_Produto"), sprItens.SpreadEventoName("vl_Produto"), sprItens.SpreadEventoName("kg_CompraItem"))
        

        
        
        Set cVendaItem = Nothing
        
    Next i
    
    Set cVendaItem = Nothing
    CarregarItens = True

    Exit Function
err_CarregarItem:
    ShowError
End Function
Private Sub Form_Load()
On Error GoTo err_Form_Load

    Call FormatarComponentes
    Call CarregarComponentes
    Call CenterForm(Me)
    mskValorTotal.Enabled = False
    mskValorVenda.Enabled = False
    
    Exit Sub
err_Form_Load:
    ShowError
End Sub
Private Function FormatarComponentes()

    Call cboCliente.Formatar("id_Pessoa,ds_Pessoa,cd_CNPJCPF", "0,1500,2000", "false,true,true", "tbdPessoa", "tp_Cliente = 'S' ", "ds_Pessoa", 2, 1500)
    Call cboProduto.Formatar("b.id_ProdutoLote,a.ds_Produto,a.cd_Produto,b.id_ProdutoLote,b.qt_ProdutoSaldo", "0,1500,2000,1000,1000", "False,True,True,True,True", "tbdProduto a inner join tbdProdutoLote b on a.id_Produto = b.id_Produto", "", "ds_Produto", 2, 1500, ",Produto,Código,Lote,Saldo")
    
    Call sprItens.NovaColunaSpread(eslNumero, True, False, "id_VendaItem", , 0, 10)
    Call sprItens.NovaColunaSpread(eslNumero, True, False, "id_Venda", , 0, 10)
    Call sprItens.NovaColunaSpread(eslNumero, True, False, "id_ProdutoLote", , 0, 10)
    Call sprItens.NovaColunaSpread(eslTexto, True, False, "cd_Produto", "Código", 15, 20)
    Call sprItens.NovaColunaSpread(eslTexto, True, False, "ds_Produto", "Produto", 32, 255)
    Call sprItens.NovaColunaSpread(eslNumero, True, False, "qt_Produto", "Quantidade", 10, 10)
    Call sprItens.NovaColunaSpread(eslValor, True, False, "vl_Produto", "Valor", 10, 10)
    Call sprItens.NovaColunaSpread(eslValor, True, False, "vl_TotalItem", "Total", 10, 10)
    Call sprItens.NovaColunaSpread(eslValor, True, False, "kg_CompraItem", "Kg.", 10, 10)
     
    Call sprItens.FormatarNovo(21)
    
End Function

Private Function CarregarComponentes()
    
    Dim strCamposItem As String
    Dim strTabelas As String
    
    Set cVenda = New clsVenda
    
    If id_Venda > 0 Then
        
        If cVenda.CarregarDados(id_Venda) Then
            
            Call cboCliente.PesquisarCombo(True, cVenda.id_Comprador, "", True)
            mskPrevisaoEntrega.Text = CDateEspecial(cVenda.dt_PrvisaoEmtrega)
            
            Call AtualizarValores
            
            txtObservacao.Text = cVenda.ds_Observacao
            
            strTabelas = " ((tbdVendaItem a " _
            & " LEFT JOIN tbdProdutoLote b ON a.id_ProdutoLote = b.id_ProdutoLote) " _
            & " LEFT JOIN tbdProduto c ON b.id_Produto = c.id_Produto) "
            
            strCamposItem = "a.id_VendaItem, a.id_Venda, b.id_ProdutoLote, c.cd_Produto, c.ds_Produto, a.qt_produto," _
            & "a.vl_Produto, ( isnull(a.vl_Produto,0) * isnull(a.qt_produto,0)) as vl_TotalItem, a.kg_Produto"
            
            Call sprItens.Carregar(Select_Table(False, strTabelas, strCamposItem, "a.id_Venda = " & id_Venda))
            
        End If
        
    End If

End Function

Private Sub Form_Unload(Cancel As Integer)
    Set cVenda = Nothing
    Set frmVendaDados = Nothing
End Sub

Private Sub cmdGravar_Click()
On Error GoTo err_cmdGravar_Click

    If Not ValidarControles(Me, , Array("fraProduto")) Then
        Exit Sub
    End If
    
    If Mensagem("Confirma Gravação?", Pergunta) = vbNo Then
        Exit Sub
    End If

    If Not CarregarPropriedades(cVenda) Then
        Exit Sub
    End If

    Call AbreTransacao
    If Not cVenda.Gravar Then
        Call VoltaTransacao
        Mensagem "Ocorreu um erro no processamento.", ErroCritico
        Exit Sub
    End If
    Call FechaTransacao

    id_Venda = cVenda.id_Venda
    
    Call FormChamador.AtualizarDados(id_Venda)
    
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
    ShowError
    Call FechaTransacao
End Sub
Private Function GravarItens() As Boolean
On Error GoTo err_Gravar
    Dim i As Long
    Dim cVendaItema As New clsVendaItem
    GravarItens = False
    
    For i = 1 To sprItens.MaxRows - 1
    
        sprItens.Row = i
        Call CarregarClasse(cVendaItema)

        If sprItens.StatusGravacao(i) = const_Insert Or sprItens.StatusGravacao(i) = const_Update Then
        
            If Not cVendaItema.Gravar Then
                Mensagem "Erro ao gravar o item:", Informacao
                Exit Function
            End If
            
            Call cVendaItema.cProdutoLote.Movimentacao(Enum_Saida, cVendaItema.qt_Produto)
            Call cVendaItema.cProdutoLote.Gravar
            
            sprItens.TextCol("id_VendaItem") = cVendaItema.id_VendaItem
            
        ElseIf sprItens.StatusGravacao(i) = const_Delete Then
        
            Call cVendaItema.cProdutoLote.Movimentacao(Enum_Entrada, cVendaItema.qt_Produto)
            Call cVendaItema.cProdutoLote.Gravar
            
            If Not cVendaItema.Excluir Then
                Exit Function
            End If
        End If
        
        Call sprItens.Atualizar(i)
    Next i
    
    Set cVendaItema = Nothing
    
    GravarItens = True
    
    Exit Function
err_Gravar:
    ShowError
End Function

Private Function CarregarClasse(cVendaItem As clsVendaItem)
On erro GoTo err_CarregarClasse

    cVendaItem.id_VendaItem = sprItens.SpreadEventoName("id_VendaItem")
    cVendaItem.id_Venda = id_Venda
    cVendaItem.qt_Produto = sprItens.SpreadEventoName("qt_Produto")
    cVendaItem.vl_Produto = sprItens.SpreadEventoName("vl_Produto")
    cVendaItem.kg_Produto = sprItens.SpreadEventoName("kg_CompraItem")
    Call cVendaItem.cProdutoLote.CarregarDados(sprItens.SpreadEventoName("id_ProdutoLote"))
         
    Exit Function
err_CarregarClasse:
    ShowError
End Function
Private Function CarregarPropriedades(cVenda As clsVenda) As Boolean
On Error GoTo err_CarregarPropriedades

    CarregarPropriedades = False

    cVenda.id_Venda = id_Venda
    cVenda.id_Comprador = cboCliente.ItemData2
    cVenda.id_Vendedor = 0
    cVenda.dt_PrvisaoEmtrega = CDateEspecial(mskPrevisaoEntrega.Text)
    cVenda.dt_Venda = DataAtual(True, False)
    cVenda.vl_Venda = CDbl1(mskValorVenda.Text)
    cVenda.vl_Frete = CDbl1(mskValorFrete.Text)
    cVenda.vl_Desconto = CDbl1(mskValorDesconto.Text)
    cVenda.vl_Total = CDbl1(mskValorTotal.Text)
    cVenda.ds_Observacao = txtObservacao.Text
        
    CarregarPropriedades = True
    
    Exit Function
err_CarregarPropriedades:
    ShowError
End Function

Private Sub cmdNovo_Click()
    Call LimparControles(Me)
    Set cVenda = New clsVenda
    id_Venda = 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Function CalcularTotalVenda()
    Call cVenda.CalcularValorTotal
End Function

Private Sub mskValorDesconto_LostFocus()
    cVenda.vl_Desconto = CDbl1(mskValorDesconto.Text)
    Call cVenda.CalcularValorTotal
    Call AtualizarValores
End Sub

Private Sub mskValorFrete_LostFocus()
    cVenda.vl_Frete = CDbl1(mskValorFrete.Text)
    Call cVenda.CalcularValorTotal
    Call AtualizarValores
End Sub

Private Sub sprItens_LostFocus()
    
    Dim i As Long

    For i = 1 To sprItens.MaxRows - 1
    
        sprItens.Row = i

        If sprItens.StatusGravacao(i) = const_Delete Then
            Call SubtrairValorItem(i)
        End If
    Next i
    
End Sub
Private Function SubtrairValorItem(lngLinha As Long)

    cVenda.vl_Venda = cVenda.vl_Venda - (sprItens.SpreadEventoName("vl_Produto") * sprItens.SpreadEventoName("qt_Produto"))
    cVenda.vl_Total = cVenda.vl_Total - (sprItens.SpreadEventoName("vl_Produto") * sprItens.SpreadEventoName("qt_Produto"))
    Call AtualizarValores

End Function
