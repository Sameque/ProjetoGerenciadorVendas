VERSION 5.00
Begin VB.Form frmProdutoItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Produtos"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   615
      Left            =   5415
      TabIndex        =   1
      Top             =   2265
      Width           =   975
   End
   Begin Transportes.SuperSpreadNovo sprItens 
      Height          =   2010
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   3545
   End
End
Attribute VB_Name = "frmProdutoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public id_Venda As Long

Dim sTabela As String
Dim sCampos As String
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    frmMDI.Arrange vbCascade
End Sub

Private Sub Form_Load()
On Error GoTo err_Form_Load

    Call IniciarComponentes
    Call CarregarComponentes
    
    Exit Sub
err_Form_Load:
    ShowError
End Sub

Private Function IniciarComponentes()

    Call sprItens.NovaColunaSpread(eslTexto, True, True, "prod.cd_Produto", "Código", 10, 50)
    Call sprItens.NovaColunaSpread(eslTexto, True, False, "prod.ds_Produto", "Descrição", 20, 50)
    Call sprItens.NovaColunaSpread(eslNumero, True, False, "item.qt_produto", "Qtd.", 10, 10)
    Call sprItens.NovaColunaSpread(eslValor, True, True, "item.vl_Produto", "Valor", 10, 10)
    Call sprItens.NovaColunaSpread(eslValor, True, True, "kg_CompraItem", "Kg", 10, 10)
    
    Call sprItens.FormatarNovo(21)
    
End Function
Private Function CarregarComponentes()
    
    sTabela = ""

    sTabela = " ((tbdVendaItem item "
    sTabela = sTabela & " LEFT JOIN tbdProdutoLote lote ON item.id_ProdutoLote = lote.id_ProdutoLote) " _
    & " LEFT JOIN tbdProduto prod ON lote.id_Produto = prod.id_Produto) "
    
    sCampos = "prod.cd_Produto, prod.ds_Produto, item.qt_produto, (isnull(item.vl_Produto,0)*isnull(item.qt_produto,0)), item.kg_Produto"
    
   Call sprItens.Carregar(Select_Table(False, sTabela, sCampos, "item.id_Venda = " & id_Venda, "prod.ds_Produto"))

End Function
