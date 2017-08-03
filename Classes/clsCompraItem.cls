VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCompraItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id_CompraItem As Long
Public id_Compra As Long
Public id_Produto As Long
Public id_ProdutoLote As Long
Public qt_Produto As Long
Public vl_Produto As Double
Public kg_CompraItem As Double

Public cProduto As clsProduto
Public cProdutoLote As clsProdutoLote

Public Function Gravar() As Boolean
On Error GoTo err_Gravar
    
    If id_CompraItem > 0 Then
        Gravar = Alterar
    Else
        Gravar = Incluir
    End If
    
    Exit Function
err_Gravar:
    ShowError
End Function

Private Function Incluir() As Boolean
On Error GoTo err_Incluir
    
    Incluir = False
    
    id_CompraItem = Insert_Table("tbdCompraItem", "id_CompraItem", Campos, Conteudo)
    
    If id_CompraItem = -1 Then
        Exit Function
    End If
    
    Incluir = True
    
    Exit Function
err_Incluir:
    ShowError
End Function

Private Function Alterar() As Boolean
On Error GoTo err_Alterar

    Alterar = False
    
    If Update_Table("tbdCompraItem", Campos, Conteudo, "id_CompraItem = " & id_CompraItem) = -1 Then
        Exit Function
    End If
    
    Alterar = True

    Exit Function
err_Alterar:
    ShowError
End Function

Public Function Excluir() As Boolean
On Error GoTo err_Excluir

    Excluir = False
    
    If VerificarSaidaEstoque Then
        Exit Function
    End If
    
    If Delete_Table("tbdCompraItem", "id_CompraItem = " & id_CompraItem) = -1 Then
        Exit Function
    End If

    Excluir = True

    Exit Function
err_Excluir:
    ShowError
End Function
Private Function VerificarSaidaEstoque() As Boolean
    
    VerificarSaidaEstoque = False
    
    Dim cMovimentacaoEstoque  As clsMovimentacaoEstoque
    
    For Each cMovimentacaoEstoque In cProdutoLote.colMovimentacaoEstoque
         If cMovimentacaoEstoque.tp_MovimentacaoEstoque = "S" Then
            Exit Function
         End If
    Next
    
End Function
Private Function VerificarEntradaEstoque() As Boolean
On Error GoTo err_VerificarEntradaEstoque
    
    VerificarEntradaEstoque = False
    
    If id_ProdutoLote <= 0 Then
        Exit Function
    End If
    
    VerificarEntradaEstoque = True
    
    Exit Function
err_VerificarEntradaEstoque:
    ShowError
End Function
Private Function Campos() As String
    Campos = "id_Compra, id_Produto, id_ProdutoLote, qt_produto, vl_Produto, kg_CompraItem"
End Function

Private Function Conteudo()
    Conteudo = Array(id_Compra, id_Produto, id_ProdutoLote, qt_Produto, vl_Produto, kg_CompraItem)
End Function


Public Function CarregarDados(id_Pesquisa As Long) As Boolean
On Error GoTo err_CarregarDados

    Dim rsado As New ADODB.Recordset
    Dim sTabela As String
    Dim sCampos As String
    Set cProduto = New clsProduto
    
    CarregarDados = False
    Call Class_Initialize
    
    sTabela = "tbdCompraItem"
    sCampos = "id_CompraItem, id_Compra, id_Produto, id_ProdutoLote, qt_produto, vl_Produto, kg_CompraItem"
    
    Call Select_Table(True, sTabela, sCampos, "id_CompraItem = " & id_Pesquisa, , , , , rsado)
    If Not rsado.EOF Then
        CarregarDados = True
        
        id_CompraItem = ReadField(rsado, "id_CompraItem")
        id_Compra = ReadField(rsado, "id_Compra")
        id_Produto = ReadField(rsado, "id_Produto")
        id_ProdutoLote = ReadField(rsado, "id_ProdutoLote")
        qt_Produto = ReadField(rsado, "qt_Produto")
        vl_Produto = ReadField(rsado, "vl_Produto")
        kg_CompraItem = ReadField(rsado, "kg_CompraItem")
                
        Call cProduto.CarregarDados(id_Produto)
        Call cProdutoLote.CarregarDados(id_ProdutoLote)
    End If
    rsado.Close
    
    If Not CarregarDados Then
        Exit Function
    End If
    
    Exit Function
err_CarregarDados:
    ShowError
End Function

Private Sub Class_Initialize()
On Error GoTo err_Class_Initialize

    id_CompraItem = 0
    id_Compra = 0
    id_Produto = 0
    qt_Produto = 0
    vl_Produto = 0
    kg_CompraItem = 0
    id_ProdutoLote = 0
    Set cProduto = New clsProduto
    Set cProdutoLote = New clsProdutoLote
    
    Exit Sub
err_Class_Initialize:
    ShowError
End Sub
