VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id_Compra As Long
Public id_Fornecedor As Long
Public id_Comprador As Long
Public dt_Compra As Date
Public dt_PrevisaoEntrega As Date
Public dt_Entrega As Date
Public vl_Compra As Double
Public vl_Frete As Double
Public vl_Adicional As Double
Public ds_Observacao As String
Public ds_BaixaRecebimento As String
Public dt_EntradaEstoque As Date

Public colCompraItem As Collection

Public Function ValidarMovimentoSaida() As Boolean
On Error GoTo err_ValidarMovimentoSaida
    Dim cCompraItem As clsCompraItem
    Dim cProdutoLote As clsProdutoLote
        
    ValidarMovimentoSaida = False
    
    For Each cCompraItem In colCompraItem
        Set cProdutoLote = New clsProdutoLote
        Call cProdutoLote.CarregarDados(cCompraItem.id_ProdutoLote)
        
        If cProdutoLote.qt_Produto > cProdutoLote.qt_ProdutoSaldo Then
            ValidarMovimentoSaida = True
            Exit Function
        End If
        Set cProdutoLote = Nothing
    Next
    
    Set cCompraItem = Nothing
    
    Exit Function
err_ValidarMovimentoSaida:
    ShowError
End Function

Public Function Gravar() As Boolean
On Error GoTo err_Gravar
       
    Call AbreTransacao
    If id_Compra > 0 Then
        Gravar = Alterar
    Else
        Gravar = Incluir
    End If
    Call FechaTransacao
        
    Exit Function
err_Gravar:
    ShowError
End Function

Public Function GravarItem() As Boolean
On Error GoTo err_Gravar
    
    For Each cCompraItem In colCompraItem
        'cCompraItem.
    Next
    
    
    
    cCompraItem As New clsCompraItem
    
    Call AbreTransacao
    If id_Compra > 0 Then
        Gravar = Alterar
    Else
        Gravar = Incluir
    End If
    Call FechaTransacao
        
    Exit Function
err_Gravar:
    ShowError
End Function

Private Function Incluir() As Boolean
On Error GoTo err_Incluir
    
    Incluir = False
    
    id_Compra = Insert_Table("tbdCompra", "id_Compra", Campos, Conteudo)
    
    If id_Compra = -1 Then
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
    If Update_Table("tbdCompra", Campos, Conteudo, "id_Compra = " & id_Compra) = -1 Then
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
    
    
    If ValidarMovimentoSaida Then
        Exit Function
    End If
    
    If id_Compra > 0 Then
        If Not ExcluirItem Then
            Mensagem "Erro ao excluir os itens do pedido de compra!", ErroCritico
            Exit Function
        End If
    Else
        Exit Function
    End If
     
    If Delete_Table("tbdCompra", "id_Compra = " & id_Compra) = -1 Then
        Exit Function
    End If

    Excluir = True

    Exit Function
err_Excluir:
    ShowError
End Function
Public Function ExcluirItem() As Boolean
On Error GoTo err_Excluir
    Dim cCompraItem As clsCompraItem
    
    ExcluirItem = False
    
    If Not ExcluirMovimentacoes Then
        Exit Function
    End If
    
'    For Each cCompraItem In colCompraItem
'        If Not cCompraItem.Excluir Then
'            Exit Function
'        End If
'    Next

    ExcluirItem = True
    Set cCompraItem = Nothing
    
    Exit Function
err_Excluir:
    ShowError
End Function

Private Function ExcluirMovimentacoes() As Boolean
    Dim cCompraItem As New clsCompraItem
    Dim cMovimentacaoEstoque As New clsMovimentacaoEstoque
    Dim id_ProdutoLote As Long
    
    ExcluirMovimentacoes = False
    
    For Each cCompraItem In colCompraItem
        For Each cMovimentacaoEstoque In cCompraItem.cProdutoLote.colMovimentacaoEstoque
            Call cMovimentacaoEstoque.Excluir
        Next
        
        cCompraItem.id_ProdutoLote = 0
        cCompraItem.Gravar
        
        Call cCompraItem.cProdutoLote.Excluir
        Call cCompraItem.Excluir
    Next
    
    ExcluirMovimentacoes = True

End Function
Private Function Campos() As String
    Campos = "id_Fornecedor, id_Comprador, dt_Compra, dt_PrevisaoEntrega, dt_Entrega, vl_Compra, vl_Frete, vl_Adicional, ds_observacao, ds_BaixaRecebimento"
End Function

Private Function Conteudo()
    Conteudo = Array(id_Fornecedor, id_Comprador, dt_Compra, dt_PrevisaoEntrega, dt_Entrega, vl_Compra, vl_Frete, vl_Adicional, ds_Observacao, ds_BaixaRecebimento)
End Function

Public Function CarregarDados(id_Pesquisa As Long) As Boolean
On Error GoTo err_CarregarDados

    Dim rsado As New ADODB.Recordset
    Dim sCampos As String
    Dim sTabela As String
    
    CarregarDados = False
    Call Class_Initialize
    
    sCampos = " id_Compra, id_Fornecedor, id_Comprador, dt_Compra, dt_PrevisaoEntrega, dt_Entrega, " _
    & "vl_Compra, vl_Frete, vl_Adicional, ds_observacao, ds_BaixaRecebimento "
    
    sTabela = "tbdCompra"
    
    Call Select_Table(True, sTabela, sCampos, "id_Compra = " & id_Pesquisa, , , , , rsado)
    If Not rsado.EOF Then
        CarregarDados = True
        
        id_Compra = ReadField(rsado, "id_Compra")
        id_Fornecedor = ReadField(rsado, "id_Fornecedor")
        id_Comprador = ReadField(rsado, "id_Comprador")
        
        dt_Compra = CDateEspecial(ReadField(rsado, "dt_Compra"))
        dt_PrevisaoEntrega = CDateEspecial(ReadField(rsado, "dt_PrevisaoEntrega"))
        dt_Entrega = CDateEspecial(ReadField(rsado, "dt_Entrega"))
        
        vl_Compra = ReadField(rsado, "vl_Compra")
        vl_Frete = ReadField(rsado, "vl_Frete")
        vl_Adicional = ReadField(rsado, "vl_Adicional")
        
        ds_Observacao = ReadField(rsado, "ds_observacao")
        ds_BaixaRecebimento = ReadField(rsado, "ds_BaixaRecebimento")
        
        Call CarregarItens
         
    End If
    rsado.Close
    
    If Not CarregarDados Then
        Exit Function
    End If
    
    Exit Function
err_CarregarDados:
    ShowError
End Function
Private Function CarregarItens()
On Error GoTo err_CarregarItens
    Dim cCompraItem As clsCompraItem
    Dim rsado As New ADODB.Recordset
    Dim sCampos As String
    Dim sTabela As String
    
    sCampos = "b.id_CompraItem, b.id_Compra, b.id_Produto, c.id_ProdutoLote, b.qt_Produto, b.vl_Produto, b.kg_CompraItem"
    
    sTabela = sTabela & " ((tbdCompraItem b " _
    & " left JOIN tbdProdutoLote c ON c.id_ProdutoLote = b.id_ProdutoLote) " _
    & " left JOIN tbdMovimentacaoEstoque d ON d.id_ProdutoLote = c.id_ProdutoLote) "
    
    Call Select_Table(True, sTabela, sCampos, "id_Compra = " & id_Compra, , , , , rsado)
    Do While Not rsado.EOF
        Set cCompraItem = New clsCompraItem
        
        cCompraItem.id_CompraItem = ReadField(rsado, "id_CompraItem")
        cCompraItem.id_Compra = ReadField(rsado, "id_Compra")
        cCompraItem.id_Produto = ReadField(rsado, "id_Produto")
        cCompraItem.id_ProdutoLote = ReadField(rsado, "id_ProdutoLote")
        cCompraItem.qt_Produto = ReadField(rsado, "qt_Produto")
        cCompraItem.vl_Produto = ReadField(rsado, "vl_Produto")
        cCompraItem.kg_CompraItem = ReadField(rsado, "kg_CompraItem")
        
        Call cCompraItem.cProdutoLote.CarregarDados(ReadField(rsado, "id_ProdutoLote"))
        Call cCompraItem.cProduto.CarregarDados(ReadField(rsado, "id_Produto"))
        
        Call AdicionarItem(cCompraItem)
        Set cCompraItem = Nothing
        
        rsado.MoveNext
    Loop
    rsado.Close
    
    Exit Function
err_CarregarItens:
    ShowError
End Function
Private Function AdicionarItem(cCompraItem As clsCompraItem)
On Error GoTo err_AdicionarItem

    Call colCompraItem.Add(cCompraItem)
    
    Exit Function
err_AdicionarItem:
    ShowError
End Function

Private Sub Class_Initialize()
On Error GoTo err_Class_Initialize

    id_Compra = 0
    id_Fornecedor = 0
    id_Comprador = 0
    dt_Compra = CDateEspecial("")
    dt_PrevisaoEntrega = CDateEspecial("")
    dt_Entrega = CDateEspecial("")
    dt_EntradaEstoque = CDateEspecial("")
    vl_Compra = 0
    vl_Frete = 0
    vl_Adicional = 0
    ds_Observacao = ""
    ds_BaixaRecebimento = ""
    
    Set colCompraItem = New Collection
    
    Exit Sub
err_Class_Initialize:
    ShowError
End Sub
