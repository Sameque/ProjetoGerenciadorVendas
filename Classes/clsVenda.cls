VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id_Venda             As Long
Public id_Comprador         As Long
Public id_Vendedor          As Long
Public vl_Venda             As Double
Public vl_Desconto          As Double
Public vl_Frete             As Double
Public vl_Total             As Double
Public dt_Venda             As Date
Public dt_PrvisaoEmtrega    As Date
Public dt_Entrega           As Date
Public ds_Observacao        As String
Public MensagemRetorno      As String

Public colVendaItem As Collection

Public Function Gravar() As Boolean
On Error GoTo err_Gravar
          
    If id_Venda > 0 Then
        Gravar = Alterar
    Else
        Gravar = Incluir
    End If
        
    If Not Gravar Then
        Exit Function
    End
    Call GravarItem
    Exit Function
err_Gravar:
    ShowError
End Function

Private Function GravarItem() As Boolean
On Error GoTo err_Gravar
    
    Dim cVendaItem As New clsVendaItem
    
    GravarItem = False
    
    For Each cVendaItem In colVendaItem
        Call cVendaItem.Gravar
    Next

    GravarItem = True
    Set cVendaItem = Nothing
    
    Exit Function
err_Gravar:
    ShowError
End Function

Private Function Incluir() As Boolean
On Error GoTo err_Incluir
    
    Incluir = False
    
    id_Venda = Insert_Table("tbdVenda", "id_Venda", Campos, Conteudo)
    
    If id_Venda = -1 Then
        MensagemRetorno = "Erro ao incluir venda."
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
    If Update_Table("tbdVenda", Campos, Conteudo, "id_Venda = " & id_Venda) = -1 Then
        MensagemRetorno = "Erro ao atualizar venda."
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
    
    If id_Venda > 0 Then
        If Not ExcluirItem Then
            Exit Function
        End If
    Else
        Exit Function
    End If
     
    If Delete_Table("tbdVenda", "id_Venda = " & id_Venda) = -1 Then
        MensagemRetorno = "Erro ao excluir venda"
        Exit Function
    End If

    Excluir = True

    Exit Function
err_Excluir:
    ShowError
End Function
Public Function ExcluirItem() As Boolean
On Error GoTo err_Excluir
    Dim cVendaItem As clsVendaItem
    
    ExcluirItem = False
    
    For Each cVendaItem In colVendaItem
    
        Call cVendaItem.cProdutoLote.Movimentacao(Enum_Entrada, cVendaItem.qt_Produto)
        If Not cVendaItem.cProdutoLote.Gravar Then
            Exit Function
        End If
        
        If Not cVendaItem.Excluir Then
            Exit Function
        End If
    Next

    ExcluirItem = True
    Set cVendaItem = Nothing
    
    Exit Function
err_Excluir:
    ShowError
End Function

Private Function Campos() As String
    Campos = "id_Comprador, id_Vendedor, vl_Venda, vl_Desconto, vl_Frete, vl_Total, dt_Venda, dt_PrvisaoEmtrega, dt_Entrega, ds_Observacao"
End Function

Private Function Conteudo()
    Conteudo = Array(id_Comprador, id_Vendedor, vl_Venda, vl_Desconto, vl_Frete, vl_Total, dt_Venda, dt_PrvisaoEmtrega, dt_Entrega, ds_Observacao)
End Function

Public Function CarregarDados(id_Pesquisa As Long) As Boolean
On Error GoTo err_CarregarDados

    Dim rsado As New ADODB.Recordset
    Dim sCampos As String
    Dim sTabela As String
    
    CarregarDados = False
    Call Class_Initialize
    
    sCampos = " id_Venda,id_Comprador,id_Vendedor,vl_Venda,vl_Desconto,vl_Frete,vl_Total,dt_Venda,dt_PrvisaoEmtrega, " _
    & "dt_Entrega,ds_Observacao "
    
    sTabela = "tbdVenda"
    
    Call Select_Table(True, sTabela, sCampos, "id_Venda = " & id_Pesquisa, , , , , rsado)
    If Not rsado.EOF Then
        CarregarDados = True
        
        id_Venda = ReadField(rsado, "id_Venda")
        id_Comprador = ReadField(rsado, "id_Comprador")
        id_Vendedor = ReadField(rsado, "id_Vendedor")
        vl_Venda = ReadField(rsado, "vl_Venda")
        vl_Desconto = ReadField(rsado, "vl_Desconto")
        vl_Frete = ReadField(rsado, "vl_Frete")
        vl_Total = ReadField(rsado, "vl_Total")
        dt_Venda = CDateEspecial(ReadField(rsado, "dt_Venda"))
        dt_PrvisaoEmtrega = CDateEspecial(ReadField(rsado, "dt_PrvisaoEmtrega"))
        dt_Entrega = CDateEspecial(ReadField(rsado, "dt_Entrega"))
        ds_Observacao = ReadField(rsado, "ds_Observacao")
        
        MensagemRetorno = "Venda Carregada."
        Call CarregarItens
         
    End If
    rsado.Close
    
    If Not CarregarDados Then
        MensagemRetorno = "Erro ao carregar venda"
        Exit Function
    End If
    
    Exit Function
err_CarregarDados:
    ShowError
End Function
Private Function CarregarItens()
On Error GoTo err_CarregarItens

On Error GoTo err_CarregarItens
    Dim cVendaItem As clsVendaItem
    Dim rsado As New ADODB.Recordset
    
    
    
    
    
    
    
    
    Call Select_Table(True, "tbdVendaItem", "id_VendaItem, id_Venda, id_ProdutoLote, qt_produto, kg_Produto, vl_Produto", "id_Venda = " & id_Venda, , , , , rsado)
    Do While Not rsado.EOF
        Set cVendaItem = New clsVendaItem
        
        cVendaItem.id_VendaItem = ReadField(rsado, "id_VendaItem")
        cVendaItem.id_Venda = ReadField(rsado, "id_Venda")
        cVendaItem.qt_Produto = ReadField(rsado, "qt_Produto")
        cVendaItem.vl_Produto = ReadField(rsado, "vl_Produto")
        cVendaItem.kg_Produto = ReadField(rsado, "kg_Produto")
        Call cVendaItem.cProdutoLote.CarregarDados(ReadField(rsado, "id_ProdutoLote"))
        
        Call AdicionarItem(cVendaItem)
        Set cVendaItem = Nothing
        
        rsado.MoveNext
    Loop
    
    rsado.Close
    
    Exit Function
err_CarregarItens:
    ShowError

    Exit Function
err_CarregarItens:
    ShowError
End Function
Public Function AdicionarItem(ByVal id_VendaItem As Long, ByVal id_Venda As Long, ByVal id_ProdutoLote As Long, ByVal qt_Produto As Long, ByVal vl_Produto As Double, ByVal kg_CompraItem As Double) As Boolean
On Error GoTo err_AdicionarItem
    
    AdicionarItem = False
    Dim cVendaItem As New clsVendaItem
    
    cVendaItem.id_VendaItem = id_VendaItem
    cVendaItem.qt_Produto = qt_Produto
    cVendaItem.vl_Produto = vl_Produto
    cVendaItem.kg_CompraItem = kg_CompraItem
    
    Call cVendaItem.cProdutoLote.CarregarDados(id_ProdutoLote)
    
    AdicionarItem = True
    
    Exit Function
err_AdicionarItem:
    ShowError
    End Function
Private Function CalcularValorProdutos()
On Error GoTo err_CalcularValorProdutos
    Dim cVendaItem As clsVendaItem
    Dim cProduto As clsProduto
    
    For Each cVendaItem In colVendaItem
        vl_Venda = vl_Venda + (cVendaItem.cProdutoLote.vl_Venda * cVendaItem.qt_Produto)
    Next
    
    Exit Function
err_CalcularValorProdutos:
    ShowError
End Function
Public Function CalcularValorTotal()
On Error GoTo err_CalcularValorTotal
    vl_Total = 0
    vl_Venda = 0
    Call CalcularValorProdutos
    
    vl_Total = (vl_Frete + vl_Venda) - vl_Desconto
    
    Exit Function
err_CalcularValorTotal:
    ShowError
End Function

Private Sub Class_Initialize()
On Error GoTo err_Class_Initialize

    id_Venda = 0
    id_Comprador = 0
    id_Vendedor = 0
    vl_Venda = 0
    vl_Desconto = 0
    vl_Frete = 0
    vl_Total = 0
    dt_Venda = CDateEspecial("")
    dt_PrvisaoEmtrega = CDateEspecial("")
    dt_Entrega = CDateEspecial("")
    ds_Observacao = ""
    
    Set colVendaItem = New Collection
    
    Exit Sub
err_Class_Initialize:
    ShowError
End Sub


Private Sub Class_Terminate()
    Set colVendaItem = Nothing
End Sub
