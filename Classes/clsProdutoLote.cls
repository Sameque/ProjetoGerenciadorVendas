VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProdutoLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id_ProdutoLote As Long
Public id_Produto As Long
Public qt_Produto As Long
Public qt_ProdutoSaldo As Long
Public vl_Compra As Double
Public vl_Venda As Double
Public dt_EntradaEstoque As Date
Public dt_Vencimento As Date

Public Enum Enum_Movimentacao
    Enum_Entrada
    Enum_Saida
End Enum

Public cProduto As clsProduto
Public colMovimentacaoEstoque As Collection

Public strMensagemRetorno  As String

Public Function Movimentacao(enuTipoMovimentacao As Enum_Movimentacao, lngQtdMovimentacao As Long)
On Error GoTo err_Movimentacao
    Dim cMovimentacaoEstoque As New clsMovimentacaoEstoque
    
    Movimentacao = False
    
    cMovimentacaoEstoque.id_ProdutoLote = id_ProdutoLote
    cMovimentacaoEstoque.tp_MovimentacaoEstoque = IIf(enuTipoMovimentacao = Enum_Entrada, "E", IIf(enuTipoMovimentacao = Enum_Saida, "S", ""))
    cMovimentacaoEstoque.qt_MovimentacaoEstoque = lngQtdMovimentacao
        
    If enuTipoMovimentacao = Enum_Entrada Then
        qt_ProdutoSaldo = qt_ProdutoSaldo + lngQtdMovimentacao
    Else
        qt_ProdutoSaldo = qt_ProdutoSaldo - lngQtdMovimentacao
    End If
    
    Call AdicionarMovimentacao(cMovimentacaoEstoque)
    
    Set cMovimentacaoEstoque = Nothing
    Movimentacao = True
    
    Exit Function
err_Movimentacao:
    ShowError
End Function

Public Function EstornoEntrada() As Boolean
On erro GoTo err_EntradaProduto
    
    EstornoEntrada = False
    
    If VerificarMovimentacaoSaida Then
        Exit Function
    End If
    
    If Not ExcluirMovimentacoes Then
        Exit Function
    End If
    
    If Not Excluir Then
        Exit Function
    End If
        
    EstornoEntrada = True
    
    Exit Function
err_EntradaProduto:
    ShowError
End Function
Private Function ExcluirMovimentacoes() As Boolean
On Error GoTo err_ExcluirMovimentacoes
    Dim cMovimentacaoEstoque As clsMovimentacaoEstoque
    
    ExcluirMovimentacoes = False
        
    For Each cMovimentacaoEstoque In colMovimentacaoEstoque
        cMovimentacaoEstoque.Excluir
    Next
    
    Set cMovimentacaoEstoque = Nothing
    
    ExcluirMovimentacoes = True
    
    Exit Function
err_ExcluirMovimentacoes:
    ShowError
End Function

Private Function VerificarMovimentacaoSaida() As Boolean
On Error GoTo err_VerificarMovimentacaoSaida
    
    Dim cMovimentacaoEstoque As New clsMovimentacaoEstoque
    
    VerificarMovimentacaoSaida = False
    
    For Each cMovimentacaoEstoque In colMovimentacaoEstoque
        If cMovimentacaoEstoque.tp_MovimentacaoEstoque = "S" Then
            VerificarMovimentacaoSaida = True
        End If
    Next
    
    Set cMovimentacaoEstoque = Nothing
    Exit Function
err_VerificarMovimentacaoSaida:
    ShowError
End Function
Public Function Gravar() As Boolean
On Error GoTo err_Gravar

    Call AbreTransacao
    
    If VerificarSaldo Then
        strMensagemRetorno = strMensagemRetorno = "Saldo insuficiente!"
    End If
    
    If id_ProdutoLote > 0 Then
        Gravar = Alterar
    Else
        Gravar = Incluir
    End If
        
    If Not Gravar Then
        Call VoltaTransacao
        Exit Function
    End If
    
    If Not GravarMovimentacao Then
        Call VoltaTransacao
    End If
    
    Call FechaTransacao
        
    Exit Function
err_Gravar:
    ShowError
End Function
Public Function GravarMovimentacao() As Boolean
On Error GoTo err_GravarMovimentacao
    Dim cMovimentacaoEstoque As New clsMovimentacaoEstoque
    
    GravarMovimentacao = False
        
    For Each cMovimentacaoEstoque In colMovimentacaoEstoque
        cMovimentacaoEstoque.id_ProdutoLote = id_ProdutoLote
        cMovimentacaoEstoque.Gravar
    Next
        
    Set cMovimentacaoEstoque = Nothing
    GravarMovimentacao = True
    
    Exit Function
err_GravarMovimentacao:
    ShowError
End Function

Public Function Excluir() As Boolean
On Error GoTo err_Excluir

    Excluir = False
        
    If Delete_Table("tbdProdutoLote", "id_ProdutoLote = " & id_ProdutoLote) = -1 Then
        Exit Function
    End If

    Excluir = True

    Exit Function
err_Excluir:
    ShowError
End Function

Public Function CarregarDados(id_Pesquisa As Long) As Boolean
On Error GoTo err_CarregarDados

    Dim rsado As New ADODB.Recordset
    
    CarregarDados = False
    
    Call Class_Initialize
    
    Call Select_Table(True, "tbdProdutoLote", "id_ProdutoLote,id_Produto,vl_Compra,dt_EntradaEstoque,dt_Vencimento,qt_Produto,qt_ProdutoSaldo", "id_ProdutoLote = " & id_Pesquisa, , , , , rsado)
    If Not rsado.EOF Then
    
        id_ProdutoLote = ReadField(rsado, "id_ProdutoLote")
        id_Produto = ReadField(rsado, "id_Produto")
        vl_Compra = ReadField(rsado, "vl_Compra")
        vl_Venda = CalcularValorVenda
        dt_EntradaEstoque = CDateEspecial(ReadField(rsado, "dt_EntradaEstoque"))
        dt_Vencimento = CDateEspecial(ReadField(rsado, "dt_Vencimento"))
        qt_Produto = ReadField(rsado, "qt_Produto")
        qt_ProdutoSaldo = ReadField(rsado, "qt_ProdutoSaldo")
        
        Call cProduto.CarregarDados(id_Produto)
        Call CarregarMovimentacao
        CarregarDados = True
        
    End If
    rsado.Close
    
    If Not CarregarDados Then
        Exit Function
    End If
        
    Exit Function
err_CarregarDados:
    ShowError
End Function
Private Function CalcularValorVenda() As Double
On Error GoTo err_CalcularValorVenda
        
    CalcularValorVenda = 0
    
    CalcularValorVenda = vl_Compra + (vl_Compra * 0.01)
    
    Exit Function
err_CalcularValorVenda:
    ShowError
End Function
Private Function Incluir() As Boolean
On Error GoTo err_Incluir
    
    Incluir = False
    
    id_ProdutoLote = Insert_Table("tbdProdutoLote", "id_ProdutoLote", Campos, Conteudo)
    
    Incluir = True

    Exit Function
err_Incluir:
    ShowError
End Function

Private Function Alterar() As Boolean
On Error GoTo err_Alterar
    
    Alterar = False

    If Update_Table("tbdProdutoLote", Campos, Conteudo, "id_ProdutoLote = " & id_ProdutoLote) = -1 Then
        Exit Function
    End If
    
    Alterar = True

    Exit Function
err_Alterar:
    ShowError
End Function

Private Function CarregarMovimentacao()
On Error GoTo err_CarregarMovimentacao
    Dim cMovimentacaoEstoque As clsMovimentacaoEstoque
    Dim rsado As New ADODB.Recordset
    
    Call Select_Table(True, "tbdMovimentacaoEstoque", "id_MovimentacaoEstoque, id_ProdutoLote, tp_MovimentacaoEstoque, qt_MovimentacaoEstoque, dt_MovimentacaoEstoque", "id_ProdutoLote = " & id_ProdutoLote, , , , , rsado)
    Do While Not rsado.EOF
        Set cMovimentacaoEstoque = New clsMovimentacaoEstoque
        
        cMovimentacaoEstoque.id_MovimentacaoEstoque = ReadField(rsado, "id_MovimentacaoEstoque")
        cMovimentacaoEstoque.id_ProdutoLote = ReadField(rsado, "id_ProdutoLote")
        cMovimentacaoEstoque.tp_MovimentacaoEstoque = ReadField(rsado, "tp_MovimentacaoEstoque")
        cMovimentacaoEstoque.qt_MovimentacaoEstoque = ReadField(rsado, "qt_MovimentacaoEstoque")
        cMovimentacaoEstoque.dt_MovimentacaoEstoque = CDateEspecial(ReadField(rsado, "dt_MovimentacaoEstoque"))
        
        Call AdicionarMovimentacao(cMovimentacaoEstoque)
        Set cMovimentacaoEstoque = Nothing
        
        rsado.MoveNext
    Loop
    rsado.Close
    
    Exit Function
err_CarregarMovimentacao:
    ShowError
End Function

Public Function AdicionarMovimentacao(cMovimentacaoEstoque As clsMovimentacaoEstoque)
    Call colMovimentacaoEstoque.Add(cMovimentacaoEstoque)
End Function

Private Function Campos()
    Campos = "id_Produto, qt_Produto, qt_ProdutoSaldo, vl_Compra, dt_EntradaEstoque, dt_Vencimento"
End Function

Private Function Conteudo()
    Conteudo = Array(id_Produto, qt_Produto, qt_ProdutoSaldo, vl_Compra, dt_EntradaEstoque, dt_Vencimento)
End Function

Private Function VerificarSaldo() As Boolean
On Error GoTo err_VerificarSaldo
    Dim rsado As New ADODB.Recordset
    Dim dblSaldoMovimentado As Double
    Dim dblSaldo As Double
    Dim cMovimentacaoEstoque As clsMovimentacaoEstoque
    
    VerificarSaldo = False
    dblSaldoMovimentado = 0
    dblSaldo = 0
    For Each cMovimentacaoEstoque In colMovimentacaoEstoque
        If cMovimentacaoEstoque.tp_MovimentacaoEstoque = "S" Then
            dblSaldoMovimentado = dblSaldoMovimentado + cMovimentacaoEstoque.qt_MovimentacaoEstoque
        End If
    Next
    
    Set cMovimentacaoEstoque = Nothing
    
    Call Select_Table(True, "tbdProdutoLote", "isnull(qt_ProdutoSaldo,0) as qt_ProdutoSaldo", "id_ProdutoLote = " & id_ProdutoLote, , , , , rsado)
    
    If Not rsado.EOF Then
        dblSaldo = ReadField(rsado, "qt_ProdutoSaldo")
        rsado.Close
        Exit Function
    End If
    rsado.Close
    
    If dblSaldoMovimentado > dblSaldo Then
        strMensagemRetorno = strMensagemRetorno & "Quantidade insuficiente!"
    End If
    
    Exit Function
err_VerificarSaldo:
    ShowError
End Function

Private Sub Class_Initialize()
On Error GoTo err_Class_Initialize

    id_ProdutoLote = 0
    id_Produto = 0
    qt_Produto = 0
    qt_ProdutoSaldo = 0
    vl_Compra = 0
    dt_EntradaEstoque = CDateEspecial("")
    dt_Vencimento = CDateEspecial("")
    strMensagemRetorno = ""
    Set cProduto = New clsProduto
    Set colMovimentacaoEstoque = New Collection
    
    Exit Sub
err_Class_Initialize:
    ShowError
End Sub
