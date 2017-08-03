VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVendaItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id_VendaItem         As Long
Public id_Venda             As Long
Public qt_Produto           As Long
Public vl_Produto           As Double
Public kg_Produto           As Double

Public MensagemRetorno      As String
Public cProdutoLote As clsProdutoLote

Public Function Gravar() As Boolean
On Error GoTo err_Gravar
    
    If id_VendaItem > 0 Then
        Gravar = Alterar
    Else
        Gravar = Incluir
    End If
    
    If Not Gravar Then
        Exit Function
    End If
    
    Exit Function
err_Gravar:
    ShowError
End Function

Private Function Incluir() As Boolean
On Error GoTo err_Incluir
    
    Incluir = False
    
    id_VendaItem = Insert_Table("tbdVendaItem", "id_VendaItem", Campos, Conteudo)
    
    If id_VendaItem = -1 Then
        MensagemRetorno = "Erro ao incluir item da venda."
        Exit Function
    End If
    
    Incluir = True
    
    Exit Function
err_Incluir:
    MsgErro = "Erro ao Incluir item da venda"
End Function

Private Function Alterar() As Boolean
On Error GoTo err_Alterar

    Alterar = False
    
    If Update_Table("tbdVendaItem", Campos, Conteudo, "id_VendaItem = " & id_VendaItem) = -1 Then
        MensagemRetorno = "Erro ao atualizar item."
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
      
    If Delete_Table("tbdVendaItem", "id_VendaItem = " & id_VendaItem) = -1 Then
        MensagemRetorno = "Erro excluir item."
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
    Set cProduto = New clsProduto
    
    CarregarDados = False
    Call Class_Initialize
        
    Call Select_Table(True, "tbdVendaItem", Campos, "id_VendaItem = " & id_Pesquisa, , , , , rsado)
    If Not rsado.EOF Then
                
        id_VendaItem = id_Pesquisa
        id_Venda = ReadField(rsado, "id_Venda")
        qt_Produto = ReadField(rsado, "qt_Produto")
        vl_Produto = ReadField(rsado, "vl_Produto")
        kg_VendaItem = ReadField(rsado, "kg_VendaItem")
                
        Call cProdutoLote.CarregarDados(ReadField(rsado, "id_ProdutoLote"))
        CarregarDados = True
        
    End If
    rsado.Close
    
    If Not CarregarDados Then
        MensagemRetorno = "Erro ao carregr item."
        Exit Function
    End If
    
    Exit Function
err_CarregarDados:
    ShowError
End Function

Private Function Campos() As String
    Campos = "id_Venda, id_ProdutoLote, qt_produto, vl_Produto, kg_Produto"
End Function

Private Function Conteudo()
    Conteudo = Array(id_Venda, cProdutoLote.id_ProdutoLote, qt_Produto, vl_Produto, kg_Produto)
End Function

Private Sub Class_Initialize()
On Error GoTo err_Class_Initialize

    id_VendaItem = 0
    id_Venda = 0
    qt_Produto = 0
    vl_Produto = 0
    kg_Produto = 0
    Set cProdutoLote = New clsProdutoLote

    Exit Sub
err_Class_Initialize:
    ShowError
End Sub
