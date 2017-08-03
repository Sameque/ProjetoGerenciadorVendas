VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMovimentacaoEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id_MovimentacaoEstoque As Long
Public id_ProdutoLote         As Long
Public tp_MovimentacaoEstoque As String
Public qt_MovimentacaoEstoque As Long
Public dt_MovimentacaoEstoque As Date
Public strMensagemRetorno     As String
Public Function Gravar() As Boolean
On Error GoTo err_Gravar

    If id_MovimentacaoEstoque > 0 Then
        Gravar = Alterar
    Else
        Gravar = Incluir
    End If
        
    If Not Gravar Then
        Exit Function
    End If
    
    Call FechaTransacao
    
    Exit Function
err_Gravar:
    ShowError
End Function

Public Function Excluir() As Boolean
On Error GoTo err_Excluir

    Excluir = False
        
    If Delete_Table("tbdMovimentacaoEstoque", "id_MovimentacaoEstoque = " & id_MovimentacaoEstoque) = -1 Then
        strMensagemRetorno = "Erro ao apagar registro!"
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
    
    Call Select_Table(True, "tbdMovimentacaoEstoque", Campos, "id_MovimentacaoEstoque = " & id_Pesquisa, , , , , rsado)
    If Not rsado.EOF Then
        
        id_MovimentacaoEstoque = ReadField(rsado, "id_MovimentacaoEstoque")
        id_ProdutoLote = ReadField(rsado, "id_ProdutoLote")
        tp_MovimentacaoEstoque = ReadField(rsado, "tp_MovimentacaoEstoque")
        qt_MovimentacaoEstoque = ReadField(rsado, "qt_MovimentacaoEstoque")
        dt_MovimentacaoEstoque = ReadField(rsado, "dt_MovimentacaoEstoque")
        
        CarregarDados = True
    End If
    rsado.Close
    
    If Not CarregarDados Then
        strMensagemRetorno = "Erro ao carregar dados"
        Exit Function
    End If
    
    Exit Function
err_CarregarDados:
    ShowError
End Function

Private Function Incluir() As Boolean
On Error GoTo err_Incluir
    
    Incluir = False
    
    id_MovimentacaoEstoque = Insert_Table("tbdMovimentacaoEstoque", "id_MovimentacaoEstoque", Campos, Conteudo)
    
    If id_MovimentacaoEstoque = -1 Then
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

    If Update_Table("tbdMovimentacaoEstoque", Campos, Conteudo, "id_MovimentacaoEstoque = " & id_MovimentacaoEstoque) = -1 Then
        Exit Function
    End If
    
    Alterar = True

    Exit Function
err_Alterar:
    ShowError
End Function

Private Sub Class_Initialize()
On Error GoTo err_Class_Initialize
    
    id_MovimentacaoEstoque = 0
    id_ProdutoLote = 0
    tp_MovimentacaoEstoque = ""
    qt_MovimentacaoEstoque = 0
    dt_MovimentacaoEstoque = CDateEspecial("")
    
    Exit Sub
err_Class_Initialize:
    ShowError
End Sub

Private Function Campos()
    Campos = "id_ProdutoLote, tp_MovimentacaoEstoque, qt_MovimentacaoEstoque, dt_MovimentacaoEstoque"
End Function
Private Function Conteudo()
    Conteudo = Array(id_ProdutoLote, tp_MovimentacaoEstoque, qt_MovimentacaoEstoque, dt_MovimentacaoEstoque)
End Function

