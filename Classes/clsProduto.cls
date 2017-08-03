VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id_Produto As Long
Public cd_Produto As String
Public ds_Produto As String
Public kg_Real As Double
Public kg_Cubado As Double
Public id_ProdutoNatureza As Long
Public qt_Minima As Long

Public Function Gravar() As Boolean
On Error GoTo err_Gravar

    If id_Produto > 0 Then
        Gravar = Alterar
    Else
        Gravar = Incluir
    End If
        
err_Gravar:
    ShowError
End Function

Private Function Incluir() As Boolean
On Error GoTo err_Incluir
    
    Dim sCampos As String
    Dim aConteudo As Variant
    
    Incluir = False

    sCampos = "cd_Produto, ds_Produto, kg_Real, kg_Cubado, id_ProdutoNatureza, qt_Minima"

    aConteudo = Array(cd_Produto, ds_Produto, kg_Real, kg_Cubado, id_ProdutoNatureza, qt_Minima)
    

    If Insert_Table("tbdProduto", "id_Produto", sCampos, aConteudo) = -1 Then
        Exit Function
    End If
    
    Incluir = True

    Exit Function
err_Incluir:
    ShowError
End Function

Private Function Alterar() As Boolean
On Error GoTo err_Alterar

    Dim sCampos As String
    Dim aConteudo As Variant
    
    Alterar = False

    sCampos = "cd_Produto, ds_Produto, kg_Real, kg_Cubado, id_ProdutoNatureza, qt_Minima"

    aConteudo = Array(cd_Produto, ds_Produto, kg_Real, kg_Cubado, id_ProdutoNatureza, qt_Minima)
    
    If Update_Table("tbdProduto", sCampos, aConteudo, "id_Produto = " & id_Produto) = -1 Then
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
    
    If Delete_Table("tbdProduto", "id_Produto = " & id_Produto) = -1 Then
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
    
    Call Select_Table(True, "tbdProduto", "id_Produto,cd_Produto,ds_Produto,kg_Real,id_ProdutoNatureza,qt_Minima", "id_Produto = " & id_Pesquisa, , , , , rsado)
    If Not rsado.EOF Then
        CarregarDados = True
        
        id_Produto = ReadField(rsado, "id_Produto")
        cd_Produto = ReadField(rsado, "cd_Produto")
        ds_Produto = ReadField(rsado, "ds_Produto")
        kg_Real = ReadField(rsado, "kg_Real")
        kg_Cubado = ReadField(rsado, "kg_Cubado")
        id_ProdutoNatureza = ReadField(rsado, "id_ProdutoNatureza")
        qt_Minima = ReadField(rsado, "qt_Minima")
        
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
    id_Produto = 0
    cd_Produto = ""
    ds_Produto = ""
    kg_Real = 0
    kg_Cubado = 0
    id_ProdutoNatureza = 0
    qt_Minima = 0
   
    Exit Sub
err_Class_Initialize:
    ShowError
End Sub

