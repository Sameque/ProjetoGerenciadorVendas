VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProdutoNatureza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id_ProdutoNatureza As Long
Public ds_ProdutoNatureza As String
Public tp_ProdutoPerigoso As String
Public tp_ControlePoliciaFederal As String

'"id_ProdutoNatureza,ds_ProdutoNatureza,tp_ProdutoPerigoso,tp_ControlePoliciaFederal"

Public Function Gravar() As Boolean
On Error GoTo err_Gravar

    If id_ProdutoNatureza > 0 Then
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
    
    id_ProdutoNatureza = Insert_Table("tbdProdutoNatureza", "id_ProdutoNatureza", Campos, Conteudo)
    If id_ProdutoNatureza = -1 Then
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

    aConteudo = Array(ds_ProdutoNatureza, tp_ProdutoPerigoso)
    
    If Update_Table("tbdProdutoNatureza", Campos, Conteudo, "id_ProdutoNatureza = " & id_ProdutoNatureza) = -1 Then
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
    
    If Delete_Table("tbdProdutoNatureza", "id_ProdutoNatureza = " & id_ProdutoNatureza) = -1 Then
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
    
    Call Select_Table(True, "tbdProdutoNatureza", "*", "id_ProdutoNatureza = " & id_Pesquisa, , , , , rsado)
    If Not rsado.EOF Then
        
        ds_ProdutoNatureza = ReadField(rsado, "ds_ProdutoNatureza")
        tp_ProdutoPerigoso = ReadField(rsado, "tp_ProdutoPerigoso")
        tp_ProdutoPerigoso = ReadField(rsado, "tp_ControlePoliciaFederal")

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
Private Function Campos()

    Dim sCampos As String

    sCampos = "ds_ProdutoNatureza,tp_ProdutoPerigoso,tp_ControlePoliciaFederal"
    Campos = sCampos

End Function
Private Function Conteudo()

    Dim aConteudo As Variant
    
    aConteudo = Array(ds_ProdutoNatureza, tp_ProdutoPerigoso, tp_ControlePoliciaFederal)
    Conteudo = aConteudo

End Function

Private Sub Class_Initialize()
On Error GoTo err_Class_Initialize
    id_ProdutoNatureza = 0
    ds_ProdutoNatureza = ""
    tp_ProdutoPerigoso = ""
    tp_ControlePoliciaFederal = ""
    Exit Sub
err_Class_Initialize:
    ShowError
End Sub
