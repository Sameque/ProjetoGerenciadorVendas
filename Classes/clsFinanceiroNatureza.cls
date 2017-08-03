VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFinanceiroNatureza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id_FinanceiroNatureza As Long
Public ds_FinanceiroNatureza As String
Public tp_Financeiro As String

Public Function Gravar() As Boolean
On Error GoTo err_Gravar

    If id_FinanceiroNatureza > 0 Then
        Gravar = Alterar
    Else
        Gravar = Incluir
    End If
    
    If Not Gravar Then
        Exit Function
    End If
    
err_Gravar:
    ShowError
End Function

Private Function Incluir() As Boolean
On Error GoTo err_Incluir
    
    Incluir = False
    
    id_FinanceiroNatureza = Insert_Table("tbdFinanceiroNatureza", "id_FinanceiroNatureza", Campos, Conteudo)
    
    Incluir = True

    Exit Function
err_Incluir:
    ShowError
End Function

Private Function Alterar() As Boolean
On Error GoTo err_Alterar
    
    Alterar = False

    If Update_Table("tbdFinanceiroNatureza", Campos, Conteudo, "id_FinanceiroNatureza= " & id_FinanceiroNatureza) = -1 Then
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
    
    If Delete_Table("tbdFinanceiroNatureza", "id_FinanceiroNatureza= " & id_FinanceiroNatureza) = -1 Then
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
    
    Call Select_Table(True, "tbdFinanceiroNatureza", "*", "id_FinanceiroNatureza = " & id_Pesquisa, , , , , rsado)
    If Not rsado.EOF Then
        id_FinanceiroNatureza = ReadField(rsado, "id_FinanceiroNatureza")
        ds_FinanceiroNatureza = ReadField(rsado, "ds_FinanceiroNatureza")
        id_FinanceiroNatureza = ReadField(rsado, "tp_Financeiro")
        
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


Private Sub Class_Initialize()
On Error GoTo err_Class_Initialize
    id_FinanceiroNatureza = 0
    ds_FinanceiroNatureza = ""
    tp_Financeiro = ""
    
    Exit Sub
err_Class_Initialize:
    ShowError
End Sub

Private Function Campos()

    Dim sCampos As String
    
    sCampos = "ds_FinanceiroNatureza,tp_Financeiro"
    Campos = sCampos

End Function
Private Function Conteudo()

    Dim aConteudo As Variant
    
    aConteudo = Array(ds_FinanceiroNatureza, tp_Financeiro)
    Conteudo = aConteudo

End Function
