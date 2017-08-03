VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFinanceiroTipoBaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id_FinanceiroTipoBaixa As Long
Public ds_TipoBaixa As String

Public Function Gravar() As Boolean
On Error GoTo err_Gravar

    Call AbreTransacao

    If id_FinanceiroTipoBaixa > 0 Then
        Gravar = Alterar
    Else
        Gravar = Incluir
    End If
    
    If Not Gravar Then
        Call VoltaTransacao
        Exit Function
    End If
    
    Call FechaTransacao
    
err_Gravar:
    ShowError
End Function

Private Function Incluir() As Boolean
On Error GoTo err_Incluir
    
    Incluir = False
    
    id_FinanceiroTipoBaixa = Insert_Table("tbdFinanceiroTipoBaixa", "id_FinanceiroTipoBaixa", Campos, Conteudo)
    
    Incluir = True

    Exit Function
err_Incluir:
    ShowError
End Function

Private Function Alterar() As Boolean
On Error GoTo err_Alterar
    
    Alterar = False

    If Update_Table("tbdFinanceiroTipoBaixa", Campos, Conteudo, "id_FinanceiroTipoBaixa = " & id_FinanceiroTipoBaixa) = -1 Then
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
    
    Call AbreTransacao
    
    If Delete_Table("tbdFinanceiroTipoBaixa", "id_FinanceiroTipoBaixa = " & id_FinanceiroTipoBaixa) = -1 Then
        Call VoltaTransacao
        Exit Function
    End If

    Call FechaTransacao

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
    
    Call Select_Table(True, "tbdFinanceiroTipoBaixa", "*", "id_FinanceiroTipoBaixa = " & id_Pesquisa, , , , , rsado)
    If Not rsado.EOF Then
        ds_TipoBaixa = ReadField(rsado, "ds_TipoBaixa")
        id_FinanceiroTipoBaixa = ReadField(rsado, "id_FinanceiroTipoBaixa")
                
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
    id_FinanceiroTipoBaixa = 0
    ds_TipoBaixa = ""
    
    Exit Sub
err_Class_Initialize:
    ShowError
End Sub

Private Function Campos()

Dim sCampos As String

sCampos = "ds_TipoBaixa"
Campos = sCampos

End Function
Private Function Conteudo()

Dim aConteudo As Variant

aConteudo = Array(ds_TipoBaixa)
Conteudo = aConteudo

End Function
