VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'[CRIAR-PROPRIEDADES]'

Public Function Gravar() As Boolean
On Error GoTo err_Gravar

    If id_Principal > 0 Then
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

    '[CAMPOS-INSERT]'
    '[CONTEUDO-INSERT]'
    
    id_Principal = Insert_Table("tbdPrincipal", "id_Principal", sCampos, aConteudo)
    If id_Principal = -1 Then
        Exit Function
    End If
    
    '[GRAVAR-ITENS]'
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

    '[CAMPOS-UPDATE]'
    '[CONTEUDO-UPDATE]'
    
    If Update_Table("tbdPrincipal", sCampos, aConteudo, "id_Principal = " & id_Principal) = -1 Then
        Exit Function
    End If
    
    '[GRAVAR-ITENS]'
    Alterar = True

    Exit Function
err_Alterar:
    ShowError
End Function

Public Function Excluir() As Boolean
On Error GoTo err_Excluir

    Excluir = False
    
    '[EXCLUIR-ITENS]'
    If Delete_Table("tbdPrincipal", "id_Principal = " & id_Principal) = -1 Then
        Exit Function
    End If

    Excluir = True

    Exit Function
err_Excluir:
    ShowError
End Function

'[FUNCAO-GRAVARITENS]'
Public Function CarregarDados(id_Pesquisa As Long) As Boolean
On Error GoTo err_CarregarDados

    Dim rsado As New ADODB.Recordset
    
    CarregarDados = False
    
    Call Class_Initialize
    
    Call Select_Table(True, "tbdPrincipal", "*", "id_Principal = " & id_Pesquisa, , , , , rsado)
    If Not rsado.EOF Then
        CarregarDados = True
        
        '[SETAR-PROPRIEDADES]'
    End If
    rsado.Close
    
    If Not CarregarDados Then
        Exit Function
    End If
    
    '[CHAMADA-CARREGARITEM]'
    Exit Function
err_CarregarDados:
    ShowError
End Function

'[CARREGAR-ITENS]'
Private Sub Class_Initialize()
On Error GoTo err_Class_Initialize
    '[INICIALIZAR-PROPRIEDADES]'
    
    Exit Sub
err_Class_Initialize:
    ShowError
End Sub

'[ZERAR-ITENS]'
