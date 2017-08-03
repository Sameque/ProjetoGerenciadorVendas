VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPessoaFuncao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id_PessoaFuncao As Long
Public ds_PessoaFuncao As String
Public qt_NivelPermissao As Long

Public Function Gravar() As Boolean
On Error GoTo err_Gravar

    If id_PessoaFuncao > 0 Then
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
    
    id_PessoaFuncao = Insert_Table("tbdPessoaFuncao", "id_PessoaFuncao", Campos, Conteudo)
    If id_PessoaFuncao = -1 Then
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

'    aConteudo = Array(ds_PessoaFuncao, qt_NivelPermissao)
    
    If Update_Table("tbdPessoaFuncao", Campos, Conteudo, "id_PessoaFuncao = " & id_PessoaFuncao) = -1 Then
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
    
    If Delete_Table("tbdPessoaFuncao", "id_PessoaFuncao = " & id_PessoaFuncao) = -1 Then
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
    
    Call Select_Table(True, "tbdPessoaFuncao", "*", "id_PessoaFuncao = " & id_Pesquisa, , , , , rsado)
    If Not rsado.EOF Then
        CarregarDados = True
        
        ds_PessoaFuncao = ReadField(rsado, "ds_PessoaFuncao")
        qt_NivelPermissao = ReadField(rsado, "qt_NivelPermissao")

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

    sCampos = "ds_PessoaFuncao, qt_NivelPermissao"
    Campos = sCampos

End Function
Private Function Conteudo()

    Dim aConteudo As Variant
    
    aConteudo = Array(ds_PessoaFuncao, qt_NivelPermissao)
    Conteudo = aConteudo

End Function

Private Sub Class_Initialize()
On Error GoTo err_Class_Initialize
    id_PessoaFuncao = 0
    ds_PessoaFuncao = ""
    qt_NivelPermissao = 0

    Exit Sub
err_Class_Initialize:
    ShowError
End Sub
