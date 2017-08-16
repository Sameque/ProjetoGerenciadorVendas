VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPessoaContato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id                       As Long
Public id_Pessoa                As Long
Public ds_Nome                  As String
Public cd_Fone                  As String
Public cd_Email                 As String

Public menumStatusGravacao           As enumStatusGravacao
Public mstrMensagemRetorno       As String

Private Sub Class_Initialize()
On Error GoTo err_Class_Initialize
    
    id = 0
    id_Pessoa = 0
    ds_Nome = ""
    cd_Fone = ""
    cd_Email = ""
        
    menumStatusGravacao = enumStatusGravacao.Nenhum
    mstrMensagemRetorno = ""
    
    Exit Sub
err_Class_Initialize:
    mstrMensagemRetorno = "Erro ao inicializar classe."
End Sub

Public Function Tabela() As String
    Tabela = "tbdPessoaContato"
End Function
Public Function Campos() As String
    Campos = "id_Pessoa, ds_Nome, cd_Fone, cd_Email"
End Function

Public Function ChavePrimaria() As String
    ChavePrimaria = "id_PessoaContato"
End Function

Public Function Conteudo()
    Conteudo = Array(id_Pessoa, ds_Nome, cd_Fone, cd_Email)
End Function

