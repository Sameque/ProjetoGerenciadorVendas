VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConexao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ds_Login As String
Public ds_Senha As String

Public Function Conectar(cnConexao As ADODB.Connection, sStringConexao As String) As Boolean
    On Error GoTo err_Conectar
    Screen.MousePointer = vbHourglass

    bConect = False
    cnConexao.CommandTimeout = 120
    cnConexao.ConnectionTimeout = 120
    cnConexao.Open sStringConexao
    bConect = True
    Conectar = True
    Screen.MousePointer = vbDefault

    Exit Function
err_Conectar:
    If bConect Then
        cnConexao.Close
    End If
    Conectar = False
    'Atencao, não pode colocar o showerro aqui. (Daniel)
    Mensagem "Ocorreu um erro na conexão: " & err.Number & " - " & err.Description, ErroCritico
End Function

Public Function logon(Optional oMDI As MDIForm) As Boolean
    On Error GoTo err_Logon
    Dim sWhere As String
    Dim cConexao As New clsConexao
    Dim sConexaoTemp As String
    Dim sIniFileBkp As String
    Dim bTerminalServerBkp As Boolean

    Screen.MousePointer = vbHourglass

    Call Conectar(cnAdo, sSistemaStringConexao)

    bUsuarioAdministrador = False

    logon = True
    bUsuarioAdministrador = True
    
    Screen.MousePointer = vbDefault

    Exit Function
err_Logon:
    If bConect Then
        cnAdo.Close
    End If
    logon = False
    ShowError
End Function
