VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNomenclatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public StatusGravacao As EnumStatusGravacao

Public ID As Long
Public ds_NomeColuna As String
Public ds_Nomenclatura As String
Public tp_Enviar As String
Public id_Pessoa As Long

Private Sub Class_Initialize()
    StatusGravacao = EnumStatusGravacao.Nenhum
    ID = 0
    ds_NomeColuna = ""
    ds_Nomenclatura = ""
    tp_Enviar = ""
    id_Pessoa = 0
    Call ZerarListas
End Sub
