VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFrete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public StatusGravacao As EnumStatusGravacao

Public ID As Long
Public id_Pessoa As Long
Public vl_FretePeso As Double
Public vl_Coleta As Double
Public vl_Entrega As Double
Public vl_FreteValor As Double
Public vl_GRIS As Double
Public vl_ITR As Double
Public vl_Despacho As Double
Public vl_SECCAT As Double
Public vl_SUFRAMA As Double
Public vl_Pedagio As Double
Public vl_TaxasDiversas As Double
Public vl_TotalFrete As Double

Private ListaNomenclatura As Collection

Private Sub Class_Initialize()
    StatusGravacao = EnumStatusGravacao.Nenhum
    ID = 0
    id_Pessoa = 0
    vl_FretePeso = 0
    vl_Coleta = 0
    vl_Entrega = 0
    vl_FreteValor = 0
    vl_GRIS = 0
    vl_ITR = 0
    vl_Despacho = 0
    vl_SECCAT = 0
    vl_SUFRAMA = 0
    vl_Pedagio = 0
    vl_TaxasDiversas = 0
    vl_TotalFrete = 0
    
    Call ZerarListas
End Sub

Public Function AdicionarNomenclatura(ByVal id_Nomenclatura As Long, ByVal ds_NomeColuna As String, ByVal ds_Nomenclatura As String, ByVal tp_Enviar As String, ByVal id_Pessoa As Long, ByVal StatusGravacao As EnumStatusGravacao) As Boolean
On Error GoTo err_Handler

    AdicionarNomenclatura = False

    Dim cNomenclatura As New clsNomenclatura

    cNomenclatura.StatusGravacao = StatusGravacao
    cNomenclatura.ID = id_Nomenclatura
    cNomenclatura.ds_NomeColuna = ds_NomeColuna
    cNomenclatura.ds_Nomenclatura = ds_Nomenclatura
    cNomenclatura.tp_Enviar = tp_Enviar
    cNomenclatura.id_Pessoa = id_Pessoa

    Call ListaNomenclatura.Add(cNomenclatura)

    Set cNomenclatura = Nothing
    AdicionarNomenclatura = True

    Exit Function
err_Handler:
    MensagemRetorno = "Ocorreu um erro ao adicionar os itens."
End Function

Public Property Get GetListaNomenclatura() As Collection
    Set GetListaNomenclatura = ListaNomenclatura
End Property

Private Sub ZerarListas()
    Set ListaNomenclatura = New Collection
End Sub
