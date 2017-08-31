VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServicoFrete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MensagemRetorno As String

Private sCampos As String
Private aConteudo As Variant


Public Function CarregarPorID(ByVal id_frete As Long, Optional ByVal blnCarregarFilhos) As clsFrete
On Error GoTo err_Handler

    Dim cFrete As New clsFrete
    Dim rsado As New ADODB.Recordset

    Call Select_Table(True, "tbdFrete", _
                     "id_Frete,id_Pessoa,vl_FretePeso,vl_Coleta,vl_Entrega,vl_FreteValor,vl_GRIS,vl_ITR,vl_Despacho,vl_SECCAT,vl_SUFRAMA,vl_Pedagio," _
                   & "vl_TaxasDiversas,vl_TotalFrete", _
                     "id_Frete =" & id_frete, , , , , rsado)

    If Not rsado.EOF Then
        With cFrete
            .ID = ReadField(rsado, "id_Frete")
            .id_Pessoa = ReadField(rsado, "id_Pessoa")
            .vl_FretePeso = ReadField(rsado, "vl_FretePeso")
            .vl_Coleta = ReadField(rsado, "vl_Coleta")
            .vl_Entrega = ReadField(rsado, "vl_Entrega")
            .vl_FreteValor = ReadField(rsado, "vl_FreteValor")
            .vl_GRIS = ReadField(rsado, "vl_GRIS")
            .vl_ITR = ReadField(rsado, "vl_ITR")
            .vl_Despacho = ReadField(rsado, "vl_Despacho")
            .vl_SECCAT = ReadField(rsado, "vl_SECCAT")
            .vl_SUFRAMA = ReadField(rsado, "vl_SUFRAMA")
            .vl_Pedagio = ReadField(rsado, "vl_Pedagio")
            .vl_TaxasDiversas = ReadField(rsado, "vl_TaxasDiversas")
            .vl_TotalFrete = ReadField(rsado, "vl_TotalFrete")
             
             Call CarregarNomenclatura(cFrete)
        End With
    Else
        MensagemRetorno = "Registro não Localizado."
    End If
    rsado.Close
    
    If Not Carregar Then
        
    End If
    
    Set CarregarPorID = cFrete
    Set cFrete = Nothing
    Set rsado = Nothing
    
    Exit Function
err_Handler:
    ShowError "Ocorreu um erro ao Carregar os Dados."
    End
End Function

Public Function CarregarNomenclatura(ByRef cFrete As clsFrete)
On Error GoTo err_CarregarNomenclatura
    
    Dim rsado As New ADODB.Recordset
    
    Call Select_Table(True, "tbdNomenclatura", "id_Nomenclatura", "id_Pessoa =" & cFrete.id_Pessoa, , , , , rsado)
    
    If ReadField(rsado, "id_Nomenclatura") > 0 Then
        Call CarregarNomenclaturaCustomizada(cFrete)
    Else
        Call CarregarNomenclaturaPadrao(cFrete)
    End If
    
    rsado.Close
    Set rsado = Nothing
    
    Exit Function
err_CarregarNomenclatura:
    ShowError "Ocorreu um erro ao Carregar os Dados."
    End
End Function

Private Function CarregarNomenclaturaPadrao(ByRef cFrete As clsFrete)
On Error GoTo err_CarregarNomenclaturaPadrao

    Call cFrete.AdicionarNomenclatura(0, "vl_FretePeso", "Frete Peso", "S", Nenhum)
    Call cFrete.AdicionarNomenclatura(0, "vl_Coleta", "Coleta", "S", Nenhum)
    Call cFrete.AdicionarNomenclatura(0, "vl_Entrega", "Entrega", "S", Nenhum)
    Call cFrete.AdicionarNomenclatura(0, "vl_FreteValor", "Frete Valor", "S", Nenhum)
    Call cFrete.AdicionarNomenclatura(0, "vl_GRIS", "GRIS", "S", Nenhum)
    Call cFrete.AdicionarNomenclatura(0, "vl_ITR", "ITR", "S", Nenhum)
    Call cFrete.AdicionarNomenclatura(0, "vl_Despacho", "Despacho", "S", Nenhum)
    Call cFrete.AdicionarNomenclatura(0, "vl_SECCAT", "SEC/CAT", "S", Nenhum)
    Call cFrete.AdicionarNomenclatura(0, "vl_SUFRAMA", "SUFRAMA", "S", Nenhum)
    Call cFrete.AdicionarNomenclatura(0, "vl_Pedagio", "Pedágio", "S", Nenhum)
    Call cFrete.AdicionarNomenclatura(0, "vl_TaxasDiversas", "Taxas Diversas  ", "S", Nenhum)
    
    Exit Function
err_CarregarNomenclaturaPadrao:
    ShowError
End Function

Private Function CarregarNomenclaturaCustomizada(ByRef cFrete As clsFrete)
On Error GoTo err_CarregarNomenclaturaCustomizada

    Dim cNomenclatura As New clsNomenclatura
    Dim rsado As New ADODB.Recordset

    Call Select_Table(True, "tbdNomenclatura", _
                     "id_Nomenclatura,ds_NomeColuna,tp_Enviar,id_Pessoa,ds_NomenclaturaConfigurada,ds_NomenclaturaPadrao,ds_Nomenclatura", _
                     "id_Pessoa =" & cFrete.id_Pessoa, , , , , rsado)
    
    Do While Not rsado.EOF
        Call cFrete.AdicionarNomenclatura(ReadField(rsado, "id_Nomenclatura"), ReadField(rsado, "ds_NomeColuna"), _
                                            ReadField(rsado, "ds_Nomenclatura"), ReadField(rsado, "tp_Enviar"), _
                                            ReadField(rsado, "id_Pessoa"), Nenhum)
        rsado.MoveNext
    Loop
    
    rsado.Close

    Set CarregarPorID = cNomenclatura
    Set cNomenclatura = Nothing
    
    Exit Function
err_CarregarNomenclaturaCustomizada:
    ShowError
End Function

Public Function Salvar(ByRef cFrete As clsFrete) As Boolean
On Error GoTo err_Handler

    Salvar = False

    Dim cServicoBase As New clsServicoBase

    Call CarregarValidacaoFrete(cServicoBase, cFrete)

    With cFrete
        sCampos = "id_Pessoa,vl_FretePeso,vl_Coleta,vl_Entrega,vl_FreteValor,vl_GRIS,vl_ITR,vl_Despacho,vl_SECCAT,vl_SUFRAMA,vl_Pedagio," _
                & "vl_TaxasDiversas,vl_TotalFrete"

        aConteudo = Array(.id_Pessoa, .vl_FretePeso, .vl_Coleta, .vl_Entrega, .vl_FreteValor, .vl_GRIS, .vl_ITR, .vl_Despacho, .vl_SECCAT, .vl_SUFRAMA, .vl_Pedagio, _
                          .vl_TaxasDiversas, .vl_TotalFrete)

        Call cServicoBase.SetarTabela("tbdFrete", "id_Frete", .ID, True, True)
        Call cServicoBase.SetarCampos(sCampos, aConteudo, .StatusGravacao)

        If Not cServicoBase.Salvar Then
            Me.MensagemRetorno = cServicoBase.MensagemRetorno
            Exit Function
        End If
        cFrete.ID = cServicoBase.ValorChavePrimaria
    End With

    Set cServicoBase = Nothing

    Salvar = True

    Exit Function
err_Handler:
    ShowError
End Function

Public Function Excluir(ByVal id_frete As Long) As Boolean
On Error GoTo err_Handler

    Excluir = False

    Dim cServicoBase As New clsServicoBase

    Call cServicoBase.SetarTabela("tbdFrete", "id_Frete", id_frete)
    If Not cServicoBase.Excluir() Then
        MensagemRetorno = cServicoBase.MensagemRetorno
        Exit Function
    End If

    Set cServicoBase = Nothing

    Excluir = True

    Exit Function
err_Handler:
    ShowError
End Function

Public Function RetornaNomenclatura(ByVal cFrete As clsFrete)
On Error GoTo err_RetornaNomenclatura
    
    
    
    Exit Function
err_RetornaNomenclatura:
    ShowError
End Function

Public Function RetornaXMLVprest(cFrete As clsFrete)
On Error GoTo err_RetornaXMLVprest

Exit Function
err_RetornaXMLVprest:
    ShowError
End Function

Public Function FormatarSpreadFrete() As clsESLFormatacaoForm
On Error GoTo err_Handler

    Dim cFormatacaoForm As New clsESLFormatacaoForm

    'Formatação do Select
    cFormatacaoForm.Tabela = "tbdFrete a"
    cFormatacaoForm.Campos = "a.id_Frete,a.vl_FretePeso,a.vl_Coleta,a.vl_Entrega,a.vl_FreteValor,a.vl_GRIS,a.vl_ITR,a.vl_Despacho,a.vl_SECCAT,a.vl_SUFRAMA," _
                           & "a.vl_Pedagio,a.vl_TaxasDiversas,a.vl_TotalFrete"
                            
    cFormatacaoForm.OrderBY = "a.id_Frete"

    cFormatacaoForm.ControlaClickSpread = True

    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.id_Frete", eslnumero, True, True, "id_Frete", "Frete", 10, 10)
    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.vl_FretePeso", eslValor, True, True, "vl_FretePeso", "Frete Peso", 10, 2)
    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.vl_Coleta", eslValor, True, True, "vl_Coleta", "Coleta", 10, 20, 2)
    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.vl_Entrega", eslValor, True, True, "vl_Entrega", "Entrega", 10, 20, 2)
    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.vl_FreteValor", eslValor, True, True, "vl_FreteValor", "Frete Valor", 10, 20, 2)
    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.vl_GRIS", eslValor, True, True, "vl_GRIS", "GRIS", 10, 20, 2)
    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.vl_ITR", eslValor, True, True, "vl_ITR", "ITR", 10, 20, 2)
    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.vl_Despacho", eslValor, True, True, "vl_Despacho", "Despacho", 10, 20, 2)
    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.vl_SECCAT", eslValor, True, True, "vl_SECCAT", "SECCAT", 10, 20, 2)
    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.vl_SUFRAMA", eslValor, True, True, "vl_SUFRAMA", "SUFRAMA", 10, 20, 2)
    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.vl_Pedagio", eslValor, True, True, "vl_Pedagio", "Pedagio", 10, 20, 2)
    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.vl_TaxasDiversas", eslValor, True, True, "vl_TaxasDiversas", "Taxas Diversas", 10, 20, 2)
    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.vl_TotalFrete", eslValor, True, True, "vl_TotalFrete", "Total Frete", 10, 20, 2)

    Set FormatarSpreadFrete = cFormatacaoForm

    Exit Function
err_Handler:
    ShowError
End Function

Public Function FormatarSpreadNomenclatura() As clsESLFormatacaoForm
On Error GoTo err_Handler

    Dim cFormatacaoForm As New clsESLFormatacaoForm

    'Formatação do Select
    cFormatacaoForm.Tabela = "tbdNomenclatura a"
    cFormatacaoForm.Campos = "a.id_Nomenclatura,a.ds_NomeColuna,a.ds_Nomenclatura,a.tp_Enviar,a.id_Pessoa"
    cFormatacaoForm.OrderBY = "a.id_Nomenclatura"

    'Formatação do spread
    cFormatacaoForm.ControlaClickSpread = True

    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.id_Nomenclatura", eslnumero, True, True, "id_Nomenclatura", "Nomenclatura")
    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.ds_NomeColuna", eslTexto, True, True, "ds_NomeColuna", "NomeColuna", 30, 80)
    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.ds_Nomenclatura", eslTexto, True, True, "ds_Nomenclatura", "Nomenclatura", 30, 80)
    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.tp_Enviar", eslTexto, True, True, "tp_Enviar", "Enviar", 1, 1)
    Call cFormatacaoForm.AdicionarColunaSpread(False, False, "a.id_Pessoa", eslnumero, True, True, "id_Pessoa", "Pessoa", 10, 10)


    Set FormatarSpreadNomenclatura = cFormatacaoForm

    Exit Function
err_Handler:
    ShowError
End Function

Private Sub CarregarValidacaoFrete(ByRef cServicoBase As clsServicoBase, ByRef cFrete As clsFrete)
    If cFrete.StatusGravacao <> EnumStatusGravacao.Excluir Then
        Call cServicoBase.AdicionarValidacao(cFrete.id_Pessoa > 0, "Pessoa", "", EnumAcaoMensagem.Bloquear)
        Call cServicoBase.AdicionarValidacao(cFrete.vl_FretePeso > 0, "FretePeso", "", EnumAcaoMensagem.Bloquear)
        Call cServicoBase.AdicionarValidacao(cFrete.vl_Coleta > 0, "Coleta", "", EnumAcaoMensagem.Bloquear)
        Call cServicoBase.AdicionarValidacao(cFrete.vl_Entrega > 0, "Entrega", "", EnumAcaoMensagem.Bloquear)
        Call cServicoBase.AdicionarValidacao(cFrete.vl_FreteValor > 0, "FreteValor", "", EnumAcaoMensagem.Bloquear)
        Call cServicoBase.AdicionarValidacao(cFrete.vl_GRIS > 0, "GRIS", "", EnumAcaoMensagem.Bloquear)
        Call cServicoBase.AdicionarValidacao(cFrete.vl_ITR > 0, "ITR", "", EnumAcaoMensagem.Bloquear)
        Call cServicoBase.AdicionarValidacao(cFrete.vl_Despacho > 0, "Despacho", "", EnumAcaoMensagem.Bloquear)
        Call cServicoBase.AdicionarValidacao(cFrete.vl_SECCAT > 0, "SECCAT", "", EnumAcaoMensagem.Bloquear)
        Call cServicoBase.AdicionarValidacao(cFrete.vl_SUFRAMA > 0, "SUFRAMA", "", EnumAcaoMensagem.Bloquear)
        Call cServicoBase.AdicionarValidacao(cFrete.vl_Pedagio > 0, "Pedagio", "", EnumAcaoMensagem.Bloquear)
        Call cServicoBase.AdicionarValidacao(cFrete.vl_TaxasDiversas > 0, "TaxasDiversas", "", EnumAcaoMensagem.Bloquear)
        Call cServicoBase.AdicionarValidacao(cFrete.vl_TotalFrete > 0, "TotalFrete", "", EnumAcaoMensagem.Bloquear)
    End If
End Sub

