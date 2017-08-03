Attribute VB_Name = "Local"
Option Explicit

Public Const lSistemaModulo = 30
Public Const sSistemaNome = "SISLOG"

Public Enum EnumTipoComboSpread
    Nenum = 0
    ItemManual = 1
'    Pessoa = 2
'    Empresa = 3
'    Transportadora = 4
'    Funcionario = 5
'    Cidade = 6
'    Estado = 7
End Enum

Public Enum EnumTipoDBComboPadrao
    Pessoa_ComCNPJ = 1
'    Pessoa_SemCNPJ = 2
'    Empresa_ComCNPJ = 3
'    Empresa_SemCNPJ = 4
'    Transportadora_ComCNPJ = 5
'    Transportadora_SemCNPJ = 6
'    Funcionario = 7
'    Cidade = 8
'    Estado = 9
End Enum


Public Function PegarFormatacaoComboSpread(nr_TipoCombo As EnumTipoComboSpread, Optional sWhereComplementar As String = "") As clsESLConfiguracaoCombo
On Error GoTo err_PegarFormatacaoComboSpread
    
    Dim cConfiguracaoCombo As New clsESLConfiguracaoCombo
    
    'If nr_TipoCombo = EnumTipoComboSpread.Empresa Then
        'cConfiguracaoCombo.Tabela = "tbd20Filial a inner join tbd20Pessoa b on a.id_Filial = b.id_Pessoa"
        'cConfiguracaoCombo.CampoID = "b.id_Pessoa"
        'cConfiguracaoCombo.CampoDS = "b.ds_Pessoa"
        'cConfiguracaoCombo.CarregarSemFiltro = True
    'End If
    
    'If cConfiguracaoCombo.Tabela <> "" Then
    '    cConfiguracaoCombo.Where = cConfiguracaoCombo.Where & IIf(cConfiguracaoCombo.Where <> "", " and ", "") & sWhereComplementar
    'End If
    
    Set PegarFormatacaoComboSpread = cConfiguracaoCombo
    
    Exit Function
err_PegarFormatacaoComboSpread:
    ShowError "PegarFormatacaoComboSpread()" & vbCrLf
End Function


Public Function RetornaConfigNaoSomarICMSISSPorTabelaPreco(id_TransportadoraServico As Long)

End Function
