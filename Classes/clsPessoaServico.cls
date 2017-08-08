VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPessoaServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public strMensagemRetorno          As String

Private Sub Class_Initialize()
On Error GoTo err_Class_Initialize
    
    strMensagemRetorno = ""
    
    Exit Sub
err_Class_Initialize:
    strMensagemRetorno = "Erro ao inicializar classe."
End Sub

Public Function CarregarClasse(ByVal id_Pesquisa As Long) As clsPessoa
On Error GoTo err_CarregarDados
    
    Dim cPessoa As New clsPessoa
    Dim rsado As New ADODB.Recordset
    Dim cServisoBase As New clsServicoBase
    
    Set CarregarClasse = Nothing
    
    Call Class_Initialize
    
    Call Select_Table(True, "tbdPessoa a", cPessoa.Campos, "a.id_Pessoa = " & id_Pesquisa, , , , , rsado)
    If Not rsado.EOF Then
        With cPessoa
            .id_Pessoa = id_Pesquisa
            .id_Cidade = ReadField(rsado, "id_Cidade")
            .cd_CEP = ReadField(rsado, "cd_CEP")
            .cd_cnpjcpf = ReadField(rsado, "cd_cnpjcpf")
            .ds_Pessoa = ReadField(rsado, "ds_Pessoa")
            .ds_RazaoSocial = ReadField(rsado, "ds_RazaoSocial")
            .ds_Endereco = ReadField(rsado, "ds_Endereco")
            .ds_Bairro = ReadField(rsado, "ds_Bairro")
            .tp_Cliente = ReadField(rsado, "tp_Cliente")
            .tp_Fornecedor = ReadField(rsado, "tp_Fornecedor")
            .tp_Funcionario = ReadField(rsado, "tp_Funcionario")
        End With
        
        Call CarregarContato(cPessoa)
        Set CarregarClasse = cPessoa
    Else
        strMensagemRetorno = "Pessoa não encontrada."
    End If
    rsado.Close
    
    Set rsado = Nothing
    Set cPessoa = Nothing
        
    Exit Function
err_CarregarDados:
    ShowError
    strMensagemRetorno = "Erro ao carregar classe."
End Function

Private Function CarregarContato(ByRef cPessoa As clsPessoa) As Boolean
On Error GoTo err_CarregarContato
    
    Dim rsado           As New ADODB.Recordset
    Dim cPessoaContato  As New clsPessoaContato
    Dim strCampos       As String
    
    strCampos = "id_PessoaContato,id_Pessoa,ds_Nome,cd_Fone,cd_Email"
    CarregarContato = False
    
    Call Select_Table(True, "tbdPessoaContato", strCampos, "id_Pessoa = " & cPessoa.id_Pessoa, , , , , rsado)
        
    Do While Not rsado.EOF
        With cPessoaContato
            .id_PessoaContato = ReadField(rsado, "id_PessoaContato")
            .id_Pessoa = ReadField(rsado, "id_Pessoa")
            .ds_Nome = ReadField(rsado, "ds_Nome")
            .cd_Fone = ReadField(rsado, "cd_Fone")
            .cd_Email = ReadField(rsado, "cd_Email")
            .enuGravacao = EnumStatusGravacao.Nenhum
        End With
        
        Call cPessoa.colPessoaContato.Add(cPessoaContato)
        Set cPessoaContato = Nothing

        rsado.MoveNext
    Loop
    
    rsado.Close
    CarregarContato = True
    Set cPessoaContato = Nothing
        
    If Not CarregarContato Then
        strMensagemRetorno = "Erro ao carregar classe."
        Exit Function
    End If
        
    Exit Function
err_CarregarContato:
    ShowError
    strMensagemRetorno = "Erro ao carregar contatos."
End Function

Public Function Salvar(ByRef cPessoa As clsPessoa) As Boolean
On Error GoTo err_Salvar

    Call AbreTransacao
    If cPessoa.enuGravacao = EnumStatusGravacao.Nenhum Then
        Exit Function
    ElseIf cPessoa.enuGravacao = EnumStatusGravacao.Excluir Then
        Salvar = Excluir(cPessoa)
    Else
        Salvar = Gravar(cPessoa)
    End If

    If Not Salvar Then
        strMensagemRetorno = "Erro no processo de gravação de pessoa."
        Exit Sub
    End If
    
    Call SalvarContato(cPessoa)
    Call FechaTransacao

    Exit Sub
err_Salvar:
    strMensagemRetorno = "Erro no processo de gravação de pessoa."
End Sub

Private Function SalvarContato(cPessoa As clsPessoa) As Boolean
On erro GoTo err_SalvarContato
    Dim cServicoBase As New clsServicoBase
    Dim cPessoaContato As New clsPessoaContato

    SalvarContato = False

    For Each cPessoaContato In cPessoa.colPessoaContato
        With cPessoaContato
            Call cServicoBase.SetarTabela(.Tabela, .ChavePrimaria, .id_Pessoa, False, True)
            Call cServicoBase.SetarCampos(.Campos, .Conteudo, .enuGravacao)
            
            If cServicoBase.Salvar Then
                SalvarContato = True
                .id_Pessoa = cServicoBase.ValorChavePrimaria
            Else
                strMensagemRetorno = cServicoBase.MensagemRetorno
            End If
        End With
        
    Next
    
    Exit Function
err_SalvarContato:
    strMensagemRetorno = "Erro no processo de gravação do contato."
End Function

Private Function Gravar(ByRef cPessoa As clsPessoa) As Boolean
On Error GoTo err_Gravar
    Dim cServicoBase As New clsServicoBase
    
    Gravar = False
    
    With cPessoa
        Call cServicoBase.SetarTabela(.Tabela, .ChavePrimaria, .id_Pessoa, False, True)
        Call cServicoBase.SetarCampos(.Campos, .Conteudo, .enuGravacao)
        
        If cServicoBase.Salvar Then
            Gravar = True
            cPessoa.id_Pessoa = cServicoBase.ValorChavePrimaria
        Else
            strMensagemRetorno = cServicoBase.MensagemRetorno
        End If
        
    End With
    
    Set cServicoBase = Nothing

    Exit Function
err_Gravar:
    strMensagemRetorno = "Erro no processo de gravação."
End Function

Private Function GravarContato(ByRef cPessoa As clsPessoa) As Boolean
On Error GoTo err_GravarContato
    Dim cPessoaContato  As New clsPessoaContato
    Dim cServicoBase As New clsServicoBase
    
    GravarContato = False
    
    For Each cPessoaContato In cPessoa.colPessoaContato
        With cPessoaContato
            .id_Pessoa = cPessoa.id_Pessoa
            Call cServicoBase.SetarTabela("tbdPessoaContato", "id_PessoaContato", .id_PessoaContato, False, True)
            Call cServicoBase.SetarCampos(.Campos, .Conteudo, .enuGravacao)
        End With
        
        If cServicoBase.Salvar Then
            GravarContato = True
            cPessoaContato.id_PessoaContato = cServicoBase.ValorChavePrimaria
        Else
            strMensagemRetorno = cServicoBase.MensagemRetorno
        End If
    Next
    
    Set cPessoaContato = Nothing
    Set cServicoBase = Nothing
    
    Exit Function
err_GravarContato:
    strMensagemRetorno = "Erro ao gravar contato."
End Function

Private Function Excluir(ByRef cPessoa As clsPessoa) As Boolean
On Error GoTo err_Excluir
    Dim cServicoBase As New clsServicoBase

    Call ExcluirContato(cPessoa, True)
    'TIRAR AS TABELAS E OS CAMPOS P VER SE EXCLUI
    With cPessoa
        Call cServicoBase.SetarTabela(.Tabela, .ChavePrimaria, .id_Pessoa, False, True)
        Call cServicoBase.SetarCampos(.Campos, .Conteudo, .enuGravacao)
        Excluir = cServicoBase.Salvar
    End With
    
    If Not GravarContato(cPessoa) Then
        strMensagemRetorno = "Erro ao excluir."
        Call VoltaTransacao
    End If
    
    Set cServicoBase = Nothing
    
    Exit Function
err_Excluir:
    strMensagemRetorno = "Erro ao excluir."
End Function

Private Function ExcluirContato(ByRef cPessoa As clsPessoa, Optional ByVal blnExcluirTodos As Boolean) As Boolean
On Error GoTo err_ExcluirContato
    Dim cPessoaContato  As New clsPessoaContato
    Dim cServicoBase As New clsServicoBase
    
    ExcluirContato = False
    
    'TIRAR AS TABELAS E OS CAMPOS P VER SE EXCLUI
    For Each cPessoaContato In cPessoa.colPessoaContato
        With cPessoaContato
            .id_Pessoa = cPessoa.id_Pessoa
            Call cServicoBase.SetarTabela(.Tabela, .ChavePrimaria, .id_PessoaContato, False, True)
            Call cServicoBase.SetarCampos(.Campos, .Conteudo, .enuGravacao)
        End With
        
        If cServicoBase.Salvar Then
            ExcluirContato = True
            cPessoaContato.id_PessoaContato = cServicoBase.ValorChavePrimaria
        Else
            strMensagemRetorno = cServicoBase.MensagemRetorno
        End If
    Next
    
    Set cPessoaContato = Nothing
    Set cServicoBase = Nothing
    
err_ExcluirContato:
    ShowError
    strMensagemRetorno = "Erro ao excluir contato"
End Function

Public Function FormatarSpreadPessoaContato() As clsESLFormatacaoForm
On Error GoTo err_FormatarSpreadPessoaContato

    Dim cFormatacaoForm As New clsESLFormatacaoForm
    Dim cPessoaContato As New clsPessoaContato
    
    'Formatação do Select
    cFormatacaoForm.Tabela = cPessoaContato.Tabela
    cFormatacaoForm.Campos = cPessoaContato.ChavePrimaria & ", " & cPessoaContato.Campos
    cFormatacaoForm.OrderBY = "ds_Nome"

    'Formatação do spread
    cFormatacaoForm.ControlaClickSpread = True
    cFormatacaoForm.ControlaQueryAdvance = True
    cFormatacaoForm.ExcluirRegistro = True

    Call cFormatacaoForm.AdicionarColunaSpread(True, False, "id_PessoaContato", eslnumero, True, True, "id_PessoaContato", "PessoaContato")
    Call cFormatacaoForm.AdicionarColunaSpread(True, False, "ds_Nome", eslTexto, False, False, "ds_Nome", "Nome", 20, 50)
    Call cFormatacaoForm.AdicionarColunaSpread(True, False, "nr_Telefone", eslTexto, False, True, "nr_Telefone", "Telefone", 20, 30)
    Call cFormatacaoForm.AdicionarColunaSpread(True, False, "cd_Email", eslTexto, False, True, "ds_Email", "Email", 25, 100)
    
    Set FormatarSpreadPessoaContato = cFormatacaoForm

    Exit Function
err_FormatarSpreadPessoaContato:
    ShowError
End Function

Private Sub CarregarValidacaoPessoa(cServicoBase As clsServicoBase, cPessoa As clsPessoa)
    If cPessoa.StatusGravacao <> EnumStatusGravacao.Excluir Then
        Call cServicoBase.AdicionarValidacao(Trim(cPessoa.cd_cnpjcpf) <> "", "CNPJ/CPF", "", EnumAcaoMensagem.Bloquear)
        Call cServicoBase.AdicionarValidacao(Trim(cPessoa.ds_Nome) <> "", "Nome Fantasia", "", EnumAcaoMensagem.Bloquear)
        Call cServicoBase.AdicionarValidacao(Trim(cPessoa.ds_RazaoSocial) <> "", "Razão Social", "", EnumAcaoMensagem.Bloquear)
    End If
End Sub
