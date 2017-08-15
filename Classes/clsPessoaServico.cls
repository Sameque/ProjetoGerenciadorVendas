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

Public mstrMensagemRetorno          As String

Private Sub Class_Initialize()
On Error GoTo err_Class_Initialize
    
    mstrMensagemRetorno = ""
    
    Exit Sub
err_Class_Initialize:
    mstrMensagemRetorno = "Erro ao inicializar classe."
End Sub

Public Function CarregarPorID(ByVal id_Pesquisa As Long, Optional ByVal blnCarregarFilho = False) As clsPessoa
On Error GoTo err_CarregarDados
    
    Dim cPessoa As New clsPessoa
    Dim rsado As New ADODB.Recordset
    Dim cServisoBase As New clsServicoBase
    
    Set CarregarPorID = Nothing
    
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
        
        If blnCarregarFilho Then
            If Not CarregarContato(cPessoa) Then
                Exit Function
            End If
        End If
        Set CarregarPorID = cPessoa
    Else
        mstrMensagemRetorno = "Pessoa não encontrada."
    End If
    rsado.Close
    
    Set rsado = Nothing
    Set cPessoa = Nothing
    mstrMensagemRetorno = ""
    
    Exit Function
err_CarregarDados:
    ShowError
    mstrMensagemRetorno = "Erro ao carregar classe."
End Function

Public Function CarregarPorCNPJCPF(ByVal strCNPJ As String, Optional ByVal blnCarregarFilho = False) As clsPessoa
On Error GoTo err_Handler
    
    Dim id_Pessoa As Long
    
    id_Pessoa = RetornaIDPorCNPJ(strCNPJ)
    
    Set CarregarPorCNPJCPF = CarregarPorID(id_Pessoa, blnCarregarFilho)
    
    mstrMensagemRetorno = ""

    Exit Function
err_Handler:
    mstrMensagemRetorno "Erro ao carregar por CNPJ"
End Function

Private Function CarregarContato(ByRef cPessoa As clsPessoa) As Boolean
On Error GoTo err_CarregarContato
    
    Dim rsado           As New ADODB.Recordset
    Dim strCampos       As String
    
    strCampos = "id_PessoaContato,id_Pessoa,ds_Nome,cd_Fone,cd_Email"
    CarregarContato = False
    
    Call Select_Table(True, "tbdPessoaContato", strCampos, " id_Pessoa = " & cPessoa.id_Pessoa, , , , , rsado)
        
    Do While Not rsado.EOF
        Call cPessoa.AdicionarContato( _
            ReadField(rsado, "id_PessoaContato"), _
            ReadField(rsado, "ds_Nome"), _
            ReadField(rsado, "cd_Fone"), _
            ReadField(rsado, "cd_Email"))
        rsado.MoveNext
    Loop
    
    rsado.Close
    CarregarContato = True
    mstrMensagemRetorno = ""
    
    Exit Function
err_CarregarContato:
    ShowError
    mstrMensagemRetorno = "Erro ao carregar contatos."
End Function

Public Function Salvar(ByRef cPessoa As clsPessoa) As Boolean
On Error GoTo err_Salvar

    Salvar = False
    
    If cPessoa.StatusGravacao = EnumStatusGravacao.Nenhum Then
        Salvar = True
        Exit Function
    ElseIf cPessoa.StatusGravacao = EnumStatusGravacao.Excluir Then
        Salvar = Excluir(cPessoa)
        Exit Function
    Else
        Salvar = Gravar(cPessoa)
    End If

    If Not SalvarContato(cPessoa) Then
        Exit Function
    End If
    
    If Not Salvar Then
        mstrMensagemRetorno = "Erro no processo de gravação de pessoa."
        Exit Function
    End If
    
     mstrMensagemRetorno = ""
    
    Exit Function
err_Salvar:
'Resume
    mstrMensagemRetorno = "Erro no processo de gravação de pessoa."
End Function

Private Function RetornaIDPorCNPJ(ByVal strCNPJ As String) As Long
On Error GoTo err_Handler

    Dim rsado As New ADODB.Recordset
    
    RetornaIDPorCNPJ = 0
    
    If Trim(cd_CNPJCPFPesquisar) <> "" Then
        Call Select_Table(True, "tbdPessoa", "id_Pessoa", "cd_CNPJCPF = '" & strCNPJ & "'", , , , , rsado)
        RetornaIDPorCNPJ = ReadField(rsado, "id_Pessoa")
        rsado.Close
    End If
    
    Set rsado = Nothing
    mstrMensagemRetorno = ""
    
    Exit Function
err_Handler:
    mstrMensagemRetorno = "Erro ao carregar por CNPJ."
End Function

Private Function SalvarContato(cPessoa As clsPessoa) As Boolean
On erro GoTo err_SalvarContato
    Dim cServicoBase As New clsServicoBase
    Dim cPessoaContato As clsPessoaContato

    SalvarContato = False

    For Each cPessoaContato In cPessoa.GetListaContatos
        With cPessoaContato
            Call cServicoBase.SetarTabela(.Tabela, .ChavePrimaria, .ID, True, True)
            Call cServicoBase.SetarCampos(.Campos, .Conteudo, .StatusGravacao)
            
            If Not cServicoBase.Salvar Then
                mstrMensagemRetorno = cServicoBase.MensagemRetorno
                Exit Function
            End If
            '///////implementar para matar a classe e removela da collecion
            .ID = cServicoBase.ValorChavePrimaria
            SalvarContato = True
            
        End With
    Next
    
    SalvarContato = True
    mstrMensagemRetorno = ""
    Exit Function
err_SalvarContato:
    mstrMensagemRetorno = "Erro no processo de gravação do contato."
End Function

Private Function Gravar(ByRef cPessoa As clsPessoa) As Boolean
On Error GoTo err_Gravar
    Dim cServicoBase As New clsServicoBase
    
    Gravar = False
    
    Call CarregarValidacaoPessoa(cServicoBase, cPessoa)
    
    With cPessoa
        Call cServicoBase.SetarTabela(.Tabela, .ChavePrimaria, .id_Pessoa, False, True)
        Call cServicoBase.SetarCampos(.Campos, .Conteudo, .StatusGravacao)
    End With
                        
    If Not cServicoBase.Salvar Then
        mstrMensagemRetorno = cServicoBase.MensagemRetorno
        Exit Function
    End If
    
    cPessoa.id_Pessoa = cServicoBase.ValorChavePrimaria
    Gravar = True
    Set cServicoBase = Nothing
    mstrMensagemRetorno = ""
    
    Exit Function
err_Gravar:
    mstrMensagemRetorno = "Erro no processo de gravação."
End Function

Private Function GravarContato(ByRef cPessoa As clsPessoa) As Boolean
On Error GoTo err_GravarContato
    Dim cPessoaContato  As clsPessoaContato
    Dim cServicoBase As New clsServicoBase
    
    GravarContato = False
    
    For Each cPessoaContato In cPessoa.colPessoaContato
        With cPessoaContato
            .id_Pessoa = cPessoa.id_Pessoa
            Call cServicoBase.SetarTabela("tbdPessoaContato", "id_PessoaContato", .id_PessoaContato, False, True)
            Call cServicoBase.SetarCampos(.Campos, .Conteudo, .StatusGravacao)
        End With
        
        If cServicoBase.Salvar Then
            GravarContato = True
            cPessoaContato.id_PessoaContato = cServicoBase.ValorChavePrimaria
        Else
            mstrMensagemRetorno = cServicoBase.MensagemRetorno
        End If
    Next
    
    Set cPessoaContato = Nothing
    Set cServicoBase = Nothing
    mstrMensagemRetorno = ""
    
    Exit Function
err_GravarContato:
    mstrMensagemRetorno = "Erro ao gravar contato."
End Function

Private Function Excluir(ByRef cPessoa As clsPessoa) As Boolean
On Error GoTo err_Excluir
    Dim cServicoBase As New clsServicoBase
    Dim str As String
    
    Excluir = False
    
    With cPessoa
        Call cServicoBase.SetarTabela(.Tabela, .ChavePrimaria, .id_Pessoa, True, True)
    End With

    If Not cServicoBase.Excluir Then
        mstrMensagemRetorno = cServicoBase.MensagemRetorno
        Exit Function
    End If
    
    Excluir = True
    Set cServicoBase = Nothing
    Set cPessoa = Nothing
    mstrMensagemRetorno = ""
    
    Exit Function
err_Excluir:
    mstrMensagemRetorno = "Erro ao excluir pessoa."
End Function

'Private Function ExcluirContato(ByRef cPessoa As clsPessoa, Optional ByVal blnExcluirTodos As Boolean = False) As Boolean
'On Error GoTo err_ExcluirContato
'    Dim cPessoaContato  As clsPessoaContato
'    Dim cServicoBase    As New clsServicoBase
'
'    ExcluirContato = False
'
'    For Each cPessoaContato In cPessoa.colPessoaContato
'        With cPessoaContato
'            .id_Pessoa = cPessoa.id_Pessoa
'            Call cServicoBase.SetarTabela(.Tabela, .ChavePrimaria, .ID, False, True)
'            Call cServicoBase.SetarCampos(.Campos, .Conteudo, IIf(blnExcluirTodos, EnumStatusGravacao.Excluir, .StatusGravacao))
'        End With
'
'        If cServicoBase.Salvar Then
'            ExcluirContato = True
'        Else
'            mstrMensagemRetorno = cServicoBase.MensagemRetorno
'            Exit Function
'        End If
'    Next
'
'    Set cPessoaContato = Nothing
'    Set cServicoBase = Nothing
'    mstrMensagemRetorno = ""
'
'    Exit Function
'err_ExcluirContato:
'    mstrMensagemRetorno = "Erro ao excluir contato"
'End Function

Public Function FormatarSpreadPessoaContato() As clsESLFormatacaoForm
On Error GoTo err_FormatarSpreadPessoaContato

    Dim cFormatacaoForm As New clsESLFormatacaoForm
    Dim cPessoaContato As New clsPessoaContato
    
    cFormatacaoForm.Tabela = cPessoaContato.Tabela
    cFormatacaoForm.Campos = cPessoaContato.ChavePrimaria & ", " & cPessoaContato.Campos
    cFormatacaoForm.OrderBY = "ds_Nome"

    cFormatacaoForm.ControlaClickSpread = True
    cFormatacaoForm.ControlaQueryAdvance = True
    cFormatacaoForm.ExcluirRegistro = True

    Call cFormatacaoForm.AdicionarColunaSpread(True, False, "id_PessoaContato", eslNumero, True, True, "id_PessoaContato", "id_PessoaContato")
    Call cFormatacaoForm.AdicionarColunaSpread(True, False, "id_Pessoa", eslNumero, True, True, "id_Pessoa", "id_Pessoa")
    Call cFormatacaoForm.AdicionarColunaSpread(True, False, "ds_Nome", eslTexto, False, False, "ds_Nome", "Nome", 20, 50)
    Call cFormatacaoForm.AdicionarColunaSpread(True, False, "nr_Telefone", eslTexto, False, True, "nr_Telefone", "Telefone", 20, 30)
    Call cFormatacaoForm.AdicionarColunaSpread(True, False, "cd_Email", eslTexto, False, True, "ds_Email", "Email", 25, 100)
    
    Set FormatarSpreadPessoaContato = cFormatacaoForm
    mstrMensagemRetorno = ""
    
    Exit Function
err_FormatarSpreadPessoaContato:
    mstrMensagemRetorno = "Erro ao formatar o spread"
End Function

Private Sub CarregarValidacaoPessoa(cServicoBase As clsServicoBase, cPessoa As clsPessoa)
    If cPessoa.StatusGravacao <> EnumStatusGravacao.Excluir And cPessoa.StatusGravacao <> EnumStatusGravacao.Nenhum Then
        Call cServicoBase.AdicionarValidacao(Trim(cPessoa.ds_Pessoa) <> "", "Nome", "", EnumAcaoMensagem.Bloquear)
        Call cServicoBase.AdicionarValidacao(Trim(cPessoa.cd_cnpjcpf) <> "", "CNPJ/CPF", "", EnumAcaoMensagem.Bloquear)
        Call cServicoBase.AdicionarValidacao(Trim(cPessoa.ds_RazaoSocial) <> "", "Razão Social", "", EnumAcaoMensagem.Alertar)
    End If
End Sub
