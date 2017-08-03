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

Public Function CarregarId(ByVal id_Pesquisa As Long) As clsPessoa
On Error GoTo err_CarregarDados
    
    Dim cPessoa As New clsPessoa
    Dim rsado As New ADODB.Recordset
    
    Set CarregarId = Nothing
    
    Call Class_Initialize
    
    Call Select_Table(True, "tbdPessoa a", "*", "a.id_Pessoa = " & id_Pesquisa, , , , , rsado)
    If Not rsado.EOF Then
        With cPessoa
            .id_Pessoa = ReadField(rsado, "id_Pessoa")
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
        Set CarregarId = cPessoa
    Else
        strMensagemRetorno = "Pessoa não encontrada."
    End If
    rsado.Close
    
    Set rsado = Nothing
    'Set cPessoa = Nothing
        
    Exit Function
err_CarregarDados:
    ShowError
    strMensagemRetorno = "Erro ao carregar classe."
End Function

Public Function CarregarFiltro(ByVal strWhere As String) As Collection
On Error GoTo err_CarregarDados
    
    Dim cPessoa As New clsPessoa
    Dim colPessoa As New Collection
    Dim rsado As New ADODB.Recordset
    
    Set CarregarFiltro = Nothing
    
    Call Class_Initialize
    
    Call Select_Table(True, "tbdPessoa", "*", strWhere, , , , , rsado)
    Do While Not rsado.EOF
    
        With cPessoa
            .id_Pessoa = ReadField(rsado, "id_Pessoa")
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
        Call colPessoa.Add(cPessoa)
        rsado.MoveNext
    Loop
    rsado.Close
    
    Set rsado = Nothing
    Set cPessoa = Nothing
    
    Exit Function
err_CarregarDados:
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
            .blnGravar = False
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
Resume
    strMensagemRetorno = "Erro ao carregar contatos."
End Function

Public Function Gravar(ByRef cPessoa As clsPessoa) As Boolean
On Error GoTo err_Gravar

    If cPessoa.id_Pessoa > 0 Then
        Gravar = Alterar(cPessoa)
    Else
        Gravar = Incluir(cPessoa)
    End If
    
    Call GravarContato(cPessoa)
    
    Exit Function
err_Gravar:
    ShowError
End Function

Private Function Incluir(ByRef cPessoa As clsPessoa) As Boolean
On Error GoTo err_Incluir
    
    Incluir = False
    
    cPessoa.id_Pessoa = Insert_Table("tbdPessoa", "id_Pessoa", cPessoa.Campos, cPessoa.Conteudo)
    
    If cPessoa.id_Pessoa = -1 Then
        strMensagemRetorno = "Erro ao incluir pessoa."
        Exit Function
    End If
    
    Call GravarContato(cPessoa)
    
    Incluir = True
    
    Exit Function
err_Incluir:
    strMensagemRetorno = "Erro ao incluir pessoa."
End Function

Private Function Alterar(ByRef cPessoa As clsPessoa) As Boolean
On Error GoTo err_Alterar

    Alterar = False

    If Update_Table("tbdPessoa", cPessoa.Campos, cPessoa.Conteudo, "id_Pessoa = " & cPessoa.id_Pessoa) = -1 Then
        strMensagemRetorno = "Erro ao atualizar pessoa."
        Exit Function
    End If

    Alterar = True

    Exit Function
err_Alterar:
    strMensagemRetorno = "Erro ao atualizar pessoa."
End Function

Private Function GravarContato(ByRef cPessoa As clsPessoa) As Boolean
On Error GoTo err_GravarContato
    Dim cPessoaContato  As New clsPessoaContato
    
    GravarContato = False
    
    For Each cPessoaContato In cPessoa.colPessoaContato
        If cPessoaContato.blnGravar Then
                cPessoaContato.id_Pessoa = cPessoa.id_Pessoa
        
                If cPessoaContato.id_PessoaContato > 0 Then
                    GravarContato = AlterarContato(cPessoaContato)
                Else
                    GravarContato = IncluirContato(cPessoaContato)
                End If
        End If
    Next
        
    Set cPessoaContato = Nothing
    GravarContato = True
    
    Exit Function
err_GravarContato:
    strMensagemRetorno = "Erro ao gravar contato"
End Function

Private Function IncluirContato(ByRef cPessoaContato As clsPessoaContato) As Boolean
On Error GoTo err_IncluirContato
    
    IncluirContato = False
    
    cPessoaContato.id_PessoaContato = Insert_Table("tbdPessoaContato", "id_PessoaContato", cPessoaContato.Campos, cPessoaContato.Conteudo)
        
    If cPessoaContato.id_PessoaContato = -1 Then
        strMensagemRetorno = "Erro ao incluir contato."
        Exit Function
    End If
                    
    IncluirContato = True
    
    Exit Function
err_IncluirContato:
    ShowError
End Function

Private Function AlterarContato(cPessoaContato As clsPessoaContato) As Boolean
On Error GoTo err_Alterar

    AlterarContato = False

    If Update_Table("tbdPessoaContato", cPessoaContato.Campos, cPessoaContato.Conteudo, "id_PessoaContato = " & cPessoaContato.id_PessoaContato) = -1 Then
        strMensagemRetorno = "Erro ao atualizar contato."
        Exit Function
    End If

    AlterarContato = True

    Exit Function
err_Alterar:
    ShowError
End Function

Public Function Excluir(ByRef cPessoa As clsPessoa) As Boolean
On Error GoTo err_Excluir

    Excluir = False
    
    Call ExcluirContato(cPessoa, True)
    
    If Delete_Table("tbdPessoa", "id_Pessoa = " & cPessoa.id_Pessoa) = -1 Then
        strMensagemRetorno = "Erro ao excluir pessoa."
        Exit Function
    End If

    Excluir = True

    Exit Function
err_Excluir:
    strMensagemRetorno = "Erro ao excluir pessoa."
End Function

Private Function ExcluirContato(ByRef cPessoa As clsPessoa, Optional ByVal blnExcluirTodos As Boolean) As Boolean
On Error GoTo err_ExcluirContato

    Dim cPessoaContato  As New clsPessoaContato
    Dim lngIndex        As Long
    
    ExcluirContato = False
    
    For Each cPessoaContato In cPessoa.colPessoaContato
        If cPessoaContato.blnExcluir Or blnExcluirTodos Then
        
            If Delete_Table("tbdPessoaContato", "id_PessoaContato = " & cPessoaContato.id_PessoaContato) = -1 Then
                strMensagemRetorno = "Erro ao excluir contato."
                Exit Function
            End If
        End If
    Next
    
    Set cPessoaContato = Nothing
    
    ExcluirContato = True
    
    Exit Function
err_ExcluirContato:
    ShowError
    strMensagemRetorno = "Erro ao excluir contato"
End Function
