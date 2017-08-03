VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPessoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id_Pessoa                As Long
Public id_Cidade                As Long

Public cd_cnpjcpf               As String
Public cd_CEP                   As String

Public ds_Pessoa                As String
Public ds_RazaoSocial           As String
Public ds_Endereco              As String
Public ds_Bairro                As String

Public tp_Cliente               As String
Public tp_Fornecedor            As String
Public tp_Funcionario           As String

Public colPessoaContato         As Collection
Public MensagemRetorno          As String

Public Function CarregarDados(id_Pesquisa As Long) As Boolean
On Error GoTo err_CarregarDados

    Dim rsado As New ADODB.Recordset
    
    CarregarDados = False
    
    Call Class_Initialize
    
    Call Select_Table(True, "tbdPessoa a", "*", "a.id_Pessoa = " & id_Pesquisa, , , , , rsado)
    If Not rsado.EOF Then
        CarregarDados = True
        
        id_Pessoa = id_Pesquisa
        id_Cidade = ReadField(rsado, "id_Cidade")
        
        cd_CEP = ReadField(rsado, "cd_CEP")
        cd_cnpjcpf = ReadField(rsado, "cd_cnpjcpf")
        
        ds_Pessoa = ReadField(rsado, "ds_Pessoa")
        ds_RazaoSocial = ReadField(rsado, "ds_RazaoSocial")
        ds_Endereco = ReadField(rsado, "ds_Endereco")
        ds_Bairro = ReadField(rsado, "ds_Bairro")
        
        tp_Cliente = ReadField(rsado, "tp_Cliente")
        tp_Fornecedor = ReadField(rsado, "tp_Fornecedor")
        tp_Funcionario = ReadField(rsado, "tp_Funcionario")
        
        Call CarregarContato
        
    End If
    rsado.Close
    Set rsado = Nothing
    
    If Not CarregarDados Then
        MensagemRetorno = "Erro ao carregar classe."
        Exit Function
    End If
    
    Exit Function
err_CarregarDados:
    MensagemRetorno = "Erro ao carregar classe."
End Function

Private Function CarregarContato() As Boolean
On Error GoTo err_CarregarContato
    
    Dim rsado As New ADODB.Recordset
    Dim cPessoaContato As clsPessoaContato
    
    Call Select_Table(True, "tbdPessoaContato", "id_PessoaContato", "id_Pessoa = " & id_Pessoa, , , , , rsado)
    
    Do While Not rsado.EOF
        Set cPessoaContato = New clsPessoaContato
        
        If Not cPessoaContato.CarregarDados(ReadField(rsado, "id_PessoaContato")) Then
            MensagemRetorno = "Erro ao carregar contato!"
        End If
        
        Call colPessoaContato.Add(colPessoaContato)
        
        Set cPessoaContato = Nothing
        rsado.MoveNext
    Loop
    
    rsado.Close
    CarregarContato = True
    Set rsado = Nothing
        
    Exit Function
err_CarregarContato:
    MensagemRetorno = "Erro ao carregar contatos."
End Function

Public Function Gravar() As Boolean
On Error GoTo err_Gravar

    If id_Pessoa > 0 Then
        Gravar = Alterar
    Else
        Gravar = Incluir
    End If
    
    Call GravarContato
    
    Exit Function
err_Gravar:
    ShowError
End Function

Private Function Incluir() As Boolean
On Error GoTo err_Incluir
    
    Incluir = False
    
    id_Pessoa = Insert_Table("tbdPessoa", "id_Pessoa", Campos, Conteudo)
    
    If id_Pessoa = -1 Then
        MensagemRetorno = "Erro ao incluir pessoa."
        Exit Function
    End If
    
    Incluir = True
    
    Exit Function
err_Incluir:
    MensagemRetorno = "Erro ao incluir pessoa."
End Function

Private Function Alterar() As Boolean
On Error GoTo err_Alterar

    Alterar = False

    If Update_Table("tbdPessoa", Campos, Conteudo, "id_Pessoa = " & id_Pessoa) = -1 Then
        MensagemRetorno = "Erro ao atualizar pessoa."
        Exit Function
    End If

    Alterar = True

    Exit Function
err_Alterar:
    MensagemRetorno = "Erro ao atualizar pessoa."
End Function

Private Function GravarContato() As Boolean
On Error GoTo err_GravarContato
    Dim cPessoaContato As New clsPessoaContato
    
    GravarContato = False
    
    For Each cPessoaContato In colPessoaContato
        cPessoaContato.id_Pessoa = id_Pessoa
        cPessoaContato.Gravar
        Set cPessoaContato = Nothing
    Next
    
    GravarContato = True
    
    Exit Function
err_GravarContato:
    MensagemRetorno = "Erro ao gravar contato"
End Function

Public Function Excluir() As Boolean
On Error GoTo err_Excluir

    Excluir = False
    
    Call ExcluirContato
    
    If Delete_Table("tbdPessoa", "id_Pessoa = " & id_Pessoa) = -1 Then
        MensagemRetorno = "Erro ao excluir pessoa."
        Exit Function
    End If

    Excluir = True

    Exit Function
err_Excluir:
    MensagemRetorno = "Erro ao excluir pessoa."
End Function

Private Function ExcluirContato() As Boolean
On Error GoTo err_ExcluirContato

    Dim cPessoaContato As New clsPessoaContato
    
    ExcluirContato = False
    
    For Each cPessoaContato In colPessoaContato
        cPessoaContato.Excluir
        Set cPessoaContato = Nothing
    Next
    
    
    ExcluirContato = True
    
    Exit Function
err_ExcluirContato:
    ShowError
    MensagemRetorno = "Erro ao excluir contato"
End Function


Private Function Campos() As String
    Campos = "cd_cnpjcpf, ds_Pessoa, ds_RazaoSocial, ds_Endereco, ds_Bairro, id_Cidade, cd_CEP, tp_Cliente, tp_Fornecedor, tp_Funcionario"
End Function

Private Function Conteudo()
    Conteudo = Array(cd_cnpjcpf, ds_Pessoa, ds_RazaoSocial, ds_Endereco, ds_Bairro, id_Cidade, cd_CEP, tp_Cliente, tp_Fornecedor, tp_Funcionario)
End Function

Public Function AdicionarContato(ds_Nome As String, cd_Fone As String, cd_Email As String) As Boolean
On Error GoTo err_AdicionarContato
    Dim cPessoaContato As New clsPessoaContato
    
    AdicionarContato = False
    
    cPessoaContato.id_Pessoa = id_Pessoa
    cPessoaContato.ds_Nome = ds_Nome
    cPessoaContato.cd_Fone = cd_Fone
    cPessoaContato.cd_Email = cd_Email
    
    Call colPessoaContato.Add(cPessoaContato)
    Set cPessoaContato = Nothing
    AdicionarContato = True
        
    Exit Function
err_AdicionarContato:
    MensagemRetorno = "Erro ao adicionar contato."
End Function
Private Sub Class_Initialize()
On Error GoTo err_Class_Initialize

    id_Pessoa = 0
    id_Cidade = 0
    cd_cnpjcpf = ""
    cd_CEP = ""
    ds_Pessoa = ""
    ds_RazaoSocial = ""
    ds_Endereco = ""
    ds_Bairro = ""
    tp_Cliente = ""
    tp_Fornecedor = ""
    tp_Funcionario = ""
    
    Set colPessoaContato = New Collection
    Exit Sub
err_Class_Initialize:
    MensagemRetorno = "Erro ao inicializar classe."
End Sub

Private Sub Class_Terminate()
    Set colPessoaContato = Nothing
End Sub
