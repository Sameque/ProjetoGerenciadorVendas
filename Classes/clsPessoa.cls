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

Public enuGravacao              As EnumStatusGravacao
Public colPessoaContato         As Collection
Public strMensagemRetorno       As String

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
    
    enuGravacao = Nenhum
    strMensagemRetorno = ""
    Set colPessoaContato = New Collection
    Exit Sub
err_Class_Initialize:
    strMensagemRetorno = "Erro ao inicializar classe."
End Sub

Public Function Campos() As String
    Campos = "cd_cnpjcpf, ds_Pessoa, ds_RazaoSocial, ds_Endereco, ds_Bairro, id_Cidade, cd_CEP, tp_Cliente, tp_Fornecedor, tp_Funcionario"
End Function

Public Function Conteudo()
    Conteudo = Array(cd_cnpjcpf, ds_Pessoa, ds_RazaoSocial, ds_Endereco, ds_Bairro, id_Cidade, cd_CEP, tp_Cliente, tp_Fornecedor, tp_Funcionario)
End Function

Public Function AdicionarContato(ByVal ds_Nome As String, ByVal cd_Fone As String, ByVal cd_Email As String)
On Error GoTo err_AdicionarContato
    Dim cPessoaContato As New clsPessoaContato
    
    cPessoaContato.id_Pessoa = id_Pessoa
    cPessoaContato.ds_Nome = ds_Nome
    cPessoaContato.cd_Fone = cd_Fone
    cPessoaContato.cd_Email = cd_Email
    cPessoaContato.enuGravacao = Incluir
    
    Call colPessoaContato.Add(cPessoaContato)
    Set cPessoaContato = Nothing
        
    Exit Function
err_AdicionarContato:
    strMensagemRetorno = "Erro ao adicionar contato."
End Function
Private Function Remover(ByVal ID As Long)
On Error GoTo err_Remover
    Dim cPessoaContato As New clsPessoaContato
    
    Call cPessoaContato.CarregarDados(ID)
    
    
    Exit Function
err_Remover:
    strMensagemRetorno = "Erro ao remover contato."
End Function

Private Sub Class_Terminate()
    Set colPessoaContato = Nothing
End Sub
