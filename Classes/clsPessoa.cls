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

Public menumStatusGravacao      As enumStatusGravacao
Public mstrMensagemRetorno      As String
Private mcolPessoaContato       As Collection

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
    
    menumStatusGravacao = enumStatusGravacao.Nenhum
    mstrMensagemRetorno = ""
    Set mcolPessoaContato = New Collection
    Exit Sub
err_Class_Initialize:
    mstrMensagemRetorno = "Erro ao inicializar classe."
End Sub

Public Function AdicionarContato(ByVal id_PessoaContato As Long, ByVal ds_Nome As String, ByVal cd_Fone As String, ByVal cd_Email As String, ByVal enumStatusGravacao As enumStatusGravacao)
On Error GoTo err_AdicionarContato
    Dim cPessoaContato As New clsPessoaContato
    
    With cPessoaContato
        .id = id_PessoaContato
        .id_Pessoa = Me.id_Pessoa
        .ds_Nome = ds_Nome
        .cd_Fone = cd_Fone
        .cd_Email = cd_Email
        .menumStatusGravacao = enumStatusGravacao
    End With
    
    Call mcolPessoaContato.Add(cPessoaContato)
    Set cPessoaContato = Nothing
        
    Exit Function
err_AdicionarContato:
    mstrMensagemRetorno = "Erro ao adicionar contato."
End Function
Public Sub ZerarListas()
    Set colPessoaContato = New Collection
End Sub

Public Property Get GetListaContatos() As Collection
    Set GetListaContatos = mcolPessoaContato
End Property

Public Function ChavePrimaria() As String
    ChavePrimaria = "id_Pessoa"
End Function

Public Function Campos() As String
    Campos = "cd_cnpjcpf, ds_Pessoa, ds_RazaoSocial, ds_Endereco, ds_Bairro, id_Cidade, cd_CEP, tp_Cliente, tp_Fornecedor, tp_Funcionario"
End Function

Public Function Tabela() As String
    Tabela = "tbdPessoa"
End Function

Public Function Conteudo()
    Conteudo = Array(cd_cnpjcpf, ds_Pessoa, ds_RazaoSocial, ds_Endereco, ds_Bairro, id_Cidade, cd_CEP, tp_Cliente, tp_Fornecedor, tp_Funcionario)
End Function

Private Sub Class_Terminate()
    Set mcolPessoaContato = Nothing
End Sub
