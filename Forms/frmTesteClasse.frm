VERSION 5.00
Begin VB.Form frmTesteClasse 
   Caption         =   "Teste Classe"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   6945
   Begin VB.CommandButton cmdClasse 
      Caption         =   "Classe PessoaContato"
      Height          =   510
      Left            =   480
      TabIndex        =   2
      Top             =   2055
      Width           =   2610
   End
   Begin VB.CommandButton cmdTestarClassePessoa 
      Caption         =   "&Testar Classe Pessoa"
      Height          =   510
      Left            =   510
      TabIndex        =   1
      Top             =   1230
      Width           =   2565
   End
   Begin VB.CommandButton cmdTestarClasseVenda 
      Caption         =   "&Testar Classe Venda"
      Height          =   540
      Left            =   525
      TabIndex        =   0
      Top             =   405
      Width           =   2580
   End
End
Attribute VB_Name = "frmTesteClasse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public MsgCarregar
Public MsgGravar
Public MsgAtualizar
Public MsgExcluir

Private Sub cmdTesteVenda_Click()
    Call TestarClasseVenda
End Sub

Private Sub cmdClasse_Click()
    Call TestarClassePessoaContato
End Sub

Private Sub cmdTestarClassePessoa_Click()
    
    Dim cPessoa As New clsPessoa
        
    Call IniciarClasse(cPessoa)
    Call AddContato(cPessoa)
    Call TestarClassePessoa(cPessoa)
    
End Sub
Private Function AddContato(ByRef cPessoa As clsPessoa)

    Call cPessoa.AdicionarContato("nome", "1234567890", "aaaaaaaa@provedor.com")
    Call cPessoa.AdicionarContato("Fulano", "1234567890", "bbbbbb@provedor.com")
    Call cPessoa.AdicionarContato("Beltranome", "1234567890", "evveveveve@provedor.com")
    Call cPessoa.AdicionarContato("Cavalo", "1234567890", "123@456.789")

End Function

Private Function TestarClasseVenda()
    Dim cVenda  As clsVenda
    Dim ID      As Long
        
    MensagemRetorno = ""
    
    Set cVenda = New clsVenda
    
    cVenda.id_Venda = 0
    cVenda.id_Comprador = 1
    cVenda.id_Vendedor = 1
    cVenda.vl_Venda = 100
    cVenda.vl_Desconto = 10
    cVenda.vl_Frete = 200
    cVenda.vl_Total = 290
    cVenda.dt_Venda = CDateEspecial("01/01/2017")
    cVenda.dt_PrvisaoEmtrega = CDateEspecial("31/12/2017")
    cVenda.dt_Entrega = CDateEspecial("02/01/2017")
    cVenda.ds_Observacao = "Teste de classe."

    If Not cVenda.Gravar Then
        MsgGravar = cVenda.MensagemRetorno
    End
        MsgGravar = " :) Gravação OK!"
    End If
        
    If Not cVenda.CarregarDados(1) Then
        MsgCarregar = cVenda.MensagemRetorno
    Else
        MsgCarregar = " :) Carregar Classe OK!"
    End If
        
    If Not cVenda.Gravar Then
        MsgAtualizar = cVenda.MensagemRetorno
    Else
        MsgAtualizar = " :) Atualizadção OK!"
    End If
    
    If Not cVenda.Excluir Then
        MsgExcluir = cVenda.MensagemRetorno
    Else
        MsgExcluir = " :) Exclusão OK!"
    End If
    
    
    Call Resultado
    Call LimparMsg
    
    Set cVenda = Nothing
End Function

Private Function TestarClassePessoa(cPessoa As clsPessoa)
    
    Dim cPessoaServico  As New clsPessoaServico
    Dim ID              As Long

    If Not cPessoaServico.Gravar(cPessoa) Then
        MsgGravar = cPessoa.MensagemRetorno
    Else
        ID = cPessoa.id_Pessoa
        MsgGravar = " :) Gravação OK!"
    End If
    
    Set cPessoa = Nothing
    Set cPessoa = cPessoaServico.CarregarId(ID)

    If IsEmpty(cPessoa) Then
        MsgCarregar = cPessoa.MensagemRetorno
    Else
        MsgCarregar = " :) Carregar Classe OK!"
    End If

    If Not cPessoaServico.Gravar(cPessoa) Then
        MsgAtualizar = cPessoa.MensagemRetorno
    Else
        MsgAtualizar = " :) Atualizadção OK!"
    End If

    If Not cPessoaServico.Excluir(cPessoa) Then
        MsgExcluir = cPessoa.MensagemRetorno
    Else
        MsgExcluir = " :) Exclusão OK!"
    End If
    
    Call Resultado
    Call LimparMsg
    Set cPessoa = Nothing
    
End Function

Private Function IniciarClasse(ByRef cPessoa As clsPessoa) As Boolean
    
    With cPessoa
        .ds_Pessoa = "Nome teste"
        .ds_RazaoSocial = "Razao Teste"
        .ds_Endereco = "Endereço teste"
        .ds_Bairro = "Bairro teste"
        .id_Cidade = 15
        .tp_Cliente = "S"
        .tp_Fornecedor = "S"
        .tp_Funcionario = "N"
        .cd_cnpjcpf = "123456789"
        .cd_CEP = "13600970"
    End With
    
End Function

Private Function TestarClassePessoaContato()


End Function


Private Function IniciarPessoa() As clsPessoa
On erro GoTo err_CarrregarPeesoa

    cPessoa As New clsPessoa
    
    Set IniciarPessoa = Nothing
    
    cPessoa.id_Pessoa = 0
    cPessoa.id_Cidade = 0
    cPessoa.id_PessoaFuncionario = 0
    cPessoa.id_PessoaFuncao = 0
    cPessoa.cd_cnpjcpf = "123456789012"
    cPessoa.cd_CEP = "13600000"
    cPessoa.ds_Pessoa = "nome"
    cPessoa.ds_RazaoSocial = "razao"
    cPessoa.ds_Endereco = "endereco"
    cPessoa.ds_Bairro = "Bairro"
    cPessoa.ds_Usuario = "usuario"
    cPessoa.ds_Senha = "senha"
    cPessoa.tp_Cliente = "S"
    cPessoa.tp_Fornecedor = "S"
    cPessoa.tp_Funcionario = "S"

    Set IniciarPessoa = cPessoa
    
err_CarrregarPeesoa:
    ShowError
End Function
Private Function Resultado()
    Mensagem MsgCarregar & vbCrLf _
            & MsgGravar & vbCrLf _
            & MsgAtualizar & vbCrLf _
            & MsgExcluir & vbCrLf _
    , Informacao
    
End Function

Private Sub Form_Load()
    Call LimparMsg
End Sub

Private Function LimparMsg()

    MsgCarregar = ""
    MsgGravar = ""
    MsgAtualizar = ""
    MsgExcluir = ""

End Function
