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
   Begin VB.CommandButton cmdFuncionalidade 
      Caption         =   "Testar Funcionalidade"
      Height          =   540
      Left            =   3945
      TabIndex        =   2
      Top             =   420
      Width           =   2700
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

Private Function TestarFuncionalidade()
    Dim cPessoa As New clsPessoa
    
    Call IniciarClasse(cPessoa)
    Call AddContato(cPessoa)
    
End Function


Private Sub cmdFuncionalidade_Click()
    Call TestarFuncionalidade
End Sub

Private Sub cmdTestarClassePessoa_Click()
    
    Dim cPessoa As New clsPessoa
    Dim cPessoaServico  As New clsPessoaServico
        
    Call IniciarClasse(cPessoa)
    Call AddContato(cPessoa)
    Call TestarClassePessoa(cPessoa)
    
    Call sprContato.FormatarPorClasse(cServicoPessoa.FormatarSpreadPessoaContato)

End Sub
Private Function AddContato(ByRef cPessoa As clsPessoa)

    Call cPessoa.AdicionarContato("nome", "1234567890", "aaaaaaaa@provedor.com")
    Call cPessoa.AdicionarContato("Fulano", "1234567890", "bbbbbb@provedor.com")
    Call cPessoa.AdicionarContato("Beltranome", "1234567890", "evveveveve@provedor.com")
    Call cPessoa.AdicionarContato("Cavalo", "1234567890", "123@456.789")

End Function

Private Function TestarClassePessoa(cPessoa As clsPessoa)
    
    Dim cPessoaServico  As New clsPessoaServico
    Dim id              As Long
    
    cPessoa.enuGravacao = EnumStatusGravacao.Incluir
    
    If Not cPessoaServico.Salvar(cPessoa) Then
        MsgGravar = cPessoaServico.strMensagemRetorno
    Else
        id = cPessoa.id_Pessoa
        MsgGravar = " :) Gravação OK!"
    End If
    
    Set cPessoa = Nothing
    Set cPessoa = cPessoaServico.CarregarClasse(id)

    cPessoa.enuGravacao = EnumStatusGravacao.Alterar
    If Not cPessoaServico.Salvar(cPessoa) Then
        MsgAtualizar = cPessoaServico.strMensagemRetorno
    Else
        MsgAtualizar = " :) Atualizadção OK!"
    End If

    cPessoa.enuGravacao = EnumStatusGravacao.Excluir
    If Not cPessoaServico.Salvar(cPessoa) Then
        MsgExcluir = cPessoaServico.strMensagemRetorno
    Else
        MsgExcluir = " :) Exclusão OK!"
    End If
    
    Call Resultado
    Call LimparMsg
    Set cPessoa = Nothing
    Set cPessoaServico = Nothing
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
        .enuGravacao = Alterar
    End With
    
End Function

Private Function Resultado()
    Mensagem MsgCarregar & vbCrLf _
            & MsgGravar & vbCrLf _
            & MsgAtualizar & vbCrLf _
            & MsgExcluir & vbCrLf _
    , Informacao
    
End Function
Private Function TestarClasseVenda()
    Dim cVenda  As clsVenda
    Dim id      As Long
        
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

Private Sub Form_Load()
    Call LimparMsg
End Sub

Private Function LimparMsg()

    MsgCarregar = ""
    MsgGravar = ""
    MsgAtualizar = ""
    MsgExcluir = ""

End Function
