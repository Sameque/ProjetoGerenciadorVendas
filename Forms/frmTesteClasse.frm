VERSION 5.00
Begin VB.Form frmTesteClasse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Teste Classe"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   9075
   Begin Transportes.SuperSpreadNovo sprContato 
      Height          =   1815
      Left            =   60
      TabIndex        =   3
      Top             =   2535
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   3201
   End
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
Public MsgCarregarSpread
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
    
    Call TestarClassePessoa(cPessoa)

End Sub

Private Function AddContato(ByRef cPessoa As clsPessoa)

    Call cPessoa.AdicionarContato(0, "nome", "1234567890", "aaaaaaaa@provedor.com")
    Call cPessoa.AdicionarContato(0, "Fulano", "1234567890", "bbbbbb@provedor.com")
    Call cPessoa.AdicionarContato(0, "Beltranome", "1234567890", "evveveveve@provedor.com")
    Call cPessoa.AdicionarContato(0, "Cavalo", "1234567890", "123@456.789")

End Function

Private Function RemoverPrimeiroContato(ByRef cPessoa As clsPessoa)
   Dim cPessoaContato As New clsPessoaContato
   
   Set cPessoaContato = cPessoa.GetListaContatos(1)
   cPessoaContato.ds_Nome = UCase(cPessoaContato.ds_Nome)
    'call cPessoa.GetListaContatos(1).StatusGravacao = EnumStatusGravacao.Alterar
End Function

Private Function TestarClassePessoa(cPessoa As clsPessoa)
    
    Dim cPessoaServico  As New clsPessoaServico
    Dim ID              As Long
    
    Call IniciarClasse(cPessoa)
    If Not cPessoaServico.Salvar(cPessoa) Then
        MsgGravar = cPessoaServico.mstrMensagemRetorno
    Else
        ID = cPessoa.id_Pessoa
        MsgGravar = "OK"
    End If
    
    Call AddContato(cPessoa)
    cPessoa.ds_Pessoa = UCase(cPessoa.ds_Pessoa)
    cPessoa.StatusGravacao = EnumStatusGravacao.Alterar
    
    If Not cPessoaServico.Salvar(cPessoa) Then
        MsgAtualizar = cPessoaServico.mstrMensagemRetorno
    Else
        MsgAtualizar = "OK"
    End If
      
    Set cPessoa = Nothing
    Set cPessoa = cPessoaServico.CarregarPorID(ID, True)
    
    Call sprContato.FormatarPorClasse(cPessoaServico.FormatarSpreadPessoaContato)
    Call sprContato.CarregarPorClasse("id_Pessoa = " & cPessoa.id_Pessoa)
    
    Call RemoverPrimeiroContato(cPessoa)
    Call cPessoaServico.Salvar(cPessoa)
    
    Call sprContato.AtualizarStatus(cPessoa.GetListaContatos)
    
    If sprContato.Col > 0 Then
        MsgCarregarSpread = "OK"
    End If
    
    cPessoa.StatusGravacao = EnumStatusGravacao.Excluir
    
    If Not cPessoaServico.Salvar(cPessoa) Then
        MsgExcluir = cPessoaServico.mstrMensagemRetorno
    Else
        MsgExcluir = "OK"
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
        .StatusGravacao = EnumStatusGravacao.Incluir
    End With
    
End Function

Private Function Resultado()
    Mensagem "Gravar: " & MsgGravar & vbCrLf _
            & "Atualizar: " & MsgAtualizar & vbCrLf _
            & "Carregar Spread: " & MsgCarregarSpread & vbCrLf _
            & "Excluir: " & MsgExcluir & vbCrLf _
    , Informacao
    
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

Private Sub Form_Load()
    Call LimparMsg
End Sub

Private Function LimparMsg()

    MsgCarregar = ""
    MsgGravar = ""
    MsgAtualizar = ""
    MsgCarregarSpread = ""
    MsgExcluir = ""
    
End Function
