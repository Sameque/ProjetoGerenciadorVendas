VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmPessoaConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Pessoa"
   ClientHeight    =   7155
   ClientLeft      =   4470
   ClientTop       =   3000
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   12720
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   750
      Left            =   10974
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6345
      Width           =   810
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "&Pesquisar"
      Height          =   750
      Left            =   11850
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Pesquisar os Dados"
      Top             =   270
      Width           =   810
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   750
      Left            =   11850
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sair da tela"
      Top             =   6345
      Width           =   810
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   750
      Left            =   10101
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Excluir o Item Selecionado"
      Top             =   6345
      Width           =   810
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Incluir"
      Height          =   750
      Left            =   8355
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Incluir novo Item"
      Top             =   6345
      Width           =   810
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   750
      Left            =   9228
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Alterar o Item Selecionado"
      Top             =   6345
      Width           =   810
   End
   Begin Threed.SSFrame fraFiltro 
      Height          =   1170
      Left            =   60
      TabIndex        =   13
      Top             =   60
      Width           =   11715
      _Version        =   65536
      _ExtentX        =   20664
      _ExtentY        =   2064
      _StockProps     =   14
      Caption         =   "Filtros para Pesquisa"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin Transportes.SuperDBCombo cboEstado 
         Height          =   510
         Left            =   9465
         TabIndex        =   3
         Top             =   285
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   900
         Label           =   "Estado"
      End
      Begin Threed.SSCheck chkFuncionario 
         Height          =   150
         Left            =   2205
         TabIndex        =   6
         Top             =   885
         Width           =   1350
         _Version        =   65536
         _ExtentX        =   2381
         _ExtentY        =   265
         _StockProps     =   78
         Caption         =   "Funcion�rio"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Transportes.SuperDBCombo cboPessoa 
         Height          =   510
         Left            =   120
         TabIndex        =   1
         Top             =   285
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   900
         CampoPesquisa2Width=   1500
         CampoPesquisa2Coluna=   2
         Label           =   "Nome"
         LabelCampoPesquisa2=   "CNPJ/CPF"
      End
      Begin Threed.SSCheck chkFornecedor 
         Height          =   150
         Left            =   990
         TabIndex        =   5
         Top             =   885
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   265
         _StockProps     =   78
         Caption         =   "Fornecedor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCheck chkCliente 
         Height          =   150
         Left            =   120
         TabIndex        =   4
         Top             =   885
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   265
         _StockProps     =   78
         Caption         =   "Cliente"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Transportes.SuperDBCombo cboCidade 
         Height          =   510
         Left            =   5385
         TabIndex        =   2
         Top             =   285
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   900
         Label           =   "Cidade"
      End
   End
   Begin Transportes.SuperSpreadNovo sprPessoa 
      Height          =   4920
      Left            =   75
      TabIndex        =   0
      Top             =   1365
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   8678
   End
   Begin Crystal.CrystalReport cryRelatorio 
      Left            =   240
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "frmPessoaConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sTabela As String
Dim sCampos As String
Dim mstrWhere As String

Private Sub Form_Activate()
    frmMDI.Arrange vbCascade
End Sub

Private Sub Form_Load()
    Call IniciarComponentes
End Sub

Private Function IniciarComponentes()
On Error GoTo err_IniciarComponentes
    Dim cServicoPessoa As New clsServicoPessoa
   
    Call cboCidade.Formatar("a.id_Cidade, a.ds_Cidade, b.ds_Estado", "0,3000,1000", "false,true,true", "tbdCidade a left join tbdEstado b on a.id_Estado = b.id_Estado", "", "ds_Cidade")
    Call cboPessoa.Formatar("id_Pessoa, ds_Pessoa, cd_cnpjcpf", "0, 3000, 1000", "false, true, true", "tbdPessoa", "", "ds_Pessoa", 2, 1500)
    Call cboEstado.Formatar("id_Estado, ds_Estado, cd_Estado", "0, 500, 500", "false, true, true", "tbdEstado", "", "ds_Estado")
    
    Call sprPessoa.FormatarPorClasse(cServicoPessoa.FormatarSpreadPessoa)
    sprPessoa.ColsFrozenName = "ds_Pessoa"

    sCampos = "a.id_Pessoa, a.cd_cnpjcpf, a.ds_Pessoa, a.ds_RazaoSocial, a.ds_Endereco, a.ds_Bairro, b.ds_Cidade," _
    & " c.cd_Estado, a.tp_Cliente, a.tp_Fornecedor, a.tp_Funcionario"
    
    sTabela = "((( tbdPessoa a " _
    & " left join tbdCidade b on a.id_Cidade = b.id_Cidade)" _
    & " left join tbdEstado c on b.id_Estado = c.id_Estado)" _
    & " left join tbdPessoaFuncionario d on a.id_Pessoa = d.id_Pessoa)"

    Set cServicoPessoa = Nothing
    
    Exit Function
err_IniciarComponentes:
    ShowError
End Function

Private Sub cmdPesquisar_Click()
On Error GoTo err_cmdPesquisar

    Call MontarWhere
    
    If mstrWhere = "" Then
        Mensagem "Favor preencher algum parametro de pesquisa!", Informacao
        Exit Sub
    End If
    
    Call CarregarSprPessoa
    
    Exit Sub
err_cmdPesquisar:
    ShowError
End Sub
Private Function MontarWhere()
On Error GoTo err_MontarWhere
    
    mstrWhere = ""
    
    If cboPessoa.ItemData2 > 0 Then
        mstrWhere = mstrWhere & "pessoa.id_Pessoa = " & cboPessoa.ItemData2 & " AND "
    End If
    
    If cboCidade.ItemData2 > 0 Then
        mstrWhere = mstrWhere & "pessoa.id_Cidade = " & cboCidade.ItemData2 & " AND "
    End If
    
    If cboEstado.ItemData2 > 0 Then
        mstrWhere = mstrWhere & "cidade.id_Estado = " & cboEstado.ItemData2 & " AND "
    End If
    
    If chkCliente.Value Then
        mstrWhere = mstrWhere & "pessoa.tp_Cliente = 'S' AND "
    End If
    
    If chkFornecedor.Value Then
        mstrWhere = mstrWhere & "pessoa.tp_Fornecedor = 'S' AND "
    End If
    
    If chkFuncionario Then
        mstrWhere = mstrWhere & "pessoa.tp_Funcionario = 'S' AND "
    End If
    
    If Len(mstrWhere) > 5 Then
        mstrWhere = Left(mstrWhere, Len(mstrWhere) - 5)
    End If
    
    Exit Function
err_MontarWhere:
    ShowError
End Function

Private Function CarregarSprPessoa()
On Error GoTo err_Pesquisar
    Dim cServicoPessoa As New clsServicoPessoa
    
    Call sprPessoa.CarregarPorClasse(mstrWhere)
    
    Exit Function
err_Pesquisar:
    ShowError
End Function

Private Sub cmdIncluir_Click()
    Set frmPessoaDados.FormChamador = Me
    frmPessoaDados.mlngPessoa = 0
    frmPessoaDados.Show vbModal
    Set frmPessoaDados = Nothing
End Sub

Private Sub cmdAlterar_Click()
On Error GoTo err_cmdAlterar_Click
    
    sprPessoa.Row = sprPessoa.ActiveRow
    If sprPessoa.RowHidden = False And Val(sprPessoa.SpreadEventoName("id_Pessoa")) > 0 Then
        Set frmPessoaDados.FormChamador = Me
        frmPessoaDados.mlngPessoa = sprPessoa.SpreadEventoName("id_Pessoa")
        frmPessoaDados.Show vbModal
    End If
    
    Exit Sub
err_cmdAlterar_Click:
    ShowError
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo err_cmdExcluir_Click

    Dim cServicoPessoa As New clsServicoPessoa
    Dim cPessoa As clsPessoa
    
    If sprPessoa.ActiveRow < 1 Then
        Mensagem "Selecione o item que ser� exclu�do.", erro
        Exit Sub
    End If

    If Mensagem("Confirma exclus�o?", Pergunta) = vbNo Then
        Exit Sub
    End If
    
    Set cPessoa = cServicoPessoa.CarregarPorID(sprPessoa.SpreadEventoName("id_Pessoa"))
    
    Call AbreTransacao
    If Not cServicoPessoa.Excluir(cPessoa) Then
        Call VoltaTransacao
        Mensagem "Ocorreu um erro na exclus�o.", ErroCritico
        Exit Sub
    End If
    Call FechaTransacao

    sprPessoa.Action = 5
    sprPessoa.MaxRows = sprPessoa.MaxRows - 1
    Mensagem "Exclus�o efetuada.", Informacao

    Exit Sub
err_cmdExcluir_Click:
    ShowError
    Call VoltaTransacao
End Sub

Private Sub cmdImprimir_Click()
On Error GoTo err_cmdImprimir
    
    Dim sFiltro As String
    
    cryRelatorio.ReportFileName = sPathReport & "\Relatorios\Pessoa.rpt"
    cryRelatorio.WindowParentHandle = frmMDI.hWnd
    cryRelatorio.SelectionFormula = mstrWhere
    cryRelatorio.Formulas(0) = "Filtro='" & sFiltro & "'"
    cryRelatorio.Connect = sStringConexaoRelatorio
    Call ChamarRelatorio(cryRelatorio)

    Exit Sub
err_cmdImprimir:
    ShowError
End Sub

Private Function MontarWhereCrystal()
On Error GoTo err_MontarWhere
    
    mstrWhere = ""
    
    If cboPessoa.ItemData2 > 0 Then
        mstrWhere = mstrWhere & "a.id_Pessoa = " & cboPessoa.ItemData2 & " AND "
        sFiltros = ""
    End If
    
    If cboCidade.ItemData2 > 0 Then
        mstrWhere = mstrWhere & "a.id_Cidade = " & cboCidade.ItemData2 & " AND "
        sFiltros = ""
    End If
    
    If cboEstado.ItemData2 > 0 Then
        mstrWhere = mstrWhere & "b.id_Estado = " & cboEstado.ItemData2 & " AND "
        sFiltros = ""
    End If
    
    If chkCliente.Value Then
        mstrWhere = mstrWhere & "a.tp_Cliente = 'S' AND "
        sFiltros = ""
    End If
    
    If chkFornecedor.Value Then
        mstrWhere = mstrWhere & "{tbdPessoa.tp_Fornecedor} = 'S' AND "
        sFiltros = ""
    End If
    
    If chkFuncionario Then
        mstrWhere = mstrWhere & "a.tp_Funcionario = 'S' AND "
        sFiltros = ""
    End If
    
    If Len(mstrWhere) > 5 Then
        mstrWhere = Left(mstrWhere, Len(mstrWhere) - 5)
        sFiltros = ""
    End If
    
    If txtCodigo.Text <> "" Then
        mstrWhere = mstrWhere & "{tbdProduto.cd_Produto} = '" & Trim(txtCodigo) & "' AND "
        sFiltros = ""
    End If
    
    If Len(mstrWhere) > 5 Then
        mstrWhere = Left(mstrWhere, Len(mstrWhere) - 5)
        sFiltros = ""
    End If
    
    Exit Function
err_MontarWhere:
    ShowError
End Function

Private Sub cmdSair_Click()
    Unload Me
End Sub

Public Sub AtualizarDados(id_Pessoa As Long)
    Call sprPessoa.AtualizarDadosSpread(id_Pessoa, "a.id_Pessoa", sTabela, sCampos)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmPessoaConsulta = Nothing
End Sub
