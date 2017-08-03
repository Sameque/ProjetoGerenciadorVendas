VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmProdutoConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Produto"
   ClientHeight    =   7155
   ClientLeft      =   4470
   ClientTop       =   3000
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      TabIndex        =   7
      Top             =   6345
      Width           =   810
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "&Pesquisar"
      Height          =   750
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Pesquisar os Dados"
      Top             =   135
      Width           =   810
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   750
      Left            =   11850
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Sair da tela"
      Top             =   6345
      Width           =   810
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   750
      Left            =   10101
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Excluir o Item Selecionado"
      Top             =   6345
      Width           =   810
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Incluir"
      Height          =   750
      Left            =   8355
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Incluir novo Item"
      Top             =   6345
      Width           =   810
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   750
      Left            =   9228
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Alterar o Item Selecionado"
      Top             =   6345
      Width           =   810
   End
   Begin Threed.SSFrame fraFiltro 
      Height          =   945
      Left            =   75
      TabIndex        =   1
      Top             =   15
      Width           =   11640
      _Version        =   65536
      _ExtentX        =   20532
      _ExtentY        =   1667
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
      Begin Transportes.SuperCombo cboNatureza 
         Height          =   315
         Left            =   8130
         TabIndex        =   12
         Top             =   480
         Width           =   3405
         _extentx        =   6006
         _extenty        =   556
      End
      Begin Transportes.SuperText txtCodigo 
         Height          =   315
         Left            =   105
         TabIndex        =   9
         Top             =   480
         Width           =   3150
         _extentx        =   5556
         _extenty        =   556
      End
      Begin Transportes.SuperText txtDescricao 
         Height          =   315
         Left            =   3345
         TabIndex        =   11
         Top             =   480
         Width           =   4695
         _extentx        =   8281
         _extenty        =   556
      End
      Begin VB.Label lblNatureza 
         Caption         =   "Natureza"
         Height          =   225
         Left            =   8130
         TabIndex        =   13
         Top             =   270
         Width           =   3315
      End
      Begin VB.Label lblCodigo 
         Caption         =   "Código do Produto"
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   270
         Width           =   3135
      End
      Begin VB.Label lblDescricao 
         Caption         =   "Descricao do Produto"
         Height          =   195
         Left            =   3345
         TabIndex        =   10
         Top             =   270
         Width           =   4680
      End
   End
   Begin Transportes.SuperSpreadNovo sprConsulta 
      Height          =   5235
      Left            =   90
      TabIndex        =   0
      Top             =   1050
      Width           =   12585
      _extentx        =   22199
      _extenty        =   9234
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
Attribute VB_Name = "frmProdutoConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sTabela As String
Dim sCampos As String

Private Sub Form_Activate()
    frmMDI.Arrange vbCascade
End Sub

Private Sub Form_Load()
On Error GoTo err_FormLoad
    
    Call IniciarComponentes
    Call CarregarComponentes
    
    Exit Sub
err_FormLoad:
    ShowError
End Sub
Private Function IniciarComponentes()

    Call sprConsulta.NovaColunaSpread(eslNumero, True, True, "id_Produto", "id_Produto", 0, 0)
    Call sprConsulta.NovaColunaSpread(eslTexto, True, True, "cd_Produto", "Código", 10, 20)
    Call sprConsulta.NovaColunaSpread(eslTexto, True, True, "ds_Produto", "Descrição", 30, 100)
    Call sprConsulta.NovaColunaSpread(eslValor, True, True, "kg_Cubado", "Peso Cubado", 12, 9)
    Call sprConsulta.NovaColunaSpread(eslValor, True, True, "kg_Real", "Peso Real", 12, 9)
    Call sprConsulta.NovaColunaSpread(eslNumero, True, True, "qt_Minima", "Estoque Mínima", 12, 4)
    Call sprConsulta.NovaColunaSpread(eslTexto, True, True, "ds_ProdutoNatureza", "Natureza", 30, 50)
    Call sprConsulta.NovaColunaSpread(eslCheck, True, True, "tp_ProdutoPerigoso", "Produto Perigoso", 10, 1)
    Call sprConsulta.NovaColunaSpread(eslCheck, True, True, "tp_ControlePoliciaFederal", "Controle Polícia Federal", 10, 1)
    Call sprConsulta.NovaColunaSpread(eslNumero, True, True, "qt_ProdutoSaldo", "Saldo Estoque", 12, 4)
    Call sprConsulta.FormatarNovo(21)

    sprConsulta.ColsFrozenName = "ds_Produto"

    Call cboNatureza.Formatar("id_ProdutoNatureza", "ds_ProdutoNatureza", "tbdProdutoNatureza", "")

End Function
Private Function CarregarComponentes()

    sCampos = "a.id_Produto AS id_Produto, max(a.cd_Produto) AS cd_Produto, max(a.ds_Produto) AS ds_Produto," _
    & "max(a.kg_Real) AS kg_Real, max(a.kg_cubado) AS kg_cubado, max(a.qt_Minima) AS qt_Minima, max(b.ds_ProdutoNatureza) AS ds_ProdutoNatureza, max(b.tp_ProdutoPerigoso) AS tp_ProdutoPerigoso," _
    & "max(b.tp_ControlePoliciaFederal) AS tp_ControlePoliciaFederal, isnull(sum(c.qt_ProdutoSaldo),0) AS qt_ProdutoSaldo"
    
    sTabela = "((tbdProduto a" _
    & " left join tbdProdutoNatureza b on a.id_ProdutoNatureza = b.id_ProdutoNatureza) " _
    & " left join tbdProdutoLote c on a.id_Produto = c.id_Produto) "

End Function
Private Sub Form_Unload(Cancel As Integer)
    Set frmProdutoConsulta = Nothing
End Sub

Private Sub cmdPesquisar_Click()
On Error GoTo err_cmdPesquisar

    Dim sWhere As String
    sWhere = MontarWhere

    If sWhere = "" Then
        Exit Sub
    End If
    
    Call sprConsulta.Carregar(Select_Table(False, sTabela, sCampos, sWhere, "max(a.id_Produto)", , , , , "a.id_Produto"))
    
    Exit Sub
err_cmdPesquisar:
    ShowError
End Sub
Private Function MontarWhere() As String
On Error GoTo err_MontarWhere
    Dim sWhere As String
    
    MontarWhere = ""
    sWhere = ""
    
    If txtCodigo.Text <> "" Then
        sWhere = sWhere & "a.cd_Produto = '" & Trim(txtCodigo) & "' AND "
    End If
    
    If txtDescricao.Text <> "" Then
        sWhere = sWhere & "a.ds_Produto like '" & Trim(txtDescricao.Text) & "%'" + " AND "
    End If
    
    If cboNatureza.ItemData2 > 0 Then
        sWhere = sWhere & "a.id_ProdutoNatureza = " & cboNatureza.ItemData2 & " AND "
    End If
    
    If sWhere = "" Then
        Mensagem "Favor preencher algum parametro de pesquisa!", Informacao
        Exit Function
    End If

    If Len(sWhere) > 5 Then
        sWhere = Left(sWhere, Len(sWhere) - 5)
    End If
    
    MontarWhere = sWhere
    
    Exit Function
err_MontarWhere:
    ShowError
End Function
Private Function MontarWhereCrystal() As String
On Error GoTo err_MontarWhere
    Dim sWhere As String
    
    MontarWhereCrystal = ""
    sWhere = ""
    
    If txtCodigo.Text <> "" Then
        sWhere = sWhere & "{tbdProduto.cd_Produto} = '" & Trim(txtCodigo) & "' AND "
    End If
    
    If txtDescricao.Text <> "" Then
        sWhere = sWhere & "{tbdProduto.ds_Produto} like '" & Trim(txtDescricao.Text) & "%'" + " AND "
    End If
    
    If cboNatureza.ItemData2 > 0 Then
        sWhere = sWhere & "{tbdProduto.id_ProdutoNatureza} = " & cboNatureza.ItemData2 & " AND "
    End If
    
    If sWhere = "" Then
        Mensagem "Favor preencher algum parametro de pesquisa!", Informacao
        Exit Function
    End If

    If Len(sWhere) > 5 Then
        sWhere = Left(sWhere, Len(sWhere) - 5)
    End If
    
    MontarWhereCrystal = sWhere
    
    Exit Function
err_MontarWhere:
    ShowError
End Function

Private Sub cmdIncluir_Click()
    Set frmProdutoDados.formChamador = Me
    frmProdutoDados.id_Produto = 0
    frmProdutoDados.Show vbModal
End Sub

Private Sub cmdAlterar_Click()
On Error GoTo err_cmdAlterar_Click
    
    sprConsulta.Row = sprConsulta.ActiveRow
    If sprConsulta.RowHidden = False And Val(sprConsulta.SpreadEventoName("id_Produto")) > 0 Then
        Set frmProdutoDados.formChamador = Me
        frmProdutoDados.id_Produto = sprConsulta.SpreadEventoName("id_Produto")
        frmProdutoDados.Show vbModal
    End If
    
    Exit Sub
err_cmdAlterar_Click:
    ShowError
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo err_cmdExcluir_Click

    Dim cProduto As New clsProduto

    If sprConsulta.ActiveRow < 1 Then
        Mensagem "Selecione o item que será excluído.", erro
        Exit Sub
    End If

    If Mensagem("Confirma exclusão?", Pergunta) = vbNo Then
        Exit Sub
    End If

    Call AbreTransacao
    
    cProduto.id_Produto = sprConsulta.SpreadEventoName("id_Produto")
    If Not cProduto.Excluir Then
        Call VoltaTransacao
        Mensagem "Ocorreu um erro na exclusão.", ErroCritico
        Exit Sub
    End If
    
    Call FechaTransacao

    sprConsulta.Action = 5
    sprConsulta.MaxRows = sprConsulta.MaxRows - 1
    Mensagem "Exclusão efetuada.", Informacao

    Exit Sub
err_cmdExcluir_Click:
    ShowError
    Call VoltaTransacao
End Sub

Private Sub cmdImprimir_Click()
On Error GoTo err_cmdImprimir
    
    Dim sWhere As String
    Dim sFiltro As String
    
    sWhere = MontarWhereCrystal

    If sWhere = "" Then
        Exit Sub
    End If

    cryRelatorio.ReportFileName = sPathReport & "\Relatorios\RelacaoProdutos.rpt"
    cryRelatorio.WindowParentHandle = frmMDI.hWnd
    cryRelatorio.SelectionFormula = sWhere
    cryRelatorio.Formulas(0) = "Filtro='" & sFiltro & "'"
    cryRelatorio.Connect = sStringConexaoRelatorio
    Call ChamarRelatorio(cryRelatorio)

    Exit Sub
err_cmdImprimir:
    ShowError
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Public Sub AtualizarDados(id_Produto As Long)
    Call sprConsulta.AtualizarDadosSpread(id_Produto, "a.id_Produto", sTabela, sCampos, "a.id_Produto")
End Sub
