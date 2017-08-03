VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmVendaConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendas"
   ClientHeight    =   7155
   ClientLeft      =   4470
   ClientTop       =   3000
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   6825
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "&Pesquisar"
      Height          =   750
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Pesquisar os Dados"
      Top             =   375
      Width           =   810
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   750
      Left            =   5910
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sair da tela"
      Top             =   6360
      Width           =   810
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   750
      Left            =   5055
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Excluir o Item Selecionado"
      Top             =   6360
      Width           =   810
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Incluir"
      Height          =   750
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Incluir novo Item"
      Top             =   6360
      Width           =   810
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   750
      Left            =   4185
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Alterar o Item Selecionado"
      Top             =   6360
      Width           =   810
   End
   Begin Threed.SSFrame fraFiltro 
      Height          =   1350
      Left            =   75
      TabIndex        =   7
      Top             =   60
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   2381
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
      Begin Transportes.SuperControlNovo mskDataVendaFinal 
         Height          =   510
         Left            =   2055
         TabIndex        =   3
         Top             =   780
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   900
         AutoTab         =   0   'False
         ToolTip         =   ""
         Mascara         =   4
         Label           =   "Data da Compra Final"
      End
      Begin Threed.SSFrame fraEntregue 
         Height          =   1050
         Left            =   4500
         TabIndex        =   4
         Top             =   210
         Width           =   1020
         _Version        =   65536
         _ExtentX        =   1799
         _ExtentY        =   1852
         _StockProps     =   14
         Caption         =   "Entregue"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optAmbos 
            Caption         =   "Ambos"
            Height          =   225
            Left            =   165
            TabIndex        =   13
            Top             =   705
            Width           =   795
         End
         Begin VB.OptionButton optNao 
            Caption         =   "Não"
            Height          =   345
            Left            =   165
            TabIndex        =   6
            Top             =   420
            Width           =   750
         End
         Begin VB.OptionButton optSim 
            Caption         =   "Sim"
            Height          =   255
            Left            =   165
            TabIndex        =   2
            Top             =   225
            Width           =   675
         End
      End
      Begin Transportes.SuperControlNovo mskDataVendaInicial 
         Height          =   510
         Left            =   195
         TabIndex        =   1
         Top             =   780
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   900
         AutoTab         =   0   'False
         ToolTip         =   ""
         Mascara         =   4
         Label           =   "Data da Compra Inicial"
      End
      Begin Transportes.SuperDBCombo cboCliente 
         Height          =   510
         Left            =   195
         TabIndex        =   0
         Top             =   225
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   900
         Label           =   "Cliente"
      End
   End
   Begin Transportes.SuperSpreadNovo sprConsulta 
      Height          =   4785
      Left            =   75
      TabIndex        =   8
      Top             =   1500
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   8440
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
Attribute VB_Name = "frmVendaConsulta"
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
       
    Call sprConsulta.NovaColunaSpread(eslBotao, False, True, "cmdItem", "Itens", 5, 5)
    Call sprConsulta.NovaColunaSpread(eslNumero, True, True, "id_Venda", "Numero", 10, 10)
    Call sprConsulta.NovaColunaSpread(eslTexto, True, True, "ds_Comprador", "Comprador", 30, 255)
    Call sprConsulta.NovaColunaSpread(eslValor, True, True, "vl_Venda", "Valor Produtos", 10, 10)
    Call sprConsulta.NovaColunaSpread(eslValor, True, True, "vl_Desconto", "Valor Desconta", 10, 10)
    Call sprConsulta.NovaColunaSpread(eslValor, True, True, "vl_Frete", "Valor Frete", 10, 10)
    Call sprConsulta.NovaColunaSpread(eslValor, True, True, "vl_Total", "Valor Total", 10, 10)
    Call sprConsulta.NovaColunaSpread(eslData, True, True, "dt_Venda", "Data Venda", 10, 10)
    Call sprConsulta.NovaColunaSpread(eslData, True, True, "dt_PrvisaoEmtrega", "Prev. entrega", 10, 10)
    Call sprConsulta.NovaColunaSpread(eslData, True, True, "dt_Entrega", "Entrega", 10, 10)
    Call sprConsulta.NovaColunaSpread(eslTexto, True, True, "ds_Observacao", "Observação", 30, 255)
    Call sprConsulta.FormatarNovo(21)

    sprConsulta.ColsFrozenName = "id_Venda"

    sCampos = " '',a.id_Venda, b.ds_Pessoa as ds_Comprador,a.vl_Venda,a.vl_Desconto,a.vl_Frete," _
    & "a.vl_Total, a.dt_Venda,a.dt_PrvisaoEmtrega,a.dt_Entrega,a.ds_Observacao"
    
    sTabela = "(tbdVenda a" _
    & " LEFT JOIN tbdPessoa b ON b.id_Pessoa = a.id_Comprador) "

    Exit Sub
err_FormLoad:
    ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmVendaConsulta = Nothing
End Sub

Private Sub cmdPesquisar_Click()
On Error GoTo err_cmdPesquisar

    Dim sWhere As String
    
    sWhere = MontarWhere
        
    If sWhere = "" Then
        Mensagem "Favor preencher algum parametro de pesquisa!", Informacao
        Exit Sub
    End If

    Call sprConsulta.Carregar(Select_Table(False, sTabela, sCampos, sWhere, "a.id_Venda"))
    
    Exit Sub
err_cmdPesquisar:
    ShowError
End Sub
Private Function MontarWhere() As String
On Error GoTo err_MontarWhere
    Dim sWhere As String
    
    MontarWhere = ""
    sWhere = ""
     
    If cboCliente.ItemData2 > 0 Then
        sWhere = sWhere & " id_Comprador = " & cboCliente.ItemData2 & " AND "
    End If
        
    If mskDataVendaInicial.ClipText <> "" Then
        sWhere = sWhere & " dt_Venda >= " & ConverterData(mskDataVendaInicial.Text) & " AND "
    End If
    
    If mskDataVendaFinal.ClipText <> "" Then
        sWhere = sWhere & " dt_Venda <= " & ConverterData(mskDataVendaFinal.Text) & " AND "
    End If
    
    If optSim.Value = True Then
        sWhere = sWhere & " dt_Entrega IS NOT NULL  AND "
    End If
    
    If optNao.Value = True Then
        sWhere = sWhere & " dt_Entrega IS NULL  AND "
    End If
    
    If Len(sWhere) > 5 Then
        sWhere = Left(sWhere, Len(sWhere) - 5)
    End If
    
    MontarWhere = sWhere
    
    Exit Function
err_MontarWhere:
    ShowError
End Function

Private Sub cmdIncluir_Click()
    Set frmVendaDados.formChamador = Me
    frmVendaDados.id_Venda = 0
    frmVendaDados.Show vbModal
End Sub

Private Sub cmdAlterar_Click()
On Error GoTo err_cmdAlterar_Click
    
    sprConsulta.Row = sprConsulta.ActiveRow
    If sprConsulta.RowHidden = False And Val(sprConsulta.SpreadEventoName("id_Venda")) > 0 Then
        Set frmVendaDados.formChamador = Me
        frmVendaDados.id_Venda = sprConsulta.SpreadEventoName("id_Venda")
        frmVendaDados.Show vbModal
    End If
    
    Exit Sub
err_cmdAlterar_Click:
    ShowError
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo err_cmdExcluir_Click

    Dim cVenda As New clsVenda

    If sprConsulta.ActiveRow < 1 Then
        Mensagem "Selecione o item que será excluído.", erro
        Exit Sub
    End If

    If Mensagem("Confirma exclusão?", Pergunta) = vbNo Then
        Exit Sub
    End If

    Call cVenda.CarregarDados(sprConsulta.SpreadEventoName("id_Venda"))
    
    Call AbreTransacao
    If Not cVenda.Excluir Then
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
        
    cryRelatorio.ReportFileName = sPathReport & "\Relatorio\Venda.rpt"
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

Public Sub AtualizarDados(id_Venda As Long)
    Call sprConsulta.AtualizarDadosSpread(id_Venda, "a.id_Venda", sTabela, sCampos)
End Sub

Private Sub sprConsulta_ButtonClickedName(ByVal ColName As String, ByVal Row As Long, ByVal ButtonDown As Integer)
On Error GoTo err_sprConsulta_ButtonClickedName

    sprConsulta.Row = Row
    If ColName = "cmdItem" Then
        frmProdutoItem.id_Venda = sprConsulta.SpreadEventoName("id_Venda")
        frmProdutoItem.Show
    End If
    
    Exit Sub
err_sprConsulta_ButtonClickedName:
    ShowError
End Sub
