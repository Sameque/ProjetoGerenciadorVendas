VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmPedidoCompraConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos de Compra"
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
   Begin VB.CommandButton cmdBaixarRecebimento 
      Caption         =   "&Baixar Entrega"
      Height          =   750
      Left            =   8355
      TabIndex        =   9
      Top             =   6345
      Width           =   810
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "&Pesquisar"
      Height          =   750
      Left            =   11865
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Pesquisar os Dados"
      Top             =   330
      Width           =   810
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   750
      Left            =   11850
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Sair da tela"
      Top             =   6345
      Width           =   810
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   750
      Left            =   10980
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Excluir o Item Selecionado"
      Top             =   6345
      Width           =   810
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Incluir"
      Height          =   750
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Incluir novo Item"
      Top             =   6345
      Width           =   810
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   750
      Left            =   10110
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Alterar o Item Selecionado"
      Top             =   6345
      Width           =   810
   End
   Begin Threed.SSFrame fraFiltro 
      Height          =   1350
      Left            =   75
      TabIndex        =   7
      Top             =   45
      Width           =   11715
      _Version        =   65536
      _ExtentX        =   20664
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
      Begin Transportes.SuperControlNovo mskDataCompraFinal 
         Height          =   510
         Left            =   8640
         TabIndex        =   3
         Top             =   735
         Width           =   1830
         _extentx        =   3228
         _extenty        =   900
         autotab         =   0
         tooltip         =   ""
         mascara         =   4
         label           =   "Data da Compra Final"
      End
      Begin Threed.SSFrame fraEntregue 
         Height          =   1050
         Left            =   10560
         TabIndex        =   14
         Top             =   180
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
            TabIndex        =   15
            Top             =   705
            Width           =   795
         End
         Begin VB.OptionButton optNao 
            Caption         =   "Não"
            Height          =   345
            Left            =   165
            TabIndex        =   5
            Top             =   420
            Width           =   750
         End
         Begin VB.OptionButton optSim 
            Caption         =   "Sim"
            Height          =   255
            Left            =   165
            TabIndex        =   4
            Top             =   225
            Width           =   675
         End
      End
      Begin Transportes.SuperControlNovo mskDataCompraInicial 
         Height          =   510
         Left            =   8640
         TabIndex        =   2
         Top             =   225
         Width           =   1830
         _extentx        =   3228
         _extenty        =   900
         autotab         =   0
         tooltip         =   ""
         mascara         =   4
         label           =   "Data da Compra Inicial"
      End
      Begin Transportes.SuperDBCombo cboComprador 
         Height          =   510
         Left            =   4500
         TabIndex        =   1
         Top             =   225
         Width           =   4080
         _extentx        =   7197
         _extenty        =   900
         label           =   "Comprador"
      End
      Begin Transportes.SuperDBCombo cboFornecedor 
         Height          =   510
         Left            =   195
         TabIndex        =   0
         Top             =   225
         Width           =   4215
         _extentx        =   7435
         _extenty        =   900
         label           =   "Fornecedor"
      End
   End
   Begin Transportes.SuperSpreadNovo sprConsulta 
      Height          =   4785
      Left            =   75
      TabIndex        =   8
      Top             =   1500
      Width           =   12585
      _extentx        =   22199
      _extenty        =   8440
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
Attribute VB_Name = "frmPedidoCompraConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sTabela As String
Dim sCampos As String

Private Sub cmdBaixarRecebimento_Click()
On Error GoTo err_cmdBaixarRecebimento_Click
    
    sprConsulta.Row = sprConsulta.ActiveRow
    If sprConsulta.RowHidden = False And Val(sprConsulta.SpreadEventoName("id_Compra")) > 0 Then
        Set frmPedidoCompraBaixaRecebimento.FormChamador = Me
        frmPedidoCompraBaixaRecebimento.id_Compra = sprConsulta.SpreadEventoName("id_Compra")
        frmPedidoCompraBaixaRecebimento.Show vbModal
    End If
    
    Exit Sub
err_cmdBaixarRecebimento_Click:
    ShowError
End Sub

Private Sub Form_Activate()
    frmMDI.Arrange vbCascade
End Sub

Private Sub Form_Load()
On Error GoTo err_FormLoad
       
    Call sprConsulta.NovaColunaSpread(eslnumero, True, True, "id_Compra", "id_Compra", 0, 15)
    Call sprConsulta.NovaColunaSpread(eslnumero, True, True, "id_Compra", "Pedido de Compra", 15, 15)
    Call sprConsulta.NovaColunaSpread(eslTexto, True, True, "ds_Pessoa", "Fornecedor", 30, 255)
    Call sprConsulta.NovaColunaSpread(eslData, True, True, "dt_Compra", "Data da Compra", 10, 10)
    Call sprConsulta.NovaColunaSpread(eslData, True, True, "dt_PrevisaoEntrega", "Data Prev. Entre", 10, 10)
    Call sprConsulta.NovaColunaSpread(eslData, True, True, "dt_Entrega", "Data Entrega", 10, 10)
    Call sprConsulta.NovaColunaSpread(eslValor, True, True, "vl_Compra", "Valor da Compra", 15, 15)
    Call sprConsulta.NovaColunaSpread(eslValor, True, True, "vl_Frete", "Valor do Frete", 15, 15)
    Call sprConsulta.NovaColunaSpread(eslValor, True, True, "vl_Adicional", "Valor Adicional", 15, 15)
    Call sprConsulta.NovaColunaSpread(eslTexto, True, True, "ds_observacao", "Observação", 30, 255)
    Call sprConsulta.FormatarNovo(21)

    sprConsulta.ColsFrozenName = "ds_Pessoa"

    sCampos = "comp.id_Compra,comp.id_Compra,fornec.ds_Pessoa,comp.dt_Compra,comp.dt_PrevisaoEntrega,comp.dt_Entrega,comp.vl_Compra,comp.vl_Frete,comp.vl_Adicional,comp.ds_observacao"
    
    sTabela = "((tbdCompra comp" _
    & " LEFT JOIN tbdpessoa fornec ON comp.id_Fornecedor = fornec.id_Pessoa) " _
    & " LEFT JOIN tbdpessoa funciona ON comp.id_Comprador = funciona.id_Pessoa) "

    Exit Sub
err_FormLoad:
    ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPedidoCompraConsulta = Nothing
End Sub

Private Sub cmdPesquisar_Click()
On Error GoTo err_cmdPesquisar

    Dim sWhere As String
    
    sWhere = MontarWhere
        
    If sWhere = "" Then
        Mensagem "Favor preencher algum parametro de pesquisa!", Informacao
        Exit Sub
    End If

    Call sprConsulta.Carregar(Select_Table(False, sTabela, sCampos, sWhere, "comp.id_Compra"))
    
    Exit Sub
err_cmdPesquisar:
    ShowError
End Sub
Private Function MontarWhere() As String
On Error GoTo err_MontarWhere
    Dim sWhere As String
    
    MontarWhere = ""
    sWhere = ""
     
    If cboFornecedor.ItemData2 > 0 Then
        sWhere = sWhere & " id_Fornecedor = " & cboFornecedor.ItemData2 & " AND "
    End If
    
    If cboComprador.ItemData2 > 0 Then
        sWhere = sWhere & " id_Comprador = " & cboComprador.ItemData2 & " AND "
    End If
    
    If mskDataCompraInicial.ClipText <> "" Then
        sWhere = sWhere & " dt_Compra >= " & ConverterData(mskDataCompraInicial.Text) & " AND "
    End If
    
    If mskDataCompraFinal.ClipText <> "" Then
        sWhere = sWhere & " dt_Compra <= " & ConverterData(mskDataCompraFinal.Text) & " AND "
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
    Set frmPedidoCompraDados.FormChamador = Me
    frmPedidoCompraDados.id_Compra = 0
    frmPedidoCompraDados.Show vbModal
End Sub

Private Sub cmdAlterar_Click()
On Error GoTo err_cmdAlterar_Click
    
    sprConsulta.Row = sprConsulta.ActiveRow
    If sprConsulta.RowHidden = False And Val(sprConsulta.SpreadEventoName("id_Compra")) > 0 Then
        Set frmPedidoCompraDados.FormChamador = Me
        frmPedidoCompraDados.id_Compra = sprConsulta.SpreadEventoName("id_Compra")
        frmPedidoCompraDados.Show vbModal
    End If
    
    Exit Sub
err_cmdAlterar_Click:
    ShowError
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo err_cmdExcluir_Click

    Dim cCompra As New clsCompra

    If sprConsulta.ActiveRow < 1 Then
        Mensagem "Selecione o item que será excluído.", erro
        Exit Sub
    End If

    If Mensagem("Confirma exclusão?", Pergunta) = vbNo Then
        Exit Sub
    End If

    Call cCompra.CarregarDados(sprConsulta.SpreadEventoName("id_Compra"))
    
    
    
    Call AbreTransacao
    If Not cCompra.Excluir Then
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
        
    cryRelatorio.ReportFileName = sPathReport & "\Relatorio\PedidoCompra.rpt"
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

Public Sub AtualizarDados(id_Compra As Long)
    Call sprConsulta.AtualizarDadosSpread(id_Compra, "comp.id_Compra", sTabela, sCampos)
End Sub

Private Sub sprConsulta_ButtonClickedName(ByVal ColName As String, ByVal Row As Long, ByVal ButtonDown As Integer)
On Error GoTo err_sprConsulta_ButtonClickedName

    sprConsulta.Row = Row
    If ColName = "cmdItem" Then
        frmProdutoItem.id_Compra = sprConsulta.SpreadEventoName("id_Compra")
        frmProdutoItem.Show
    End If
    
    Exit Sub
err_sprConsulta_ButtonClickedName:
    ShowError
End Sub
