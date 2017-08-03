VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmModeloConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modelo Consulta"
   ClientHeight    =   7155
   ClientLeft      =   4470
   ClientTop       =   3000
   ClientWidth     =   12720
   Icon            =   "frmModeloConsulta.frx":0000
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
      Picture         =   "frmModeloConsulta.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6345
      Width           =   810
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "&Pesquisar"
      Height          =   750
      Left            =   11850
      Picture         =   "frmModeloConsulta.frx":0E4E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Pesquisar os Dados"
      Top             =   210
      Width           =   810
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   750
      Left            =   11850
      Picture         =   "frmModeloConsulta.frx":1A10
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
      Picture         =   "frmModeloConsulta.frx":1D1A
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
      Picture         =   "frmModeloConsulta.frx":25E4
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
      Picture         =   "frmModeloConsulta.frx":2EAE
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
      Width           =   11715
      _Version        =   65536
      _ExtentX        =   20664
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
   End
   Begin Transportes.SuperSpreadNovo sprConsulta 
      Height          =   5235
      Left            =   75
      TabIndex        =   0
      Top             =   1050
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   9234
      ControlaQueryAdvance=   0   'False
      EsconderUltimaLinha=   -1  'True
   End
   Begin Crystal.CrystalReport cryRelatorio 
      Left            =   240
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "frmModeloConsulta"
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

    '[FORMATAÇÃO-COMBOS]'
        
    '[FORMATAÇÃO-SPREAD]'

    sprConsulta.ColsFrozenName = "id_Principal"

    '[FORMATAÇÃO-CAMPOS]'
    
    '[FORMATAÇÃO-TABELA]'

    Exit Sub
err_FormLoad:
    ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Me = Nothing
End Sub

Private Sub cmdPesquisar_Click()
On Error GoTo err_cmdPesquisar

    Dim sWhere As String
    
        
    'select id_Pessoa, ds_Pessoa, id_Cidade, '', ds_Logradouro,nr_Telefone from tbdPauloPessoa
    Call sprConsulta.Carregar(Select_Table(False, "tbdPauloPessoa a inner join tbdCidade b o (a.id_Cidade= b.id_Cidade)", "a.id_Pessoa, a.ds_Pessoa, , a.id_Cidade, b.ds_Cidade, a.ds_Logradouro, a.nr_Telefone", sWhere, "a.id_Pessoa, a.ds_Pessoa"))
    
    Exit Sub
err_cmdPesquisar:
    ShowError
End Sub

Private Sub cmdIncluir_Click()
    Set frmModeloDados.formChamador = Me
    frmModeloDados.id_Principal = 0
    frmModeloDados.Show vbModal
End Sub

Private Sub cmdAlterar_Click()
On Error GoTo err_cmdAlterar_Click
    
    sprConsulta.Row = sprConsulta.ActiveRow
    If sprConsulta.RowHidden = False And Val(sprConsulta.SpreadEventoName("id_Principal")) > 0 Then
        Set frmModeloDados.formChamador = Me
        frmModeloDados.id_Principal = sprConsulta.SpreadEventoName("id_Principal")
        frmModeloDados.Show vbModal
    End If
    
    Exit Sub
err_cmdAlterar_Click:
    ShowError
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo err_cmdExcluir_Click

    Dim cModelo As New clsModelo

    If sprConsulta.ActiveRow < 1 Then
        Mensagem "Selecione o item que será excluído.", erro
        Exit Sub
    End If

    If Mensagem("Confirma exclusão?", Pergunta) = vbNo Then
        Exit Sub
    End If

    Call AbreTransacao
    
    cModelo.id_Principal = sprConsulta.SpreadEventoName("id_Principal")
    If Not cModelo.Excluir Then
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
        
    cryRelatorio.ReportFileName = sPathReport & "\Principal.rpt"
    cryRelatorio.WindowParentHandle = frmMDI.hwnd
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

Public Sub AtualizarDados(id_Principal As Long)
    Call sprConsulta.AtualizarDadosSpread(id_Principal, "id_Principal", sTabela, sCampos)
End Sub
