VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGerarCodigoVB6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração de Código VB6"
   ClientHeight    =   8760
   ClientLeft      =   3480
   ClientTop       =   2400
   ClientWidth     =   9300
   Icon            =   "frmGerarCodigoVB6.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   9300
   Begin TabDlg.SSTab SSTab 
      Height          =   8070
      Left            =   75
      TabIndex        =   2
      Top             =   75
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   14235
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tabelas Utilizadas"
      TabPicture(0)   =   "frmGerarCodigoVB6.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTabelasFilhas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTabelaPrincipal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraJoin"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Configurações da Tela de Consulta"
      TabPicture(1)   =   "frmGerarCodigoVB6.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraConfiguracaoSpread"
      Tab(1).Control(1)=   "SSFrame"
      Tab(1).ControlCount=   2
      Begin Threed.SSFrame SSFrame 
         Height          =   2775
         Left            =   -74895
         TabIndex        =   24
         Top             =   390
         Width           =   8925
         _Version        =   65536
         _ExtentX        =   15743
         _ExtentY        =   4895
         _StockProps     =   14
         Caption         =   "Ordem dos Filtros"
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
         Begin Transportes.SuperSpreadNovo sprFiltros 
            Height          =   2385
            Left            =   120
            TabIndex        =   25
            Top             =   270
            Width           =   8700
            _ExtentX        =   15346
            _ExtentY        =   4207
            ControlaQueryAdvance=   0   'False
            EsconderUltimaLinha=   -1  'True
         End
      End
      Begin Threed.SSFrame fraJoin 
         Height          =   2685
         Left            =   105
         TabIndex        =   3
         Top             =   5145
         Width           =   8910
         _Version        =   65536
         _ExtentX        =   15716
         _ExtentY        =   4736
         _StockProps     =   14
         Caption         =   "Join das Tabelas para o spread de Consulta"
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
         Begin VB.TextBox txtJoin 
            Height          =   1665
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   7
            Top             =   930
            Width           =   8715
         End
         Begin VB.ComboBox cboCampoJoin 
            Height          =   315
            Left            =   90
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   495
            Width           =   4155
         End
         Begin VB.ComboBox cboTabelaJoin 
            Height          =   315
            Left            =   4275
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   495
            Width           =   3765
         End
         Begin VB.CommandButton cmdAdicionarJoin 
            Height          =   570
            Left            =   8175
            Picture         =   "frmGerarCodigoVB6.frx":0044
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Adicionar"
            Top             =   240
            Width           =   630
         End
         Begin VB.Label label3 
            AutoSize        =   -1  'True
            Caption         =   "ID para Join"
            Height          =   195
            Left            =   90
            TabIndex        =   9
            Top             =   285
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tabela para o Join"
            Height          =   195
            Left            =   4275
            TabIndex        =   8
            Top             =   285
            Width           =   1320
         End
      End
      Begin Threed.SSFrame fraTabelaPrincipal 
         Height          =   4740
         Left            =   105
         TabIndex        =   10
         Top             =   345
         Width           =   4425
         _Version        =   65536
         _ExtentX        =   7805
         _ExtentY        =   8361
         _StockProps     =   14
         Caption         =   "Tabela Principal"
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
         Begin VB.ComboBox cboTabela 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   525
            Width           =   4230
         End
         Begin VB.ListBox lstCampos 
            Height          =   3435
            Left            =   120
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   11
            Top             =   1155
            Width           =   4200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tabela"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   315
            Width           =   600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Campos para Filtros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   930
            Width           =   1680
         End
      End
      Begin Threed.SSFrame fraTabelasFilhas 
         Height          =   4740
         Left            =   4590
         TabIndex        =   15
         Top             =   345
         Width           =   4425
         _Version        =   65536
         _ExtentX        =   7805
         _ExtentY        =   8361
         _StockProps     =   14
         Caption         =   "Tabelas Filhas"
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
         Begin VB.ListBox lstTabelasFilhas 
            Height          =   4335
            Left            =   105
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   16
            Top             =   270
            Width           =   4200
         End
      End
      Begin Threed.SSFrame fraConfiguracaoSpread 
         Height          =   4740
         Left            =   -74895
         TabIndex        =   17
         Top             =   3180
         Width           =   8925
         _Version        =   65536
         _ExtentX        =   15743
         _ExtentY        =   8361
         _StockProps     =   14
         Caption         =   "Colunas do Spread"
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
         Begin VB.CommandButton cmdAdicionarCampoSpread 
            Height          =   570
            Left            =   6495
            Picture         =   "frmGerarCodigoVB6.frx":034E
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Adicionar"
            Top             =   225
            Width           =   630
         End
         Begin VB.ComboBox cboTabelaSpread 
            Height          =   315
            Left            =   105
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   480
            Width           =   3315
         End
         Begin VB.ComboBox cboCampoSpread 
            Height          =   315
            Left            =   3450
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   465
            Width           =   2985
         End
         Begin Transportes.SuperSpreadNovo sprConsulta 
            Height          =   3750
            Left            =   105
            TabIndex        =   21
            Top             =   870
            Width           =   8715
            _ExtentX        =   15372
            _ExtentY        =   6615
            ControlaQueryAdvance=   0   'False
            EsconderUltimaLinha=   -1  'True
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Tabela"
            Height          =   195
            Left            =   105
            TabIndex        =   23
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Campo"
            Height          =   195
            Left            =   3450
            TabIndex        =   22
            Top             =   270
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7440
      TabIndex        =   1
      Top             =   8190
      Width           =   1800
   End
   Begin VB.CommandButton cmdGerarCodigo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gerar Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5580
      TabIndex        =   0
      Top             =   8190
      Width           =   1800
   End
End
Attribute VB_Name = "frmGerarCodigoVB6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sfrmModeloConsulta As String
Dim sfrmModeloDados As String
Dim sclsModelo As String

Dim sfrmModeloConsultaAux As String
Dim sfrmModeloDadosAux As String
Dim sclsModeloAux As String

Dim ascAlias As String
Dim WidthFrameFiltros As Long
Dim WidthFrameDados  As Long
Dim HeightFrameDados As Long
Dim sPKPrincipal As String
Dim sNomeComponentesGeral As String

Private Sub cboCampoSpread_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        cboCampoSpread.ListIndex = -1
    End If
End Sub

Private Sub cboTabela_LostFocus()
On Error GoTo err_cboTabela_LostFocus
    
    Dim rsado As New ADODB.Recordset
    Dim scolid As String
    
    If ItemComboLista(cboTabela) <> Val(cboTabela.tag) Then
        If ItemComboLista(cboTabela) > 0 Then
            Call Carregar_ComboLista(lstCampos, Select_Table(False, "SysColumns", "colid,name", "id = " & ItemComboLista(cboTabela), "name"), 1, False)
            Call Carregar_ComboLista(cboCampoJoin, Select_Table(False, "SysColumns", "colid,name", "id = " & ItemComboLista(cboTabela) & " and mid(name,1,3) = 'id_'", "name"), 1, False)
        Else
            lstCampos.Clear
            cboCampoJoin.Clear
        End If
        
        Call Select_Table(True, "SysColumns", "colid, name", "id = " & ItemComboLista(cboTabela) & " and type = 56", , , , , rsado)
        sPKPrincipal = ReadField(rsado, "name")
        scolid = ReadField(rsado, "colid")
        rsado.Close
        
        If sPKPrincipal = "" Then
            sPKPrincipal = "id_" & Mid(cboTabela.Text, 4)
        End If
            
        txtJoin.Text = ""
        ascAlias = 97 'código do 'a'
        cboTabela.tag = ItemComboLista(cboTabela)
        txtJoin.Text = cboTabela.Text & " " & Chr(ascAlias)
        
        cboTabelaSpread.Clear
        cboTabelaSpread.AddItem cboTabela.Text & " " & Chr(ascAlias)
        cboTabelaSpread.ItemData(cboTabelaSpread.NewIndex) = ItemComboLista(cboTabela)
        
        sprConsulta.MaxRows = sprConsulta.MaxRows + 1
        sprConsulta.Row = sprConsulta.MaxRows
        sprConsulta.TextCol("id_Coluna") = scolid
        sprConsulta.TextCol("ds_Coluna") = sPKPrincipal
        sprConsulta.TextCol("ds_Cabecalho") = sPKPrincipal
        sprConsulta.TextCol("ds_Conteudo") = Chr(ascAlias) & "." & sPKPrincipal
        sprConsulta.TextCol("nr_Tamanho") = 0
        sprConsulta.TextCol("nr_Tipo") = 56
    End If
    
    Exit Sub
err_cboTabela_LostFocus:
    ShowError
End Sub

Private Sub cboTabelaSpread_LostFocus()
    If ItemComboLista(cboTabelaSpread) <> Val(cboTabelaSpread.tag) Then
        If ItemComboLista(cboTabelaSpread) > 0 Then
            Call Carregar_ComboLista(cboCampoSpread, Select_Table(False, "SysColumns", "colid,name", "id = " & ItemComboLista(cboTabelaSpread), "name"), 1, False)
        Else
            cboCampoSpread.Clear
        End If
    End If
End Sub

Private Sub cmdAdicionarCampoSpread_Click()
On Error GoTo err_cmdAdicionarCampoSpread_Click
    
    Dim bAdicionarTodos As Boolean
    Dim i As Integer
    Dim rsado As New ADODB.Recordset
    
    bAdicionarTodos = False
    
    If cboTabelaSpread.ListIndex < 0 Then
        Mensagem "Selecione a Tabela.", erro
        cboTabelaSpread.SetFocus
        Exit Sub
    End If
    
    If cboCampoSpread.ListIndex < 0 Then
        If Mensagem("Adicionar todos os campos da tabela " & cboTabelaSpread & "?", Pergunta) = vbNo Then
            cboCampoSpread.SetFocus
            Exit Sub
        End If
        bAdicionarTodos = True
    End If
    
    If Not bAdicionarTodos Then
        sprConsulta.MaxRows = sprConsulta.MaxRows + 1
        sprConsulta.Row = sprConsulta.MaxRows
        sprConsulta.TextCol("id_Coluna") = ItemComboLista(cboCampoSpread)
        sprConsulta.TextCol("ds_Coluna") = cboCampoSpread.Text
        sprConsulta.TextCol("ds_Cabecalho") = Mid(cboCampoSpread.Text, 4)
        sprConsulta.TextCol("ds_Conteudo") = Right(cboTabelaSpread.Text, 1) & "." & cboCampoSpread.Text
        
        Call Select_Table(True, "SysColumns", "type, length", "id = " & ItemComboLista(cboTabelaSpread) & " and colid = " & ItemComboLista(cboCampoSpread), , , , , rsado)
        sprConsulta.TextCol("nr_Tamanho") = ReadField(rsado, "length")
        sprConsulta.TextCol("nr_Tipo") = ReadField(rsado, "type")
        rsado.Close
    Else
        For i = 0 To cboCampoSpread.ListCount - 1
            
            sprConsulta.MaxRows = sprConsulta.MaxRows + 1
            sprConsulta.Row = sprConsulta.MaxRows
            
            sprConsulta.TextCol("id_Coluna") = cboCampoSpread.ItemData(i)
            sprConsulta.TextCol("ds_Coluna") = cboCampoSpread.List(i)
            sprConsulta.TextCol("ds_Cabecalho") = Mid(cboCampoSpread.List(i), 4)
            sprConsulta.TextCol("ds_Conteudo") = Right(cboTabelaSpread.Text, 1) & "." & cboCampoSpread.List(i)
            
            Call Select_Table(True, "SysColumns", "type, length", "id = " & ItemComboLista(cboTabelaSpread) & " and colid = " & cboCampoSpread.ItemData(i), , , , , rsado)
            sprConsulta.TextCol("nr_Tamanho") = ReadField(rsado, "length")
            sprConsulta.TextCol("nr_Tipo") = ReadField(rsado, "type")
            rsado.Close
        Next i
    End If
    
    Exit Sub
err_cmdAdicionarCampoSpread_Click:
    ShowError
End Sub

Private Sub cmdAdicionarJoin_Click()
On Error GoTo err_cmdAdicionarJoin_Click
    
    Dim rsado As New ADODB.Recordset
    Dim sPK As String
    Dim i As Integer
    Dim bAchou As Boolean
    
    If cboCampoJoin.ListIndex < 0 Then
        Mensagem "Selecione o Campo para montar o Join.", Informacao
        cboCampoJoin.SetFocus
        Exit Sub
    End If
    If cboTabelaJoin.ListIndex < 0 Then
        Mensagem "Selecione a tabela para montar o Join.", Informacao
        cboTabelaJoin.SetFocus
        Exit Sub
    End If
    
    Call Select_Table(True, "SysColumns", "name", "id = " & ItemComboLista(cboTabelaJoin) & " and type = 56", , , , , rsado)
    sPK = ReadField(rsado, "name")
    rsado.Close
    
    If sPK = "" Then
        sPK = "id_" & Mid(cboTabelaJoin.Text, 4)
    End If
        
    ascAlias = ascAlias + 1
    
    txtJoin.Text = "(" & txtJoin.Text & vbCrLf _
    & " left join " & cboTabelaJoin.Text & " " & Chr(ascAlias) & " on a." & cboCampoJoin.Text & " = " & Chr(ascAlias) & "." & sPK & ") "
    
    bAchou = False
    For i = 0 To cboTabelaSpread.ListCount - 1
        If cboTabelaSpread.ItemData(i) = ItemComboLista(cboTabelaJoin) Then
            bAchou = True
            Exit For
        End If
    Next i
    
    If Not bAchou Then
        cboTabelaSpread.AddItem cboTabelaJoin.Text & " " & Chr(ascAlias)
        cboTabelaSpread.ItemData(cboTabelaSpread.NewIndex) = ItemComboLista(cboTabelaJoin)
    End If
        
    Exit Sub
err_cmdAdicionarJoin_Click:
    ShowError
End Sub

Private Sub cmdGerarCodigo_Click()
On Error GoTo err_cmdGerarCodigo_Click
    Dim sCombos As String
    Dim sColunasSpread As String
    Dim sCampos As String
    Dim sCamposUpdate As String
    Dim sPropriedades As String
    Dim sConteudo As String
    Dim sConteudoUpdate As String
    Dim sInicializar As String
    Dim sInicializarItem As String
    Dim sSetarPropriedades As String
    Dim sGravarItens As String
    Dim sDeletarItens As String
    Dim sFuncaoItens As String
    Dim snomeArrayItem As String
    Dim sParametrosCarregarItem As String
    Dim sCamposItem As String
    Dim sConteudoItem As String
    Dim iLenCampos As Long
    Dim i As Long
    Dim iCont As Integer
    Dim sComponentes As String
    Dim qt_Item As Integer
    Dim iLeft As Long
    Dim iTop As Long
    Dim iTabIndex As Integer
    Dim iPos1 As Long
    Dim iPos2 As Long
    Dim bAuxiliar As Boolean
    Dim sPKItem As String
    Dim aJoin As Variant
    Dim sJoin As String
    Dim sSetarItem As String
    Dim sCarregarItens As String
    Dim sChamadaCarregarItens As String
    Dim sNomeComponenteAtual As String
    Dim sCamposCarregarItem As String
    Dim sZerarItens As String
    Dim sAtualizarSpread As String
    
    Call sprFiltros.SpreadClickName(0, "nr_Ordem", True)
    sfrmModeloConsulta = sfrmModeloConsultaAux
    sfrmModeloDados = sfrmModeloDadosAux
    sclsModelo = sclsModeloAux
            
    '-------------------------------------------------------------------------------------------------------------------------------------------
    'Código para criar o form de consulta
    '-------------------------------------------------------------------------------------------------------------------------------------------
    sfrmModeloConsulta = Replace(sfrmModeloConsulta, "frmModeloConsulta", "frm" & Mid(cboTabela.Text, 4) & "Consulta", , , vbTextCompare)
    sfrmModeloConsulta = Replace(sfrmModeloConsulta, "Modelo Consulta", "Cadastro de " & Mid(cboTabela.Text, 4), , , vbTextCompare)
    sfrmModeloConsulta = Replace(sfrmModeloConsulta, "id_Principal", sPKPrincipal, , , vbTextCompare)
    sfrmModeloConsulta = Replace(sfrmModeloConsulta, "Principal.rpt", Mid(cboTabela.Text, 4) & ".rpt", , , vbTextCompare)
    sfrmModeloConsulta = Replace(sfrmModeloConsulta, "frmModeloDados", "frm" & Mid(cboTabela.Text, 4) & "Dados", , , vbTextCompare)
    sfrmModeloConsulta = Replace(sfrmModeloConsulta, "clsModelo", "cls" & Mid(cboTabela.Text, 4), , , vbTextCompare)
    sfrmModeloConsulta = Replace(sfrmModeloConsulta, "cModelo", "c" & Mid(cboTabela.Text, 4), , , vbTextCompare)
    sfrmModeloConsulta = Replace(sfrmModeloConsulta, "Set Me = Nothing", "Set frm" & Mid(cboTabela.Text, 4) & "Consulta = Nothing", , , vbTextCompare)
    
    'Formata os filtros
    iLeft = 0
    iTop = 0
    iTabIndex = 8
    sCombos = ""
    sColunasSpread = ""
    sCampos = ""
    sCamposUpdate = ""
    sPropriedades = ""
    sConteudo = ""
    sConteudoUpdate = ""
    sInicializar = ""
    sSetarPropriedades = ""
    sGravarItens = ""
    sDeletarItens = ""
    sFuncaoItens = ""
    snomeArrayItem = ""
    sCamposItem = ""
    sConteudoItem = ""
    sComponentes = ""
    sParametrosCarregarItem = ""
    sSetarItem = ""
    sCarregarItens = ""
    sChamadaCarregarItens = ""
    sInicializarItem = ""
    sNomeComponentesGeral = ""
    sZerarItens = ""
    sCamposCarregarItem = ""
    sAtualizarSpread = ""
    
    For i = 1 To sprFiltros.MaxRows
        sprFiltros.Row = i
        
        sComponentes = sComponentes & CriarComponente(Mid(sprFiltros.SpreadEventoName("ds_Campo"), 3), iTop, iLeft, iTabIndex, WidthFrameFiltros, True, sNomeComponenteAtual)
        If InStr(1, sprFiltros.SpreadEventoName("ds_Filtro"), "Combo", vbTextCompare) > 0 Then
            sCombos = sCombos & IIf(sCombos <> "", "    ", "") & "Call " & sNomeComponenteAtual & ".FormatarComboPadrao(" & PegarTipoComboFiltro(Mid(sprFiltros.SpreadEventoName("ds_Campo"), 6)) & ")" & vbCrLf
        End If
    Next i
    sfrmModeloConsulta = Replace(sfrmModeloConsulta, "'[FORMATAÇÃO-COMBOS]'", sCombos, , , vbTextCompare)
            
    'Cria a variavel sCampos da consulta e colunas do spread
    sCampos = ""
    sColunasSpread = ""
    iLenCampos = 0
    For i = 1 To sprConsulta.MaxRows
        sprConsulta.Row = i
        
        iLenCampos = iLenCampos + Len(sprConsulta.SpreadEventoName("ds_Conteudo"))
        
        If iLenCampos < 130 Then
            sCampos = sCampos & sprConsulta.SpreadEventoName("ds_Conteudo") & ", "
        Else
            sCampos = sCampos & Chr(34) & " _" & vbCrLf & "    & " & Chr(34) & sprConsulta.SpreadEventoName("ds_Conteudo") & ", "
            iLenCampos = 0
        End If
        
        sColunasSpread = sColunasSpread & IIf(sColunasSpread <> "", "    ", "") & "Call sprConsulta.NovaColunaSpread(" & RetornarTipoColunaSpread(Mid(sprConsulta.SpreadEventoName("ds_Coluna"), 1, 2)) & ",True, True," & Chr(34) & sprConsulta.SpreadEventoName("ds_Coluna") & Chr(34) & "," & Chr(34) & sprConsulta.SpreadEventoName("ds_Cabecalho") & Chr(34) & "," & RetornarTamanhoColunaSpread(Mid(sprConsulta.SpreadEventoName("ds_Coluna"), 1, 2)) & "," & sprConsulta.SpreadEventoName("nr_Tamanho") & ")" & vbCrLf
    Next i
    If sCampos <> "" Then
        sColunasSpread = sColunasSpread & "    Call sprConsulta.FormatarNovo(21)" & vbCrLf
        
        sCampos = "sCampos = " & Chr(34) & Left(sCampos, Len(sCampos) - 2) & Chr(34)
        sfrmModeloConsulta = Replace(sfrmModeloConsulta, "'[FORMATAÇÃO-CAMPOS]'", sCampos, , , vbTextCompare)
        sfrmModeloConsulta = Replace(sfrmModeloConsulta, "'[FORMATAÇÃO-SPREAD]'", sColunasSpread, , , vbTextCompare)
    End If
    
    'Cria a variável sTabela
    aJoin = Split(txtJoin.Text, vbCrLf)
    For i = 0 To UBound(aJoin)
        If i = 0 Then
            sJoin = Chr(34) & aJoin(i) & Chr(34) & " _" & vbCrLf
        Else
            sJoin = sJoin & "    & " & Chr(34) & aJoin(i) & Chr(34) & " _" & vbCrLf
        End If
    Next i
    If sJoin <> "" Then
        sJoin = Left(sJoin, Len(sJoin) - 3)
    End If
    sfrmModeloConsulta = Replace(sfrmModeloConsulta, "'[FORMATAÇÃO-TABELA]'", "sTabela = " & sJoin, , , vbTextCompare)
        
    'Cria os componentes
    iPos1 = InStr(1, sfrmModeloConsulta, "fraFiltro", vbTextCompare)
    If iPos1 > 0 Then
        iPos2 = InStr(iPos1, sfrmModeloConsulta, "ShadowStyle", vbTextCompare)
        If iPos2 = 0 Then
            iPos2 = iPos1
        End If
        iPos1 = InStr(iPos2, sfrmModeloConsulta, "End", vbTextCompare)
        If iPos1 > 0 Then
            sfrmModeloConsulta = Mid(sfrmModeloConsulta, 1, iPos1 - 3) & vbCrLf & sComponentes & Mid(sfrmModeloConsulta, iPos1)
        End If
    End If
    
    '-------------------------------------------------------------------------------------------------------------------------------------------
    'Gerar a Classe de Gravação
    '-------------------------------------------------------------------------------------------------------------------------------------------
    sclsModelo = Replace(sclsModelo, "clsModelo", "cls" & Mid(cboTabela.Text, 4), , , vbTextCompare)
    sclsModelo = Replace(sclsModelo, "id_Principal", sPKPrincipal, , , vbTextCompare)
    sclsModelo = Replace(sclsModelo, "tbdPrincipal", cboTabela.Text, , , vbTextCompare)
    
    sInicializar = ""
    sPropriedades = ""
    sCampos = ""
    sConteudo = ""
    iLenCampos = 0
    Call Select_Table(True, "syscolumns", "name, type", "id = " & ItemComboLista(cboTabela), "colid")
    Do While Not rsADOGlobal.EOF
        
        If ReadField(rsADOGlobal, "name") <> sPKPrincipal Then
            iLenCampos = iLenCampos + Len(ReadField(rsADOGlobal, "name"))
            If iLenCampos < 130 Then
                sCampos = sCampos & ReadField(rsADOGlobal, "name") & ", "
                sConteudo = sConteudo & ReadField(rsADOGlobal, "name") & ", "
            Else
                sCampos = sCampos & Chr(34) & " _" & vbCrLf & "    & " & Chr(34) & ReadField(rsADOGlobal, "name") & ", "
                sConteudo = sConteudo & " _" & vbCrLf & "    " & ReadField(rsADOGlobal, "name") & ", "
                iLenCampos = 0
            End If
            
            sSetarPropriedades = sSetarPropriedades & IIf(sSetarPropriedades <> "", "        ", "") & ReadField(rsADOGlobal, "name") & " = ReadField(rsado, " & Chr(34) & ReadField(rsADOGlobal, "name") & Chr(34) & ")" & vbCrLf
        End If
        
        sInicializar = sInicializar & IIf(sInicializar <> "", "    ", "") & ReadField(rsADOGlobal, "name") & " = " & PegarValorPadrao(ReadField(rsADOGlobal, "type")) & vbCrLf
        
        sPropriedades = sPropriedades & "Public " & ReadField(rsADOGlobal, "name") & " as " & PegarTipoPropriedade(ReadField(rsADOGlobal, "type")) & vbCrLf
        rsADOGlobal.MoveNext
    Loop
    rsADOGlobal.Close
        
    bAuxiliar = True
    sGravarItens = ""
    sDeletarItens = ""
    sFuncaoItens = ""
    
    For i = 0 To lstTabelasFilhas.ListCount - 1
        If lstTabelasFilhas.Selected(i) Then
            If bAuxiliar Then
                sPropriedades = sPropriedades & vbCrLf
                sInicializar = sInicializar & vbCrLf
                bAuxiliar = False
            End If
            
            sPKItem = ""
            sCamposItem = ""
            sConteudoItem = ""
            snomeArrayItem = "a" & Mid(lstTabelasFilhas.List(i), 4) & "Item"
            sParametrosCarregarItem = ""
            sCamposCarregarItem = ""
            sSetarItem = ""
            iCont = 2
            
            Call Select_Table(True, "syscolumns", "name, type", "id = " & lstTabelasFilhas.ItemData(i), "colid")
            qt_Item = rsADOGlobal.RecordCount - 1
            Do While Not rsADOGlobal.EOF
            
                If UCase(ReadField(rsADOGlobal, "name")) <> UCase(sPKPrincipal) Then
                    sCamposCarregarItem = sCamposCarregarItem & "ReadField(rsado," & Chr(34) & ReadField(rsADOGlobal, "name") & Chr(34) & "), "
                    sParametrosCarregarItem = sParametrosCarregarItem & ReadField(rsADOGlobal, "name") & " as " & PegarTipoPropriedade(ReadField(rsADOGlobal, "type")) & ", "
                End If
                    
                If ReadField(rsADOGlobal, "type") = 56 Then
                    sPKItem = ReadField(rsADOGlobal, "name")
                Else
                    sCamposItem = sCamposItem & ReadField(rsADOGlobal, "name") & ", "
                                        
                    sSetarItem = sSetarItem & "    " & snomeArrayItem & "(" & iCont & ", UBound(" & snomeArrayItem & ", 2)) = " & ReadField(rsADOGlobal, "name") & vbCrLf
                                                            
                    If ReadField(rsADOGlobal, "type") = 56 Or ReadField(rsADOGlobal, "type") = 36 Then
                        sConteudoItem = sConteudoItem & "Val(" & snomeArrayItem & "(" & iCont & ",iCont)), "
                    ElseIf ReadField(rsADOGlobal, "type") = 106 Then
                        sConteudoItem = sConteudoItem & "Cdbl1(" & snomeArrayItem & "(" & iCont & ",iCont)), "
                    ElseIf ReadField(rsADOGlobal, "type") = 111 Then
                        sConteudoItem = sConteudoItem & "CDateEspecial(" & snomeArrayItem & "(" & iCont & ",iCont)), "
                    Else
                        sConteudoItem = sConteudoItem & snomeArrayItem & "(" & iCont & ",iCont), "
                    End If
                                       
                    iCont = iCont + 1
                End If
                
                rsADOGlobal.MoveNext
            Loop
            rsADOGlobal.Close
            
            If sPKItem = "" Then
                sPKItem = "id_" & Mid(lstTabelasFilhas.List(i), 4)
            End If
            If sCamposItem <> "" Then
                sCamposCarregarItem = Left(sCamposCarregarItem, Len(sCamposCarregarItem) - 2)
                sCamposItem = Left(sCamposItem, Len(sCamposItem) - 2)
                sConteudoItem = Left(sConteudoItem, Len(sConteudoItem) - 2)
                sParametrosCarregarItem = Left(sParametrosCarregarItem, Len(sParametrosCarregarItem) - 2)
            End If
                                    
            sPropriedades = sPropriedades & "Public " & snomeArrayItem & " as Variant" & vbCrLf
            sInicializarItem = sInicializarItem & "    Redim " & snomeArrayItem & "(" & qt_Item & ",0)" & vbCrLf
                        
            sGravarItens = sGravarItens & IIf(sGravarItens <> "", "    ", "") & "If Not Gravar" & Mid(lstTabelasFilhas.List(i), 4) & "Item() Then" & vbCrLf
            sGravarItens = sGravarItens & "        Exit function" & vbCrLf
            sGravarItens = sGravarItens & "    End If" & vbCrLf & vbCrLf
            
            sDeletarItens = sDeletarItens & IIf(sDeletarItens <> "", "    ", "") & "If Delete_Table(" & Chr(34) & lstTabelasFilhas.List(i) & Chr(34) & "," & Chr(34) & sPKPrincipal & " = " & Chr(34) & " & " & sPKPrincipal & ") = -1 Then" & vbCrLf
            sDeletarItens = sDeletarItens & "        Exit function" & vbCrLf
            sDeletarItens = sDeletarItens & "    End If" & vbCrLf & vbCrLf
                        
            sChamadaCarregarItens = sChamadaCarregarItens & IIf(sChamadaCarregarItens <> "", "    ", "") & "Call Select_Table(True, " & Chr(34) & lstTabelasFilhas.List(i) & Chr(34) & ",""*""," & Chr(34) & sPKPrincipal & " = " & Chr(34) & " & " & sPKPrincipal & ", , , , , rsado)" & vbCrLf
            sChamadaCarregarItens = sChamadaCarregarItens & "    Do While Not rsado.EOF" & vbCrLf
            sChamadaCarregarItens = sChamadaCarregarItens & "        Call Carregar" & Mid(lstTabelasFilhas.List(i), 4) & "Item(const_Inicial, " & sCamposCarregarItem & ")" & vbCrLf
            sChamadaCarregarItens = sChamadaCarregarItens & "        rsado.MoveNext" & vbCrLf
            sChamadaCarregarItens = sChamadaCarregarItens & "    Loop" & vbCrLf
            sChamadaCarregarItens = sChamadaCarregarItens & "    rsado.Close" & vbCrLf & vbCrLf
            
            sCarregarItens = sCarregarItens & "Public Function Carregar" & Mid(lstTabelasFilhas.List(i), 4) & "Item(statusAtualizacao As Integer," & sParametrosCarregarItem & ")" & vbCrLf
            sCarregarItens = sCarregarItens & "On Error GoTo err_Carregar" & Mid(lstTabelasFilhas.List(i), 4) & "Item" & vbCrLf & vbCrLf
            sCarregarItens = sCarregarItens & "    ReDim Preserve " & snomeArrayItem & "(" & qt_Item & ",UBound(" & snomeArrayItem & ", 2) + 1)" & vbCrLf & vbCrLf
            sCarregarItens = sCarregarItens & "    " & snomeArrayItem & "(0, UBound(" & snomeArrayItem & ", 2)) = statusAtualizacao" & vbCrLf
            sCarregarItens = sCarregarItens & "    " & snomeArrayItem & "(1, UBound(" & snomeArrayItem & ", 2)) = " & sPKItem & vbCrLf
            sCarregarItens = sCarregarItens & sSetarItem & vbCrLf
            sCarregarItens = sCarregarItens & "    Exit Function" & vbCrLf
            sCarregarItens = sCarregarItens & "err_Carregar" & Mid(lstTabelasFilhas.List(i), 4) & "Item:" & vbCrLf
            sCarregarItens = sCarregarItens & "    ShowError" & vbCrLf
            sCarregarItens = sCarregarItens & "End Function" & vbCrLf & vbCrLf
            
            sFuncaoItens = sFuncaoItens & "Private Function Gravar" & Mid(lstTabelasFilhas.List(i), 4) & "Item() As Boolean" & vbCrLf
            sFuncaoItens = sFuncaoItens & "On Error GoTo err_Gravar" & Mid(lstTabelasFilhas.List(i), 4) & "Item" & vbCrLf & vbCrLf
            sFuncaoItens = sFuncaoItens & "    Dim iCont As Long" & vbCrLf
            sFuncaoItens = sFuncaoItens & "    Dim sWhere As String" & vbCrLf & vbCrLf
            sFuncaoItens = sFuncaoItens & "    Gravar" & Mid(lstTabelasFilhas.List(i), 4) & "Item = False" & vbCrLf & vbCrLf
            sFuncaoItens = sFuncaoItens & "    For iCont = 1 To UBound(" & snomeArrayItem & ", 2)" & vbCrLf
            sFuncaoItens = sFuncaoItens & "        If " & snomeArrayItem & "(0, iCont) = const_Insert Then" & vbCrLf
            sFuncaoItens = sFuncaoItens & "            " & snomeArrayItem & "(1, iCont) = Insert_Table(" & Chr(34) & lstTabelasFilhas.List(i) & Chr(34) & ", " & Chr(34) & sPKItem & Chr(34) & ", " & Chr(34) & sCamposItem & Chr(34) & ", Array(" & sConteudoItem & "))" & vbCrLf
            sFuncaoItens = sFuncaoItens & "            If " & snomeArrayItem & "(1, iCont) = -1 Then" & vbCrLf
            sFuncaoItens = sFuncaoItens & "                Exit Function" & vbCrLf
            sFuncaoItens = sFuncaoItens & "            End If" & vbCrLf
            sFuncaoItens = sFuncaoItens & "        ElseIf " & snomeArrayItem & "(0, iCont) = const_Update Then" & vbCrLf
            sFuncaoItens = sFuncaoItens & "            If Update_Table(" & Chr(34) & lstTabelasFilhas.List(i) & Chr(34) & ", " & Chr(34) & sCamposItem & Chr(34) & ", Array(" & sConteudoItem & "), " & Chr(34) & sPKItem & " = " & Chr(34) & " & " & snomeArrayItem & "(1, iCont)) = -1 Then" & vbCrLf
            sFuncaoItens = sFuncaoItens & "                Exit Function" & vbCrLf
            sFuncaoItens = sFuncaoItens & "            End If" & vbCrLf
            sFuncaoItens = sFuncaoItens & "        ElseIf " & snomeArrayItem & "(0, iCont) = const_Delete Then" & vbCrLf
            sFuncaoItens = sFuncaoItens & "            If Delete_Table(" & Chr(34) & lstTabelasFilhas.List(i) & Chr(34) & ", " & Chr(34) & sPKItem & " = " & Chr(34) & " & " & snomeArrayItem & "(1, iCont)) = -1 Then" & vbCrLf
            sFuncaoItens = sFuncaoItens & "                Exit Function" & vbCrLf
            sFuncaoItens = sFuncaoItens & "            End If" & vbCrLf
            sFuncaoItens = sFuncaoItens & "        End If" & vbCrLf
            sFuncaoItens = sFuncaoItens & "    Next iCont" & vbCrLf & vbCrLf
            sFuncaoItens = sFuncaoItens & "    Gravar" & Mid(lstTabelasFilhas.List(i), 4) & "Item = True" & vbCrLf & vbCrLf
            sFuncaoItens = sFuncaoItens & "    Exit Function" & vbCrLf
            sFuncaoItens = sFuncaoItens & "err_Gravar" & Mid(lstTabelasFilhas.List(i), 4) & "Item:" & vbCrLf
            sFuncaoItens = sFuncaoItens & "    ShowError" & vbCrLf
            sFuncaoItens = sFuncaoItens & "End Function" & vbCrLf & vbCrLf
        End If
    Next i
    
    If sDeletarItens <> "" Then
        sDeletarItens = Left(sDeletarItens, Len(sDeletarItens) - 2)
        sGravarItens = Left(sGravarItens, Len(sGravarItens) - 2)
    End If
    
    If sCampos <> "" Then
        sCampos = Left(sCampos, Len(sCampos) - 2)
        sCampos = "sCampos = " & Chr(34) & sCampos & Chr(34) & vbCrLf
        sCamposUpdate = sCampos
        
        sConteudo = Left(sConteudo, Len(sConteudo) - 2)
        sConteudo = "aConteudo = Array(" & sConteudo & ")"
        sConteudoUpdate = sConteudo
        
        sConteudo = Replace(sConteudo, "ds_Inclusao,", "sUsuario + '-' + DataAtual(true, true),")
        sConteudo = Replace(sConteudo, "ds_Alteracao,", "sUsuario + '-' + DataAtual(true, true),")
        
        sCamposUpdate = Replace(sCamposUpdate, "ds_Inclusao,", "")
        sConteudoUpdate = Replace(sConteudoUpdate, "ds_Inclusao,", "")
        sConteudoUpdate = Replace(sConteudoUpdate, "ds_Alteracao,", "sUsuario + '-' + DataAtual(true, true),")
    End If
    
    If sInicializarItem <> "" Then
        sInicializar = sInicializar & "    Call ZerarItens()"
        sZerarItens = "Public Function ZerarItens() as Boolean" & vbCrLf
        sZerarItens = sZerarItens & "On Error GoTo err_ZerarItens" & vbCrLf & vbCrLf
        sZerarItens = sZerarItens & "    ZerarItens = False" & vbCrLf & vbCrLf
        sZerarItens = sZerarItens & sInicializarItem & vbCrLf
        sZerarItens = sZerarItens & "    ZerarItens = True" & vbCrLf
        sZerarItens = sZerarItens & "    Exit Function" & vbCrLf
        sZerarItens = sZerarItens & "err_ZerarItens:" & vbCrLf
        sZerarItens = sZerarItens & "    ShowError" & vbCrLf
        sZerarItens = sZerarItens & "End Function" & vbCrLf
    End If
    
    sclsModelo = Replace(sclsModelo, "'[CRIAR-PROPRIEDADES]'", sPropriedades, , , vbTextCompare)
    sclsModelo = Replace(sclsModelo, "'[CAMPOS-INSERT]'", sCampos, , , vbTextCompare)
    sclsModelo = Replace(sclsModelo, "'[CONTEUDO-INSERT]'", sConteudo, , , vbTextCompare)
    sclsModelo = Replace(sclsModelo, "'[CAMPOS-UPDATE]'", sCamposUpdate, , , vbTextCompare)
    sclsModelo = Replace(sclsModelo, "'[CONTEUDO-UPDATE]'", sConteudoUpdate, , , vbTextCompare)
    sclsModelo = Replace(sclsModelo, "'[INICIALIZAR-PROPRIEDADES]'", sInicializar, , , vbTextCompare)
    sclsModelo = Replace(sclsModelo, "'[SETAR-PROPRIEDADES]'", sSetarPropriedades, , , vbTextCompare)
    sclsModelo = Replace(sclsModelo, "'[GRAVAR-ITENS]'", sGravarItens, , , vbTextCompare)
    sclsModelo = Replace(sclsModelo, "'[EXCLUIR-ITENS]'", sDeletarItens, , , vbTextCompare)
    sclsModelo = Replace(sclsModelo, "'[FUNCAO-GRAVARITENS]'", sFuncaoItens, , , vbTextCompare)
    sclsModelo = Replace(sclsModelo, "'[CARREGAR-ITENS]'", sCarregarItens, , , vbTextCompare)
    sclsModelo = Replace(sclsModelo, "'[CHAMADA-CARREGARITEM]'", sChamadaCarregarItens, , , vbTextCompare)
    sclsModelo = Replace(sclsModelo, "'[ZERAR-ITENS]'", sZerarItens, , , vbTextCompare)
    
    Call SalvarArquivo("frm" & Mid(cboTabela, 4) & "Consulta.frm", sfrmModeloConsulta)
    Call SalvarArquivo("cls" & Mid(cboTabela, 4) & ".cls", sclsModelo)
    
    
    '-------------------------------------------------------------------------------------------------------------------------------------------
    'Código para criar o form de dados
    '-------------------------------------------------------------------------------------------------------------------------------------------
    sfrmModeloDados = Replace(sfrmModeloDados, "frmModeloDados", "frm" & Mid(cboTabela.Text, 4) & "Dados", , , vbTextCompare)
    sfrmModeloDados = Replace(sfrmModeloDados, "Modelo Dados", "Cadastro de " & Mid(cboTabela.Text, 4), , , vbTextCompare)
    sfrmModeloDados = Replace(sfrmModeloDados, "id_Principal", sPKPrincipal, , , vbTextCompare)
    sfrmModeloDados = Replace(sfrmModeloDados, "clsModelo", "cls" & Mid(cboTabela.Text, 4), , , vbTextCompare)
    sfrmModeloDados = Replace(sfrmModeloDados, "cModelo", "c" & Mid(cboTabela.Text, 4), , , vbTextCompare)
    sfrmModeloDados = Replace(sfrmModeloDados, "Set Me = Nothing", "Set frm" & Mid(cboTabela.Text, 4) & "Dados = Nothing", , , vbTextCompare)
    
    sCombos = ""
    sComponentes = ""
    iTabIndex = 4
    iTop = 0
    iLeft = 0
    sNomeComponentesGeral = ""
    sSetarPropriedades = "c" & Mid(cboTabela.Text, 4) & "." & sPKPrincipal & " = " & sPKPrincipal & vbCrLf
    sCarregarItens = ""
    Call Select_Table(True, "syscolumns", "name, type, length", "id = " & ItemComboLista(cboTabela), "name")
    Do While Not rsADOGlobal.EOF
        If ReadField(rsADOGlobal, "type") <> 56 Then 'diferente da chave primária
            sComponentes = sComponentes & CriarComponente(ReadField(rsADOGlobal, "name"), iTop, iLeft, iTabIndex, WidthFrameDados, False, sNomeComponenteAtual)
            If Mid(ReadField(rsADOGlobal, "name"), 1, 2) = "id" Then
                sCombos = sCombos & IIf(sCombos <> "", "    ", "") & "Call " & sNomeComponenteAtual & ".FormatarComboPadrao(" & PegarTipoComboFiltro(Mid(ReadField(rsADOGlobal, "name"), 4)) & ")" & vbCrLf
            End If
            
            If Mid(ReadField(rsADOGlobal, "name"), 1, 2) = "id" Then
                sCarregarItens = sCarregarItens & IIf(sCarregarItens <> "", "            ", "") & "Call " & sNomeComponenteAtual & ".PesquisarCombo(True, " & "c" & Mid(cboTabela.Text, 4) & "." & ReadField(rsADOGlobal, "name") & "," & Chr(34) & Chr(34) & ", True)" & vbCrLf
            ElseIf ReadField(rsADOGlobal, "type") = 106 Then
                sCarregarItens = sCarregarItens & IIf(sCarregarItens <> "", "            ", "") & sNomeComponenteAtual & ".Text = " & "Format(c" & Mid(cboTabela.Text, 4) & "." & ReadField(rsADOGlobal, "name") & "," & Chr(34) & "###,###0.00" & Chr(34) & ")" & vbCrLf
            Else
                sCarregarItens = sCarregarItens & IIf(sCarregarItens <> "", "            ", "") & sNomeComponenteAtual & ".Text = " & "c" & Mid(cboTabela.Text, 4) & "." & ReadField(rsADOGlobal, "name") & vbCrLf
            End If
            
            sSetarPropriedades = sSetarPropriedades & "    c" & Mid(cboTabela.Text, 4) & "." & ReadField(rsADOGlobal, "name") & " = "
            If Mid(ReadField(rsADOGlobal, "name"), 1, 2) = "id" Then
                sSetarPropriedades = sSetarPropriedades & sNomeComponenteAtual & ".ItemData2" & vbCrLf
            ElseIf ReadField(rsADOGlobal, "type") = 56 Or ReadField(rsADOGlobal, "type") = 38 Then
                sSetarPropriedades = sSetarPropriedades & "Val(" & sNomeComponenteAtual & ".Text)" & vbCrLf
            ElseIf ReadField(rsADOGlobal, "type") = 106 Then
                sSetarPropriedades = sSetarPropriedades & "CDbl1(" & sNomeComponenteAtual & ".Text)" & vbCrLf
            ElseIf ReadField(rsADOGlobal, "type") = 111 Then
                sSetarPropriedades = sSetarPropriedades & "CDateEspecial(" & sNomeComponenteAtual & ".Text)" & vbCrLf
            Else
                sSetarPropriedades = sSetarPropriedades & sNomeComponenteAtual & ".Text" & vbCrLf
            End If
        End If
        rsADOGlobal.MoveNext
    Loop
    rsADOGlobal.Close
    sfrmModeloDados = Replace(sfrmModeloDados, "'[FORMATAÇÃO-COMBOS]'", sCombos, , , vbTextCompare)
    sfrmModeloDados = Replace(sfrmModeloDados, "'[CARREGAR-DADOS]'", sCarregarItens, , , vbTextCompare)
    
    'Carregar propriedades e formatação do spread
    sColunasSpread = ""
    sCarregarItens = ""
    sCamposItem = ""
    iTop = iTop + 555
    iLeft = 105
    bAuxiliar = True
    For i = 0 To lstTabelasFilhas.ListCount - 1
        If lstTabelasFilhas.Selected(i) Then
                
            sComponentes = sComponentes & CriarComponente(lstTabelasFilhas.List(i), iTop, iLeft, iTabIndex, WidthFrameDados, False, sNomeComponenteAtual)
                                    
            sPKItem = ""
            sCampos = ""
            sCamposItem = ""
            If bAuxiliar Then
                bAuxiliar = False
                sSetarPropriedades = sSetarPropriedades & vbCrLf & "    If Not c" & Mid(cboTabela.Text, 4) & ".ZerarItens() Then " & vbCrLf
                sSetarPropriedades = sSetarPropriedades & "        Exit function" & vbCrLf
                sSetarPropriedades = sSetarPropriedades & "    End If" & vbCrLf & vbCrLf
            End If
                        
            Call Select_Table(True, "syscolumns", "name, type, length", "id = " & lstTabelasFilhas.ItemData(i), "colid")
            Do While Not rsADOGlobal.EOF
                sColunasSpread = sColunasSpread & IIf(sColunasSpread <> "", "    ", "") & "Call " & sNomeComponenteAtual & ".NovaColunaSpread(" & RetornarTipoColunaSpread(Mid(ReadField(rsADOGlobal, "name"), 1, 2)) & ",False, True," & Chr(34) & ReadField(rsADOGlobal, "name") & Chr(34) & "," & Chr(34) & Mid(ReadField(rsADOGlobal, "name"), 4) & Chr(34) & "," & RetornarTamanhoColunaSpread(Mid(ReadField(rsADOGlobal, "name"), 1, 2)) & "," & ReadField(rsADOGlobal, "length") & ")" & vbCrLf
                
                If UCase(ReadField(rsADOGlobal, "name")) <> UCase(sPKPrincipal) Then
                    sCampos = sCampos & ".SpreadEventoName(" & Chr(34) & ReadField(rsADOGlobal, "name") & Chr(34) & "), "
                    sCamposItem = sCamposItem & ReadField(rsADOGlobal, "name") & ", "
                End If
                If ReadField(rsADOGlobal, "type") = 56 Then
                    sPKItem = ReadField(rsADOGlobal, "name")
                End If
                
                If ReadField(rsADOGlobal, "type") = 56 Then
                    sAtualizarSpread = sAtualizarSpread & IIf(sAtualizarSpread <> "", "    ", "") & "Call " & sNomeComponenteAtual & ".AtualizarStatusSpread(" & Chr(34) & ReadField(rsADOGlobal, "name") & Chr(34) & ", c" & Mid(cboTabela.Text, 4) & "." & "a" & Mid(lstTabelasFilhas.List(i), 4) & "Item" & ")" & vbCrLf
                End If
                
                rsADOGlobal.MoveNext
            Loop
            rsADOGlobal.Close
            
            sColunasSpread = sColunasSpread & "    Call " & sNomeComponenteAtual & ".FormatarNovo(21)" & vbCrLf & vbCrLf
                                                
            If sCampos <> "" Then
                If sPKItem = "" Then
                    sPKItem = "id_" & Mid(lstTabelasFilhas.List(i), 4)
                End If
                sCampos = Left(sCampos, Len(sCampos) - 2)
                sCamposItem = Left(sCamposItem, Len(sCamposItem) - 2)
                
                sCarregarItens = sCarregarItens & IIf(sCarregarItens <> "", "        ", "") & "Call " & sNomeComponenteAtual & ".Carregar(Select_Table(False," & Chr(34) & lstTabelasFilhas.List(i) & Chr(34) & "," & Chr(34) & sCamposItem & Chr(34) & "," & Chr(34) & sPKPrincipal & " = " & Chr(34) & " & " & sPKPrincipal & "," & Chr(34) & sPKItem & Chr(34) & "))" & vbCrLf
                
                sSetarPropriedades = sSetarPropriedades & "    With " & sNomeComponenteAtual & vbCrLf
                sSetarPropriedades = sSetarPropriedades & "        For i = 1 To .MaxRows - 1" & vbCrLf
                sSetarPropriedades = sSetarPropriedades & "            .Row = i" & vbCrLf
                sSetarPropriedades = sSetarPropriedades & "            if .StatusGravacao(i) <> const_Inicial Then" & vbCrLf
                sSetarPropriedades = sSetarPropriedades & "                If Not c" & Mid(cboTabela.Text, 4) & ".Carregar" & Mid(lstTabelasFilhas.List(i), 4) & "Item(.StatusGravacao(i), " & sCampos & ") Then" & vbCrLf
                sSetarPropriedades = sSetarPropriedades & "                    Exit Function" & vbCrLf
                sSetarPropriedades = sSetarPropriedades & "                End If" & vbCrLf
                sSetarPropriedades = sSetarPropriedades & "            End If" & vbCrLf
                sSetarPropriedades = sSetarPropriedades & "        Next i" & vbCrLf
                sSetarPropriedades = sSetarPropriedades & "    End With" & vbCrLf & vbCrLf
            End If
        End If
    Next i
    
    sfrmModeloDados = Replace(sfrmModeloDados, "'[CARREGAR-DADOSITEM]'", sCarregarItens, , , vbTextCompare)
    sfrmModeloDados = Replace(sfrmModeloDados, "'[SETAR-PROPRIEDADES]'", sSetarPropriedades, , , vbTextCompare)
    sfrmModeloDados = Replace(sfrmModeloDados, "'[FORMATAÇÃO-SPREAD]'", sColunasSpread, , , vbTextCompare)
    sfrmModeloDados = Replace(sfrmModeloDados, "'[ATUALIZAR-SPREAD]'", sAtualizarSpread, , , vbTextCompare)
    
    'Cria os componentes
    iPos1 = InStr(1, sfrmModeloDados, "fraDados", vbTextCompare)
    If iPos1 > 0 Then
        iPos2 = InStr(iPos1, sfrmModeloDados, "ShadowStyle", vbTextCompare)
        If iPos2 = 0 Then
            iPos2 = iPos1
        End If
        iPos1 = InStr(iPos2, sfrmModeloDados, "End", vbTextCompare)
        If iPos1 > 0 Then
            sfrmModeloDados = Mid(sfrmModeloDados, 1, iPos1 - 3) & vbCrLf & sComponentes & Mid(sfrmModeloDados, iPos1)
        End If
    End If
    
    Call SalvarArquivo("frm" & Mid(cboTabela, 4) & "Dados.frm", sfrmModeloDados)

    Exit Sub
err_cmdGerarCodigo_Click:
    ShowError
End Sub

Private Function PegarTipoPropriedade(Tipo As Integer) As String
On Error GoTo err_PegarTipoPropriedade

    If Tipo = 56 Or Tipo = 38 Then
        PegarTipoPropriedade = "Long"
    ElseIf Tipo = 106 Then
        PegarTipoPropriedade = "Double"
    ElseIf Tipo = 111 Then
        PegarTipoPropriedade = "Date"
    Else
        PegarTipoPropriedade = "String"
    End If

    Exit Function
err_PegarTipoPropriedade:
    ShowError
End Function

Private Function PegarValorPadrao(Tipo As Integer) As String
On Error GoTo err_PegarValorPadrao

    If Tipo = 56 Or Tipo = 38 Or Tipo = 106 Then
        PegarValorPadrao = "0"
    ElseIf Tipo = 111 Then
        PegarValorPadrao = "CDateEspecial(" & Chr(34) & Chr(34) & ")"
    Else
        PegarValorPadrao = """"
    End If

    Exit Function
err_PegarValorPadrao:
    ShowError
End Function

Private Function CriarComponente(sCampo As String, iTop As Long, iLeft As Long, iTabIndex As Integer, WidthFrame As Long, bConsulta As Boolean, ByRef sNomeComponenteAtual As String) As String
On Error GoTo err_CriarComponente

    Dim sComponentes  As String
    Dim Tipo As String
    Dim espaco As Integer
    Dim sLabel As String
    Dim sFiltroData As String
    
    If iLeft = 0 Then
        iLeft = 105
    End If
    If iTop = 0 Then
        iTop = 285
    End If
    
    espaco = 80
    Tipo = UCase(Mid(sCampo, 1, 3))
    sLabel = Mid(sCampo, 4)
    
    If bConsulta And Tipo = "DT" Then
        sFiltroData = "Inicial"
    End If
    
    Select Case Tipo
        Case "ID_"
            If iLeft + 2805 + espaco > WidthFrame Then
                iTop = iTop + 555
                iLeft = 105
            End If
            sNomeComponenteAtual = VerificarNomeComponente("cbo" & sLabel, iTabIndex)
            sComponentes = "      Begin Transportes.SuperDBCombo " & sNomeComponenteAtual & vbCrLf
            sComponentes = sComponentes & "         Height = 510" & vbCrLf
            sComponentes = sComponentes & "         Left = " & iLeft & vbCrLf
            sComponentes = sComponentes & "         TabIndex = " & iTabIndex & vbCrLf
            sComponentes = sComponentes & "         Top = " & iTop & vbCrLf
            sComponentes = sComponentes & "         Width = 2805" & vbCrLf
            sComponentes = sComponentes & "         _ExtentX        =   0" & vbCrLf
            sComponentes = sComponentes & "         _ExtentY        =   0" & vbCrLf
            sComponentes = sComponentes & "         Label = " & Chr(34) & sLabel & Chr(34) & vbCrLf
            sComponentes = sComponentes & "      End" & vbCrLf
            
            iTabIndex = iTabIndex + 1
            iLeft = iLeft + 2805 + espaco
        Case "QT_", "KG_", "VL_", "PC_", "DT_", "HR_"
            If iLeft + 1300 + espaco > WidthFrame Then
                iTop = iTop + 555
                iLeft = 105
            End If
            
            sNomeComponenteAtual = VerificarNomeComponente("lbl" & sLabel & sFiltroData, iTabIndex)
            sComponentes = "      Begin VB.Label " & sNomeComponenteAtual & vbCrLf
            sComponentes = sComponentes & "         Caption = " & Chr(34) & sLabel & sFiltroData & Chr(34) & vbCrLf
            sComponentes = sComponentes & "         Left = " & iLeft & vbCrLf
            sComponentes = sComponentes & "         TabIndex = " & iTabIndex & vbCrLf
            sComponentes = sComponentes & "         Top = " & iTop & vbCrLf
            sComponentes = sComponentes & "         Height = 195" & vbCrLf
            sComponentes = sComponentes & "         Width = 1300" & vbCrLf
            sComponentes = sComponentes & "      End" & vbCrLf
            iTabIndex = iTabIndex + 1
            
            sNomeComponenteAtual = VerificarNomeComponente("msk" & sLabel & sFiltroData, iTabIndex)
            sComponentes = sComponentes & "      Begin Transportes.SuperControlNovo " & sNomeComponenteAtual & vbCrLf
            sComponentes = sComponentes & "         Height = 315" & vbCrLf
            sComponentes = sComponentes & "         Left = " & iLeft & vbCrLf
            sComponentes = sComponentes & "         TabIndex = " & iTabIndex & vbCrLf
            sComponentes = sComponentes & "         Top = " & iTop + 190 & vbCrLf
            sComponentes = sComponentes & "         Width = 1300" & vbCrLf
            sComponentes = sComponentes & "         _ExtentX        =   0" & vbCrLf
            sComponentes = sComponentes & "         _ExtentY        =   0" & vbCrLf
            sComponentes = sComponentes & "         Mascara = " & IIf(Tipo = "QT", "4", IIf(Tipo = "DT", "1", IIf(Tipo = "HR", "6", "2"))) & vbCrLf
            sComponentes = sComponentes & "      End" & vbCrLf
            
            iTabIndex = iTabIndex + 1
            iLeft = iLeft + 1300 + espaco
            
            If bConsulta And Tipo = "DT" Then
                If iLeft + 1300 + espaco > WidthFrame Then
                    iTop = iTop + 555
                    iLeft = 105
                End If
                
                sNomeComponenteAtual = VerificarNomeComponente("lbl" & sLabel & "Final", iTabIndex)
                sComponentes = "      Begin VB.Label " & sNomeComponenteAtual & vbCrLf
                sComponentes = sComponentes & "         Caption = " & Chr(34) & sLabel & " Final" & Chr(34) & vbCrLf
                sComponentes = sComponentes & "         Left = " & iLeft & vbCrLf
                sComponentes = sComponentes & "         TabIndex = " & iTabIndex & vbCrLf
                sComponentes = sComponentes & "         Top = " & iTop & vbCrLf
                sComponentes = sComponentes & "         Height = 195" & vbCrLf
                sComponentes = sComponentes & "         Width = 1300" & vbCrLf
                sComponentes = sComponentes & "      End" & vbCrLf
                iTabIndex = iTabIndex + 1
                
                sNomeComponenteAtual = VerificarNomeComponente("msk" & sLabel & "Final", iTabIndex)
                sComponentes = sComponentes & "      Begin Transportes.SuperControlNovo " & sNomeComponenteAtual & vbCrLf
                sComponentes = sComponentes & "         Height = 315" & vbCrLf
                sComponentes = sComponentes & "         Left = " & iLeft & vbCrLf
                sComponentes = sComponentes & "         TabIndex = " & iTabIndex & vbCrLf
                sComponentes = sComponentes & "         Top = " & iTop + 190 & vbCrLf
                sComponentes = sComponentes & "         Width = 1300" & vbCrLf
                sComponentes = sComponentes & "         _ExtentX        =   0" & vbCrLf
                sComponentes = sComponentes & "         _ExtentY        =   0" & vbCrLf
                sComponentes = sComponentes & "         Mascara = 1" & vbCrLf
                sComponentes = sComponentes & "      End" & vbCrLf
                
                iTabIndex = iTabIndex + 1
                iLeft = iLeft + 1300 + espaco
            End If
        Case "TBD"
            If iLeft + 3735 + espaco > WidthFrame Then
                iTop = iTop + 1600
                iLeft = 105
            End If
            
            sNomeComponenteAtual = VerificarNomeComponente("lbl" & sLabel, iTabIndex)
            sComponentes = "      Begin VB.Label " & sNomeComponenteAtual & vbCrLf
            sComponentes = sComponentes & "         Caption = " & Chr(34) & sLabel & Chr(34) & vbCrLf
            sComponentes = sComponentes & "         Left = " & iLeft & vbCrLf
            sComponentes = sComponentes & "         TabIndex = " & iTabIndex & vbCrLf
            sComponentes = sComponentes & "         Top = " & iTop & vbCrLf
            sComponentes = sComponentes & "         Height = 195" & vbCrLf
            sComponentes = sComponentes & "         Width = 1300" & vbCrLf
            sComponentes = sComponentes & "      End" & vbCrLf
            iTabIndex = iTabIndex + 1
            
            sNomeComponenteAtual = VerificarNomeComponente("spr" & sLabel, iTabIndex)
            sComponentes = sComponentes & "      Begin Transportes.SuperSpreadNovo " & sNomeComponenteAtual & vbCrLf
            sComponentes = sComponentes & "         Height = 1545" & vbCrLf
            sComponentes = sComponentes & "         Left = " & iLeft & vbCrLf
            sComponentes = sComponentes & "         TabIndex = " & iTabIndex & vbCrLf
            sComponentes = sComponentes & "         Top = " & iTop + 190 & vbCrLf
            sComponentes = sComponentes & "         Width = 3735" & vbCrLf
            sComponentes = sComponentes & "         _ExtentX        =   0" & vbCrLf
            sComponentes = sComponentes & "         _ExtentY        =   0" & vbCrLf
            sComponentes = sComponentes & "      End" & vbCrLf
            
            iTabIndex = iTabIndex + 1
            iLeft = iLeft + 3735 + espaco
        Case Else
            If iLeft + 1300 + espaco > WidthFrame Then
                iTop = iTop + 555
                iLeft = 105
            End If
            
            sNomeComponenteAtual = VerificarNomeComponente("lbl" & sLabel, iTabIndex)
            sComponentes = "      Begin VB.Label " & sNomeComponenteAtual & vbCrLf
            sComponentes = sComponentes & "         Caption = " & Chr(34) & sLabel & Chr(34) & vbCrLf
            sComponentes = sComponentes & "         Left = " & iLeft & vbCrLf
            sComponentes = sComponentes & "         TabIndex = " & iTabIndex & vbCrLf
            sComponentes = sComponentes & "         Top = " & iTop & vbCrLf
            sComponentes = sComponentes & "         Height = 195" & vbCrLf
            sComponentes = sComponentes & "         Width = 1300" & vbCrLf
            sComponentes = sComponentes & "      End" & vbCrLf
            iTabIndex = iTabIndex + 1
            
            sNomeComponenteAtual = VerificarNomeComponente("txt" & sLabel, iTabIndex)
            sComponentes = sComponentes & "      Begin Transportes.SuperText " & sNomeComponenteAtual & vbCrLf
            sComponentes = sComponentes & "         Height = 315" & vbCrLf
            sComponentes = sComponentes & "         Left = " & iLeft & vbCrLf
            sComponentes = sComponentes & "         TabIndex = " & iTabIndex & vbCrLf
            sComponentes = sComponentes & "         Top = " & iTop + 190 & vbCrLf
            sComponentes = sComponentes & "         Width = 1300" & vbCrLf
            sComponentes = sComponentes & "         _ExtentX        =   0" & vbCrLf
            sComponentes = sComponentes & "         _ExtentY        =   0" & vbCrLf
            sComponentes = sComponentes & "      End" & vbCrLf
            
            iTabIndex = iTabIndex + 1
            iLeft = iLeft + 1300 + espaco
    End Select

    'Ajusta a tela
    If bConsulta And iTop > 285 Then
        sfrmModeloConsulta = Replace(sfrmModeloConsulta, "Height          =   945", "Height          =   1450") 'Altura do Frame
        sfrmModeloConsulta = Replace(sfrmModeloConsulta, "Top             =   210", "Top             =   720") 'Top botão pesquisar
        sfrmModeloConsulta = Replace(sfrmModeloConsulta, "Top             =   1050", "Top             =   1560") 'Top do Spread
        sfrmModeloConsulta = Replace(sfrmModeloConsulta, "Height          =   5235", "Height          =   4725") 'Altura do Spread
    End If
    CriarComponente = sComponentes

    Exit Function
err_CriarComponente:
    ShowError
End Function

Private Function VerificarNomeComponente(Nome As String, iTabIndex As Integer)

    VerificarNomeComponente = Nome
    If InStr(1, "," & sNomeComponentesGeral, "," & Nome & ",", vbTextCompare) > 0 Then
        VerificarNomeComponente = Nome & iTabIndex
    End If
    
    sNomeComponentesGeral = sNomeComponentesGeral & Nome & ","
End Function

Private Sub SalvarArquivo(sNomeArquivo As String, txtArquivo As Variant)
On Error GoTo err_SalvarArquivo

On Error GoTo err_SalvarArquivo

On Error GoTo err_cmdSalvar_Click
    Dim nFile As Integer
    Dim sTexto As String
    
    nFile = FreeFile
    Open "C:\Modelos\" & sNomeArquivo For Binary Access Write As #nFile
    Put #nFile, , txtArquivo
    Close #nFile
    
'    sTexto = PegarModelo(sNomeArquivo)
    
        
    Exit Sub
err_cmdSalvar_Click:
    ShowError

    Exit Sub
err_SalvarArquivo:
    ShowError
    
End Sub

Private Function PegarTipoComboFiltro(campo As String) As String
    campo = UCase(campo)
    If InStr(1, campo, "PESSOA") > 0 Or InStr(1, campo, "REMETENTE") Or InStr(1, campo, "DESTINATARIO") > 0 Or InStr(1, campo, "FATURADO") > 0 Or InStr(1, campo, "REDESPACHO") > 0 Or InStr(1, campo, "CONSIGNATARIO") > 0 Or InStr(1, campo, "EXPEDIDOR") > 0 Or InStr(1, campo, "RECEBEDOR") > 0 Then
        PegarTipoComboFiltro = "Combo_Pessoas_Sem_CNPJ"
    ElseIf InStr(1, campo, "VEICULO") > 0 Then
        PegarTipoComboFiltro = "Combo_Veiculo_Sem_Placa_Com_Inativos"
    ElseIf InStr(1, campo, "VENDEDOR") > 0 Then
        PegarTipoComboFiltro = "Combo_Vendedor_Sem_CNPJ"
    ElseIf InStr(1, campo, "TRANSPORTADORA") > 0 Then
        PegarTipoComboFiltro = "Combo_Transportadora"
    ElseIf InStr(1, campo, "CIAAEREA") > 0 Then
        PegarTipoComboFiltro = "Combo_CiaAerea"
    ElseIf InStr(1, campo, "TIPOVEICULO") > 0 Then
        PegarTipoComboFiltro = "Combo_TipoVeiculo"
    ElseIf InStr(1, campo, "FUNCIONARIO") > 0 Then
        PegarTipoComboFiltro = "Combo_Funcionario_Sem_CNPJ"
    ElseIf InStr(1, campo, "FORNECEDOR") > 0 Then
        PegarTipoComboFiltro = "Combo_Fornecedor_Sem_CNPJ"
    ElseIf InStr(1, campo, "CONTA") > 0 Then
        PegarTipoComboFiltro = "Combo_ContaBancaria"
    ElseIf InStr(1, campo, "CLIENTE") > 0 Then
        PegarTipoComboFiltro = "Combo_ClienteSistema_Sem_CNPJ"
    ElseIf InStr(1, campo, "EMPRESA") > 0 Then
        PegarTipoComboFiltro = "Combo_Empresa_Sem_CNPJ"
    ElseIf InStr(1, campo, "CIDADE") > 0 Then
        PegarTipoComboFiltro = "Combo_Cidade_Sem_Sigla"
    ElseIf InStr(1, campo, "CARRETA") > 0 Then
        PegarTipoComboFiltro = "Combo_Carreta_Sem_Placa"
    ElseIf InStr(1, campo, "AGENTE") > 0 Then
        PegarTipoComboFiltro = "Combo_Agentes_Sem_CNPJ"
    ElseIf InStr(1, campo, "MOTORISTA") > 0 Then
        PegarTipoComboFiltro = "Combo_Motoristas_Sem_CNPJ"
    Else
        PegarTipoComboFiltro = "Padrão_Não_Localizado"
    End If
End Function

Private Function RetornarTipoColunaSpread(Tipo As String) As String
On Error GoTo err_RetornarTipoColunaSpread
    
    Select Case UCase(Tipo)
        Case "ID", "QT"
            RetornarTipoColunaSpread = "eslNumero"
        Case "DT"
            RetornarTipoColunaSpread = "eslData"
        Case "KG", "VL", "PC"
            RetornarTipoColunaSpread = "eslValor"
        Case "TP"
            RetornarTipoColunaSpread = "eslCheck"
        Case Else
            RetornarTipoColunaSpread = "eslTexto"
    End Select

    Exit Function
err_RetornarTipoColunaSpread:
    ShowError
End Function

Private Function RetornarTamanhoColunaSpread(Tipo As String) As String
On Error GoTo err_RetornarTamanhoColunaSpread
    
    Select Case UCase(Tipo)
        Case "ID"
            RetornarTamanhoColunaSpread = "0"
        Case "QT", "KG", "VL", "PC"
            RetornarTamanhoColunaSpread = "12"
        Case "DT", "TP", "NR", "CD"
            RetornarTamanhoColunaSpread = "10"
        Case "HR"
            RetornarTamanhoColunaSpread = "8"
        Case "CM"
            RetornarTamanhoColunaSpread = "80"
        Case Else
            RetornarTamanhoColunaSpread = "30"
    End Select

    Exit Function
err_RetornarTamanhoColunaSpread:
    ShowError
End Function

Private Sub Form_Activate()
    frmMDI.Arrange vbCascade
End Sub

Private Sub Form_Load()
On Error GoTo err_Form_Load

    Dim i As Integer
    Dim iPos As Integer
    
    Call Carregar_ComboLista(cboTabela, Select_Table(False, "SysObjects", "id,name", "type = 'U' and (mid(name,1,2) = 'tb' or mid(name,1,4) = 'LOG_')", "name"), 1, False)
    
    For i = 0 To cboTabela.ListCount - 1
        lstTabelasFilhas.AddItem cboTabela.List(i)
        lstTabelasFilhas.ItemData(lstTabelasFilhas.NewIndex) = cboTabela.ItemData(i)
        
        cboTabelaJoin.AddItem cboTabela.List(i)
        cboTabelaJoin.ItemData(cboTabelaJoin.NewIndex) = cboTabela.ItemData(i)
    Next i
    
    Call sprConsulta.Formatar(Array("id_Coluna,0,I,S,S", "ds_Coluna|Nome Coluna,25,100,S,S,N", "ds_Cabecalho|Cabeçalho Coluna,25,100,N,S,S", "ds_Conteudo|Conteúdo,25,100,N,S,S", "nr_Tamanho,0,S,I,N", "nr_Tipo,0,S,I,N"), 21)
    Call sprFiltros.Formatar(Array("ds_Campo,25,100,S,S,S", "ds_Filtro|Filtro,40,100,S,S,S", "nr_Ordem|Ordem,8,N,I,S"))
        
    sfrmModeloConsulta = PegarModelo("frmModeloConsulta.frm")
    sfrmModeloDados = PegarModelo("frmModeloDados.frm")
    sclsModelo = PegarModelo("clsModelo.cls")
    
    sfrmModeloConsultaAux = sfrmModeloConsulta
    sfrmModeloDadosAux = sfrmModeloDados
    sclsModeloAux = sclsModelo
    
    WidthFrameFiltros = 11715
    WidthFrameDados = 12750
    HeightFrameDados = 7485
    
    Exit Sub
err_Form_Load:
    ShowError
End Sub

Private Function PegarModelo(sArquivo As String) As String
    Dim nFile As Integer
    Dim sRetorno As String
    Dim sPath As String
    
    sPath = "c:\Modelos\"
    nFile = FreeFile
    Open sPath & sArquivo For Input As #nFile
    sRetorno = Input$(LOF(nFile), nFile)
    Close #nFile

    PegarModelo = sRetorno
End Function

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub lstCampos_ItemCheck(Item As Integer)
    Dim i As Integer
    Dim nr_Ordem As Integer
    
    nr_Ordem = 0
    
    If lstCampos.Selected(Item) Then
    
        For i = 1 To sprFiltros.MaxRows
            sprFiltros.Row = i
            
            If Val(sprFiltros.SpreadEventoName("nr_Ordem")) > nr_Ordem Then
                nr_Ordem = sprFiltros.SpreadEventoName("nr_Ordem")
            End If
        Next i
        
        nr_Ordem = nr_Ordem + 1
        
        sprFiltros.MaxRows = sprFiltros.MaxRows + 1
        sprFiltros.Row = sprFiltros.MaxRows
        
        sprFiltros.TextCol("ds_Campo") = "a." & lstCampos.List(Item)
        sprFiltros.TextCol("nr_Ordem") = nr_Ordem
        
        Select Case UCase(Mid(lstCampos.List(Item), 1, 2))
            Case "ID"
                sprFiltros.TextCol("ds_Filtro") = "Combo de " & Mid(lstCampos.List(Item), 4)
            Case "NR", "CD"
                sprFiltros.TextCol("ds_Filtro") = "Número de " & Mid(lstCampos.List(Item), 4)
            Case "DT"
                sprFiltros.TextCol("ds_Filtro") = "Período de " & Mid(lstCampos.List(Item), 4)
            Case "HR"
                sprFiltros.TextCol("ds_Filtro") = "Horário de " & Mid(lstCampos.List(Item), 4)
            Case "KG"
                sprFiltros.TextCol("ds_Filtro") = "Peso de " & Mid(lstCampos.List(Item), 4)
            Case "QT"
                sprFiltros.TextCol("ds_Filtro") = "Quantidade de " & Mid(lstCampos.List(Item), 4)
            Case "TP"
                sprFiltros.TextCol("ds_Filtro") = "Tipo de " & Mid(lstCampos.List(Item), 4)
            Case "DS"
                sprFiltros.TextCol("ds_Filtro") = "Descrição de " & Mid(lstCampos.List(Item), 4)
            Case "PC"
                sprFiltros.TextCol("ds_Filtro") = "Percentual de " & Mid(lstCampos.List(Item), 4)
            Case "CM"
                sprFiltros.TextCol("ds_Filtro") = "Comentário de " & Mid(lstCampos.List(Item), 4)
            Case "VL"
                sprFiltros.TextCol("ds_Filtro") = "Valor de " & Mid(lstCampos.List(Item), 4)
        End Select
    Else
        For i = 1 To sprFiltros.MaxRows
            sprFiltros.Row = i
            
            If sprFiltros.SpreadEventoName("ds_Campo") = "a." & lstCampos.List(Item) Then
                nr_Ordem = sprFiltros.SpreadEventoName("nr_Ordem")
                sprFiltros.Action = 5
                sprFiltros.MaxRows = sprFiltros.MaxRows - 1
                Exit For
            End If
        Next i
        
        For i = 1 To sprFiltros.MaxRows
            sprFiltros.Row = i
            
            If Val(sprFiltros.SpreadEventoName("nr_Ordem")) > nr_Ordem Then
                sprFiltros.TextCol("nr_Ordem") = Val(sprFiltros.SpreadEventoName("nr_Ordem")) - 1
            End If
        Next i
    End If
End Sub

Private Sub sprConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        sprConsulta.Row = sprConsulta.ActiveRow
        sprConsulta.MaxRows = sprConsulta.MaxRows + 1
        sprConsulta.Action = 7
    End If
End Sub
