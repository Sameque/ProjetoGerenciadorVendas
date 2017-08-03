VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmProdutoDados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Produto"
   ClientHeight    =   3030
   ClientLeft      =   3030
   ClientTop       =   4905
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4500
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Height          =   750
      HelpContextID   =   23
      Left            =   2715
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Novo lançamento "
      Top             =   2160
      Width           =   810
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   750
      Left            =   1845
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Gravar os Dados"
      Top             =   2160
      Width           =   810
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   750
      Left            =   3585
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Sair da tela"
      Top             =   2160
      Width           =   810
   End
   Begin Threed.SSFrame fraDados 
      Height          =   2055
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   4320
      _Version        =   65536
      _ExtentX        =   7620
      _ExtentY        =   3625
      _StockProps     =   14
      Caption         =   "Dados"
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
      Begin Transportes.SuperText txtCodigo 
         Height          =   315
         Left            =   105
         TabIndex        =   1
         Top             =   475
         Width           =   1300
         _ExtentX        =   0
         _ExtentY        =   556
      End
      Begin Transportes.SuperText txtDescricao 
         Height          =   315
         Left            =   1485
         TabIndex        =   2
         Top             =   480
         Width           =   2745
         _ExtentX        =   0
         _ExtentY        =   556
      End
      Begin Transportes.SuperDBCombo cboProdutoNatureza 
         Height          =   510
         Left            =   90
         TabIndex        =   6
         Top             =   1440
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   900
         Label           =   "ProdutoNatureza"
      End
      Begin Transportes.SuperControlNovo mskCubado 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   1065
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         ToolTip         =   ""
      End
      Begin Transportes.SuperControlNovo mskReal 
         Height          =   315
         Left            =   1470
         TabIndex        =   4
         Top             =   1065
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         ToolTip         =   ""
      End
      Begin Transportes.SuperControlNovo mskQtMinima 
         Height          =   315
         Left            =   2850
         TabIndex        =   5
         Top             =   1065
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         ToolTip         =   ""
      End
      Begin VB.Label lblProduto 
         Caption         =   "Código"
         Height          =   195
         Left            =   105
         TabIndex        =   10
         Top             =   285
         Width           =   1300
      End
      Begin VB.Label lblProduto6 
         Caption         =   "Descrição"
         Height          =   195
         Left            =   1485
         TabIndex        =   11
         Top             =   285
         Width           =   1300
      End
      Begin VB.Label lblPesoCubado 
         Caption         =   "Peso Cubado"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label lblPesoReal 
         Caption         =   "Peso Real"
         Height          =   195
         Left            =   1485
         TabIndex        =   13
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label lblQtMinima 
         Caption         =   "Qt. Minima"
         Height          =   195
         Left            =   2880
         TabIndex        =   14
         Top             =   840
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmProdutoDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public id_Produto As Long
Public formChamador As Form

Dim cProduto As clsProduto

Private Sub Form_Activate()
    frmMDI.Arrange vbCascade
End Sub

Private Sub Form_Load()
On Error GoTo err_Form_Load
    
    Call CarregarTela
    Call CenterForm(Me)
    
    Exit Sub
err_Form_Load:
    ShowError
End Sub
Private Function CarregarTela()
    Set cProduto = New clsProduto

    mskCubado.Mascara = Valor
    mskReal.Mascara = Valor
    mskQtMinima.Mascara = Numero
    Call cboProdutoNatureza.Formatar("id_ProdutoNatureza,ds_ProdutoNatureza", "0,200", "false,true", "tbdProdutoNatureza", "", "ds_ProdutoNatureza", , , "Teste cabe")

    If id_Produto > 0 Then

        If cProduto.CarregarDados(id_Produto) Then
            txtCodigo.Text = cProduto.cd_Produto
            txtDescricao.Text = cProduto.ds_Produto
            Call cboProdutoNatureza.Formatar("id_ProdutoNatureza,ds_ProdutoNatureza", "0,200", "false,true", "tbdProdutoNatureza", "", "ds_ProdutoNatureza", , , "Teste cabe")
            Call cboProdutoNatureza.PesquisarCombo(True, cProduto.id_ProdutoNatureza, "", True)
            mskCubado.Text = Format(cProduto.kg_Cubado, "#,###0.00")
            mskReal.Text = Format(cProduto.kg_Real, "#,###0.00")
            mskQtMinima.Text = cProduto.qt_Minima

        End If
        
    End If
    
    mskCubado.Mascara = Valor
    mskReal.Mascara = Valor
    mskQtMinima.Mascara = Numero
    
End Function
Private Sub Form_Unload(Cancel As Integer)
    Set cProduto = Nothing
    Set frmProdutoDados = Nothing
End Sub

Private Sub cmdGravar_Click()
On Error GoTo err_cmdGravar_Click

    If Not ValidarControles(Me) Then
        Exit Sub
    End If
    
    If Mensagem("Confirma Gravação?", Pergunta) = vbNo Then
        Exit Sub
    End If

    If Not CarregarPropriedades(cProduto) Then
        Exit Sub
    End If

    Call AbreTransacao
    If Not cProduto.Gravar Then
        Call VoltaTransacao
        Mensagem "Ocorreu um erro no processamento.", ErroCritico
        Exit Sub
    End If
    Call FechaTransacao

    id_Produto = cProduto.id_Produto
    
    Call formChamador.AtualizarDados(id_Produto)
    
    Call Mensagem("Gravação Efetuada", Informacao)
    cmdNovo.SetFocus

    Exit Sub
err_cmdGravar_Click:
    ShowError
    Call FechaTransacao
End Sub

Private Function CarregarPropriedades(cProduto As clsProduto) As Boolean
On Error GoTo err_CarregarPropriedades

    CarregarPropriedades = False
    
    Dim i As Long

    cProduto.id_Produto = id_Produto
    cProduto.cd_Produto = Trim(txtCodigo.Text)
    cProduto.ds_Produto = Trim(txtDescricao.Text)
    cProduto.id_ProdutoNatureza = cboProdutoNatureza.ItemData2
    cProduto.kg_Cubado = CDbl1(mskCubado.Text)
    cProduto.kg_Real = CDbl1(mskReal.Text)
    cProduto.qt_Minima = Val(mskQtMinima.Text)

    CarregarPropriedades = True
    
    Exit Function
err_CarregarPropriedades:
    ShowError
End Function

Private Sub cmdNovo_Click()
    Call LimparControles(Me)
    Set cProduto = New clsProduto
    id_Produto = 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub txtCodigo_LostFocus()
    If Not ValidaCodigo Then
        txtCodigo.SetFocus
    End If
    
End Sub
Private Function ValidaCodigo() As Boolean
    
    ValidaCodigo = False
    
    Call Select_Table(True, "tbdProduto", "cd_Produto", "cd_Produto = '" & txtCodigo.Text & "'")
    
    If (Not rsADOGlobal.EOF) And id_Produto = 0 Then
        Mensagem "Você já possui um produto com esse código!", Informacao
        Exit Function
    End If
    
    ValidaCodigo = True
    
End Function
