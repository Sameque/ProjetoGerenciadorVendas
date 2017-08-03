VERSION 5.00
Begin VB.Form frmProdutoNatureza 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Natureza"
   ClientHeight    =   3900
   ClientLeft      =   4470
   ClientTop       =   3000
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Sair da tela"
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   855
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Incluir novo Item"
      Top             =   2880
      Width           =   855
   End
   Begin Transportes.SuperSpreadNovo sprNatureza 
      Height          =   2670
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   4710
   End
End
Attribute VB_Name = "frmProdutoNatureza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

    Call sprNatureza.NovaColunaSpread(eslNumero, True, True, "id_ProdutoNatureza", "id_ProdutoNatureza", 0, 0)
    Call sprNatureza.NovaColunaSpread(eslTexto, False, False, "ds_ProdutoNatureza", "Produto", 30, 50)
    Call sprNatureza.NovaColunaSpread(eslCheck, False, False, "tp_ProdutoPerigoso", "Produto Perigoso", 14, 4)
    Call sprNatureza.NovaColunaSpread(eslCheck, False, False, "tp_ControlePoliciaFederal", "Controle da Polícia Federal", 14, 4)
    Call sprNatureza.FormatarNovo(21)
    
End Function
Private Function CarregarComponentes()

    Call sprNatureza.Carregar(Select_Table(False, Tabelas, Campos, "", "ds_ProdutoNatureza"))
    
End Function
Private Function Campos()
   
    Dim sCampos As String
    sCampos = "id_ProdutoNatureza,ds_ProdutoNatureza,tp_ProdutoPerigoso,tp_ControlePoliciaFederal"
    Campos = sCampos
    
End Function
Private Function Tabelas()
   
    Dim sTabelas As String
    sTabelas = "tbdProdutoNatureza"
    Tabelas = sTabelas
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frmProdutoNatureza = Nothing
End Sub

Private Sub cmdGravar_Click()
On Error GoTo err_cmdGravar_Click

    If Not ValidarControles(Me) Then
        Exit Sub
    End If

    If Mensagem("Confirma Gravação?", Pergunta) = vbNo Then
        Exit Sub
    End If
    
    If Not Gravar Then
        Mensagem "Erro ao gravar registro!", ErroCritico
        Exit Sub
    End If
        
    Call sprNatureza.DeletarLinha
    Mensagem "Gravação efetuada", Informacao

    Exit Sub
err_cmdGravar_Click:
    ShowError
End Sub
Private Function Gravar()
    Dim i As Long
    Dim cProdutoNatureza As New clsProdutoNatureza

    Call AbreTransacao
    Gravar = False
    
    For i = 1 To sprNatureza.MaxRows - 1
    
        sprNatureza.Row = i
        Call CarregarClasse(cProdutoNatureza)

        If sprNatureza.StatusGravacao(i) = const_Insert Or sprNatureza.StatusGravacao(i) = const_Update Then
                       
            If Not cProdutoNatureza.Gravar Then
                Call VoltaTransacao
                Exit Function
            End If
            sprNatureza.TextCol("id_ProdutoNatureza") = cProdutoNatureza.id_ProdutoNatureza
        ElseIf sprNatureza.StatusGravacao(i) = const_Delete Then
            If Not cProdutoNatureza.Excluir Then
                Call VoltaTransacao
                Exit Function
            End If
        End If
        
        Call sprNatureza.Atualizar(i)
    Next i
    
    Call FechaTransacao
    Gravar = True
    Set cProdutoNatureza = Nothing

End Function
Private Function CarregarClasse(cProdutoNatureza As clsProdutoNatureza)
    
    cProdutoNatureza.id_ProdutoNatureza = sprNatureza.SpreadEventoName("id_ProdutoNatureza")
    cProdutoNatureza.ds_ProdutoNatureza = sprNatureza.SpreadEventoName("ds_ProdutoNatureza")
    cProdutoNatureza.tp_ProdutoPerigoso = sprNatureza.SpreadEventoName("tp_ProdutoPerigoso")
    cProdutoNatureza.tp_ControlePoliciaFederal = sprNatureza.SpreadEventoName("tp_ControlePoliciaFederal")

End Function

Private Sub cmdSair_Click()
    Unload Me
End Sub

