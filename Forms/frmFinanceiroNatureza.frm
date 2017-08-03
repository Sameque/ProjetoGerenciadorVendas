VERSION 5.00
Begin VB.Form frmFinanceiroNatureza 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Natureza"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   615
      Left            =   2460
      TabIndex        =   2
      Top             =   2265
      Width           =   975
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   2265
      Width           =   975
   End
   Begin Transportes.SuperSpreadNovo sprFinanceiroNatureza 
      Height          =   2010
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   3545
   End
End
Attribute VB_Name = "frmFinanceiroNatureza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTabela As String
Dim sCampos As String
Dim aTipo() As String

Private Sub Form_Activate()
    frmMDI.Arrange vbCascade
End Sub

Private Sub Form_Load()
On Error GoTo err_Form_Load

    Call IniciarComponentes
    Call CarregarComponentes
    
    Exit Sub
err_Form_Load:
    ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFinanceiroNatureza = Nothing
End Sub

Private Sub cmdGravar_Click()
On Error GoTo err_cmdGravar_Click
    
    If Not ValidarControles(Me) Then
        Exit Sub
    End If
    
    If Mensagem("Confirma Gravação?", Pergunta) = vbNo Then
        Exit Sub
    End If
    
    Call AbreTransacao
    
    If Not Gravar Then
        Call VoltaTransacao
        Mensagem "Erro ao gravar registro!", ErroCritico
    End If
    
    Call FechaTransacao
    
    Call sprFinanceiroNatureza.DeletarLinha
    
    Mensagem "Gravação efetuada", Informacao
           
    Exit Sub
err_cmdGravar_Click:
    ShowError
End Sub
Private Function Gravar() As Boolean
On Error GoTo err_Gravar
    Dim i As Long
    Dim cFinanceiroNatureza As New clsFinanceiroNatureza
    
    Gravar = False
    
    For i = 1 To sprFinanceiroNatureza.MaxRows - 1
    
        sprFinanceiroNatureza.Row = i
        Call CarregarClasse(cFinanceiroNatureza)

        If sprFinanceiroNatureza.StatusGravacao(i) = const_Insert Or sprFinanceiroNatureza.StatusGravacao(i) = const_Update Then
            If Not cFinanceiroNatureza.Gravar Then
                Exit Function
            End If
            sprFinanceiroNatureza.TextCol("id_FinanceiroNatureza") = cFinanceiroNatureza.id_FinanceiroNatureza
        ElseIf sprFinanceiroNatureza.StatusGravacao(i) = const_Delete Then
            If Not cFinanceiroNatureza.Excluir Then
                Exit Function
            End If
        End If
        
        Call sprFinanceiroNatureza.Atualizar(i)
    Next i
    
    Set cFinanceiroNatureza = Nothing
    
    Gravar = True
    
    Exit Function
err_Gravar:
    ShowError
    
End Function
Private Function CarregarClasse(cFinanceiroNatureza As clsFinanceiroNatureza)
    
    cFinanceiroNatureza.id_FinanceiroNatureza = sprFinanceiroNatureza.SpreadEventoName("id_FinanceiroNatureza")
    cFinanceiroNatureza.ds_FinanceiroNatureza = sprFinanceiroNatureza.SpreadEventoName("ds_FinanceiroNatureza")
    cFinanceiroNatureza.tp_Financeiro = sprFinanceiroNatureza.SpreadComboEventoName("cboTipo", aTipo)

End Function
Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Function IniciarComponentes()
    
    Call sprFinanceiroNatureza.NovaColunaSpread(eslNumero, True, False, "id_FinanceiroNatureza", , 0, 10)
    Call sprFinanceiroNatureza.NovaColunaSpread(eslTexto, False, False, "ds_FinanceiroNatureza", "Natureza", 26, 50)
    Call sprFinanceiroNatureza.NovaColunaSpread(eslTexto, False, False, "cboTipo", "Tipo", 15)
    Call sprFinanceiroNatureza.FormatarNovo(21)
    
End Function
Private Function CarregarComponentes()

    sCampos = "id_FinanceiroNatureza, ds_FinanceiroNatureza, tp_Financeiro"
    sTabela = "tbdFinanceiroNatureza"

    Call sprFinanceiroNatureza.Carregar(Select_Table(False, sTabela, sCampos, "", "ds_FinanceiroNatureza"))
    Call sprFinanceiroNatureza.Combo_SpreadName("", "cboTipo", aTipo, Array("Credito", "Debito", 1))
    
End Function
