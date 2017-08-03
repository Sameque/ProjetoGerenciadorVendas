VERSION 5.00
Begin VB.Form frmFinanceiroTipoBaixa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipo de Baixa"
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
   Begin Transportes.SuperSpreadNovo sprFinanceiroBaixaParcelaTipo 
      Height          =   2010
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   3545
   End
End
Attribute VB_Name = "frmFinanceiroTipoBaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sTabelas As String
Dim sCampos As String

Private Sub cmdGravar_Click()
On Error GoTo err_cmdGravar_Click
    
    If Not ValidarControlesNovo(Me) Then
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
    
    Call sprFinanceiroBaixaParcelaTipo.DeletarLinha
    
    Mensagem "Gravação efetuada", Informacao
           
    Exit Sub
err_cmdGravar_Click:
    ShowError
End Sub
Private Function Gravar() As Boolean
On Error GoTo err_Gravar
    Dim i As Long
    Dim cFinanceiroTipoBaixa As New clsFinanceiroTipoBaixa
    
    Gravar = False
    
    For i = 1 To sprFinanceiroBaixaParcelaTipo.MaxRows - 1
    
        sprFinanceiroBaixaParcelaTipo.Row = i
                    
        Call CarregarClasse(cFinanceiroTipoBaixa)
                            
        If sprFinanceiroBaixaParcelaTipo.StatusGravacao(i) = const_Insert Or sprFinanceiroBaixaParcelaTipo.StatusGravacao(i) = const_Update Then
        
            If Not cFinanceiroTipoBaixa.Gravar Then
                Exit Function
            End If
            
            sprFinanceiroBaixaParcelaTipo.TextCol("id_FinanceiroTipoBaixa") = cFinanceiroTipoBaixa.id_FinanceiroTipoBaixa
            
        ElseIf sprFinanceiroBaixaParcelaTipo.StatusGravacao(i) = const_Delete Then
        
            If Not cFinanceiroTipoBaixa.Excluir Then
                Exit Function
            End If
        End If
        
        Call sprFinanceiroBaixaParcelaTipo.Atualizar(i)
    Next i
    
    Set cFinanceiroTipoBaixa = Nothing
    Gravar = True
    
    Exit Function
err_Gravar:
    ShowError
End Function
Private Function CarregarClasse(cFinanceiroTipoBaixa As clsFinanceiroTipoBaixa)

    cFinanceiroTipoBaixa.id_FinanceiroTipoBaixa = sprFinanceiroBaixaParcelaTipo.SpreadEventoName("id_FinanceiroTipoBaixa")
    cFinanceiroTipoBaixa.ds_TipoBaixa = sprFinanceiroBaixaParcelaTipo.SpreadEventoName("ds_TipoBaixa")

End Function
Private Sub cmdSair_Click()
    Unload Me
End Sub

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

Private Function IniciarComponentes()
        
    Call sprFinanceiroBaixaParcelaTipo.NovaColunaSpread(eslNumero, True, False, "id_FinanceiroTipoBaixa", , 0, 10)
    Call sprFinanceiroBaixaParcelaTipo.NovaColunaSpread(eslTexto, False, False, "ds_TipoBaixa", "Tipo de Baixa", 40, 50)
    
    Call sprFinanceiroBaixaParcelaTipo.FormatarNovo(21)
    
End Function
Private Function CarregarComponentes()

    sCampos = "id_FinanceiroTipoBaixa,ds_TipoBaixa"
    sTabela = "tbdFinanceiroTipoBaixa"

    sprFinanceiroBaixaParcelaTipo.Carregar (Select_Table(False, sTabela, sCampos, , "ds_TipoBaixa"))

End Function
