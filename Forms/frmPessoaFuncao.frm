VERSION 5.00
Begin VB.Form frmPessoaFuncao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de PessoaFuncao"
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
   Begin Transportes.SuperSpreadNovo sprFuncao 
      Height          =   2670
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   4710
   End
End
Attribute VB_Name = "frmPessoaFuncao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCampos As String
Dim sTabelas As String

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
    
    Call sprFuncao.NovaColunaSpread(eslNumero, True, True, "id_PessoaFuncao", "id_PessoaFuncao", 0, 0)
    Call sprFuncao.NovaColunaSpread(eslTexto, False, False, "ds_PessoaFuncao", "Função", 30, 50)
    Call sprFuncao.NovaColunaSpread(eslNumero, False, False, "qt_NivelPermissao", "Nível de Permissão", 14, 4)
    Call sprFuncao.FormatarNovo(21)
    
End Function
Private Function CarregarComponentes()

    sCampos = "a.id_PessoaFuncao, a.ds_PessoaFuncao, a.qt_NivelPermissao"
    sTabelas = "tbdPessoaFuncao a"

    Call sprFuncao.Carregar(Select_Table(False, sTabelas, sCampos, "", "ds_PessoaFuncao"))
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frmPessoaFuncao = Nothing
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
        Exit Sub
    End If
    
    Call FechaTransacao
        
    Call sprFuncao.DeletarLinha
    Mensagem "Gravação efetuada", Informacao

    Exit Sub
err_cmdGravar_Click:
    ShowError
End Sub
Private Function Gravar()
    Dim i As Long
    Dim cPessoaFuncao As New clsPessoaFuncao

    Gravar = False
    
    For i = 1 To sprFuncao.MaxRows - 1
    
        sprFuncao.Row = i
        Call CarregarClasse(cPessoaFuncao)

        If sprFuncao.StatusGravacao(i) = const_Insert Or sprFuncao.StatusGravacao(i) = const_Update Then
            
            If Not cPessoaFuncao.Gravar Then
                Exit Function
            End If
            
            sprFuncao.TextCol("id_PessoaFuncao") = cPessoaFuncao.id_PessoaFuncao
        ElseIf sprFuncao.StatusGravacao(i) = const_Delete Then
            
            If Not cPessoaFuncao.Excluir Then
                Exit Function
            End If
        End If
        
        Call sprFuncao.Atualizar(i)
    Next i
    
    Gravar = True
    Set cPessoaFuncao = Nothing

End Function
Private Function CarregarClasse(cPessoaFuncao As clsPessoaFuncao)
    
    cPessoaFuncao.id_PessoaFuncao = sprFuncao.SpreadEventoName("id_PessoaFuncao")
    cPessoaFuncao.ds_PessoaFuncao = sprFuncao.SpreadEventoName("ds_PessoaFuncao")
    cPessoaFuncao.qt_NivelPermissao = sprFuncao.SpreadEventoName("qt_NivelPermissao")

End Function

Private Sub cmdSair_Click()
    Unload Me
End Sub

