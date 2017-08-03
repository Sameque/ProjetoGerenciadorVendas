VERSION 5.00
Begin VB.Form frmModeloCadastroSpread 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modelo Cadastro"
   ClientHeight    =   6330
   ClientLeft      =   4905
   ClientTop       =   3075
   ClientWidth     =   10455
   Icon            =   "frmModeloCadastroSpread.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   10455
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   750
      Left            =   9570
      Picture         =   "frmModeloCadastroSpread.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Sair da tela"
      Top             =   5520
      Width           =   810
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   750
      Left            =   8700
      Picture         =   "frmModeloCadastroSpread.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Gravar os dados"
      Top             =   5520
      Width           =   810
   End
   Begin Transportes.SuperSpreadNovo sprCadastro 
      Height          =   5370
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   9472
   End
End
Attribute VB_Name = "frmModeloCadastroSpread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    frmMDI.Arrange vbCascade
End Sub

Private Sub Form_Load()
On Error GoTo err_Form_Load

    '[FORMATAÇÃO-SPREAD]'
    
    '[CARREGAR-SPREAD]'

    Exit Sub
err_Form_Load:
    ShowError
End Sub

Private Sub cmdGravar_Click()
On Error GoTo err_cmdGravar

    Dim i As Long
    
    If Not sprCadastro.ValidaGravacao() Then
        Exit Sub
    End If
    
    If Mensagem("Confirma Gravação?", Informacao) = vbNo Then
        Exit Sub
    End If
    
    
    If Not CarregarPropriedades Then
       Mensagem "Ocorreu um erro na carga das propriedades.", ErroCritico
       Exit Sub
    End If
    
    Call AbreTransacao
    If Not cPrincipal.Gravar Then
        Call VoltaTransacao
        Call Mensagem("Ocorreu um erro na Gravação!", ErroCritico)
        Exit Sub
    End If
    Call FechaTransacao
               
    Call sprCadastro.Atualizar(0)
    Call sprCadastro.DeletarLinha
    
    Call Mensagem("Gravação Efetuada.", Informacao)
    Exit Sub
err_cmdGravar:
    Call VoltaTransacao
    ShowError
End Sub

Private Function CarregarPropriedades() As Boolean
On Error GoTo err_CarregarPropriedades

    CarregarPropriedades = False
    
    Dim i As Long

    '[SETAR-PROPRIEDADES]'
    CarregarPropriedades = True
    
    Exit Function
err_CarregarPropriedades:
    ShowError
End Function


Private Sub cmdSair_Click()
    Unload Me
End Sub
