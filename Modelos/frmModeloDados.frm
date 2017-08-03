VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmModeloDados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modelo Dados"
   ClientHeight    =   8370
   ClientLeft      =   3540
   ClientTop       =   1950
   ClientWidth     =   12840
   Icon            =   "frmModeloDados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   12840
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Height          =   750
      HelpContextID   =   23
      Left            =   11085
      Picture         =   "frmModeloDados.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Novo lançamento "
      Top             =   7560
      Width           =   810
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   750
      Left            =   10215
      Picture         =   "frmModeloDados.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Gravar os Dados"
      Top             =   7560
      Width           =   810
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   750
      Left            =   11955
      Picture         =   "frmModeloDados.frx":0BE0
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Sair da tela"
      Top             =   7560
      Width           =   810
   End
   Begin Threed.SSFrame fraDados 
      Height          =   7485
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   12705
      _Version        =   65536
      _ExtentX        =   22410
      _ExtentY        =   13203
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
   End
End
Attribute VB_Name = "frmModeloDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public id_Principal As Long
Public formChamador As Form

'Dim cModelo As clsModelo

Private Sub Form_Load()
On Error GoTo err_Form_Load

    Set cModelo = New clsModelo
    
    '[FORMATAÇÃO-COMBOS]'
    '[FORMATAÇÃO-SPREAD]'
    If id_Principal > 0 Then
        If cModelo.Carregardados(id_Principal) Then
            '[CARREGAR-DADOS]'
        End If
        
        '[CARREGAR-DADOSITEM]'
    End If

    Call CenterForm(Me)
    
    Exit Sub
err_Form_Load:
    ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cModelo = Nothing
    Set Me = Nothing
End Sub

Private Sub cmdGravar_Click()
On Error GoTo err_cmdGravar_Click

    If Not ValidarControles(Me) Then
        Exit Sub
    End If
    
    If Mensagem("Confirma Gravação?", Pergunta) = vbNo Then
        Exit Sub
    End If

    If Not CarregarPropriedades(cModelo) Then
        Exit Sub
    End If

    Call AbreTransacao
    If Not cModelo.Gravar Then
        Call VoltaTransacao
        Mensagem "Ocorreu um erro no processamento.", ErroCritico
        Exit Sub
    End If
    Call FechaTransacao

    id_Principal = cModelo.id_Principal
    
    '[ATUALIZAR-SPREAD]'
    Call formChamador.AtualizarDados(id_Principal)
    
    Call Mensagem("Gravação Efetuada", Informacao)
    cmdNovo.SetFocus

    Exit Sub
err_cmdGravar_Click:
    ShowError
    Call FechaTransacao
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

Private Sub cmdNovo_Click()
    Call LimparControles(Me)
    Set cModelo = New clsModelo
    id_Principal = 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub
