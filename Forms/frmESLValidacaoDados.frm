VERSION 5.00
Object = "{C8165CE7-70BA-4DB0-93DB-54254E1E7849}#8.0#0"; "ComponentesESL.ocx"
Begin VB.Form frmESLValidacaoDados 
   BackColor       =   &H00A8FAFD&
   BorderStyle     =   0  'None
   Caption         =   "Validação de Dados"
   ClientHeight    =   6030
   ClientLeft      =   3225
   ClientTop       =   2175
   ClientWidth     =   6600
   Icon            =   "frmESLValidacaoDados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   Begin ComponentesESL.SuperBotaoNovo cmdOK 
      Height          =   375
      Left            =   5595
      TabIndex        =   4
      Top             =   5520
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   661
      Caption         =   "OK"
      UseSound        =   0
      Align           =   0
      Picture         =   "frmESLValidacaoDados.frx":000C
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   611929
      ForeColorOver   =   611929
   End
   Begin Transportes.SuperSpreadNovo sprErro 
      Height          =   4500
      Left            =   180
      TabIndex        =   3
      Top             =   930
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   7938
      ControlaQueryAdvance=   0   'False
      ControlaClick   =   0   'False
      ExcluirRegistro =   0   'False
      EsconderUltimaLinha=   -1  'True
   End
   Begin ComponentesESL.SuperBotaoNovo cmdFechar 
      Height          =   240
      Left            =   6255
      TabIndex        =   0
      Top             =   150
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      Caption         =   ""
      UseSound        =   0
      Align           =   0
      Picture         =   "frmESLValidacaoDados.frx":0689
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
   End
   Begin VB.Image imgAlerta 
      Height          =   240
      Left            =   5685
      Picture         =   "frmESLValidacaoDados.frx":0E7D
      Top             =   555
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgErro 
      Height          =   240
      Left            =   5415
      Picture         =   "frmESLValidacaoDados.frx":12F2
      Top             =   555
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblAviso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Favor verificar as ocorrências abaixo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005D8D8F&
      Height          =   240
      Left            =   1740
      TabIndex        =   2
      Top             =   495
      Width           =   3315
   End
   Begin VB.Label lblErro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem de Alerta do Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005D8D8F&
      Height          =   300
      Left            =   1485
      TabIndex        =   1
      Top             =   150
      Width           =   3915
   End
   Begin VB.Image Image 
      Height          =   615
      Left            =   165
      Picture         =   "frmESLValidacaoDados.frx":1767
      Top             =   150
      Width           =   660
   End
   Begin VB.Shape borda 
      BorderColor     =   &H002DAAAF&
      Height          =   570
      Left            =   900
      Top             =   225
      Width           =   570
   End
End
Attribute VB_Name = "frmESLValidacaoDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call cmdFechar_Click
End Sub

Private Sub Form_Load()
On Error GoTo err_Handler

    Call RedimensionarBorda
    Call FormatarSpread
    Call CenterForm(Me)

    Exit Sub
err_Handler:
    ShowError
End Sub

Private Sub RedimensionarBorda()
    borda.Top = 0
    borda.Left = 0
    borda.Width = Me.Width
    borda.Height = Me.Height
End Sub

Private Sub FormatarSpread()
On Error GoTo err_FormatarSpread

    Call sprErro.NovaColunaSpread(eslImagem, True, True, "imgStatus", "", 3)
    Call sprErro.NovaColunaSpread(eslTexto, True, True, "ds_NomeCampo", "ds_NomeCampo", 27, 100)
    Call sprErro.NovaColunaSpread(eslfrase, True, True, "ds_Erro", "Descrição", 32, 500)
    sprErro.FormatarNovo
    
    sprErro.ControlaClick = False
    sprErro.objSpread.BorderStyle = BorderStyleNone
    sprErro.Row = -1
    sprErro.Col = -1
    sprErro.fontBold = True
    sprErro.ForeColor = 5980455 '&H666666
    sprErro.objSpread.SelModeSelected = True
    sprErro.objSpread.OperationMode = OperationModeRow
    sprErro.objSpread.ActiveCellHighlightStyle = ActiveCellHighlightStyleNormal
    sprErro.objSpread.SelBackColor = vbWhite
    sprErro.objSpread.SelForeColor = 5980455 '&H666666
    sprErro.objSpread.NoBorder = True
    sprErro.RowHeight = 25
    sprErro.BackColor = Me.BackColor
    sprErro.GrayAreaBackColor = Me.BackColor
    sprErro.GridShowHoriz = True
    sprErro.GridShowVert = False
    sprErro.DestacarRegistro = False
    sprErro.BackColorCellAtiva = vbWhite
    sprErro.objSpread.ColHeadersShow = False
    sprErro.objSpread.RowHeadersShow = False
    sprErro.Row = -1
    sprErro.ColName = "ds_NomeCampo"
    sprErro.objSpread.TypeVAlign = TypeVAlignCenter
    sprErro.Row = -1
    sprErro.ColName = "ds_Erro"
    sprErro.objSpread.TypeVAlign = TypeVAlignCenter
    sprErro.fontBold = False
    sprErro.Font = "Times New Roman"
    sprErro.fontSize = 10
    
    'Call sprErro.objSpread.SetOddEvenRowColor(&HA3A3EF, vbWhite, &HB2B2EF, vbWhite)
    'Call sprErro.objSpread.SetOddEvenRowColor(&HA8FAFD, &H666666, &HC1FAFC, &H666666)
    
    
    Exit Sub
err_FormatarSpread:
    ShowError
End Sub

Public Sub CarregarInconsistencias(aInconsistencias As Variant)
On Error GoTo err_CarregarInconsistencias

    Dim contador As Integer
    
    If IsEmpty(aInconsistencias) Then
        Exit Sub
    End If
    
    For contador = 1 To UBound(aInconsistencias, 2)
        sprErro.MaxRows = sprErro.MaxRows + 1
        sprErro.Row = sprErro.MaxRows
        
        If aInconsistencias(2, contador) = EnumAcaoMensagem.Bloquear Then
            Call sprErro.SetarPicture(imgErro, Val(contador), "imgStatus")
        Else
            Call sprErro.SetarPicture(imgAlerta, Val(contador), "imgStatus")
        End If
        
        sprErro.TextCol("ds_NomeCampo") = aInconsistencias(0, contador)
        sprErro.TextCol("ds_Erro") = aInconsistencias(1, contador)
    Next contador
    
    Exit Sub
err_CarregarInconsistencias:
    ShowError
End Sub

Private Sub MoverForm(Button As Integer)
On Error GoTo err_MoverForm

    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If

    Exit Sub
err_MoverForm:
    ShowError "MoverForm()" & vbCrLf
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoverForm(Button)
End Sub

Private Sub Image_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoverForm(Button)
End Sub

Private Sub lblAviso_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoverForm(Button)
End Sub

Private Sub lblErro_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoverForm(Button)
End Sub

