VERSION 5.00
Begin VB.Form frmFrete 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frete"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleMode       =   0  'User
   ScaleWidth      =   9870
   Begin Transportes.SuperTextMultiline txtVprestCustomizada 
      Height          =   3195
      Left            =   5280
      TabIndex        =   2
      Top             =   1980
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5636
      SkinESL         =   -1  'True
      BackColor       =   16119285
      Label           =   "Tag configurada pelo usuário"
   End
   Begin Transportes.SuperTextMultiline txtVprestNormal 
      Height          =   3195
      Left            =   120
      TabIndex        =   1
      Top             =   1980
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5636
      SkinESL         =   -1  'True
      BackColor       =   16119285
      Label           =   "Tag padrão sistema"
   End
   Begin Transportes.SuperSpreadNovo sprFrete 
      Height          =   1755
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   3096
      BackColorCellAtiva=   14733514
      GrayAreaBackColor=   14670555
      SkinESL         =   -1  'True
      Label           =   "Frete"
   End
End
Attribute VB_Name = "frmFrete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call FormatarComponentes
    Call CarregarComponentes
End Sub

Private Function FormatarComponentes()
On Error GoTo err_FormatarComponentes
    Dim cServicoFrete As New clsServicoFrete
    
    Call sprFrete.FormatarPorClasse(cServicoFrete.FormatarSpreadFrete)
    txtVprestNormal.Enabled = False
    txtVprestCustomizada.Enabled = False
    
    Exit Function
err_FormatarComponentes:
    ShowError
End Function

Private Function CarregarComponentes()
On Error GoTo err_CarregarComponentes

    Call sprFrete.CarregarPorClasse("", True)
    
    Exit Function
err_CarregarComponentes:
    ShowError
End Function

Private Function CarregartxtVprestNormal()
On Error GoTo err_CarregartxtVprestNormal

    Exit Function
err_CarregartxtVprestNormal:
    ShowError
End Function






Private Sub sprFrete_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
End Sub

Private Sub sprFrete_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Mensagem "teste", erro
End Sub

Private Sub sprFrete_Change(ByVal Col As Long, ByVal Row As Long)
    Mensagem "teste", erro
End Sub

Private Sub sprFrete_ChangeName(ByVal ColName As String, ByVal Row As Long)
    Mensagem "teste", erro
End Sub

Private Sub sprFrete_Click(ByVal Col As Long, ByVal Row As Long)
    'Mensagem "teste", erro
End Sub

Private Sub sprFrete_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)
    Mensagem "teste ColWidthChange", erro
End Sub

Private Sub sprFrete_DataFill(ByVal Col As Long, ByVal Row As Long, ByVal DataType As Integer, ByVal fGetData As Integer, Cancel As Integer)
    Mensagem "teste date file", erro
End Sub

Private Sub sprFrete_GotFocus()
    'Mensagem "teste gotoFocus", erro
End Sub

Private Sub sprFrete_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    'Mensagem "teste", erro
End Sub

Private Sub sprFrete_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    Mensagem "teste LeaveRow", erro
End Sub

Private Sub sprFrete_LostFocus()
    'Mensagem "teste LostFocus", erro
End Sub

Private Sub sprFrete_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Mensagem "teste", erro
End Sub

Private Sub sprFrete_RowHeightChange(ByVal Row1 As Long, ByVal Row2 As Long)
    Mensagem "teste RowHeightChange", erro
End Sub

Private Sub sprFrete_SelChange(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long, ByVal CurCol As Long, ByVal CurRow As Long)
    Mensagem "teste", erro
End Sub




