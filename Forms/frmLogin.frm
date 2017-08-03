VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2310
   ClientLeft      =   8130
   ClientTop       =   5130
   ClientWidth     =   4065
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame fraLogin 
      Height          =   2070
      Left            =   120
      TabIndex        =   5
      Top             =   105
      Width           =   3000
      _Version        =   65536
      _ExtentX        =   5292
      _ExtentY        =   3651
      _StockProps     =   14
      Caption         =   "Login p/ o sistema"
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
      Begin VB.ComboBox cboAmbiente 
         Height          =   315
         ItemData        =   "frmLogin.frx":000C
         Left            =   1005
         List            =   "frmLogin.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   390
         Width           =   1830
      End
      Begin VB.TextBox txtSenha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1005
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1500
         Width           =   1830
      End
      Begin VB.TextBox txtUsuario 
         Height          =   315
         Left            =   1005
         MaxLength       =   20
         TabIndex        =   0
         Top             =   945
         Width           =   1830
      End
      Begin VB.Label lblAmbiente 
         AutoSize        =   -1  'True
         Caption         =   "Ambiente:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   450
         Width           =   705
      End
      Begin VB.Label lblSenha 
         AutoSize        =   -1  'True
         Caption         =   "Senha:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   510
      End
      Begin VB.Label lblUsuario 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   990
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   750
      Left            =   3195
      Picture         =   "frmLogin.frx":0010
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   810
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   750
      Left            =   3195
      Picture         =   "frmLogin.frx":031A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   645
      Width           =   810
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cConexao As New clsConexao
Dim iControle As Integer

Private Sub cmdGravar_Click()
    On Error GoTo err_cmdGravarClick

    Dim iCont As Integer
    Dim bOK As Boolean
    Dim lRetorno As Long
    Dim sPath As String

    If Trim(txtUsuario.Text) = "" Then
        Mensagem "Informe o login.", erro
        txtUsuario.SetFocus
        Exit Sub
    End If
    If Trim(txtSenha.Text) = "" Then
        Mensagem "Informe a senha.", erro
        txtSenha.SetFocus
        Exit Sub
    End If

    'Call InicializaSistemaContainer(, IIf(cboAmbiente.ListCount > 0, cboAmbiente.ListIndex, 0))

    '    id_Matriz = Val(GetINIString("ISL", "Matriz", ""))
    '    tp_BancoDados = Val(Trim(GetINIString("DataBase", "Type", "")))
    '    sPath = Trim(GetINIString("DataBase", "Path", ""))
    '    sPathZip = Trim(GetINIString("Diversos", "PathZip", ""))
    '    sPathReport = Trim(GetINIString("Diversos", "PathReport", ""))
    '    sSenhaBanco = ""
    '    If Trim(sPathReport) = "" Then
    '        sPathReport = App.Path
    '    End If
    '    ds_BancoDados = sPath & "dtbTransporte.mdb"

    cConexao.ds_Login = txtUsuario
    cConexao.ds_Senha = txtSenha

    If cboAmbiente.ListCount > 0 Then
        If cboAmbiente.ListIndex = -1 Then
            Call Mensagem("Escolha um ambiente.", erro)
            cboAmbiente.SetFocus
            Exit Sub
        End If
    End If
    sAmbiente = Trim(cboAmbiente.Text)
    'Não pode tirar esta linha daqui, devido aos ambientes (Daniel)
    Call InicializaSistemaGeral(, IIf(cboAmbiente.ListCount > 0, cboAmbiente.ListIndex, 0), , True)
    bConect = False

    frmMDI.ds_Senha = txtSenha.Text
    frmMDI.ds_Usuario = txtUsuario.Text
    bOK = False
    iControle = iControle + 1
    If cConexao.logon Then
        bOK = True
    End If
    If bOK Then
        If Trim(sAmbiente) <> "" Then
            frmMDI.StatusBarMDI.Panels(3).Text = "Ambiente: " & sAmbiente
        End If
        Unload Me
    Else
        If iControle > 3 Then
            End
        End If
    End If

    Exit Sub
err_cmdGravarClick:
    ShowError
End Sub

Private Sub cmdSair_Click()
    End
End Sub

Private Sub Form_Activate()
    txtSenha.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo err_FormLoad
    Dim lSalto As Long
    Dim sUser As String
    Dim aBancos() As String
    Dim i As Integer

    aBancos = Split(GetINIString(sSistemaNome, "Bancos", ""), "||")
    cboAmbiente.Clear
    For i = 0 To UBound(aBancos)
        cboAmbiente.AddItem StrZero((i + 1), Len(CStr(UBound(aBancos) + 1))) & " - " & aBancos(i)    'O "i" é para deixar os dados ordenados corretamente
    Next
    If cboAmbiente.ListCount > 0 Then
        i = CLng1(GetINIString(sSistemaNome, "Ambiente", ""))
        If i < cboAmbiente.ListCount Then
            cboAmbiente.ListIndex = i
        Else
            cboAmbiente.ListIndex = 0
        End If
        tp_ConfigAmbiente = True
    Else
        tp_ConfigAmbiente = False
        lblAmbiente.Visible = False
        cboAmbiente.Visible = False
        lSalto = txtUsuario.Top - cboAmbiente.Top

        lblUsuario.Top = lblUsuario.Top - lSalto
        txtUsuario.Top = txtUsuario.Top - lSalto
        lblSenha.Top = lblSenha.Top - lSalto
        txtSenha.Top = txtSenha.Top - lSalto
        cmdGravar.Top = cmdGravar.Top - lSalto
        cmdSair.Top = cmdSair.Top - lSalto
        fraLogin.Height = fraLogin.Height - lSalto
        Me.Height = Me.Height - lSalto
    End If


    bConect = False

    bTerminalServer = False
    sUser = Trim(GetINIString(sSistemaNome, "Usuario", ""))
    If sUser = "" Then
        sUser = Trim(GetINIString("ISL", "Usuario", ""))
    End If
    txtUsuario.Text = sUser

    'If Not VerificarTrava Then
    '    End
    'End If
    
    txtSenha.Text = "fgh201706esl"
    cboAmbiente.ListIndex = 0
    'cmdGravar.SetFocus
    
    Call CenterForm(Me)
    
    cmdGravar.Value = True

    Exit Sub
err_FormLoad:
    ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not bConect Then
        End
    End If
End Sub
