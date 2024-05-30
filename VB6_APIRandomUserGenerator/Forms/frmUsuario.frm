VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{DC81D4AD-48D8-4DD6-A8B5-228CB11C1826}#1.0#0"; "prjXTab.ocx"
Begin VB.Form frmUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Cadápio"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   360
      Left            =   6135
      TabIndex        =   20
      Top             =   4065
      Width           =   1005
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   360
      Left            =   5055
      TabIndex        =   19
      Top             =   4065
      Width           =   1005
   End
   Begin VB.TextBox txtIdUsuario 
      Height          =   285
      Left            =   1575
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   150
      Visible         =   0   'False
      Width           =   150
   End
   Begin prjXTab.XTab xtbUsuario 
      Height          =   3180
      Left            =   90
      TabIndex        =   2
      Top             =   765
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   5609
      TabCount        =   1
      TabCaption(0)   =   "  Usuário "
      TabContCtrlCnt(0)=   2
      Tab(0)ContCtrlCap(1)=   "fraTab1"
      Tab(0)ContCtrlCap(2)=   "lblCodigo"
      TabStyle        =   1
      TabTheme        =   1
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   16514555
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   15397104
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   10198161
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      Begin VB.Frame fraTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   2535
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   480
         Width           =   6765
         Begin VB.TextBox txtGenero 
            Height          =   315
            Left            =   1305
            MaxLength       =   30
            TabIndex        =   18
            Tag             =   "0"
            Top             =   2100
            Width           =   2415
         End
         Begin VB.TextBox txtTelefone 
            Height          =   315
            Left            =   1305
            MaxLength       =   30
            TabIndex        =   17
            Tag             =   "0"
            Top             =   1695
            Width           =   2415
         End
         Begin VB.TextBox txtEmail 
            Height          =   315
            Left            =   1305
            MaxLength       =   250
            TabIndex        =   16
            Tag             =   "0"
            Top             =   1290
            Width           =   5385
         End
         Begin VB.TextBox txtSenha 
            Height          =   315
            Left            =   1305
            MaxLength       =   60
            TabIndex        =   15
            Tag             =   "0"
            Top             =   870
            Width           =   3585
         End
         Begin VB.TextBox txtSobrenome 
            Height          =   315
            Left            =   1305
            MaxLength       =   30
            TabIndex        =   14
            Tag             =   "0"
            Top             =   435
            Width           =   1635
         End
         Begin VB.TextBox txtNome 
            Height          =   315
            Left            =   1305
            MaxLength       =   30
            TabIndex        =   0
            Tag             =   "0"
            Top             =   0
            Width           =   1635
         End
         Begin VB.Label lblGEnero 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Genero:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   13
            Top             =   2160
            Width           =   705
         End
         Begin VB.Label lblTelefone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   12
            Top             =   1740
            Width           =   810
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   15
            TabIndex        =   7
            Top             =   1305
            Width           =   540
         End
         Begin VB.Label lblSenha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Senha:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   6
            Top             =   900
            Width           =   615
         End
         Begin VB.Label lblNome 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   4
            Top             =   75
            Width           =   570
         End
         Begin VB.Label lblSobrenome 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobrenome:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   15
            TabIndex        =   5
            Top             =   480
            Width           =   1065
         End
      End
      Begin VB.TextBox Text1 
         Height          =   1575
         Left            =   -74940
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   5080
      End
      Begin VB.Label lblCodigo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1020
         TabIndex        =   9
         Top             =   15
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lvwUsuario 
      Height          =   360
      Left            =   7770
      TabIndex        =   11
      Top             =   1515
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   635
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IdUsuario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sobrenome"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Senha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Email"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Telefone"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Genero"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   225
      Picture         =   "frmUsuario.frx":0000
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Left            =   -1980
      Picture         =   "frmUsuario.frx":08CA
      Top             =   675
      Width           =   10740
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   17955
   End
End
Attribute VB_Name = "frmUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Dim vop_UsuarioNegocios As New clsUsuarioNegocios
'Variaveis de controle do form
Public vbp_Usuario As Boolean                             'Verifica uma inclusao ou alteracao


'Eventos
Private Sub Form_Activate()

    Me.Refresh
   
End Sub

Public Sub Form_Load()

    'Verifica uma inclusao ou alteracao
    vbp_Usuario = False
    
    'Inicializa entrada e saida
    Call InicializaEntradaSaida
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo TrataErros

    'Tecla de sair do form
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
TrataErros:
    If Err.Number <> 0 Then
        Err.Clear
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        KeyAscii = 0
        Unload Me
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Set frmUsuario = Nothing
   
End Sub

Private Sub txtNome_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtSobrenome_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtSobrenome_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtSenhaKeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtTelefone_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtTelefone_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtGenero_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtGenero_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub

Private Sub cmdGravar_Click()
Dim vsp_Mensagem As String

   'Valida mensagem
   If vbp_Usuario = False Then
      vsp_Mensagem = "Confirma a Inclusão ?"
   Else
      vsp_Mensagem = "Confirma a Alteração ?"
   End If

   'Valida entrada de dados
   If VerCampos = False Then Exit Sub
   
   If MsgBox(vsp_Mensagem, vbQuestion + vbYesNo, "Confirme !") = vbYes Then
   
      Set vop_UsuarioNegocios = New clsUsuarioNegocios
          vop_UsuarioNegocios.IdUsuario = lblCodigo.Caption
          vop_UsuarioNegocios.Nome = txtNome.text
          vop_UsuarioNegocios.Sobrenome = txtSobrenome.text
          vop_UsuarioNegocios.Senha = txtSenha.text
          vop_UsuarioNegocios.Email = txtEmail.text
          vop_UsuarioNegocios.Telefone = txtTelefone.text
          vop_UsuarioNegocios.Genero = txtGenero.text
          If vbp_Usuario = False Then
             If vop_UsuarioNegocios.IncluirUsuario() = True Then
                MsgBox "Usuario cadastrado com sucesso !", vbExclamation, "Usuario"
             End If
          Else
             If vop_UsuarioNegocios.AlterarUsuario() = True Then
                MsgBox "Usuario alterado com sucesso !", vbExclamation, "Usuario"
             End If
          End If
      Set vop_UsuarioNegocios = Nothing
      
   End If
   
   'Atualiza grid
   Call frmUsuarioLista.CarregarGrid
  
    'Inicializa entrada e saida
    Call InicializaEntradaSaida
    
   'Valida entrada de dados
   If vbp_Usuario = True Then
      Call cmdFechar_Click
   End If
   
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

'Funcoes
Function Editar(ByVal pIdUsuario As Integer) As Boolean

    'Verifica uma inclusao ou alteracao do cliente
    vbp_Usuario = True
    'Controle de exibicao
    lblCodigo.Visible = True
    
    Set vop_UsuarioNegocios = New clsUsuarioNegocios
        
        If vop_UsuarioNegocios.PesquisarUsuario(lvwUsuario, pIdUsuario) = True Then
           lblCodigo.Caption = pIdUsuario
           txtNome.text = vop_UsuarioNegocios.Nome
           txtSobrenome.text = vop_UsuarioNegocios.Sobrenome
           txtSenha.text = vop_UsuarioNegocios.Senha
           txtEmail.text = vop_UsuarioNegocios.Email
           txtTelefone.text = vop_UsuarioNegocios.Telefone
           txtGenero.text = vop_UsuarioNegocios.Genero
        Else
            MsgBox "Não foi possível encontrar o Usuario !", vbCritical, "Usuario"
        End If
          
    Set vop_UsuarioNegocios = Nothing
    
    Me.Show vbModal

End Function

Function VerCampos() As Boolean
    
    If Trim$(txtNome.text) = Empty Then
        MsgBox "Informe o nome do Usuario !", vbExclamation, "Usuario"
        If txtNome.text <> Empty Then txtNome.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(txtSobrenome.text) = Empty Then
        MsgBox "Informe o Sobrenome do Usuário !", vbExclamation, "Usuario"
        If txtSobrenome.text <> Empty Then txtSobrenome.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(txtSenha.text) = Empty Then
        MsgBox "Informe a Senha do Usuario !", vbExclamation, "Usuario"
        If txtSenha.text <> Empty Then txtSenha.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(txtEmail.text) = Empty Then
        MsgBox "Informe o Email do Usuario !", vbExclamation, "Usuario"
        If txtEmail.text <> Empty Then txtEmail.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(txtTelefone.text) = Empty Then
        MsgBox "Informe o Telefone do Usuario !", vbExclamation, "Usuario"
        If txtTelefone.text <> Empty Then txtTelefone.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(txtGenero.text) = Empty Then
        MsgBox "Informe o Genero do Usuario !", vbExclamation, "Usuario"
        If txtGenero.text <> Empty Then txtGenero.SetFocus
        VerCampos = False
        Exit Function
    End If
    
    
    VerCampos = True

End Function

Private Function InicializaEntradaSaida() As Boolean

    'Limpa entrada de dados
    Call LimpaCampos(Me)
    
    'Inicializa entrada de dados
    Call DefaultCampos
    
End Function

Private Function DefaultCampos() As BookmarkEnum

    txtNome.text = Empty
    txtSobrenome.text = Empty
    txtSenha.text = Empty
    txtEmail.text = Empty
    txtTelefone.text = Empty
    txtGenero.text = Empty
    
End Function


