VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUsuarioLista 
   Caption         =   "Form1"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   12345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRelatorio 
      Caption         =   "Relatório"
      Height          =   360
      Left            =   1560
      TabIndex        =   10
      Top             =   4545
      Width           =   1260
   End
   Begin MSAdodcLib.Adodc adoUsuario 
      Height          =   330
      Left            =   6120
      Top             =   4620
      Visible         =   0   'False
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PaschoalottoDesafio"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PaschoalottoDesafio"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT IdUsuario, Nome, Sobrenome, Senha, Email, Telefone, Genero FROM Usuario (NOLOCK)"
      Caption         =   "adoUsuario"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAPIRandom 
      Caption         =   "API Rondom"
      Height          =   360
      Left            =   150
      TabIndex        =   8
      Top             =   4545
      Width           =   1260
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   360
      Left            =   10080
      TabIndex        =   7
      Top             =   4575
      Width           =   1005
   End
   Begin VB.CommandButton cmFechar 
      Caption         =   "Fechar"
      Height          =   360
      Left            =   11115
      TabIndex        =   6
      Top             =   4575
      Width           =   1005
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      Height          =   360
      Left            =   8955
      TabIndex        =   4
      Top             =   4590
      Width           =   1005
   End
   Begin VB.ComboBox cmbLocalizar 
      Height          =   315
      ItemData        =   "frmUsuarioLista.frx":0000
      Left            =   4500
      List            =   "frmUsuarioLista.frx":000A
      TabIndex        =   3
      Top             =   750
      Width           =   1425
   End
   Begin VB.TextBox txtLocalizar 
      Height          =   315
      Left            =   855
      MaxLength       =   50
      TabIndex        =   1
      Top             =   750
      Width           =   3045
   End
   Begin MSDataGridLib.DataGrid dtgUsuario 
      Bindings        =   "frmUsuarioLista.frx":001F
      Height          =   3345
      Left            =   135
      TabIndex        =   9
      Top             =   1110
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   5900
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "IdUsuario"
         Caption         =   "IdUsuario"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Nome"
         Caption         =   "Nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Sobrenome"
         Caption         =   "Sobrenome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Senha"
         Caption         =   "Senha"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Email"
         Caption         =   "Email"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Telefone"
         Caption         =   "Telefone"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Genero"
         Caption         =   "Genero"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   2640,189
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitulo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cadastro de Usuário"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   7665
      TabIndex        =   5
      Tag             =   "175"
      Top             =   135
      Width           =   4380
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image3 
      Height          =   795
      Left            =   4305
      Picture         =   "frmUsuarioLista.frx":0038
      Top             =   -165
      Width           =   8250
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      Height          =   195
      Index           =   1
      Left            =   4065
      TabIndex        =   2
      Top             =   825
      Width           =   360
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Localizar:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   810
      Width           =   675
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   225
      Picture         =   "frmUsuarioLista.frx":088C
      Top             =   60
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   45
      Left            =   -1515
      Picture         =   "frmUsuarioLista.frx":1156
      Top             =   630
      Width           =   10740
   End
   Begin VB.Image imgBarra 
      Height          =   795
      Left            =   0
      Picture         =   "frmUsuarioLista.frx":1ADA
      Top             =   -165
      Width           =   8250
   End
End
Attribute VB_Name = "frmUsuarioLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Dim vop_UsuarioNegocios As New clsUsuarioNegocios
'Variaveis de controle do form
Dim vil_IdUsuario As Long              'Identificador do Usuario



'Eventos
Private Sub Form_Activate()
   
   Me.Refresh
   
End Sub

Private Sub Form_Load()
    
    cmbLocalizar.ListIndex = 0
    Call CarregarGrid
    
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
    
    Set frmUsuarioLista = Nothing
    
End Sub

Private Sub cmdExcluir_Click()
    If vil_IdUsuario = 0 Then Exit Sub
       
    If MsgBox("Confirma a Exclusão ?", vbQuestion + vbYesNo, "Confirme !") = vbYes Then
      Set vop_UsuarioNegocios = New clsUsuarioNegocios
          vop_UsuarioNegocios.IdUsuario = vil_IdUsuario
          If vop_UsuarioNegocios.ExcluirUsuario() = True Then
             txtLocalizar.text = Empty
             Call CarregarGrid
          End If
      Set vop_UsuarioNegocios = Nothing
    End If
End Sub

Private Sub cmdNovo_Click()
    frmUsuario.Show vbModal
End Sub

Private Sub lblFechar_Click()
   Unload Me
End Sub

Private Sub cmFechar_Click()
On Error GoTo TrataErros
    If MsgBox(MSG01, Style10, Title01) = vbYes Then
       Set frmUsuarioLista = Nothing
       End
    End If
TrataErros:
    If Err.Number = 3420 Then End
End Sub

Private Sub cmdAPIRandom_Click()
Dim vbl_Carregar As Boolean
    
On Error GoTo TrataErros

   If MsgBox("Adicionar usuários aleatórios ?", vbQuestion + vbYesNo, "Confirme !") = vbYes Then
   
      Set vop_UsuarioNegocios = New clsUsuarioNegocios
          vbl_Carregar = vop_UsuarioNegocios.InserirUsuarioAleatorio()
          If vbl_Carregar = True Then
             Call CarregarGrid
          End If
      Set vop_UsuarioNegocios = Nothing
   
   End If
   
TrataErros:
    If Err.Number <> 0 Then
       Set vop_UsuarioNegocios = Nothing
    End If
End Sub

Private Sub cmdRelatorio_Click()
Dim vbl_Carregar As Boolean
    
On Error GoTo TrataErros

   If MsgBox("Imprimir relatório de usuários ?", vbQuestion + vbYesNo, "Confirme !") = vbYes Then
      Set vop_UsuarioNegocios = New clsUsuarioNegocios
          vbl_Carregar = vop_UsuarioNegocios.ImprimirUsuarios()
          If vbl_Carregar = True Then
             dtrUsuario.Show vbModal
          End If
      Set vop_UsuarioNegocios = Nothing
    End If

TrataErros:
    If Err.Number <> 0 Then
       Set vop_UsuarioNegocios = Nothing
    End If
End Sub


Private Sub cmbLocalizar_Click()
   
On Error GoTo TrataErros

   Set vop_UsuarioNegocios = New clsUsuarioNegocios
       Call vop_UsuarioNegocios.LocalizarUsuario(adoUsuario, txtLocalizar.text, cmbLocalizar.ListIndex)
   Set vop_UsuarioNegocios = Nothing
    
   txtLocalizar.text = Empty
   
TrataErros:
    If Err.Number <> 0 Then
       Set vop_UsuarioNegocios = Nothing
       Exit Sub
    End If
   
End Sub

Private Sub dtgUsuario_DblClick()
Dim vvl_BookMark As Variant
Dim vil_RowIndex As Long

On Error GoTo TrataErros

    vil_RowIndex = dtgUsuario.Row
    vvl_BookMark = dtgUsuario.RowBookmark(vil_RowIndex)
    If vvl_BookMark = Empty Then Exit Sub

    Call frmUsuario.Form_Load
    Call frmUsuario.Editar(vil_IdUsuario)
        
    dtgUsuario.SelBookmarks.Remove (0)
    dtgUsuario.Bookmark = vvl_BookMark
    dtgUsuario.Scroll 0, dtgUsuario.RowContaining(vvl_BookMark)
    dtgUsuario.SelBookmarks.Add vvl_BookMark
    dtgUsuario.Refresh
    
TrataErros:
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

End Sub

Private Sub dtgUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = vbKeyDown Then
        'MsgBox "Seta para baixo !", vbExclamation
        If dtgUsuario.SelBookmarks.Count > 0 Then
           dtgUsuario.SelBookmarks.Remove 0
        End If
        adoUsuario.Recordset.MoveNext
        'Valida fim do DataGrid
        If adoUsuario.Recordset.EOF = True Then
           adoUsuario.Recordset.MovePrevious
        End If
        
    End If


End Sub

Private Sub dtgUsuario_KeyUp(KeyCode As Integer, Shift As Integer)

    'Tecla de sair do form
    If KeyCode = vbKeyUp Then
        'MsgBox "Seta para cima !", vbExclamation
        If dtgUsuario.SelBookmarks.Count > 0 Then
           dtgUsuario.SelBookmarks.Remove 0
        End If
        adoUsuario.Recordset.MovePrevious
        'Valida inicio do DataGrid
        If adoUsuario.Recordset.BOF = True Then
           adoUsuario.Recordset.MoveNext
        End If
        
    End If
    
End Sub

Private Sub dtgUsuario_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   
   If dtgUsuario.Bookmark > 0 Then
   
      dtgUsuario.SelBookmarks.Add dtgUsuario.Bookmark
      If dtgUsuario.Columns(0).text = "" Then
         vil_IdUsuario = 0
      Else
         vil_IdUsuario = CInt(dtgUsuario.Columns(0).text)
      End If
   
   End If
   
End Sub

Private Sub txtLocalizar_Change()
  
On Error GoTo TrataErros

   Set vop_UsuarioNegocios = New clsUsuarioNegocios
       Call vop_UsuarioNegocios.LocalizarUsuario(adoUsuario, txtLocalizar.text, cmbLocalizar.ListIndex)
   Set vop_UsuarioNegocios = Nothing
    
TrataErros:
    If Err.Number <> 0 Then
       Set vop_UsuarioNegocios = Nothing
       Exit Sub
    End If
    
End Sub

'Metodos
Public Sub CarregarGrid()
Dim vbl_Carregar As Boolean
    
On Error GoTo TrataErros

    Set vop_UsuarioNegocios = New clsUsuarioNegocios
        
        vbl_Carregar = vop_UsuarioNegocios.CarregarGridUsuarioRS(adoUsuario, cmbLocalizar.ListIndex)
        If vbl_Carregar = False Then
           Set vop_UsuarioNegocios = Nothing
           Exit Sub
        End If
        adoUsuario.Refresh
    Set vop_UsuarioNegocios = Nothing

TrataErros:
    If Err.Number <> 0 Then
       Set vop_UsuarioNegocios = Nothing
       Exit Sub
    End If

End Sub



