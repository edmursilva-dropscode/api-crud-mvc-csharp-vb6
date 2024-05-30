VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsuarioNegocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Inicializa variáveis da classe
Private vop_UsuarioDados As New clsUsuarioDados      'Variável de acesso classe clsUsuarioDados

'Variáveis das Propriedades
Private mvarIdUsuario As Long
Private mvarNome As String
Private mvarSobrenome As String
Private mvarSenha As String
Private mvarEmail As String
Private mvarTelefone As String
Private mvarGenero As String


'Propriedades
'IdUsuario
Public Property Let IdUsuario(ByVal vData As Long)
    mvarIdUsuario = vData
End Property

Public Property Get IdUsuario() As Long
    IdUsuario = mvarIdUsuario
End Property

'Nome
Public Property Let Nome(ByVal vData As String)
    mvarNome = vData
End Property

Public Property Get Nome() As String
    Nome = mvarNome
End Property

'Sobrenome
Public Property Let Sobrenome(ByVal vData As String)
    mvarSobrenome = vData
End Property

Public Property Get Sobrenome() As String
    Sobrenome = mvarSobrenome
End Property

'Senha
Public Property Let Senha(ByVal vData As String)
    mvarSenha = vData
End Property

Public Property Get Senha() As String
    Senha = mvarSenha
End Property

'Email
Public Property Let Email(ByVal vData As String)
    mvarEmail = vData
End Property

Public Property Get Email() As String
    Email = mvarEmail
End Property

'Telefone
Public Property Let Telefone(ByVal vData As String)
    mvarTelefone = vData
End Property

Public Property Get Telefone() As String
    Telefone = mvarTelefone
End Property

'Genero
Public Property Let Genero(ByVal vData As String)
    mvarGenero = vData
End Property

Public Property Get Genero() As String
    Genero = mvarGenero
End Property


'Metodos
Public Function PesquisarUsuario(ByRef plvwUsuario As ListView, ByVal pIdUsuario As Integer) As Boolean

    PesquisarUsuario = False
    Screen.MousePointer = vbHourglass
      If vop_UsuarioDados.PesquisarUsuario(plvwUsuario, pIdUsuario) = True Then
         If plvwUsuario.ListItems.Count > 0 Then
            With plvwUsuario
               IdUsuario = Trim$(.ListItems(plvwUsuario.ListItems.Count).text)
               Nome = Trim$(.ListItems(plvwUsuario.ListItems.Count).SubItems(1))
               Sobrenome = Trim$(.ListItems(plvwUsuario.ListItems.Count).SubItems(2))
               Senha = Trim$(.ListItems(plvwUsuario.ListItems.Count).SubItems(3))
               Email = Trim$(.ListItems(plvwUsuario.ListItems.Count).SubItems(4))
               Telefone = Trim$(.ListItems(plvwUsuario.ListItems.Count).SubItems(5))
               Genero = Trim$(.ListItems(plvwUsuario.ListItems.Count).SubItems(6))
            End With
            PesquisarUsuario = True
         End If
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function LocalizarUsuario(ByRef padoUsuario As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean

    LocalizarUsuario = False
    Screen.MousePointer = vbHourglass
      If vop_UsuarioDados.LocalizarUsuario(padoUsuario, ptxtLocalizar, pcmbLocalizar) = True Then
         LocalizarUsuario = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function CarregarGridUsuarioRS(ByRef padoUsuario As Adodc, ByVal pcmbLocalizar As Integer) As Boolean
     
    Screen.MousePointer = vbHourglass
      CarregarGridUsuarioRS = vop_UsuarioDados.CarregarGridUsuarioRS(padoUsuario, pcmbLocalizar)
    Screen.MousePointer = vbDefault
    
End Function

Public Function IncluirUsuario() As Boolean
On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      IncluirUsuario = vop_UsuarioDados.IncluirUsuario(Nome, Sobrenome, Senha, Email, Telefone, Genero)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function AlterarUsuario() As Boolean
On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      AlterarUsuario = vop_UsuarioDados.AlterarUsuario(IdUsuario, Nome, Sobrenome, Senha, Email, Telefone, Genero)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function ExcluirUsuario() As Boolean
   Screen.MousePointer = vbHourglass
      ExcluirUsuario = vop_UsuarioDados.ExcluirUsuario(IdUsuario)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
End Function

Public Function InserirUsuarioAleatorio() As Boolean

    Screen.MousePointer = vbHourglass
      InserirUsuarioAleatorio = vop_UsuarioDados.InserirUsuarioAleatorioDE()
    Screen.MousePointer = vbDefault

End Function

Public Function ImprimirUsuarios() As Boolean

    Screen.MousePointer = vbHourglass
      ImprimirUsuarios = vop_UsuarioDados.ImprimirUsuariosDE()
    Screen.MousePointer = vbDefault

End Function













