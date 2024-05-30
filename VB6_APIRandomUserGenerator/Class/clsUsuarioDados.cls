VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsuarioDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Variáveis do ADO
Private vol_Conexao As New clsConexao
Private vol_Command As New ADODB.Command



Public Property Set Conexao(ByVal vData As ADODB.Connection)
    Set cConexao = vData
End Property


'Metodos
Public Function CarregarGridUsuarioRS(ByRef padoUsuario As Adodc, ByVal pcmbLocalizar As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim i As Integer
On Error GoTo TrataErros
     
   CarregarGridUsuarioRS = True
   vol_Conexao.AbrirConexao
      padoUsuario.RecordSource = "SELECT IdUsuario, Nome, Sobrenome, Senha, Email, Telefone, Genero FROM Usuario (NOLOCK) ORDER BY " + IIf(pcmbLocalizar <= 0, "IdUsuario", IIf(pcmbLocalizar = 1, "Nome", "Sobrenome")) + " "
      padoUsuario.Refresh
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        CarregarGridUsuarioRS = False
        MsgBox "Não foi possível carregar informações do Usuario !", vbExclamation
    End If
End Function

Public Function PesquisarUsuario(ByRef plvwUsuario As ListView, ByVal pIdUsuario As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim i As Integer
On Error GoTo TrataErros
     
   PesquisarUsuario = True
   vol_Conexao.AbrirConexao
      Set vol_Recordset = vol_Conexao.RetornaRs("SELECT IdUsuario, Nome, Sobrenome, Senha, Email, Telefone, Genero FROM Usuario WHERE IdUsuario = " & pIdUsuario & " ")
      With vol_Recordset
         If .RecordCount <> 0 Then
            Do Until .EOF
               plvwUsuario.ListItems.Add , , .Fields(0).Value
               plvwUsuario.ListItems(plvwUsuario.ListItems.Count).SubItems(1) = .Fields(1).Value
               plvwUsuario.ListItems(plvwUsuario.ListItems.Count).SubItems(2) = .Fields(2).Value
               plvwUsuario.ListItems(plvwUsuario.ListItems.Count).SubItems(3) = .Fields(3).Value
               plvwUsuario.ListItems(plvwUsuario.ListItems.Count).SubItems(4) = .Fields(4).Value
               plvwUsuario.ListItems(plvwUsuario.ListItems.Count).SubItems(5) = .Fields(5).Value
               plvwUsuario.ListItems(plvwUsuario.ListItems.Count).SubItems(6) = .Fields(6).Value
               .MoveNext
            Loop
         End If
      End With
      Set vol_Recordset = Nothing
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        PesquisarUsuario = False
        MsgBox "Não foi possível encontrar informações do Usuario !", vbExclamation
    End If
End Function

Public Function LocalizarUsuario(ByRef padoUsuario As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim i As Integer
On Error GoTo TrataErros
     
   LocalizarUsuario = True
   vol_Conexao.AbrirConexao
      If Trim$(ptxtLocalizar) = Empty Then
         padoUsuario.RecordSource = "SELECT IdUsuario, Nome, Sobrenome, Senha, Email, Telefone, Genero FROM Usuario (NOLOCK) ORDER BY " + IIf(pcmbLocalizar < 0, "IdUsuario", IIf(pcmbLocalizar = 0, "Nome", "Sobrenome")) + " "
      Else
         padoUsuario.RecordSource = "SELECT IdUsuario, Nome, Sobrenome, Senha, Email, Telefone, Genero FROM Usuario (NOLOCK) WHERE " + IIf(pcmbLocalizar < 0, "IdUsuario", IIf(pcmbLocalizar = 0, "Nome", "Sobrenome")) + " LIKE '%" + Trim$(ptxtLocalizar) + "%' ORDER BY " + IIf(pcmbLocalizar = 0, "Nome", "Sobrenome") + " "
      End If
      padoUsuario.Refresh
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        LocalizarUsuario = False
        MsgBox "Não foi possível encontrar informações do Usuario !", vbExclamation
    End If
End Function

Public Function IncluirUsuario(ByVal pNome As String, _
                                ByVal pSobrenome As String, _
                                ByVal pSenha As String, _
                                ByVal pEmail As String, _
                                ByVal pTelefone As String, _
                                ByVal pGenero As String) As Boolean
Dim Incluir As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
         
         IncluirUsuario = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
   
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_I_Usuario"
         vol_Command.Parameters.Refresh
      
         With vol_Command.Parameters
           .Item(1).Value = pNome
           .Item(2).Value = pSobrenome
           .Item(3).Value = pSenha
           .Item(4).Value = pEmail
           .Item(5).Value = pTelefone
           .Item(6).Value = pGenero
         End With
   
         vol_Command.Execute
         
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       IncluirUsuario = False
       MsgBox "Não foi possível incluir informações do Usuario !", vbExclamation
    End If
End Function

Public Function AlterarUsuario(ByVal pIdUsuario As Integer, _
                                ByVal pNome As String, _
                                ByVal pSobrenome As String, _
                                ByVal pSenha As String, _
                                ByVal pEmail As String, _
                                ByVal pTelefone As String, _
                                ByVal pGenero As String) As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
      
         AlterarUsuario = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_U_Usuario"
         vol_Command.Parameters.Refresh
      
         With vol_Command.Parameters
           .Item(1).Value = pIdUsuario
           .Item(2).Value = pNome
           .Item(3).Value = pSobrenome
           .Item(4).Value = pSenha
           .Item(5).Value = pEmail
           .Item(6).Value = pTelefone
           .Item(7).Value = pGenero
         End With
      
         vol_Command.Execute
      
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       AlterarUsuario = False
       MsgBox "Não foi possível alterar informações do Usuario !", vbExclamation
    End If
End Function

Public Function ExcluirUsuario(ByVal pIdUsuario As Integer) As Boolean
Dim Incluir As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
      
         ExcluirUsuario = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "SP_Teste_D_Usuario"
         vol_Command.Parameters.Refresh
      
         With vol_Command.Parameters
           .Item(1).Value = pIdUsuario
         End With
      
         vol_Command.Execute
   
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       ExcluirUsuario = False
       MsgBox "Não foi possível excluir informações do Usuario !", vbExclamation
    End If
End Function

Public Function InserirUsuarioAleatorioDE() As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim xmlhttp As Object
Dim jsonResponse As String
Dim parsedResponse As Object
Dim novoUsuario As Object
Dim conn As Object
Dim strSQL As String
Dim i As Integer

On Error GoTo TrataErros
     
   InserirUsuarioAleatorioDE = True
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
       
          ' Criar um objeto XMLHTTP para a requisição HTTP
          Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
           
          ' Fazer a requisição GET para a API Random User Generator para obter 20 usuários
          xmlhttp.Open "GET", "https://randomuser.me/api/?results=20", False
          xmlhttp.send
       
          ' Capturar a resposta da API
          jsonResponse = xmlhttp.responseText
           
          ' Parsear o JSON de resposta
          Set parsedResponse = JsonConverter.ParseJson(jsonResponse)
                        
          For i = 1 To 20
              ' Extrair o usuário gerado do JSON
              Set novoUsuario = parsedResponse("results")(i)
               
              ' Dados do novo usuário
              Dim pNome As String
              Dim pSobrenome As String
              Dim pSenha As String
              Dim pEmail As String
              Dim pTelefone As String
              Dim pGenero As String
      
              pNome = novoUsuario("name")("first")
              pSobrenome = novoUsuario("name")("last")
              pSenha = novoUsuario("login")("password")
              pEmail = novoUsuario("email")
              pTelefone = novoUsuario("phone")
              pGenero = novoUsuario("gender")
               
              Set vol_Command = New ADODB.Command
              Set vol_Command.ActiveConnection = vol_Conexao.cConexao
          
              vol_Command.CommandType = adCmdStoredProc
              vol_Command.CommandText = "SP_Teste_I_Usuario"
              vol_Command.Parameters.Refresh
             
              With vol_Command.Parameters
                 .Item(1).Value = pNome
                 .Item(2).Value = pSobrenome
                 .Item(3).Value = pSenha
                 .Item(4).Value = pEmail
                 .Item(5).Value = pTelefone
                 .Item(6).Value = pGenero
              End With
          
              vol_Command.Execute
           
          Next i
       
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        InserirUsuarioAleatorioDE = False
        MsgBox "Não foi possível carregar informações de Usuários !", vbExclamation
    End If
End Function

Public Function ImprimirUsuariosDE() As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim i As Integer
On Error GoTo TrataErros
     
   ImprimirUsuariosDE = True
   vol_Conexao.AbrirConexao
      'Usuarios
      If dteSistemaUsuario.rscmdUsuario.State = adStateOpen Then
         dteSistemaUsuario.rscmdUsuario.Close
      End If
      dteSistemaUsuario.rscmdUsuario.Open "SELECT IdUsuario, Nome, Sobrenome, Senha, Email, Telefone, Genero FROM Usuario (NOLOCK) ORDER BY IdUsuario  "
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        ImprimirUsuariosDE = False
        MsgBox "Não foi possível carregar informações de Usuários !", vbExclamation
    End If
End Function




























