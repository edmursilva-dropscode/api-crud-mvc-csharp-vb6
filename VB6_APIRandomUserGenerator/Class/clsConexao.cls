VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConexao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Declara um objeto ADO connection para usar
Public cConexao As ADODB.Connection

'String de conexao
Public Function AbrirConexao() As Boolean
   If cConexao.State = adStateClosed Then
       cConexao.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PaschoalottoDesafio;Data Source=EDMURSILVA-PC"
       'cConexao.ConnectionString = "Host=localhost;Port=5432;Pooling=true;Database=PaschoalottoDesafio;User Id=postgres;Password=admin;"
       cConexao.Open
   End If
End Function

Public Function FecharConexao() As Boolean
   If cConexao.State = adStateOpen Then
      cConexao.Close
   End If
End Function

Public Function RetornaRs(ByVal strSQL As String) As ADODB.Recordset
    'Retorna o registro do banco de dados
    Dim ADORecordset As ADODB.Recordset

    Set ADORecordset = New ADODB.Recordset
      ADORecordset.Open strSQL, cConexao, adOpenForwardOnly, adLockReadOnly
      Set RetornaRs = ADORecordset
    Set ADORecordset = Nothing
End Function

Public Sub Update(ByVal strSQL As String)
     cConexao.Execute strSQL
End Sub

Private Sub Class_Initialize()
   'instancia o objeto Ado Connection
   Set cConexao = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
    'Fechando a instancia do objeto Ado Connection
    Set cConexao = Nothing
End Sub

