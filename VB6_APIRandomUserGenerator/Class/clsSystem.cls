VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Variáveis do ADO
Private vol_Conexao As New clsConexao

'Declaração para acessar o Registro
Private mvarColuna_01 As String
Private mvarColuna_02 As String
Private mvarColuna_03 As String
Private mvarFormataColuna_01 As String



Public Property Let FormataColuna_01(ByVal vData As String)
    mvarFormataColuna_01 = vData
End Property

Public Property Get FormataColuna_01() As String
    FormataColuna_01 = mvarFormataColuna_01
End Property

Public Property Let Coluna_01(ByVal vData As String)
    mvarColuna_01 = vData
End Property

Public Property Get Coluna_01() As String
    Coluna_01 = mvarColuna_01
End Property

Public Property Let Coluna_02(ByVal vData As String)
    mvarColuna_02 = vData
End Property

Public Property Get Coluna_02() As String
    Coluna_02 = mvarColuna_02
End Property

Public Property Let Coluna_03(ByVal vData As String)
    mvarColuna_03 = vData
End Property

Public Property Get Coluna_03() As String
    Coluna_03 = mvarColuna_03
End Property

