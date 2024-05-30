Attribute VB_Name = "modFuncoes"
Option Explicit

'Variáveis do ADO
Private vol_Conexao As New clsConexao
Private vol_System As New clsSystem

'Variaveis pricate
Private vcl_WinDir As String

'Constantes usadas para acessar o Registro
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_SZ = 1
Private Const ERROR_SUCCESS = 0&

'Declaração para acessar o Registro
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

'Mensagens de Erro
Global Const ERR01 = "A memória livre disponível é insuficiênte !"
Global Const ERR02 = "O espaço livre em disco é insuficiênte !"
Global Const ERR03 = "Resolução de vídeo inválida !"
Global Const ERR04 = "Não foi possível abrir a Base de Dados !"
Global Const ERR05 = "Número de FILES no arquivo CONFIG.SYS é insuficiênte !"
Global Const ERR06 = "Relatório não encontrado !"
Global Const ERR07 = "Registro bloqueado por outro Usuário."
Global Const ERR08 = "Registro excluído por outro usuário !"
Global Const ERR09 = "Erro de Disco ou drive de Rede !"
Global Const ERR10 = "Erro ao acessar a unidade de Rede !"
Global Const ERR11 = "A base de dados está sendo usada por outro usuário !"
Global Const ERR12 = "Não foi possível encontrar o arquivo de Dicas !"
Global Const ERR13 = "Não foi possível iniciar a calculadora do Windows."
Global Const ERR14 = "Arquivo não encontrado !"
Global Const ERR15 = "Falha na transação dos dados !"
Global Const ERR16 = "Falha na exclusão !"
Global Const ERR17 = "Verifique se a unidade está protegida."

'Mensagens Padrão do Sistema
Global Const MSG01 = "Deseja realmente Sair do Sistema ?"
Global Const MSG02 = "Não há dados para o intervalo !"
Global Const MSG03 = "Utilize uma resolução de 800 x 600 pixels !"
Global Const MSG04 = "Confirma a Inclusão dos Dados ? "
Global Const MSG05 = "Confirma a Alteração dos Dados ?"
Global Const MSG06 = "Confirma a Exclusão ?"
Global Const MSG07 = "Deseja tentar novamente ?"
Global Const MSG08 = "Impressão Cancelada !"
Global Const MSG09 = " Aguarde... Iniciando Impressão"
Global Const MSG10 = "Verifique se você possui direitos para acessar esta unidade."
Global Const MSG11 = "Erro Grave com Banco de Dados !"
Global Const MSG12 = "Impossível Continuar a Operação !"
Global Const MSG13 = "Deseja reoganizar a Base de Dados ?"
Global Const MSG14 = " Aguarde... Reorganizando Arquivos."
Global Const MSG15 = "Contate o Suporte Técnico e informe o Código: "
Global Const MSG16 = " Aguarde... Gerando Relatório"
Global Const MSG17 = "Verifique se você possui direitos para acessar a unidade de rede."
Global Const MSG18 = " Aguarde... Gerando Planilha"
Global Const ERR19 = "O Registro atual está relacionado com outra record !"

'Estilos dos Botões do MsgBox
Global Const Style01 = vbCritical
Global Const Style02 = vbQuestion
Global Const Style03 = vbInformation
Global Const Style04 = vbExclamation
Global Const Style05 = vbCritical + vbMsgBoxHelpButton
Global Const Style06 = vbQuestion + vbMsgBoxHelpButton
Global Const Style07 = vbInformation + vbMsgBoxHelpButton
Global Const Style08 = vbExclamation + vbMsgBoxHelpButton
Global Const Style09 = vbCritical + vbYesNo
Global Const Style10 = vbQuestion + vbYesNo
Global Const Style11 = vbInformation + vbYesNo
Global Const Style12 = vbExclamation + vbYesNo
Global Const Style13 = vbCritical + vbYesNoCancel
Global Const Style14 = vbQuestion + vbYesNoCancel
Global Const Style15 = vbInformation + vbYesNoCancel
Global Const Style16 = vbExclamation + vbYesNoCancel
Global Const Style17 = vbCritical + vbRetryCancel
Global Const Style18 = vbQuestion + vbRetryCancel
Global Const Style19 = vbInformation + vbRetryCancel
Global Const Style20 = vbExclamation + vbRetryCancel

'Títulos de Mensagens
Global Const Title01 = "Confirme !"
Global Const Title02 = "Aviso !"
Global Const Title03 = "Atenção !"
Global Const Title04 = "Alerta !"

'Títulos de Opções de Procua
Global Const Opt01 = "Localizar por &Código"
Global Const Opt02 = "Localizar por &Descrição"

'API Random
Global Const API_URL As String = "https://randomuser.me/api/?results=20"

Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(text), wait
   Set WshShell = Nothing
End Sub

Public Function LimpaCampos(ByVal Tela As Form)
Dim vil_Contador As Integer
    
    vil_Contador = 0
    For vil_Contador = 0 To Tela.Controls.Count - 1
        If TypeOf Tela.Controls(vil_Contador) Is TextBox Then
            Tela.Controls(vil_Contador).text = Empty
            Tela.Controls(vil_Contador).ForeColor = vbBlack
        End If
        If TypeOf Tela.Controls(vil_Contador) Is Label Then
            If Tela.Controls(vil_Contador).BorderStyle = 1 Then
               Tela.Controls(vil_Contador).Caption = Empty
               Tela.Controls(vil_Contador).ForeColor = vbBlack
            End If
        End If
        If TypeOf Tela.Controls(vil_Contador) Is CheckBox Then
            Tela.Controls(vil_Contador).Value = 0
        End If
        If TypeOf Tela.Controls(vil_Contador) Is ComboBox Then
            If Tela.Controls(vil_Contador).Style <> 2 Then
                Tela.Controls(vil_Contador).text = Empty
                Tela.Controls(vil_Contador).ForeColor = vbBlack
            Else
                If Tela.Controls(vil_Contador).Tag <> Empty Then
                    Tela.Controls(vil_Contador).ForeColor = vbBlack
                End If
            End If
        End If
        'If TypeOf Tela.Controls(vil_Contador) Is DTPicker Then
        '    If Tela.Controls(vil_Contador).Format = dtpShortDate Then
        '        Tela.Controls(vil_Contador).Value = Date
        '    End If
        'End If
    Next vil_Contador

End Function

Public Function TrataErros() As Boolean
    Screen.MousePointer = vbDefault
    Select Case Err
        Case 0
        Case 5
            Exit Function
        Case 6, 7, -2147221394   'Memória Insuficiênte
            MsgBox ERR01, Style05, "Erro 01"
            End
        Case 75   'Disco Protegido
            MsgBox ERR10 & Chr(13) & ERR17, Style05, "Erro 12"
            bolRetorno = False
            Exit Function
        Case 91
            Exit Function
        Case 3006 'Banco de dados aberto em Modo Exclusivo
            MsgBox ERR14 & Chr(13) & ERR11, Style05, "Erro 11"
            End
        Case 3024  'Ocorre se o sistema não conseguir abrir o banco de dados
            MsgBox ERR04 & Chr(13) & MSG15 & Err, Style01, "Erro " & Err
            End
        Case 3026 'Espaço livre de disco insuficiênte
            MsgBox ERR02, Style05, "Erro 02"
            Exit Function
        Case 3042 'Número de Files insuficiênte
            MsgBox ERR05, Style05, "Erro 05"
            Exit Function
        Case 3043 'Erro de disco ou de Rede
            MsgBox ERR09, Style05, "Erro 09"
            Exit Function
        Case 3044 'Erro de disco ou de Rede
            MsgBox ERR04 & Chr(13) & MSG17 & Chr(13) & MSG15 & Err, Style01, "Erro " & Err
            End
        Case 3046 'Registro Bloqueado, ao Salvar
            MsgBox ERR07 & Chr(13) & MSG07, Style16, "Erro 07"
            Exit Function
        Case 3050 'Sem permissão para Ler/Gravar na unidade de disco
            MsgBox ERR10 & Chr(13) & MSG10, Style05, "Erro 10", wHelp, 1001
            bolRetorno = False
            Exit Function
        Case 3051 'Sem permissão para Ler/Gravar na unidade de disco
            MsgBox ERR10 & Chr(13) & MSG10, Style05, "Erro 10", wHelp, 1001
            bolRetorno = False
            Exit Function
        Case 3078 'Erro de disco ou de Rede
            MsgBox ERR09, Style05, "Erro 09"
            Exit Function
        Case 3158 'Registro Bloqueado, ao Salvar
            MsgBox ERR07 & Chr(13) & MSG07, Style16, "Erro 07"
            Exit Function
        Case 3167 'Registro excluído por outro Usuário
            MsgBox ERR08, Style05, "Erro 08"
            Exit Function
        Case 3186 'Registro Bloqueado, ao Salvar
            MsgBox ERR07 & Chr(13) & MSG07, Style16, "Erro 07"
            Exit Function
        Case 3187 'Registro Bloqueado, ao Ler
            MsgBox ERR07 & Chr(13) & MSG07, Style16, "Erro 07"
            Exit Function
        Case 3188 'Registro Bloqueado, ao Atualizar
            MsgBox ERR07 & Chr(13) & MSG07, Style16, "Erro 07"
            Exit Function
        Case 3265 'Itens não encontrados em uma Coleção
            Exit Function
        Case 3356 'Base de Dados aberta em modo exclusivo
            MsgBox MSG12 & Chr(13) & ERR11, Style05, Title02
            Exit Function
        Case -2147217887 'Relacionamento
            MsgBox MSG12 & Chr(13) & ERR19, Style05, Title02
            Exit Function
        Case -2147467259
            MsgBox MSG12 & Chr(13) & ERR04, Style05
            End
        Case Else 'Demais Erros
            MsgBox MSG15 & Err.Number & " !", Style05, "Erro Inesperado"
            bolRetorno = False
            Exit Function
    End Select
End Function

Sub Main()
    
    
    DoEvents
    Load frmUsuarioLista
    DoEvents
    frmUsuarioLista.Show
    
End Sub



