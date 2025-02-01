VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Gerar Detalhamento de Vertical"
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10080
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Variável global para armazenar o caption original do Label1
Private originalLabel1Caption As String
Private originalLabel1Italic As Boolean
Private isCancelled As Boolean

Private Sub UserForm_Activate()
    ' Armazena o caption original do Label1
    With UserForm1.Frame2.Controls("Label1")
        originalLabel1Caption = .Caption
        originalLabel1Italic = .Font.Italic
    End With
    ' Garantir que a Label1 comece sem itálico
      UserForm1.Frame2.Controls("Label1").Font.Italic = False
End Sub

Private Sub GerarDetalhamento_Click()
    ' Verifica se o caminho do arquivo Excel foi selecionado
    Dim FilePathTextBox As Object
    On Error Resume Next
        Set FilePathTextBox = UserForm1.Frame1.Controls("FilePath")
    On Error GoTo 0

    If FilePathTextBox Is Nothing Then
        MsgBox "A TextBox FilePath não foi encontrada no Frame1 do UserForm1.", vbCritical
        Exit Sub
    End If
    If FilePathTextBox.Value = "" Then
        MsgBox "Selecione um arquivo Excel antes de gerar o detalhamento.", vbExclamation
        Exit Sub
    End If

    ' Inicializa a flag de cancelamento
    isCancelled = False
    
    ' Alterar o caption do Label1 para "Processando" e itálico
    SetStatusLabel "Processando", True
    'Altera o capition da Label1 para processando
     SetStatusLabel1 "Processando... Pode demorar se houverem muitos dados."

    ' Primeiro, exporta os nomes dos arquivos para o TXT
    If Not isCancelled Then Call ExportarNomesArquivosParaTXT
    ' Em seguida, processa as placas
    If Not isCancelled Then Call ProcessarPlacas

    ' Após finalizar o processo ou cancelamento, restaurar o caption do label
    If isCancelled Then
        SetStatusLabel "Cancelado", True
         'restaura o caption original da label1 e o estado do itálico se o processo for cancelado
         SetStatusLabel1 originalLabel1Caption
         UserForm1.Frame2.Controls("Label1").Font.Italic = False
    Else
        SetStatusLabel "Aguardando...", False
         'restaura o caption original da label1 e o estado do itálico se o processo não for cancelado
         SetStatusLabel1 originalLabel1Caption
        UserForm1.Frame2.Controls("Label1").Font.Italic = False
    End If
    isCancelled = False
End Sub

Private Sub SelectExcelFile_Click()
    SelectAndShowExcelPath
End Sub

' Função para alterar o caption do Label1 no Frame2 e definir o itálico
Private Sub SetStatusLabel(statusText As String, isItalic As Boolean)
    On Error Resume Next
      With UserForm1.Frame2.Controls("Label1")
        .Caption = statusText
        .Font.Italic = isItalic
    End With
    On Error GoTo 0
    UserForm1.Repaint ' Forçar a atualização visual do userform
End Sub

' Função para alterar o caption do Label1 no Frame2
Private Sub SetStatusLabel1(statusText As String)
    On Error Resume Next
        UserForm1.Frame2.Controls("Label1").Caption = statusText
    On Error GoTo 0
    UserForm1.Repaint ' Forçar a atualização visual do userform
End Sub

Private Sub Cancel_Click()
    ' Define a flag de cancelamento como verdadeira
    isCancelled = True
      ' Alterar o caption do Label1 para "Cancelando" e itálico
    SetStatusLabel "Cancelando...", True
    'Define o capition da label2 como "Cancelando..."
    SetStatusLabel1 "Cancelando..."
    UserForm1.Frame2.Controls("Label1").Font.Italic = True
    UserForm1.Repaint
    ' Descarrega o UserForm1 ao clicar no botão cancelar
   Unload UserForm1
End Sub
