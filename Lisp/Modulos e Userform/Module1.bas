Attribute VB_Name = "Module1"
Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Boolean
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As LongPtr
    lpTemplateName As String
    'if windows 2000 or greater
    pvReserved  As LongPtr
    dwReserved  As Long
    FlagsEx     As Long

End Type

Option Explicit 'Obriga a declaração de variáveis

'Variável global para verificar se o processo foi cancelado
Public isCancelled As Boolean
Sub SelectAndShowExcelPath()

    Dim FileToOpen As String
    Dim FilePathTextBox As Object
    Dim OpenFile As OPENFILENAME
    Dim Ret As Boolean
    Dim strFileFilter As String
    
    ' Verifica se o UserForm1 está carregado
    If UserForms.count = 0 Then
        MsgBox "O UserForm1 não está carregado. Carregue o formulário antes de usar este script.", vbCritical
        Exit Sub
    End If

    ' Tenta obter a TextBox FilePath do Frame1 no UserForm1
    On Error Resume Next
        Set FilePathTextBox = UserForm1.Frame1.Controls("FilePath")
    On Error GoTo 0

    If FilePathTextBox Is Nothing Then
        MsgBox "A TextBox FilePath não foi encontrada no Frame1 do UserForm1.", vbCritical
        Exit Sub
    End If

    ' Define o filtro de arquivos
    strFileFilter = "Arquivos Excel (*.xls;*.xlsx)" & Chr(0) & "*.xls;*.xlsx" & Chr(0)

    With OpenFile
      .lStructSize = Len(OpenFile)
      .hwndOwner = GetActiveWindow()
      .lpstrFilter = strFileFilter
      .nFilterIndex = 1
      .lpstrFile = String(257, 0)
      .nMaxFile = 256
      .lpstrFileTitle = String(257, 0)
      .nMaxFileTitle = 256
      .lpstrTitle = "Selecionar Arquivo Excel"
      .Flags = 0  'OFN_FILEMUSTEXIST
      .Flags = .Flags Or &H4  ' OFN_EXPLORER
      .Flags = .Flags Or &H200000    'OFN_PATHMUSTEXIST
     ' .Flags = .Flags Or &H800  ' OFN_HIDEREADONLY
     '.Flags = .Flags Or &H1000   ' OFN_CREATEPROMPT
      '.Flags = .Flags Or &H10   'OFN_ALLOWMULTISELECT
      
       
       Ret = GetOpenFileName(OpenFile)
       If Ret Then
        FileToOpen = Left(OpenFile.lpstrFile, InStr(OpenFile.lpstrFile, Chr(0)) - 1)
       Else
            FileToOpen = ""
       End If
      
    End With
    ' Verifica se o usuário selecionou um arquivo
    If FileToOpen = "" Then
        ' Se o usuário cancelou, não faz nada.
        Exit Sub
    End If

    ' Se um arquivo foi selecionado, preenche a caixa de texto com o caminho
    FilePathTextBox.Value = FileToOpen

End Sub

Sub ProcessarPlacas()
    Dim logFilePath As String
    logFilePath = "C:\Users\marco\OneDrive\Clientes\Projevias\Sinalização vertical\Lista de Placas\log_processar_placas.txt" ' Define o caminho do log

    ' Exclui o arquivo caso ele exista
    On Error Resume Next
    If Dir(logFilePath) <> "" Then
        Kill logFilePath
    End If
    On Error GoTo 0

    LogMessage "Início da Sub ProcessarPlacas"
    Debug.Print "-------------------" ' Adicionado separador no Immediate Window
    On Error GoTo ErrorHandler

    Dim FilePath As String
    Dim caminhoListaDwg As String
    Dim caminhoSaidaTxt As String
    Dim objExcel As Object, objWorkbook As Object, objWorksheet As Object
    Dim lastRow As Long, i As Long, idPlacaColumn As Long, tipoSuporteColumn As Long  ' REMOVE fileNum aqui
    Dim listaDwg As Variant, placasUnicas As Object, placaValue As Variant, tipoSuporteValue As Variant
    Dim placaFormatada As String, startTime As Double, duration As Double
    Dim arrValues As Variant, arrSuporteValues As Variant
    Dim correspondencias As Object ' Declaração da variável correspondencias
    Dim wbOpen As Boolean

    ' Configurações
    FilePath = UserForm1.Frame1.Controls("FilePath").Value
    caminhoListaDwg = "C:\Users\marco\OneDrive\Clientes\Projevias\Sinalização vertical\Lista de Placas\lista_dwg.txt"
    caminhoSaidaTxt = "C:\Users\marco\OneDrive\Clientes\Projevias\Sinalização vertical\Lista de Placas\filtro_placas.txt"

    LogMessage "  FilePath: " & FilePath
    LogMessage "  Lista DWG: " & caminhoListaDwg
    LogMessage "  Arquivo TXT de saída: " & caminhoSaidaTxt

    If FilePath = "" Then
        MsgBox "Nenhum caminho de arquivo informado.", vbExclamation
        Exit Sub
    End If

    ' Leitura da lista de DWG
    If isCancelled Then GoTo CancelProcess
    startTime = Timer
    listaDwg = LerListaDeArquivo(caminhoListaDwg)
    If IsEmpty(listaDwg) Then
        MsgBox "Erro ao ler a lista de DWG do arquivo txt!", vbExclamation
        GoTo ErrorHandler
    End If
    duration = Timer - startTime
    LogMessage "  Lista de DWG lida em " & duration & " segundos."

    ' Inicializar o Excel
    If isCancelled Then GoTo CancelProcess
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False

    ' Verificar se o arquivo já está aberto
    wbOpen = False
    On Error Resume Next
    Set objWorkbook = objExcel.Workbooks(ExtractFileName(FilePath))
    If Err.Number = 0 Then wbOpen = True
    On Error GoTo 0

    If wbOpen Then
        MsgBox "O arquivo Excel '" & ExtractFileName(FilePath) & "' já está aberto. Por favor, feche-o e tente novamente.", vbCritical
        objExcel.Quit
        Set objExcel = Nothing
        GoTo Cleanup
    End If

    Set objWorkbook = objExcel.Workbooks.Open(FilePath)
    Set objWorksheet = objWorkbook.Sheets(1)
    If isCancelled Then GoTo CancelProcess
    If objWorksheet Is Nothing Then
        MsgBox "Erro: Objeto objWorksheet não foi definido.", vbCritical
        LogMessage "  Erro: Objeto objWorksheet não foi definido."
        objWorkbook.Close SaveChanges:=False
        objExcel.Quit
        Set objWorksheet = Nothing
        Set objWorkbook = Nothing
        Set objExcel = Nothing
        GoTo Cleanup
    End If

    ' Encontrar a coluna ID_PLACA
    If isCancelled Then GoTo CancelProcess
    idPlacaColumn = GetColumnNumber(objWorksheet, "ID_PLACA")
    If idPlacaColumn = 0 Then
        MsgBox "Coluna 'ID_PLACA' não encontrada no cabeçalho!", vbCritical
        LogMessage "  Coluna 'ID_PLACA' não encontrada no cabeçalho."
        objWorkbook.Close SaveChanges:=False
        objExcel.Quit
        Set objWorksheet = Nothing
        Set objWorkbook = Nothing
        Set objExcel = Nothing
        GoTo Cleanup
    End If
    LogMessage "  Coluna ID_PLACA encontrada: " & Split(objWorksheet.Cells(1, idPlacaColumn).Address, "$")(1)

    ' Encontrar a coluna TIPO_SUPORTE
    If isCancelled Then GoTo CancelProcess
    tipoSuporteColumn = GetColumnNumber(objWorksheet, "TIPO_SUPORTE")
    If tipoSuporteColumn = 0 Then
        MsgBox "Coluna 'TIPO_SUPORTE' não encontrada no cabeçalho!", vbCritical
        LogMessage "  Coluna 'TIPO_SUPORTE' não encontrada no cabeçalho."
        objWorkbook.Close SaveChanges:=False
        objExcel.Quit
        Set objWorksheet = Nothing
        Set objWorkbook = Nothing
        Set objExcel = Nothing
        GoTo Cleanup
    End If
    LogMessage "  Coluna TIPO_SUPORTE encontrada: " & Split(objWorksheet.Cells(1, tipoSuporteColumn).Address, "$")(1)

    ' Remover espaços no inicio e fim de todos os dados da coluna ID_PLACA
    If isCancelled Then GoTo CancelProcess
    Call TrimColumn(objWorksheet, idPlacaColumn)
    ' Remover espaços no inicio e fim de todos os dados da coluna TIPO_SUPORTE
    If isCancelled Then GoTo CancelProcess
    Call TrimColumn(objWorksheet, tipoSuporteColumn)

    ' Verificação para garantir que a planilha não está vazia e que possua dados abaixo do cabeçalho nas colunas ID_PLACA e TIPO_SUPORTE
    If isCancelled Then GoTo CancelProcess
    If objExcel.WorksheetFunction.CountA(objWorksheet.Range(Split(objWorksheet.Cells(1, idPlacaColumn).Address, "$")(1) & "2:" & Split(objWorksheet.Cells(1, idPlacaColumn).Address, "$")(1) & objWorksheet.Rows.count)) = 0 Or _
        objExcel.WorksheetFunction.CountA(objWorksheet.Range(Split(objWorksheet.Cells(1, tipoSuporteColumn).Address, "$")(1) & "2:" & Split(objWorksheet.Cells(1, tipoSuporteColumn).Address, "$")(1) & objWorksheet.Rows.count)) = 0 Then
        MsgBox "A planilha está vazia ou não contém dados abaixo do cabeçalho nas colunas ID_PLACA ou TIPO_SUPORTE!", vbExclamation
        LogMessage "  A planilha está vazia ou não contém dados abaixo do cabeçalho nas colunas ID_PLACA ou TIPO_SUPORTE!"
        objWorkbook.Close SaveChanges:=False
        objExcel.Quit
        Set objWorksheet = Nothing
        Set objWorkbook = Nothing
        Set objExcel = Nothing
        GoTo Cleanup
    End If

    ' Obter a última linha com dados nas colunas ID_PLACA e TIPO_SUPORTE
    If isCancelled Then GoTo CancelProcess
    lastRow = GetLastRowWithData(objWorksheet, idPlacaColumn)
    If lastRow = 0 Then
        MsgBox "Nenhuma linha com dados encontrada na coluna 'ID_PLACA'!", vbExclamation
        LogMessage "  Nenhuma linha com dados encontrada na coluna 'ID_PLACA'!"
        objWorkbook.Close SaveChanges:=False
        objExcel.Quit
        Set objWorksheet = Nothing
        Set objWorkbook = Nothing
        Set objExcel = Nothing
        GoTo Cleanup
    End If
    Dim lastRowTipoSuporte As Long
    If isCancelled Then GoTo CancelProcess
    lastRowTipoSuporte = GetLastRowWithData(objWorksheet, tipoSuporteColumn)
    If lastRowTipoSuporte = 0 Then
        MsgBox "Nenhuma linha com dados encontrada na coluna 'TIPO_SUPORTE'!", vbExclamation
        LogMessage "  Nenhuma linha com dados encontrada na coluna 'TIPO_SUPORTE'!"
        objWorkbook.Close SaveChanges:=False
        objExcel.Quit
        Set objWorksheet = Nothing
        Set objWorkbook = Nothing
        Set objExcel = Nothing
        GoTo Cleanup
    End If
    If lastRow <> lastRowTipoSuporte Then
        MsgBox "As colunas 'ID_PLACA' e 'TIPO_SUPORTE' não possuem a mesma quantidade de linhas preenchidas com dados!", vbExclamation
        LogMessage "  As colunas 'ID_PLACA' e 'TIPO_SUPORTE' não possuem a mesma quantidade de linhas preenchidas com dados!"
        objWorkbook.Close SaveChanges:=False
        objExcel.Quit
        Set objWorksheet = Nothing
        Set objWorkbook = Nothing
        Set objExcel = Nothing
        GoTo Cleanup
    End If
    LogMessage "  Última linha com dados nas colunas 'ID_PLACA' e 'TIPO_SUPORTE' encontrada: " & lastRow

    ' Ler os dados do Excel para um array, linha por linha
    startTime = Timer
    If isCancelled Then GoTo CancelProcess
    Call ReadDataFromExcelWithSuporte(objWorksheet, idPlacaColumn, tipoSuporteColumn, lastRow, arrValues, arrSuporteValues)
    duration = Timer - startTime
    LogMessage "  Dados das colunas 'ID_PLACA' e 'TIPO_SUPORTE' lidos para o array em " & duration & " segundos."
    If IsEmpty(arrValues) Or IsEmpty(arrSuporteValues) Then
        MsgBox "Erro ao ler os dados das colunas 'ID_PLACA' ou 'TIPO_SUPORTE'!", vbCritical
        LogMessage "  Erro ao ler os dados das colunas 'ID_PLACA' ou 'TIPO_SUPORTE'!"
        objWorkbook.Close SaveChanges:=False
        objExcel.Quit
        Set objWorksheet = Nothing
        Set objWorkbook = Nothing
        Set objExcel = Nothing
        GoTo Cleanup
    End If

    ' Obter valores únicos da coluna ID_PLACA usando um dicionário
    startTime = Timer
    If isCancelled Then GoTo CancelProcess
    Set placasUnicas = CreateObject("Scripting.Dictionary")
    For i = LBound(arrValues, 1) To UBound(arrValues, 1)
         If isCancelled Then Exit For
         If Not IsEmpty(arrValues(i, 1)) Then
            Dim chaveCombinada As String
            chaveCombinada = NormalizarPlaca(CStr(arrValues(i, 1))) & " " & NormalizarSuporte(CStr(arrSuporteValues(i, 1)))
            If Not placasUnicas.Exists(chaveCombinada) Then
                placasUnicas.Add chaveCombinada, 1
                LogMessage "   Dados lidos do Excel. Linha: " & i + 1 & ", Valor: " & arrValues(i, 1) & ", Tipo Suporte: " & arrSuporteValues(i, 1) & ", Chave Combinada: " & chaveCombinada
            End If
        End If
        DoEvents
    Next i
    duration = Timer - startTime
    LogMessage "  Valores únicos da coluna 'ID_PLACA' obtidos em " & duration & " segundos."
    If placasUnicas.count = 0 Then
        MsgBox "Nenhuma linha com dados válida encontrada nas colunas 'ID_PLACA' e 'TIPO_SUPORTE'!", vbExclamation
        LogMessage "  Nenhuma linha com dados válida encontrada nas colunas 'ID_PLACA' e 'TIPO_SUPORTE'!"
        objWorkbook.Close SaveChanges:=False
        objExcel.Quit
        Set objWorksheet = Nothing
        Set objWorkbook = Nothing
        Set objExcel = Nothing
        GoTo Cleanup
    End If

    ' Criar dicionário de correspondências
    If isCancelled Then GoTo CancelProcess
    Set correspondencias = CriarDicionarioCorrespondencias(placasUnicas) ' Passar apenas as chaves
    If correspondencias Is Nothing Then
        LogMessage "  Erro: A função CriarDicionarioCorrespondencias retornou Nothing. Encerrando a sub ProcessarPlacas."
        GoTo Cleanup
    End If

    ' Abrir arquivo de saída
    Dim fileNum As Integer ' REMOVE fileNum aqui
    fileNum = FreeFile ' REMOVE fileNum aqui
    If isCancelled Then GoTo CancelProcess
    Open caminhoSaidaTxt For Output As #fileNum
    LogMessage "  Arquivo TXT de saída aberto: " & caminhoSaidaTxt

    ' Processar cada placa
    startTime = Timer
    If isCancelled Then GoTo CancelProcess
    Dim key As Variant
    For Each key In correspondencias.Keys
         If isCancelled Then Exit For
        If Not IsEmpty(key) Then
           Print #fileNum, correspondencias(key)
           LogMessage "  Placa formatada e escrita no TXT: " & correspondencias(key) & " (Placa: " & key & ")"
        End If
       DoEvents
    Next key
    duration = Timer - startTime
    LogMessage "  Dados processados e escritos em " & duration & " segundos."

CancelProcess:
    Call FinalizarProcesso
    GoTo Cleanup

Cleanup:
    ' Fechar e liberar recursos
    On Error Resume Next
    If Not objWorkbook Is Nothing Then
        objWorkbook.Close SaveChanges:=False
        Set objWorkbook = Nothing
    End If
    If Not objExcel Is Nothing Then
        objExcel.Quit
        Set objExcel = Nothing
    End If
    If Not objWorksheet Is Nothing Then
        Set objWorksheet = Nothing
    End If
    If Not correspondencias Is Nothing Then
        Set correspondencias = Nothing
    End If
    On Error GoTo 0
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    If Not isCancelled Then MsgBox "Filtro salvo em: " & caminhoSaidaTxt, vbInformation
    
    Debug.Print "-------------------" ' Adicionado separador no Immediate Window
    ' A chamada a CloseLog deve ser feita aqui
    LogMessage "Fim da Sub ProcessarPlacas"
    GoTo FinalizaSub

FinalizaSub:
    CloseLog
    Unload UserForm1
    Exit Sub

ErrorHandler:
    Debug.Print "-------------------" ' Adicionado separador no Immediate Window
    LogMessage "Erro não tratado na Sub ProcessarPlacas!"
    LogMessage "Número do Erro: " & Err.Number
    LogMessage "Descrição do Erro: " & Err.Description
    On Error Resume Next
    If Not objWorkbook Is Nothing Then
        objWorkbook.Close SaveChanges:=False
        Set objWorkbook = Nothing
    End If
    If Not objExcel Is Nothing Then
        objExcel.Quit
        Set objExcel = Nothing
    End If
    If Not objWorksheet Is Nothing Then
        Set objWorksheet = Nothing
    End If
    If Not correspondencias Is Nothing Then
        Set correspondencias = Nothing
    End If
    On Error GoTo 0
    On Error Resume Next
    If fileNum > 0 Then
        On Error Resume Next
        Close #fileNum
        On Error GoTo 0
    End If
    
    MsgBox "Ocorreu um erro não tratado. Verifique as mensagens do Debug.", vbCritical
    LogMessage "  Sub ProcessarPlacas encerrada devido a um erro."
    ' A chamada a FinalizarProcesso deve ser feita aqui mesmo em caso de erro
    Call FinalizarProcesso
End Sub

Public Sub FinalizarProcesso()
    'Verifica se a subrotina foi cancelada ou não para exibir a mensagem
    If isCancelled Then
        LogMessage "Processo cancelado pelo usuário."
    End If
    'Chama a função para fechar o log
    CloseLog
    'Descarrega o UserForm1
    Unload UserForm1
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ' O usuário clicou no X
        isCancelled = True
        Call FinalizarProcesso
    End If
End Sub

Sub ExportarNomesArquivosParaTXT()
    ' Define o caminho da pasta
    Dim caminhoPasta As String
    caminhoPasta = "C:\Users\marco\OneDrive\Clientes\Projevias\Sinalização vertical\Arquivos de Placas" ' Substitua pelo caminho real da sua pasta

    ' Define o caminho e nome do arquivo TXT
    Dim caminhoArquivoTXT As String
    caminhoArquivoTXT = "C:\Users\marco\OneDrive\Clientes\Projevias\Sinalização vertical\Lista de Placas\lista_dwg.txt" ' Substitua pelo caminho e nome desejados para o arquivo TXT

    Dim fileNum As Integer  ' REMOVE fileNum aqui

    On Error GoTo ErrorHandler
    
    ' Cria um objeto FileSystemObject
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Cria um objeto Folder
    Dim pasta As Object
    Set pasta = fso.GetFolder(caminhoPasta)

    ' Declara uma variável para armazenar os nomes dos arquivos
    Dim nomesArquivos As String
    nomesArquivos = ""

    ' Loop através dos arquivos na pasta
    Dim arquivo As Object
    For Each arquivo In pasta.Files
        If isCancelled Then GoTo CancelProcess
        nomesArquivos = nomesArquivos & arquivo.Name & vbCrLf
        DoEvents
    Next arquivo

    If Not isCancelled Then
        ' Abre o arquivo TXT para escrita
         fileNum = FreeFile  ' REMOVE fileNum aqui
         Open caminhoArquivoTXT For Output As #fileNum

        ' Escreve os nomes dos arquivos no arquivo TXT
        Print #fileNum, nomesArquivos

        ' Fecha o arquivo TXT
        Close #fileNum

       ' Exibe uma mensagem de confirmação
       MsgBox "Lista de placas exportada para: " & caminhoArquivoTXT
    End If

CancelProcess:
    Call CleanupExportarNomesArquivosParaTXT(fileNum)
    Exit Sub
    
ErrorHandler:
    LogMessage "Erro não tratado na Sub ExportarNomesArquivosParaTXT!"
    LogMessage "Número do Erro: " & Err.Number
    LogMessage "Descrição do Erro: " & Err.Description
    Call CleanupExportarNomesArquivosParaTXT(fileNum)
    MsgBox "Ocorreu um erro não tratado em ExportarNomesArquivosParaTXT. Verifique as mensagens do Debug.", vbCritical
    LogMessage "  Sub ExportarNomesArquivosParaTXT encerrada devido a um erro."

End Sub


Sub CleanupExportarNomesArquivosParaTXT(fileNum As Integer)
    On Error Resume Next
    If fileNum > 0 Then
        Close #fileNum
    End If
    On Error GoTo 0
    CloseLog
End Sub

'Função para extrair apenas o nome do arquivo com a extensão
Function ExtractFileName(fullPath As String) As String
   Dim startPos As Long
   startPos = InStrRev(fullPath, "\")
    If startPos > 0 Then
        ExtractFileName = Mid(fullPath, startPos + 1)
    Else
        ExtractFileName = fullPath
    End If
End Function

Function CriarDicionarioCorrespondencias(placasUnicas As Object) As Object
    LogMessage " Início da Função CriarDicionarioCorrespondencias"
    On Error GoTo ErrorHandler
    Dim mapDict As Object
    Set mapDict = CreateMappingDictionary() ' Obtém o dicionário de mapeamento

    Dim correspondencias As Object
    Set correspondencias = CreateObject("Scripting.Dictionary")

    Dim placaUnica As Variant
    For Each placaUnica In placasUnicas.Keys
        If mapDict.Exists(placaUnica) Then
            correspondencias(placaUnica) = mapDict(placaUnica) ' Usa atribuição em vez de Add
            LogMessage "   Correspondência definida! Chave: '" & placaUnica & "', Valor: '" & mapDict(placaUnica) & "'."
        Else
            LogMessage "   Aviso! Nenhuma correspondência encontrada para a chave: '" & placaUnica & "'. Verifique o dicionário de mapeamento!"
        End If
    Next placaUnica

    Set CriarDicionarioCorrespondencias = correspondencias
    LogMessage " Fim da Função CriarDicionarioCorrespondencias"
    Exit Function

ErrorHandler:
    LogMessage "Erro não tratado na Função CriarDicionarioCorrespondencias!"
    LogMessage "Número do Erro: " & Err.Number
    LogMessage "Descrição do Erro: " & Err.Description
    Set CriarDicionarioCorrespondencias = Nothing
    Exit Function
End Function
Function CreateMappingDictionary() As Object
    Dim mapDict As Object
    Set mapDict = CreateObject("Scripting.Dictionary")

    mapDict.Add Trim("ADVERTENCIA - COMP 0,80 X 1,00 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.50M"), "ADV_COMP_0,80X1,00_ECO_80X80_3,50M.dwg"
    mapDict.Add Trim("ADVERTENCIA - COMP 0,80 X 1,00 U SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 4.00M"), ""
    mapDict.Add Trim("ADVERTENCIA - COMP 1,20 X 1,50 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.50M"), ""
    mapDict.Add Trim("ADVERTENCIA - COMP 1,20 X 1,50 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 4.00M"), "ADV_COMP_1,20X1,50_ECO_80X80_4,00M.dwg"
    mapDict.Add Trim("ADVERTENCIA - COMP 3,00 X 1,50 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 4.00M"), "ADV_COMP_3,00X1,50_ECO_80X80_4,00M.dwg"
    mapDict.Add Trim("ADVERTENCIA - QUADRADA 0,5 U SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.50M"), ""
    mapDict.Add Trim("ADVERTENCIA - QUADRADA 0,5 U SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 5.00M"), ""
    mapDict.Add Trim("ADVERTENCIA - QUADRADA 1 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 4.00M"), "ADV_QUAD_1_ECO_80X80_4,00M.dwg"
    mapDict.Add Trim("DELINEADOR 0,50 X 0,60 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 2.50M"), "DEL_0,50X0,60_ECO_80X80_2,50M.dwg"
    mapDict.Add Trim("IDENTIFICACAO DE ROD. ESTADUAIS 0,60 X 0,76 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80. 2.50M"), ""
    mapDict.Add Trim("IDENTIFICACAO DE ROD. ESTADUAIS 0,75 X 0,95 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 2.50M"), ""
    mapDict.Add Trim("IDENTIFICACAO DE ROD. ESTADUAIS 0,76 X 0,60 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 2.50M"), ""
    mapDict.Add Trim("IDENTIFICACAO DE ROD. ESTADUAIS 0,95 X 0,75 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 2.50M"), ""
    mapDict.Add Trim("IDENTIFICACAO DE ROD. ESTADUAIS 0,95 X 0,75 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.00M"), ""
    mapDict.Add Trim("IND 1,50 X 0,50 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.00M"), ""
    mapDict.Add Trim("IND 1,50 X 0,75 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.50M"), ""
    mapDict.Add Trim("IND 1,50 X 1,00 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.50M"), ""
    mapDict.Add Trim("IND 1,50 X 1,00 SEMIPORTICO EXISTENTE"), ""
    mapDict.Add Trim("IND 1,50 X 1,50 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 4.00M"), ""
    mapDict.Add Trim("IND 2,00 X 0,50 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.50M"), ""
    mapDict.Add Trim("IND 2,00 X 1,00 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 2.50M"), "IND_2,00X1,00_ECO_80X80_2,50M.dwg"
    mapDict.Add Trim("IND 2,00 X 1,00 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.50M"), "IND_2,00X1,00_ECO_80X80_3,50M.dwg"
    mapDict.Add Trim("IND 2,00 X 1,00 SEMI-PORTICO METALICO TIPO I CONICO CONTINUO. PROJECAO DE 5.50 M - VENTO 35M/S - AREA DE EXPOSICAO ATE 2.4 M²"), ""
    mapDict.Add Trim("IND 2,00 X 1,00 PORTICO EXISTENTE"), ""
    mapDict.Add Trim("IND 2,00 X 1,00 SEMI-PORTICO METALICO TIPO II CONICO CONTINUO. PROJECAO DE 6.50 M - VENTO 35M/S - AREA DE EXPOSICAO ATE 4.5 M²"), ""
    mapDict.Add Trim("IND 2,00 X 1,00 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 4.00M"), ""
    mapDict.Add Trim("IND 2,00 X 1,50 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 4.00M"), "IND_2,00X1,50_ECO_80X80_4,00M.dwg"
    mapDict.Add Trim("IND 2,00 X 1,50 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 110X70X25X2.00. 4.00M"), "IND_2,00X1,50_GAL_110X70_4,00M.dwg"
    mapDict.Add Trim("IND 2,00 X 1,50 SEMI-PORTICO METALICO TIPO II CONICO CONTINUO. PROJECAO DE 6.50 M - VENTO 35M/S - AREA DE EXPOSICAO ATE 4.5 M²"), ""
    mapDict.Add Trim("IND 2,00 X 1,50 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.00M"), "IND_2,00X1,50_GAL_250X85_5,00M.dwg"
    mapDict.Add Trim("IND 2,00 X 2,00 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.00M"), "IND_2,00X2,00_GAL_250X85_5,00M.dwg"
    mapDict.Add Trim("IND 2,00 X 2,50 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.00M"), "IND_2,00X2,50_GAL_250X85_5,00M.dwg"
    mapDict.Add Trim("IND 2,00 X 3,00 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.00M"), "IND_2,00X3,00_GAL_250X85_5,00M.dwg"
    mapDict.Add Trim("IND 2,40 X 2,00 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.00M"), "IND_2,40X2,00_GAL_250X85_5,00M.dwg"
    mapDict.Add Trim("IND 2,50 X 1,00 SEMI-PORTICO METALICO TIPO I CONICO CONTINUO. PROJECAO DE 5.50 M - VENTO 35M/S - AREA DE EXPOSICAO ATE 2.4 M²"), ""
    mapDict.Add Trim("IND 2,50 X 1,00 PORTICO EXISTENTE"), ""
    mapDict.Add Trim("IND 2,50 X 1,00 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 110X70X25X2.00. 4.00M"), "IND_2,50X1,00_GAL_110X70_4,00M.dwg"
    mapDict.Add Trim("IND 2,50 X 1,00 SEMI-PORTICO METALICO TIPO II CONICO CONTINUO. PROJECAO DE 6.50 M - VENTO 35M/S - AREA DE EXPOSICAO ATE 4.5M²"), ""
    mapDict.Add Trim("IND 2,50 X 1,20 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 110X70X25X2.00. 4.00M"), ""
    mapDict.Add Trim("IND 2,50 X 1,50 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 110X70X25X2.00. 4.00M"), "IND_2,50X1,50_GAL_110X70_4,00M.dwg"
    mapDict.Add Trim("IND 2,50 X 1,50 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.00M"), "IND_2,50X1,50_GAL_250X85_5,00M.dwg"
    mapDict.Add Trim("IND 2,50 X 1,50 SEMI-PORTICO METALICO TIPO II CONICO CONTINUO. PROJECAO DE 6.50 M - VENTO 35M/S - AREA DE EXPOSICAO ATE 4.5M²"), ""
    mapDict.Add Trim("IND 2,50 X 2,00 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.00M"), "IND_2,50X2,00_GAL_250X85_5,00M.dwg"
    mapDict.Add Trim("IND 2,50 X 2,00 SEMI-PORTICO METALICO TIPO II CONICO CONTINUO. PROJECAO DE 6.50 M - VENTO 35M/S - AREA DE EXPOSICAO ATE 4.5 M²"), ""
    mapDict.Add Trim("IND 3,00 X 1,00 PORTICO EXISTENTE"), ""
    mapDict.Add Trim("IND 3,00 X 1,00 SEMI-PORTICO METALICO TIPO II CONICO CONTINUO. PROJECAO DE 6.50M - VENTO 35M/S - AREA DE EXPOSICAO ATE 4.5 M²"), ""
    mapDict.Add Trim("IND 3,00 X 1,00 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 110X70X25X2.00. 4.00M"), "IND_3,00X1,00_GAL_110X70_4,00M.dwg"
    mapDict.Add Trim("IND 3,00 X 1,50 SEMI-PORTICO METALICO TIPO II CONICO CONTINUO. PROJECAO DE 6.50 M - VENTO 35M/S - AREA DE EXPOSICAO ATE 4.5 M²"), ""
    mapDict.Add Trim("IND 3,00 X 1,50 SEMI-PORTICO METALICO TIPO LI CONICO CONTINUO DUPLO. VAO DE 6.00 M - VENTO 35M/S - AREA DE EXPOSICAO ATE 9.0 M²"), ""
    mapDict.Add Trim("IND 3,00 X 1,50 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 110X70X25X2.00. 3.00M"), "IND_3,00X1,50_GAL_110X70_3,00M.dwg"
    mapDict.Add Trim("IND 3,00 X 1,50 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 110X70X25X2.00. 4.00M"), "IND_3,00X1,50_GAL_110X70_4,00M.dwg"
    mapDict.Add Trim("IND 3,00 X 1,50 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.00M"), "IND_3,00X1,50_GAL_250X85_5,00M.dwg"
    mapDict.Add Trim("IND 3,00 X 2,00 SEMIPORTICO EXISTENTE"), ""
    mapDict.Add Trim("IND 3,00 X 2,00 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.00M"), "IND_3,00X2,00_GAL_250X85_5,00M.dwg"
    mapDict.Add Trim("IND 3,00 X 2,00 SEMI-PORTICO METALICO TIPO II CONICO CONTINUO. PROJECAO DE 6.50 M - VENTO 35M/S - AREA DE EXPOSICAO ATE 4.5 M²"), ""
    mapDict.Add Trim("IND 3,50 X 1,00 SEMI-PORTICO METALICO TIPO II CONICO CONTINUO. PROJECAO DE 6.50 M - VENTO 35M/S - AREA DE EXPOSICAO ATE 4.5 M²"), ""
    mapDict.Add Trim("IND 3,50 X 1,00 SEMI-PORTICO SIMPLES 6.00M P/PLACA ATE 12 M²"), ""
    mapDict.Add Trim("IND 3,50 X 1,50 SEMI-PORTICO METALICO TIPO II CONICO CONTINUO. PROJECAO DE 6.50 M - VENTO 35M/S - AREA DE EXPOSICAO ATE 4.5 M²"), ""
    mapDict.Add Trim("IND 3,50 X 1,50 SEMI-PORTICO SIMPLES 6.00M P/PLACA ATE 12 M²"), ""
    mapDict.Add Trim("IND 3,50 X 1,50 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 110X70X25X2.00. 4.00M"), "IND_3,50X1,50_GAL_110X70_4,00M.dwg"
    mapDict.Add Trim("IND 3,50 X 1,50 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.00M"), "IND_3,50X1,50_GAL_250X85_5,00M.dwg"
    mapDict.Add Trim("IND 3,50 X 2,00 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.00M"), "IND_3,50X2,00_GAL_250X85_5,00M.dwg"
    mapDict.Add Trim("IND 3,50 X 2,50 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.00M"), "IND_3,50X2,50_GAL_250X85_5,00M.dwg"
    mapDict.Add Trim("IND 4,00 X 1,50 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.00M"), ""
    mapDict.Add Trim("IND 4,00 X 1,50 SEMI-PORTICO SIMPLES 6.00M P/PLACA ATE 12 M²"), ""
    mapDict.Add Trim("IND 4,00 X 2,00 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.00M"), "IND_4,00X2,00_GAL_250X85_5,00M.dwg"
    mapDict.Add Trim("IND 4,50 X 2,00 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.00M"), ""
    mapDict.Add Trim("MARCO QUILOMETRICO 0,60 X 0,85 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 2.50M"), ""
    mapDict.Add Trim("MARCO QUILOMETRICO 0,60 X 0,85 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.00M"), "MQ_0,6X0,85_ECO_80X80_3,00M.dwg"
    mapDict.Add Trim("MP 0,30 X 0,90 SUPORTE ECOLOGICO COLAPSIVEL 55 X 55 2.50M"), "MP_0,30X0,90_ECO_55X55_2,50M.dwg"
    mapDict.Add Trim("REGULAMENTACAO - CIRCULAR 0,5 U SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.00M"), "REG_CIRC_0,5_U_ECO_80X80_3,00M.dwg"
    mapDict.Add Trim("REGULAMENTACAO - CIRCULAR 0,5 U SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.50M"), ""
    mapDict.Add Trim("REGULAMENTACAO - CIRCULAR 0,50 U + OCTOGONAL 0,25 U SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 4.00M"), ""
    mapDict.Add Trim("REGULAMENTACAO - CIRCULAR 1 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.00M"), "REG_CIRC_1_ECO_80X80_3,00M.dwg"
    mapDict.Add Trim("REGULAMENTACAO - COMP 1,20 X 1,50 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.50M"), ""
    mapDict.Add Trim("REGULAMENTACAO - COMPLEMENTAR 0,80 X 1,00 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 4.00M"), ""
    mapDict.Add Trim("REGULAMENTACAO - COMPLEMENTAR 2,00 X 2,50 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.00M"), "REG_COMP_2,00X2,50_GAL_250X85_5,00M.dwg"
    mapDict.Add Trim("REGULAMENTACAO - COMPLEMENTAR 3,50 X 3,00 SUPORTE METALICO GALVANIZADO A FOGO PERFIL C 250X85X25X2.70. 5.50M"), "REG_COMP_3,50X3,00_GAL_250X85_5,50M.dwg"
    mapDict.Add Trim("REGULAMENTACAO - OCTOGONAL 0,25 U SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.50M"), ""
    mapDict.Add Trim("REGULAMENTACAO - OCTOGONAL 0,414 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.50M"), "REG_OCT_0,414_ECO_80X80_3,50M.dwg"
    mapDict.Add Trim("REGULAMENTACAO + MP 1 + (0,30X0,90) SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.00M"), ""
    mapDict.Add Trim("REGULAMENTACAO + MP 1 + (0,30X0,90) SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.50M"), ""
    mapDict.Add Trim("REGULAMENTACAO - TRIANGULAR 0,75 U SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.50M"), ""
    mapDict.Add Trim("REGULAMENTACAO - TRIANGULAR 1 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.50M"), "REG_TRI_1_ECO_80X80_3,50M.dwg"
    mapDict.Add Trim("SERVICOS AUXILIARES 0,62 X 1,00 SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 3.50M"), "SERV_0,62X1,00_ECO_80X80_3,50M.dwg"
    mapDict.Add Trim("2 DELINEADORES 2(0,50 X 0,60) SUPORTE ECOLOGICO COLAPSIVEL 80 X 80 2.50M"), "2DEL_0,50X0,60_ECO_80X80_2,50M.dwg"
    
    Set CreateMappingDictionary = mapDict
End Function
Function NormalizarPlaca(idPlaca As String) As String
    On Error GoTo ErrorHandler
    Dim temp As String
    temp = idPlaca
    ' Remover acentos
    temp = RemoverAcentos(temp)
    'Mapeamento das abreviações
    temp = Replace(temp, "SERVIÇOS AUXILIARES", "SERV", , , vbTextCompare)
    temp = Replace(temp, "COMPOSTA", "COMP", , , vbTextCompare)
    temp = Replace(temp, "REGULAMENTAÇÃO - CIRCULAR", "REG CIRC", , , vbTextCompare)
    temp = Replace(temp, "MARCADOR DE PERIGO", "MP", , , vbTextCompare)
    temp = Replace(temp, "REGULAMENTAÇÃO - OCTOGONAL", "REG OCT", , , vbTextCompare)
    temp = Replace(temp, "PLACA INDICATIVA", "IND", , , vbTextCompare)
    temp = Replace(temp, "IDENTIFICAÇÃO DE ROD. ESTADUAIS", "", , , vbTextCompare)
    temp = Replace(temp, "MARCO QUILOMÉTRICO", "MQ", , , vbTextCompare)
    temp = Replace(temp, "REGULAMENTAÇÃO - TRIANGULAR", "REG TRI", , , vbTextCompare)
    ' Converter para MAIUSCULO
    temp = UCase(temp)
    'Remover espaços no inicio e no fim
    temp = Trim(temp)
    NormalizarPlaca = temp
    Exit Function
ErrorHandler:
    LogMessage "Erro não tratado na Função NormalizarPlaca!"
    LogMessage "Número do Erro: " & Err.Number
    LogMessage "Descrição do Erro: " & Err.Description
    NormalizarPlaca = ""
End Function

Function NormalizarSuporte(tipoSuporte As String) As String
    On Error GoTo ErrorHandler
    Dim temp As String
    temp = tipoSuporte
    ' Remover acentos
    temp = RemoverAcentos(temp)
    ' Remover caracteres especiais e converter para MAIUSCULO
    temp = Replace(temp, "SUPORTE ECOLÓGICO COLAPSÍVEL ", "ECO", , , vbTextCompare)
    temp = Replace(temp, "SUPORTE METÁLICO GALVANIZADO A FOGO PERFIL C 110X70X25X2,00MM, ", "GAL_", , , vbTextCompare)
    temp = Replace(temp, "SUPORTE METÁLICO GALVANIZADO A FOGO PERFIL C 250X85X25X2,70MM, ", "GAL_", , , vbTextCompare)
    temp = Replace(temp, "H=", "", , , vbTextCompare)
    temp = Replace(temp, "MM", "", , , vbTextCompare)
    temp = Replace(temp, ",", ".", , , vbTextCompare)
    temp = UCase(temp)
    'Remover espaços no inicio e no fim
    temp = Trim(temp)
    NormalizarSuporte = temp
    Exit Function
ErrorHandler:
    LogMessage "Erro não tratado na Função NormalizarSuporte!"
    LogMessage "Número do Erro: " & Err.Number
    LogMessage "Descrição do Erro: " & Err.Description
    NormalizarSuporte = ""
End Function

Function LerListaDeArquivo(caminhoArquivo As String) As Variant
    LogMessage " Início da Função LerListaDeArquivo. Caminho: " & caminhoArquivo
    Dim fso As Object, ts As Object, fileText As String, arrLinhas() As String, i As Long
    On Error GoTo ErrorHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(caminhoArquivo, 1) ' 1 for reading

    fileText = ts.ReadAll
    ts.Close
    LogMessage "  Arquivo TXT lido com FileSystemObject."

    arrLinhas = Split(fileText, vbLf) ' Split by new line

   'Remove linhas vazias do array
    Dim filteredLinhas() As String
    Dim count As Long
    count = 0
    For i = LBound(arrLinhas) To UBound(arrLinhas)
      If Len(Trim(arrLinhas(i))) > 0 Then
        ReDim Preserve filteredLinhas(count)
        filteredLinhas(count) = arrLinhas(i)
         count = count + 1
      End If
   Next i

   LogMessage "  Linhas vazias removidas."
   LerListaDeArquivo = filteredLinhas
   LogMessage " Fim da Função LerListaDeArquivo. Linhas encontradas: " & UBound(filteredLinhas) + 1
   Set ts = Nothing
   Set fso = Nothing
    Exit Function

ErrorHandler:
   LogMessage "Erro não tratado na Função LerListaDeArquivo!"
   LogMessage "Número do Erro: " & Err.Number
   LogMessage "Descrição do Erro: " & Err.Description
   LerListaDeArquivo = Empty
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Function

Function RemoverAcentos(texto As String) As String
    Dim comAcento As String, semAcento As String
    Dim i As Long, char As String
    On Error GoTo ErrorHandler
    comAcento = "áàãâäéèêëíìîïóòõôöúùûüçÁÀÃÂÄÉÈÊËÍÌÎÏÓÒÕÔÖÚÙÛÜÇ"
    semAcento = "aaaaaeeeeiiiiooooouuuucAAAAAEEEEIIIIOOOOOUUUUC"
    For i = 1 To Len(texto)
        char = Mid(texto, i, 1)
        Dim pos As Long
        pos = InStr(1, comAcento, char, vbBinaryCompare)
        If pos > 0 Then
            Mid(texto, i, 1) = Mid(semAcento, pos, 1)
        End If
    Next i
    RemoverAcentos = texto
    Exit Function

ErrorHandler:
    LogMessage "Erro não tratado na Função RemoverAcentos!"
    LogMessage "Número do Erro: " & Err.Number
    LogMessage "Descrição do Erro: " & Err.Description
    RemoverAcentos = texto
End Function

Function GetColumnNumber(objWorksheet As Object, columnName As String) As Long
    LogMessage "  Início da Função GetColumnNumber. Coluna: " & columnName
    On Error GoTo ErrorHandler
    Dim idPlacaColumn As Long
    Dim maxColuna As Long
    maxColuna = 26
    For idPlacaColumn = 1 To maxColuna
        If objWorksheet.Cells(1, idPlacaColumn).Value = columnName Then
            LogMessage "  Coluna " & columnName & " encontrada na coluna " & idPlacaColumn
            GetColumnNumber = idPlacaColumn
            Exit Function
        End If
    Next idPlacaColumn
    GetColumnNumber = 0
    LogMessage "  Coluna " & columnName & " não encontrada."
    Exit Function

ErrorHandler:
    LogMessage "Erro não tratado na Função GetColumnNumber!"
    LogMessage "Número do Erro: " & Err.Number
    LogMessage "Descrição do Erro: " & Err.Description
    GetColumnNumber = 0
End Function
' Function para obter a ultima linha de dados em uma coluna
Function GetLastRowWithData(objWorksheet As Object, idPlacaColumn As Long) As Long
    LogMessage " Início da Função GetLastRowWithData. Coluna: " & idPlacaColumn
    Dim lastRow As Long
    Dim usedRange As Object
    Dim i As Long

    On Error Resume Next
    'Obtém o UsedRange da planilha
    Set usedRange = objWorksheet.usedRange

    lastRow = 0
    'Percorre todas as linhas do UsedRange, a partir da segunda linha para não considerar o cabeçalho
    If usedRange.Rows.count > 1 Then
        For i = 2 To usedRange.Rows.count
            'Verifica se na coluna ID_PLACA da linha atual existe algum tipo de valor
            If Not IsEmpty(objWorksheet.Cells(i, idPlacaColumn).Value) Then
                lastRow = i
            End If
        Next i
    End If

    On Error GoTo 0
    LogMessage "  Fim da Função GetLastRowWithData. Última linha: " & lastRow
    GetLastRowWithData = lastRow
End Function
Function TrimColumn(objWorksheet As Object, column As Long)
    Dim lastRow As Long
    Dim i As Long

     'Obtem a ultima linha da coluna
    lastRow = GetLastRowWithData(objWorksheet, column)
    'Itera sobre todas as células da coluna e remove os espaços do inicio e fim
    For i = 2 To lastRow
        objWorksheet.Cells(i, column).Value = Trim(objWorksheet.Cells(i, column).Value)
    Next i
    LogMessage "  Espaços removidos das células da coluna " & Split(objWorksheet.Cells(1, column).Address, "$")(1)
End Function
Sub ReadDataFromExcelWithSuporte(objWorksheet As Object, idPlacaColumn As Long, tipoSuporteColumn As Long, lastRow As Long, ByRef arrValues As Variant, ByRef arrSuporteValues As Variant)
    On Error GoTo ErrorHandler
     Dim i As Long
    'Redimensionar os arrays
    ReDim arrValues(1 To lastRow - 1, 1 To 1)
    ReDim arrSuporteValues(1 To lastRow - 1, 1 To 1)

    'Leitura dos dados
    For i = 2 To lastRow
        arrValues(i - 1, 1) = objWorksheet.Cells(i, idPlacaColumn).Value
        arrSuporteValues(i - 1, 1) = objWorksheet.Cells(i, tipoSuporteColumn).Value
         LogMessage "   Dados lidos do Excel. Linha: " & i & ", Valor: " & arrValues(i - 1, 1) & ", Tipo Suporte: " & arrSuporteValues(i - 1, 1)
    Next i
    Exit Sub
ErrorHandler:
    LogMessage "Erro não tratado na Sub ReadDataFromExcelWithSuporte!"
    LogMessage "Número do Erro: " & Err.Number
    LogMessage "Descrição do Erro: " & Err.Description
    arrValues = Empty
    arrSuporteValues = Empty
    Exit Sub
End Sub
'Funções de Log
Public Sub LogMessage(message As String)
    Static fileNum As Integer ' Mantém o número do arquivo aberto entre chamadas
    Dim logFilePath As String
    
    logFilePath = "C:\Users\marco\OneDrive\Clientes\Projevias\Sinalização vertical\Lista de Placas\log_processar_placas.txt" ' Define o caminho do log

    ' Se o arquivo não estiver aberto, abre ele
    If fileNum = 0 Then
        fileNum = FreeFile
        On Error Resume Next ' Tratamento de erro
        Open logFilePath For Append As #fileNum
        If Err.Number <> 0 Then
            Debug.Print "Erro ao abrir arquivo de log para escrita: " & Err.Description
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0
    End If
     ' Verifica se o arquivo está aberto antes de tentar escrever
    If fileNum > 0 Then
        ' Escreve no arquivo
        On Error Resume Next ' Tratamento de erro
        Print #fileNum, Now & " - " & message
        If Err.Number <> 0 Then
            Debug.Print "Erro ao escrever no arquivo de log: " & Err.Description
            Err.Clear
        End If
       On Error GoTo 0
   End If
    ' Escreve no Immediate Window
    Debug.Print Now & " - " & message

End Sub

Public Sub CloseLog()
    Static fileNum As Integer
    Dim logFilePath As String

    logFilePath = "Z:\Projetos\Sinalização vertical\Lista de Placas\log_processar_placas.txt" ' Define o caminho do log

    On Error Resume Next ' Tratamento de erro
    If fileNum > 0 Then
        Unlock #fileNum
        Close #fileNum
        fileNum = 0
    End If

    If Err.Number <> 0 Then
        Debug.Print "Erro ao fechar o arquivo de log: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

End Sub

Sub DetalharVertical()

    ' Exibe o UserForm
    UserForm1.Show

End Sub
