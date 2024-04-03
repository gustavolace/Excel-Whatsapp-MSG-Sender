Attribute VB_Name = "Módulo11"
Sub ExecutarExe()
    Dim caminhoExe As String
    Dim caminhoTxt As String
    Dim objFSO As Object
    Dim objArquivo As Object
    Dim conteudoArquivo As String
    Dim objShell As Object

    ' Configurar objeto FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' Mudar para o diret�rio da pasta de trabalho
    ChDir ThisWorkbook.Path

    ' Especificar caminhos
    caminhoExe = "main_headless.exe"
    caminhoTxt = "logs/log.log"
    
    MsgBox "Iniciando o processo...", vbInformation, "Inicio do Processo"

    ' Configurar objeto Shell
    Set objShell = CreateObject("WScript.Shell")

    ' Executar o arquivo .exe e esperar at� que o processo seja conclu�do
    objShell.Run caminhoExe, vbHide, True

    ' Ler o conte�do do arquivo de log
    Set objArquivo = objFSO.OpenTextFile(caminhoTxt, 1) ' 1 indica modo de leitura
    conteudoArquivo = objArquivo.ReadAll
    objArquivo.Close

    ' Exibir o conte�do em um MsgBox
    MsgBox conteudoArquivo, vbInformation, "operacao concluida"
End Sub


