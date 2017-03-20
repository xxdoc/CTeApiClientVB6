Attribute VB_Name = "Geral"
Public cte As String
Public Const titleCTeAPI = "NS API - CTe"
Function lerJson(txt As Boolean)
On Error GoTo SAI
    cte = ""
    Select Case txt
    Case False
        Open App.Path & "\json\rodo.json" For Input As #1
            cte = input(FileLen(App.Path & "\json\rodo.json"), #1)
        Close #1
    Case True
        Open App.Path & "\txt\cte.txt" For Input As #1
            cte = input(FileLen(App.Path & "\txt\cte.txt"), #1)
        Close #1
    End Select
    Exit Function
SAI:
    MsgBox ("Problemas ao Ler o Arquivo JSON" & vbNewLine & Err.Description), vbInformation, titleCTeAPI
End Function
Public Function Salvar_Arquivo(fileName As String, conteudo As String) As Boolean
    Open fileName For Output As #1
        Print #1, conteudo
    Close #1
End Function
Function lerArquivo(fileName As String) As String
On Error GoTo SAI
    Open fileName For Input As #1
        lerArquivo = input(FileLen(fileName), #1)
    Close #1
    Exit Function
SAI:
    MsgBox ("Problemas ao Ler o Arquivo JSON" & vbNewLine & Err.Description), vbInformation, titleCTeAPI
End Function
