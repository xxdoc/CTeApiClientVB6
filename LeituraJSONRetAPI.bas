Attribute VB_Name = "LeituraJSONRetAPI"
'activate microsoft script control 1.0 in references
Public Function JSON2(sJsonString As String, Key1 As String, Key2 As String, key3 As String) As String
On Error GoTo err_handler
    Dim oScriptEngine As ScriptControl
    Set oScriptEngine = New ScriptControl
    oScriptEngine.Language = "JScript"
    Dim objJSON As Object
    Set objJSON = oScriptEngine.Eval("(" + sJsonString + ")")
    If Key1 <> "" And Key2 <> "" And key3 <> "" Then
        JSON2 = VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, Key1, VbGet), Key2, VbGet), key3, VbGet)
    ElseIf Key1 <> "" And Key2 <> "" Then
        JSON2 = VBA.CallByName(VBA.CallByName(objJSON, Key1, VbGet), Key2, VbGet)
    ElseIf Key1 <> "" Then
        JSON2 = VBA.CallByName(objJSON, Key1, VbGet)
    End If
Err_Exit:
    Exit Function
err_handler:
    JSON2 = "Error: " & Err.Description
    Resume Err_Exit
End Function
