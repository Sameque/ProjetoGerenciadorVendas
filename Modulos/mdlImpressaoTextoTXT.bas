Attribute VB_Name = "mdlImpressaoTextoTXT"
Option Explicit

Public Function ImprimirTexto(Texto As String, Pasta As String)

'para testar em arquivo texto

    Dim nFile As Variant
    
    nFile = FreeFile

    Open Pasta & Format(Date, "ddmmyy") & Format(Time, "hhmmss") & ".txt" For Append As #nFile

    Print #nFile, Texto

    Close #nFile

End Function
