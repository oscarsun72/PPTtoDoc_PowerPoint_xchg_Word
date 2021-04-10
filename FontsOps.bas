Attribute VB_Name = "FontsOps"
Option Explicit
Sub removeNoFont()
Dim fontname As String
fontname = ThisDocument.Name
fontname = Mid(fontname, 1, InStr(fontname, "(") - 1)
Dim a
For Each a In ThisDocument.Characters
     If a.Font.NameFarEast = fontname And a.Font.Name = fontname Then
     Else
        a.Delete
     End If
Next a
End Sub
