Attribute VB_Name = "FontsOps"
Option Explicit
Sub removeNoFont()
Const fontname As String = "台大說文小篆"
Dim a
For Each a In ThisDocument.Characters
     If a.Font.NameFarEast <> fontname Then
        a.Delete
     End If
Next a
End Sub
