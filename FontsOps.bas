Attribute VB_Name = "FontsOps"
Option Explicit
Sub removeNoFont()
Const fontname As String = "教育部隸書"
Dim a
For Each a In ThisDocument.Characters
     If a.Font.NameFarEast <> "教育部隸書" Then
        a.Delete
     End If
Next a
End Sub
