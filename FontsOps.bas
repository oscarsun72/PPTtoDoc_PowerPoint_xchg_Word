Attribute VB_Name = "FontsOps"
Option Explicit
Public Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim fontOk
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
ThisDocument.Save
Beep
playSound
End Sub

'https://analystcave.com/vba-status-bar-progress-bar-sounds-emails-alerts-vba/#:~:text=The%20VBA%20Status%20Bar%20is%20a%20panel%20that,Bar%20we%20need%20to%20Enable%20it%20using%20Application.DisplayStatusBar%3A
Sub playSound() 'Public Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
On Error Resume Next
    '播放聲音、音效、音樂
        sndPlaySound32 "c:\Windows\Media\Alarm08.wav", &H0 '"C:\Windows\Media\Chimes.wav", &H0
'        sndPlaySound32 "C:\Program Files (x86)\Microsoft Office\Office16\MEDIA\LYNC_ringtone2.wav", &H0
'        sndPlaySound32 "C:\Program Files (x86)\Microsoft Office\Office16\MEDIA\LYNC_fsringing.wav", &H0
End Sub


Sub FontIterator()
Dim fnt 'As String
For Each fnt In Application.FontNames
    If (InStr(fnt, "隸") Or InStr(1, fnt, "li", vbTextCompare)) And InStr(1, fnt, "@", vbTextCompare) = 0 And InStr(1, fnt, "lian", vbTextCompare) = 0 And InStr(1, fnt, "Libre", vbTextCompare) = 0 And InStr(1, fnt, "Lith", vbTextCompare) = 0 And InStr(1, fnt, "Liber", vbTextCompare) = 0 And InStr(1, fnt, "light", vbTextCompare) = 0 And InStr(1, fnt, "Franklin", vbTextCompare) = 0 And InStr(1, fnt, "Italic", vbTextCompare) = 0 Then
        ThisDocument.Range.Font.Name = fnt
        Debug.Print fnt
        Stop
    End If
Next fnt
playSound
Beep
'Dim strFont As String
'Dim intResponse As Integer
'
'For Each strFont In FontNames
' intResponse = MsgBox(Prompt:=strFont, Buttons:=vbOKCancel)
' If intResponse = vbCancel Then Exit For
'Next strFont
End Sub


Sub FontsListView()
Dim fnt 'As String
Dim fontCount As Integer, x As String, i As Integer, xp As String
fontCount = Application.FontNames.Count
x = Chr(13) & Left(ThisDocument.Paragraphs(1).Range.Text, Len(ThisDocument.Paragraphs(1).Range.Text) - 1)
For i = 2 To fontCount
    xp = xp & x
Next i
ThisDocument.Range.InsertAfter xp
i = 0
For Each fnt In Application.FontNames
    i = i + 1
    ThisDocument.Paragraphs(i).Range.Font.Name = fnt
Next fnt
Dim e
fontokList
For Each fnt In ThisDocument.Paragraphs
        For Each e In fontOk
            If e = fnt.Range.Font.NameFarEast Then fnt.Range.Delete
        Next e
Next fnt
playSound
Beep
End Sub


Sub fontokList()
fontOk = Array("標楷體", "新細明體", "微軟正黑體", "新細明體 (本文中文字型)", "+本文中文字型" _
                , "細明體_HKSCS", "細明體", "細明體_HKSCS-ExtB", "細明體-ExtB", _
                 "教育部隸書", _
 _
                "64卦圖", "行書", _
                "小篆", "甲骨文", "金文", "隸書", "文鼎隸書B", "文鼎隸書DB", "文鼎隸書HKM", "文鼎隸書M", _
                "華康行書體", "文鼎行楷L", "DFGGyoSho-W7", "DFPGyoSho-W7", "DFPOYoJun-W5", "DFPPenJi-W4", _
                "文鼎魏碑B", "文鼎行楷碑體B", "文鼎鋼筆行楷M", _
 _
                "FangSong", "Adobe 仿宋 Std R", "文鼎仿宋B", "文鼎仿宋L", _
                "教育部標準楷書", "Adobe 楷体 Std R", "KaiTi", "文鼎標準楷體ProM", _
                "文鼎顏楷H", "文鼎顏楷U", "文鼎毛楷B", "文鼎毛楷EB", "文鼎毛楷H", _
                "DFMinchoP-W5", _
                "DFGothicP-W5", _
                "DFGKanTeiRyu-W11", "文鼎古印體B", _
                "文鼎雕刻體B", "DFKinBun-W3", _
                "DFGFuun-W7", _
 _
                "華康行書體(P)", "DFPFuun-W7", "DFGyoSho-W7") '華康行書體(P)以下為沒必要做的
End Sub

