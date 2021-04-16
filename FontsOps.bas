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
    '冀羘贾
        sndPlaySound32 "c:\Windows\Media\Alarm08.wav", &H0 '"C:\Windows\Media\Chimes.wav", &H0
'        sndPlaySound32 "C:\Program Files (x86)\Microsoft Office\Office16\MEDIA\LYNC_ringtone2.wav", &H0
'        sndPlaySound32 "C:\Program Files (x86)\Microsoft Office\Office16\MEDIA\LYNC_fsringing.wav", &H0
End Sub


Sub FontIterator()
Dim fnt 'As String
For Each fnt In Application.FontNames
    If (InStr(fnt, "刘") Or InStr(1, fnt, "li", vbTextCompare)) And InStr(1, fnt, "@", vbTextCompare) = 0 And InStr(1, fnt, "lian", vbTextCompare) = 0 And InStr(1, fnt, "Libre", vbTextCompare) = 0 And InStr(1, fnt, "Lith", vbTextCompare) = 0 And InStr(1, fnt, "Liber", vbTextCompare) = 0 And InStr(1, fnt, "light", vbTextCompare) = 0 And InStr(1, fnt, "Franklin", vbTextCompare) = 0 And InStr(1, fnt, "Italic", vbTextCompare) = 0 Then
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
fontOk = Array("夹发砰", "穝灿砰", "稬硁タ堵砰", "穝灿砰 (セゅいゅ)", "+セゅいゅ" _
                , "灿砰_HKSCS", "灿砰", "灿砰_HKSCS-ExtB", "灿砰-ExtB", _
                 "毙▅场刘", _
 _
                "64瓜", "︽", _
                "絝", "ヒ癌ゅ", "ゅ", "刘", "ゅ供刘B", "ゅ供刘DB", "ゅ供刘HKM", "ゅ供刘M", _
                "地眃︽砰", "ゅ供︽发L", "DFGGyoSho-W7", "DFPGyoSho-W7", "DFPOYoJun-W5", "DFPPenJi-W4", _
                "ゅ供肣窸B", "ゅ供︽发窸砰B", "ゅ供葵掸︽发M", _
 _
                "FangSong", "Adobe ラШ Std R", "ゅ供ラШB", "ゅ供ラШL", _
                "毙▅场夹非发", "Adobe 发蔨 Std R", "KaiTi", "ゅ供夹非发砰ProM", _
                "ゅ供肅发H", "ゅ供肅发U", "ゅ供を发B", "ゅ供を发EB", "ゅ供を发H", _
                "DFMinchoP-W5", _
                "DFGothicP-W5", _
                "DFGKanTeiRyu-W11", "ゅ供砰B", _
                "ゅ供繨ㄨ砰B", "DFKinBun-W3", _
                "DFGFuun-W7", _
 _
                "地眃︽砰(P)", "DFPFuun-W7", "DFGyoSho-W7") '地眃︽砰(P)⊿ゲ璶暗
End Sub

