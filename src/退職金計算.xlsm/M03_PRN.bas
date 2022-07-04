Attribute VB_Name = "M03_PRN"
Option Explicit

Sub Print_OK()
    
    Call ‘ŞE‹àŒvZ
    
    If MsgBox("ŠTZ•ª‚Ìˆóü‚Å‚·‚©H", vbYesNo, "‘ŞE‹àŒvZ‚Ìˆóü") = vbYes Then
        Call Print_Page
    Else
        If UCase(Sheets("Main").Range("G10").Value) = "Y" Then
            If MsgBox("ˆÔ˜J‹à‚Í‘ŞEx•¥‚¢‚Å‚·‚©H", vbYesNo, "ˆÔ˜J‹à") = vbYes Then
                Call Print_Page3
            Else
                Call Print_Page2
            End If
        Else
            Call Print_Page1
        End If
    End If
End Sub

'ŠTZ”Åˆóü
Sub Print_Page()
    ActiveSheet.PageSetup.PrintArea = "$A$61:$H$105"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End Sub

'‡@‰Á‹‹‚È‚µŒˆ’è”Åˆóü
Sub Print_Page1()
    Dim DateA As Date  '‘ŞE“ú
    Dim DateB As Date  'x‹‹“ú
    Dim lngYY As Long
    Dim lngMM As Long
    Dim lngDD As Long
    Dim lngKG As Long
    Dim lngK1 As Long
    Dim lngK2 As Long
    
    DateA = Range("C10")
    lngKG = Range("Q22")
    lngYY = CLng(Format(DateA, "yyyy"))
    lngMM = CLng(Format(DateA, "m"))
    lngDD = CLng(Format(DateA, "d"))
    '—‚Œ
    lngMM = lngMM + 1
    If lngMM = 13 Then
        lngMM = 1
        lngYY = lngYY + 1
    End If
    DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
    If lngMM = 5 Then DateB = DateB + 1 '‚±‚Ç‚à‚Ì“ú
    If Weekday(DateB) = 1 Then '“ú—j“ú
        DateB = DateB + 1
    ElseIf Weekday(DateB) = 7 Then '“y—j“ú
        DateB = DateB + 2
    End If
    Range("C139") = DateB
    If lngKG > 1000000 Then
         lngK1 = Application.RoundUp(lngKG / 3, -4)
    Else
        lngK1 = lngKG
    End If
    Range("D139") = lngK1
    
     If lngK1 = lngKG Then
        Range("C140") = ""
        Range("D140") = ""
        Range("C141") = ""
        Range("D141") = ""
     Else
        '—‚XŒ
        lngMM = lngMM + 1
        If lngMM = 13 Then
            lngMM = 1
            lngYY = lngYY + 1
        End If
        DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
        If lngMM = 5 Then DateB = DateB + 1 '‚±‚Ç‚à‚Ì“ú
        If Weekday(DateB) = 1 Then '“ú—j“ú
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '“y—j“ú
            DateB = DateB + 2
        End If
        Range("C140") = DateB
        Range("D140") = lngK1
        
        lngMM = lngMM + 1
        If lngMM = 13 Then
            lngMM = 1
            lngYY = lngYY + 1
        End If
        DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
        If lngMM = 5 Then DateB = DateB + 1 '‚±‚Ç‚à‚Ì“ú
        If Weekday(DateB) = 1 Then '“ú—j“ú
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '“y—j“ú
            DateB = DateB + 2
        End If
        Range("C141") = DateB
        Range("D141") = lngKG - (lngK1 * 2)
    End If
        
    ActiveSheet.PageSetup.PrintArea = "$A$111:$H$151"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    
End Sub

'‡A‰Á‹‹‚ ‚èŒˆ’è”Åˆóü
Sub Print_Page2()
    
    Dim DateA As Date '‘ŞE“ú
    Dim DateB As Date 'x‹‹“ú
    Dim lngYY As Long
    Dim lngMM As Long
    Dim lngDD As Long
    Dim lngKG As Long
    Dim lngK1 As Long
    Dim lngK2 As Long
    
    DateA = Range("C10")
    lngKG = Range("Q23")
    lngYY = CLng(Format(DateA, "yyyy"))
    lngMM = CLng(Format(DateA, "m"))
    lngDD = CLng(Format(DateA, "d"))
    '—‚Œ
    lngMM = lngMM + 1
    If lngMM = 13 Then
        lngMM = 1
        lngYY = lngYY + 1
    End If
    DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
    If lngMM = 5 Then DateB = DateB + 1 '‚±‚Ç‚à‚Ì“ú
    If Weekday(DateB) = 1 Then '“ú—j“ú
        DateB = DateB + 1
    ElseIf Weekday(DateB) = 7 Then '“y—j“ú
        DateB = DateB + 2
    End If
    Range("C194") = DateB
    If lngKG > 1000000 Then
         lngK1 = Application.RoundUp(lngKG / 3, -4)
    Else
        lngK1 = lngKG
    End If
    Range("D194") = lngK1
    
     If lngK1 = lngKG Then
        Range("C195") = ""
        Range("D195") = ""
        Range("C196") = ""
        Range("D196") = ""
    Else
        '—‚XŒ
        lngMM = lngMM + 1
        If lngMM = 13 Then
            lngMM = 1
            lngYY = lngYY + 1
        End If
        DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
        If lngMM = 5 Then DateB = DateB + 1 '‚±‚Ç‚à‚Ì“ú
        If Weekday(DateB) = 1 Then '“ú—j“ú
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '“y—j“ú
            DateB = DateB + 2
        End If
        Range("C195") = DateB
        Range("D195") = lngK1
        
        lngMM = lngMM + 1
        If lngMM = 13 Then
            lngMM = 1
            lngYY = lngYY + 1
        End If
        DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
        If lngMM = 5 Then DateB = DateB + 1 '‚±‚Ç‚à‚Ì“ú
        If Weekday(DateB) = 1 Then '“ú—j“ú
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '“y—j“ú
            DateB = DateB + 2
        End If
        Range("C196") = DateB
        Range("D196") = lngKG - (lngK1 * 2)
    End If
    
    ActiveSheet.PageSetup.PrintArea = "$A$161:$H$205"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    
End Sub

'‡B‰Á‹‹‚ ‚èŒˆ’è”ÅiˆÔ˜J‹à‘ŞEjˆóü
Sub Print_Page3()
    
    Dim DateA As Date '‘ŞE“ú
    Dim DateB As Date 'x‹‹“ú
    Dim lngYY As Long
    Dim lngMM As Long
    Dim lngDD As Long
    Dim lngKG As Long
    Dim lngK1 As Long
    Dim lngK2 As Long
    
    DateA = Range("C10")
    lngKG = Range("Q22")
    lngYY = CLng(Format(DateA, "yyyy"))
    lngMM = CLng(Format(DateA, "m"))
    lngDD = CLng(Format(DateA, "d"))
    
    '—‚Œ
    lngMM = lngMM + 1
    If lngMM = 13 Then
        lngMM = 1
        lngYY = lngYY + 1
    End If
    DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
    If lngMM = 5 Then DateB = DateB + 1 '‚±‚Ç‚à‚Ì“ú
    If Weekday(DateB) = 1 Then '“ú—j“ú
        DateB = DateB + 1
    ElseIf Weekday(DateB) = 7 Then '“y—j“ú
        DateB = DateB + 2
    End If
    Range("C239") = DateB
    If lngKG > 1000000 Then
         lngK1 = Application.RoundUp(lngKG / 3, -4)
    Else
        lngK1 = lngKG
    End If
    Range("D239") = lngK1
    
    If lngK1 = lngKG Then
        Range("C240") = ""
        Range("D240") = ""
        Range("C241") = ""
        Range("D241") = ""
    Else
        '—‚XŒ
        lngMM = lngMM + 1
        If lngMM = 13 Then
            lngMM = 1
            lngYY = lngYY + 1
        End If
        DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
        If lngMM = 5 Then DateB = DateB + 1 '‚±‚Ç‚à‚Ì“ú
        If Weekday(DateB) = 1 Then '“ú—j“ú
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '“y—j“ú
            DateB = DateB + 2
        End If
        Range("C240") = DateB
        Range("D240") = lngK1
        '—‚XXŒ
        lngMM = lngMM + 1
        If lngMM = 13 Then
            lngMM = 1
            lngYY = lngYY + 1
        End If
        DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
        If lngMM = 5 Then DateB = DateB + 1 '‚±‚Ç‚à‚Ì“ú
        If Weekday(DateB) = 1 Then '“ú—j“ú
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '“y—j“ú
            DateB = DateB + 2
        End If
        Range("C241") = DateB
        Range("D241") = lngKG - (lngK1 * 2)
    End If
    
    ActiveSheet.PageSetup.PrintArea = "$A$206:$H$248"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    
End Sub
