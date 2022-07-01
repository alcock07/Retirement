Attribute VB_Name = "M01_Main"
Option Explicit

' 退職金計算
' 2000/ 6/14  作成 : Shigeo ITOI
' 2006/10/27  修正 : takazawa
' 2022/07/01  修正 : takazawa

Public Const MYPROVIDERE = "Provider=SQLOLEDB;"
Public Const MYSERVER = "Data Source=HB14\SQLEXPRESS;"
Public Const MYSERVER9 = "Data Source=192.168.128.9\SQLEXPRESS;"
Public Const USER = "User ID=sa;"
Public Const PSWD = "Password=admin;"
Public Const PSWD9 = "Password=ALCadmin!;"

'事業所を選択した時に社員名を取得

Sub コード検索()
    
    Dim cnA As New ADODB.Connection
    Dim rsA As New ADODB.Recordset
    Dim strSQL As String
    Dim strSTN As String
    Dim strDB  As String
    Dim i      As Long
    
    strDB = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strDB & USER & PSWD
    cnA.Open
    strSTN = Sheets("退職金計算").Range("AD5").Value
    strSQL = ""
    strSQL = strSQL & "SELECT SCODE"
    strSQL = strSQL & "     , Trim(SNAME)"
    strSQL = strSQL & "  FROM KYUMTA"
    strSQL = strSQL & "       WHERE KBN = '" & strSTN & "'"
    strSQL = strSQL & "  ORDER BY SCODE"
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    i = 19
    If rsA.EOF = False Then
        Range("AC19:AD65").ClearContents
        Range("AB16") = 1
        rsA.MoveFirst
    End If
    Do Until rsA.EOF
        Cells(i, 29) = rsA(0).Value
        Cells(i, 30) = rsA(1).Value
        i = i + 1
        rsA.MoveNext
    Loop
    
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    
End Sub
    
Sub マスター読込()
Attribute マスター読込.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim cnA As New ADODB.Connection
    Dim rsA As New ADODB.Recordset
    Dim strSQL As String
    Dim strSTN As String
    Dim strCD  As String
    Dim strDB  As String
    
    
    strDB = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strDB & USER & PSWD
    cnA.Open
    
    strSTN = Sheets("退職金計算").Range("AD5").Value  '事業所区分（RH,RO,RT,TA,KA）
    strCD = Strings.Format(Sheets("退職金計算").Range("G7").Value, "00000") '社員ｺｰﾄﾞ
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "  FROM KYUMTA"
    strSQL = strSQL & "       WHERE KBN = '" & strSTN & "'"
    strSQL = strSQL & "       AND SCODE = '" & strCD & "'"
    strSQL = strSQL & "  ORDER BY SCODE"
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF Then
        Sheets("退職金計算").Range("X15").Value = "未登録です"  '名前
        Sheets("退職金計算").Range("X16").Value = 0  '本給
        Sheets("退職金計算").Range("X17").Value = 0  '加給
        Sheets("退職金計算").Range("X18").Value = "" '生年月日
        Sheets("退職金計算").Range("X19").Value = "" '入社日
    Else
        If IsNull(rsA![SNAME]) Then
            Sheets("退職金計算").Range("X15").Value = ""
        Else
            Sheets("退職金計算").Range("X15").Value = Trim(rsA![SNAME])
        End If
        If IsNull(rsA![PAY1]) Then
            Sheets("退職金計算").Range("X16").Value = ""
        Else
            Sheets("退職金計算").Range("X16").Value = rsA![PAY1]
        End If
        If IsNull(rsA![PAY2]) Then
            Sheets("退職金計算").Range("X17").Value = ""
        Else
            Sheets("退職金計算").Range("X17").Value = rsA![PAY2]
        End If
        If IsNull(rsA![DATE1]) Then
            Sheets("退職金計算").Range("X18").Value = ""
        Else
            Sheets("退職金計算").Range("X18").Value = Format(rsA![DATE1], "yyyy/mm/dd")
        End If
        If IsNull(rsA![DATE2]) Then
            Sheets("退職金計算").Range("X19").Value = ""
        Else
            Sheets("退職金計算").Range("X19").Value = Format(rsA![DATE2], "yyyy/mm/dd")
        End If
    End If
    
    Call 退職金計算
    
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    
End Sub

Sub Print_OK()
Attribute Print_OK.VB_ProcData.VB_Invoke_Func = " \n14"
    
    If MsgBox("概算分の印刷ですか？", vbYesNo, "退職金計算の印刷") = vbYes Then
        Call Print_Page
    Else
        If UCase(Sheets("退職金計算").Range("G10").Value) = "Y" Then
            If MsgBox("慰労金は退職時支払いですか？", vbYesNo, "慰労金") = vbYes Then
                Call Print_Page3
            Else
                Call Print_Page2
            End If
        Else
            Call Print_Page1
        End If
    End If
End Sub

'概算版印刷
Sub Print_Page()
Attribute Print_Page.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveSheet.PageSetup.PrintArea = "$A$61:$H$105"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End Sub

'�@加給なし決定版印刷
Sub Print_Page1()
Attribute Print_Page1.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim DateA As Date  '退職日
    Dim DateB As Date  '支給日
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
    '翌月
    lngMM = lngMM + 1
    If lngMM = 13 Then
        lngMM = 1
        lngYY = lngYY + 1
    End If
    DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
    If lngMM = 5 Then DateB = DateB + 1 'こどもの日
    If Weekday(DateB) = 1 Then '日曜日
        DateB = DateB + 1
    ElseIf Weekday(DateB) = 7 Then '土曜日
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
        '翌々月
        lngMM = lngMM + 1
        If lngMM = 13 Then
            lngMM = 1
            lngYY = lngYY + 1
        End If
        DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
        If lngMM = 5 Then DateB = DateB + 1 'こどもの日
        If Weekday(DateB) = 1 Then '日曜日
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '土曜日
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
        If lngMM = 5 Then DateB = DateB + 1 'こどもの日
        If Weekday(DateB) = 1 Then '日曜日
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '土曜日
            DateB = DateB + 2
        End If
        Range("C141") = DateB
        Range("D141") = lngKG - (lngK1 * 2)
    End If
        
    ActiveSheet.PageSetup.PrintArea = "$A$111:$H$151"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End Sub

'�A加給あり決定版印刷
Sub Print_Page2()
Attribute Print_Page2.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim DateA As Date '退職日
    Dim DateB As Date '支給日
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
    '翌月
    lngMM = lngMM + 1
    If lngMM = 13 Then
        lngMM = 1
        lngYY = lngYY + 1
    End If
    DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
    If lngMM = 5 Then DateB = DateB + 1 'こどもの日
    If Weekday(DateB) = 1 Then '日曜日
        DateB = DateB + 1
    ElseIf Weekday(DateB) = 7 Then '土曜日
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
        '翌々月
        lngMM = lngMM + 1
        If lngMM = 13 Then
            lngMM = 1
            lngYY = lngYY + 1
        End If
        DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
        If lngMM = 5 Then DateB = DateB + 1 'こどもの日
        If Weekday(DateB) = 1 Then '日曜日
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '土曜日
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
        If lngMM = 5 Then DateB = DateB + 1 'こどもの日
        If Weekday(DateB) = 1 Then '日曜日
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '土曜日
            DateB = DateB + 2
        End If
        Range("C196") = DateB
        Range("D196") = lngKG - (lngK1 * 2)
    End If
    
    ActiveSheet.PageSetup.PrintArea = "$A$161:$H$205"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End Sub

'�B加給あり決定版（慰労金退職時）印刷
Sub Print_Page3()
    
    Dim DateA As Date '退職日
    Dim DateB As Date '支給日
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
    
    '翌月
    lngMM = lngMM + 1
    If lngMM = 13 Then
        lngMM = 1
        lngYY = lngYY + 1
    End If
    DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
    If lngMM = 5 Then DateB = DateB + 1 'こどもの日
    If Weekday(DateB) = 1 Then '日曜日
        DateB = DateB + 1
    ElseIf Weekday(DateB) = 7 Then '土曜日
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
        '翌々月
        lngMM = lngMM + 1
        If lngMM = 13 Then
            lngMM = 1
            lngYY = lngYY + 1
        End If
        DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
        If lngMM = 5 Then DateB = DateB + 1 'こどもの日
        If Weekday(DateB) = 1 Then '日曜日
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '土曜日
            DateB = DateB + 2
        End If
        Range("C240") = DateB
        Range("D240") = lngK1
        '翌々々月
        lngMM = lngMM + 1
        If lngMM = 13 Then
            lngMM = 1
            lngYY = lngYY + 1
        End If
        DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
        If lngMM = 5 Then DateB = DateB + 1 'こどもの日
        If Weekday(DateB) = 1 Then '日曜日
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '土曜日
            DateB = DateB + 2
        End If
        Range("C241") = DateB
        Range("D241") = lngKG - (lngK1 * 2)
    End If
    
    ActiveSheet.PageSetup.PrintArea = "$A$206:$H$248"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
End Sub

Sub CLR_表示部()
Attribute CLR_表示部.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("G7").Value = "": Range("G10").Value = "": Range("X15").Value = ""
    Range("X16").Value = "": Range("X17").Value = "": Range("X18").Value = ""
    Range("X19").Value = "": Range("AB5").Value = "": Range("AI5").Value = ""
    Range("C10").Value = "": Range("Y15").Value = "": Range("AB16") = 1
    Range("G7").Select
End Sub

Sub AP_END()
'==================
' 終了処理　Ver2.0
'==================

    Dim myBook As Workbook
    Dim strFN As String
    Dim boolB As Boolean
    
    Application.ReferenceStyle = xlA1
    Application.MoveAfterReturnDirection = xlDown
    Application.DisplayAlerts = False
    
    strFN = ThisWorkbook.Name 'このブックの名前
    boolB = False
    For Each myBook In Workbooks
        If myBook.Name <> strFN Then boolB = True
    Next
    If boolB Then
        ThisWorkbook.Close False  'ファイルを閉じる
    Else
        Application.Quit  'Excellを終了
        ThisWorkbook.Saved = True
        ThisWorkbook.Close False
    End If
    
End Sub
