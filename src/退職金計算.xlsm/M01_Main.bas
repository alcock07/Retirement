Attribute VB_Name = "M01_Main"
Option Explicit

' �ސE���v�Z
' 2000/ 6/14  �쐬 : Shigeo ITOI
' 2006/10/27  �C�� : takazawa

Public P_maisu
Public C_name$, Lg_name$, Shain_Name$
Public Const dbS = "\\192.168.128.4\hb\kyuyo\�O���[�v����.accdb"
Public Const dbT = "\\192.168.128.4\hb\ta\���^�V�X�e��\�O���[�v����.accdb"

'���ʎg�p��..........�t�@�C���̑��݌���
Function FileExists(FileName) As Boolean
Attribute FileExists.VB_ProcData.VB_Invoke_Func = " \n14"
    FileExists = (Dir(FileName) <> "")
End Function

Sub �R�[�h����()
    
    Const SQL1 = "SELECT �Ј��R�[�h, �Ј��� FROM �O���[�v�Ј��}�X�^�[ WHERE (((���Ə��敪)='"
    Const SQL2 = "')) ORDER BY �Ј��R�[�h"

    Dim cnA As New ADODB.Connection
    Dim rsA As New ADODB.Recordset
    Dim strSQL As String
    Dim strSTN As String
    Dim strDB  As String
    Dim i      As Long
    
    Range("AC18:AD47").ClearContents
    strSTN = Sheets("�ސE���v�Z").Range("AD5").Value
    If strSTN = "TA" Or strSTN = "KA" Then
        strDB = dbT
    Else
        strDB = dbS
    End If
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strDB
    cnA.Open
    strSQL = SQL1 & strSTN & SQL2
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    i = 18
    Do Until rsA.EOF
        Cells(i, 29) = rsA(0).Value
        Cells(i, 30) = rsA(1).Value
        i = i + 1
        rsA.MoveNext
    Loop
    rsA.Close
    cnA.Close
    Set rsA = Nothing
    Set cnA = Nothing
    
End Sub

Sub �Ј��I��()

    If Range("G7").Value = "" Or Range("G7").Value = 0 Then Exit Sub
    If Range("G7").Value > 0 And Range("G7").Value <= 99999 Then
            Call �}�X�^�[�Ǎ�
        Else
            Range("X15").Value = "���o�^�ł�"
            Range("Y15").Value = ""
            Range("X16").Value = 0
            Range("X17").Value = 0
            Range("X18").Value = ""
            Range("X19").Value = ""
    End If
    
End Sub
    
Sub �}�X�^�[�Ǎ�()
Attribute �}�X�^�[�Ǎ�.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Const SQL1 = "SELECT * FROM �O���[�v�Ј��}�X�^�[ WHERE (((���Ə��敪)='"
    Const SQL2 = "') AND ((�Ј��R�[�h)='"
    Const SQL3 = "'))"

    Dim cnA As New ADODB.Connection
    Dim rsA As New ADODB.Recordset
    Dim strSQL As String
    Dim strSTN As String
    Dim strCD  As String
    Dim strDB  As String
    
    strSTN = Sheets("�ސE���v�Z").Range("AD5").Value
    If strSTN = "TA" Or strSTN = "KA" Then
        strDB = dbT
    Else
        strDB = dbS
    End If
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strDB
    cnA.Open
    
    strCD = Strings.Format(Sheets("�ސE���v�Z").Range("G7").Value, "00000")
    strSQL = SQL1 & strSTN & SQL2 & strCD & SQL3
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF Then
        Sheets("�ސE���v�Z").Range("X15").Value = "���o�^�ł�"
        Sheets("�ސE���v�Z").Range("Y15").Value = ""
        Sheets("�ސE���v�Z").Range("X16").Value = 0
        Sheets("�ސE���v�Z").Range("X17").Value = 0
        Sheets("�ސE���v�Z").Range("X18").Value = ""
        Sheets("�ސE���v�Z").Range("X19").Value = ""
    Else
        If IsNull(rsA![�Ј���]) Then Sheets("�ސE���v�Z").Range("X15").Value = "" Else Sheets("�ސE���v�Z").Range("X15").Value = rsA![�Ј���]
        If IsNull(rsA![����]) Then Sheets("�ސE���v�Z").Range("Y15").Value = "" Else Sheets("�ސE���v�Z").Range("Y15").Value = rsA![����]
        If IsNull(rsA![��{���P]) Then Sheets("�ސE���v�Z").Range("X16").Value = "" Else Sheets("�ސE���v�Z").Range("X16").Value = rsA![��{���P]
        If IsNull(rsA![��{���Q]) Then Sheets("�ސE���v�Z").Range("X17").Value = "" Else Sheets("�ސE���v�Z").Range("X17").Value = rsA![��{���Q]
        If IsNull(rsA![���N����]) Then
            Sheets("�ސE���v�Z").Range("X18").Value = ""
        Else
            Sheets("�ސE���v�Z").Range("X18").Value = rsA![���N����]
        End If
        If IsNull(rsA![���ДN����]) Then
            Sheets("�ސE���v�Z").Range("X19").Value = ""
        Else
            Sheets("�ސE���v�Z").Range("X19").Value = rsA![���ДN����]
        End If
    End If
    Shain_Name$ = Sheets("�ސE���v�Z").Range("X15").Value
    
    rsA.Close
    cnA.Close
    Set rsA = Nothing
    Set cnA = Nothing
    
End Sub

Sub Print_OK()
Attribute Print_OK.VB_ProcData.VB_Invoke_Func = " \n14"
    Call MAISU_ent
    If P_maisu = 0 Or P_maisu = "" Then Exit Sub
    If MsgBox("�T�Z���̈���ł����H", vbYesNo, "�ސE���v�Z�̈��") = vbYes Then
        Call Print_Page
    Else
        
        If UCase(Sheets("�ސE���v�Z").Range("G10").Value) = "Y" Then
            If MsgBox("�ԘJ���͑ސE���x�����ł����H", vbYesNo, "�ԘJ��") = vbYes Then
                Call Print_Page3
            Else
                Call Print_Page2
            End If
        Else
            Call Print_Page1
        End If
    End If
End Sub

Sub Print_Page()       '�T�Z�ň��
Attribute Print_Page.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveSheet.PageSetup.PrintArea = "$A$61:$H$105"
    ActiveWindow.SelectedSheets.PrintOut Copies:=P_maisu, Collate:=True
End Sub

Sub Print_Page1()      '�@�����Ȃ�����ň��
Attribute Print_Page1.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim DateA As Date '�ސE��
    Dim DateB As Date '�x����
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
    '����
    lngMM = lngMM + 1
    If lngMM = 13 Then
        lngMM = 1
        lngYY = lngYY + 1
    End If
    DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
    If lngMM = 5 Then DateB = DateB + 1 '���ǂ��̓�
    If Weekday(DateB) = 1 Then '���j��
        DateB = DateB + 1
    ElseIf Weekday(DateB) = 7 Then '�y�j��
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
        '���X��
        lngMM = lngMM + 1
        If lngMM = 13 Then
            lngMM = 1
            lngYY = lngYY + 1
        End If
        DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
        If lngMM = 5 Then DateB = DateB + 1 '���ǂ��̓�
        If Weekday(DateB) = 1 Then '���j��
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '�y�j��
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
        If lngMM = 5 Then DateB = DateB + 1 '���ǂ��̓�
        If Weekday(DateB) = 1 Then '���j��
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '�y�j��
            DateB = DateB + 2
        End If
        Range("C141") = DateB
        Range("D141") = lngKG - (lngK1 * 2)
    End If
        
    ActiveSheet.PageSetup.PrintArea = "$A$111:$H$151"
    ActiveWindow.SelectedSheets.PrintOut Copies:=P_maisu, Collate:=True
End Sub

Sub Print_Page2()      '�A�������茈��ň��
Attribute Print_Page2.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim DateA As Date '�ސE��
    Dim DateB As Date '�x����
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
    '����
    lngMM = lngMM + 1
    If lngMM = 13 Then
        lngMM = 1
        lngYY = lngYY + 1
    End If
    DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
    If lngMM = 5 Then DateB = DateB + 1 '���ǂ��̓�
    If Weekday(DateB) = 1 Then '���j��
        DateB = DateB + 1
    ElseIf Weekday(DateB) = 7 Then '�y�j��
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
        '���X��
        lngMM = lngMM + 1
        If lngMM = 13 Then
            lngMM = 1
            lngYY = lngYY + 1
        End If
        DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
        If lngMM = 5 Then DateB = DateB + 1 '���ǂ��̓�
        If Weekday(DateB) = 1 Then '���j��
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '�y�j��
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
        If lngMM = 5 Then DateB = DateB + 1 '���ǂ��̓�
        If Weekday(DateB) = 1 Then '���j��
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '�y�j��
            DateB = DateB + 2
        End If
        Range("C196") = DateB
        Range("D196") = lngKG - (lngK1 * 2)
    End If
    
    ActiveSheet.PageSetup.PrintArea = "$A$161:$H$205"
    ActiveWindow.SelectedSheets.PrintOut Copies:=P_maisu, Collate:=True
End Sub

Sub Print_Page3()      '�B�������茈��Łi�ԘJ���ސE���j���
    
    Dim DateA As Date '�ސE��
    Dim DateB As Date '�x����
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
    
    '����
    lngMM = lngMM + 1
    If lngMM = 13 Then
        lngMM = 1
        lngYY = lngYY + 1
    End If
    DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
    If lngMM = 5 Then DateB = DateB + 1 '���ǂ��̓�
    If Weekday(DateB) = 1 Then '���j��
        DateB = DateB + 1
    ElseIf Weekday(DateB) = 7 Then '�y�j��
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
        '���X��
        lngMM = lngMM + 1
        If lngMM = 13 Then
            lngMM = 1
            lngYY = lngYY + 1
        End If
        DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
        If lngMM = 5 Then DateB = DateB + 1 '���ǂ��̓�
        If Weekday(DateB) = 1 Then '���j��
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '�y�j��
            DateB = DateB + 2
        End If
        Range("C240") = DateB
        Range("D240") = lngK1
        '���X�X��
        lngMM = lngMM + 1
        If lngMM = 13 Then
            lngMM = 1
            lngYY = lngYY + 1
        End If
        DateB = CDate(Format(lngYY, "0000") & "/" & Format(lngMM, "00") & "/05")
        If lngMM = 5 Then DateB = DateB + 1 '���ǂ��̓�
        If Weekday(DateB) = 1 Then '���j��
            DateB = DateB + 1
        ElseIf Weekday(DateB) = 7 Then '�y�j��
            DateB = DateB + 2
        End If
        Range("C241") = DateB
        Range("D241") = lngKG - (lngK1 * 2)
    End If
    
    ActiveSheet.PageSetup.PrintArea = "$A$206:$H$248"
    ActiveWindow.SelectedSheets.PrintOut Copies:=P_maisu, Collate:=True
End Sub

Sub MAISU_ent()
Attribute MAISU_ent.VB_ProcData.VB_Invoke_Func = " \n14"
'    P_maisu = InputBox("�����������͂��ĉ�����", "", "1")
    P_maisu = 1
End Sub

Sub CLR_�\����()
Attribute CLR_�\����.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("G7").Value = "": Range("G10").Value = "": Range("X15").Value = ""
    Range("X16").Value = "": Range("X17").Value = "": Range("X18").Value = ""
    Range("X19").Value = "": Range("AB5").Value = "": Range("AI5").Value = ""
    Range("C10").Value = "": Range("Y15").Value = ""
    Range("G7").Select
End Sub

Sub AP_END()
'==================
' �I�������@Ver2.0
'==================

    Dim myBook As Workbook
    Dim strFN As String
    Dim boolB As Boolean
    
    Application.ReferenceStyle = xlA1
    Application.MoveAfterReturnDirection = xlDown
    Application.DisplayAlerts = False
    
    strFN = ThisWorkbook.Name '���̃u�b�N�̖��O
    boolB = False
    For Each myBook In Workbooks
        If myBook.Name <> strFN Then boolB = True
    Next
    If boolB Then
        ThisWorkbook.Close False  '�t�@�C�������
    Else
        Application.Quit  'Excell���I��
        ThisWorkbook.Saved = True
        ThisWorkbook.Close False
    End If
    
End Sub
