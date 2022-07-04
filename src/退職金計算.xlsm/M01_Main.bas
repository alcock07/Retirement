Attribute VB_Name = "M01_Main"
Option Explicit

' �ސE���v�Z
' 2000/ 6/14  �쐬 : Shigeo ITOI
' 2006/10/27  �C�� : takazawa
' 2022/07/01  �C�� : takazawa

Public Const MYPROVIDERE = "Provider=SQLOLEDB;"
Public Const MYSERVER = "Data Source=HB14\SQLEXPRESS;"
Public Const MYSERVER9 = "Data Source=192.168.128.9\SQLEXPRESS;"
Public Const USER = "User ID=sa;"
Public Const PSWD = "Password=admin;"
Public Const PSWD9 = "Password=ALCadmin!;"


' ���Ə��敪�R���{�{�b�N�X�I�������� ===
Sub �R�[�h����()
    
    '���Ə���I���������ɎЈ������擾
    
    Dim cnA As New ADODB.Connection
    Dim rsA As New ADODB.Recordset
    Dim strSQL As String
    Dim strSTN As String
    Dim strDB  As String
    Dim i      As Long
    
    strDB = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strDB & USER & PSWD
    cnA.Open
    strSTN = Sheets("Main").Range("AD5").Value
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

' �Ј����R���{�{�b�N�X�I�������� ===
Sub �}�X�^�[�Ǎ�()
Attribute �}�X�^�[�Ǎ�.VB_ProcData.VB_Invoke_Func = " \n14"
    
    '�I�������Ј��̃f�[�^�����^�}�X�^�������Ă���
    
    Dim cnA As New ADODB.Connection
    Dim rsA As New ADODB.Recordset
    Dim strSQL As String
    Dim strSTN As String
    Dim strCD  As String
    Dim strDB  As String
    
    
    strDB = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strDB & USER & PSWD
    cnA.Open
    
    strSTN = Sheets("Main").Range("AD5").Value  '���Ə��敪�iRH,RO,RT,TA,KA�j
    strCD = Strings.Format(Sheets("Main").Range("G7").Value, "00000") '�Ј�����
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "  FROM KYUMTA"
    strSQL = strSQL & "       WHERE KBN = '" & strSTN & "'"
    strSQL = strSQL & "       AND SCODE = '" & strCD & "'"
    strSQL = strSQL & "  ORDER BY SCODE"
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF Then
        Sheets("Main").Range("X15").Value = "���o�^�ł�"  '���O
        Sheets("Main").Range("X16").Value = 0  '�{��
        Sheets("Main").Range("X17").Value = 0  '����
        Sheets("Main").Range("X18").Value = "" '���N����
        Sheets("Main").Range("X19").Value = "" '���Г�
    Else
        If IsNull(rsA![SNAME]) Then
            Sheets("Main").Range("X15").Value = ""
        Else
            Sheets("Main").Range("X15").Value = Trim(rsA![SNAME])
        End If
        If IsNull(rsA![PAY1]) Then
            Sheets("Main").Range("X16").Value = ""
        Else
            Sheets("Main").Range("X16").Value = rsA![PAY1]
        End If
        If IsNull(rsA![PAY2]) Then
            Sheets("Main").Range("X17").Value = ""
        Else
            Sheets("Main").Range("X17").Value = rsA![PAY2]
        End If
        If IsNull(rsA![DATE1]) Then
            Sheets("Main").Range("X18").Value = ""
        Else
            Sheets("Main").Range("X18").Value = Format(rsA![DATE1], "yyyy/mm/dd")
        End If
        If IsNull(rsA![DATE2]) Then
            Sheets("Main").Range("X19").Value = ""
        Else
            Sheets("Main").Range("X19").Value = Format(rsA![DATE2], "yyyy/mm/dd")
        End If
    End If
    
    Call �ސE���v�Z
    
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    
End Sub

Sub CLR_�\����()
Attribute CLR_�\����.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("G7").Value = "": Range("G10").Value = "": Range("X15").Value = ""
    Range("X16").Value = "": Range("X17").Value = "": Range("X18").Value = ""
    Range("X19").Value = "": Range("AB5").Value = "": Range("AI5").Value = ""
    Range("C10").Value = "": Range("Y15").Value = "": Range("AB16") = 1
    Range("G7").Select
End Sub

Sub Selected_Add()

    '1.���ȓs��
    '2.����
    '3.��N
    '4.�����A�C
    
    Dim lngADD As Long
    
    lngADD = Sheets("Main").Range("AI5")
    Select Case lngADD
        Case 1
            Sheets("Main").Range("G10") = "N"
        Case 2
            Sheets("Main").Range("G10") = "Y"
        Case 3
            Sheets("Main").Range("G10") = "Y"
        Case 4
            Sheets("Main").Range("G10") = "Y"
        Case Else
            Sheets("Main").Range("G10") = "N"
    End Select
    
    Call �ސE���v�Z
    
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
