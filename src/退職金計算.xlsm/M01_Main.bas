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


' 事業所区分コンボボックス選択時処理 ===
Sub コード検索()
    
    '事業所を選択した時に社員名を取得
    
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

' 社員名コンボボックス選択時処理 ===
Sub マスター読込()
Attribute マスター読込.VB_ProcData.VB_Invoke_Func = " \n14"
    
    '選択した社員のデータを給与マスタから取ってくる
    
    Dim cnA As New ADODB.Connection
    Dim rsA As New ADODB.Recordset
    Dim strSQL As String
    Dim strSTN As String
    Dim strCD  As String
    Dim strDB  As String
    
    
    strDB = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strDB & USER & PSWD
    cnA.Open
    
    strSTN = Sheets("Main").Range("AD5").Value  '事業所区分（RH,RO,RT,TA,KA）
    strCD = Strings.Format(Sheets("Main").Range("G7").Value, "00000") '社員ｺｰﾄﾞ
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "  FROM KYUMTA"
    strSQL = strSQL & "       WHERE KBN = '" & strSTN & "'"
    strSQL = strSQL & "       AND SCODE = '" & strCD & "'"
    strSQL = strSQL & "  ORDER BY SCODE"
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF Then
        Sheets("Main").Range("X15").Value = "未登録です"  '名前
        Sheets("Main").Range("X16").Value = 0  '本給
        Sheets("Main").Range("X17").Value = 0  '加給
        Sheets("Main").Range("X18").Value = "" '生年月日
        Sheets("Main").Range("X19").Value = "" '入社日
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

Sub CLR_表示部()
Attribute CLR_表示部.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("G7").Value = "": Range("G10").Value = "": Range("X15").Value = ""
    Range("X16").Value = "": Range("X17").Value = "": Range("X18").Value = ""
    Range("X19").Value = "": Range("AB5").Value = "": Range("AI5").Value = ""
    Range("C10").Value = "": Range("Y15").Value = "": Range("AB16") = 1
    Range("G7").Select
End Sub

Sub Selected_Add()

    '1.自己都合
    '2.解雇
    '3.定年
    '4.役員就任
    
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
    
    Call 退職金計算
    
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
