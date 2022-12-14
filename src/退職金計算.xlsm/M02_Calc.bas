Attribute VB_Name = "M02_Calc"
Option Explicit

Sub 退職金計算()
Attribute 退職金計算.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim Kihon1   As Long    ' 本給
    Dim Kihon2   As Long    ' 加給
    Dim S_ritu   As Single  ' 支給率(自己都合の場合年数で変る）
    Dim G_ritu   As Single  ' 基本率(0.8)
    Dim lngKIN1  As Long
    Dim lngKIN2  As Long
    Dim JOB_kin  As Long
    Dim lngY     As Single
    Dim JOB_gak  As Long
    
    Kihon1 = Range("X16")
    Kihon2 = Range("X17")
    S_ritu = Range("U18")
    G_ritu = Range("AE5")
    lngY = Range("U19")
    
    lngKIN1 = Round((Kihon1 + Kihon2) * G_ritu, 0) '基本給(本給+加給) × 基本率
    lngKIN2 = Round(Kihon1 * G_ritu, 0) '基本給(本給) × 基本率
    
    If lngY >= 35 Then
        lngY = 35 '勤続年数が上限を超えていたら上限年数をｾｯﾄ
    End If
    
    '退職金ｾｯﾄ
    JOB_kin = Round(lngKIN1 * lngY * S_ritu, 0) '上記基本額に年数を掛けて支給率を掛ける
    '1000円未満を整理
    lngKIN1 = Int(JOB_kin / 1000) '支給額を1000で割る（小数点以下切り捨て）
    JOB_gak = lngKIN1 * 1000 '支給額に1000を掛ける
    If (JOB_kin - JOB_gak) > 0 Then lngKIN1 = lngKIN1 + 1 '整理した額が元の支給額より少なかったら1000足す
    Range("Q21") = lngKIN1 * 1000
    Range("Q23") = lngKIN1 * 1000
    Range("Q26") = JOB_kin
    
    '?@の場合ｾｯﾄ
    JOB_kin = Round(lngKIN2 * lngY * S_ritu, 0)
    lngKIN2 = Int(JOB_kin / 1000) '支給額を1000で割る（小数点以下切り捨て）
    JOB_gak = lngKIN2 * 1000 '支給額に1000を掛ける
    If (JOB_kin - JOB_gak) > 0 Then lngKIN2 = lngKIN2 + 1 '整理した額が元の支給額より少なかったら1000足す
    Range("Q22") = lngKIN2 * 1000
    Range("Q25") = JOB_kin
    
End Sub
