Attribute VB_Name = "ユーザー関数"
Option Explicit

Function 退職金計算(Kihon1 As Long, Kihon2 As Long, nensu As Integer, tuki As Integer, S_ritu As Single, G_ritu As Single, max_nen As Integer)
Attribute 退職金計算.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim Job_kin As Long, Moto_kin As Long, JOB_nen As Single, JOB_gak As Long
    Moto_kin = Round((Kihon1 + Kihon2) * G_ritu, 0)
    If nensu >= max_nen Then
            JOB_nen = max_nen
        Else
            JOB_nen = nensu + Application.RoundDown((tuki / 12), 2)
            '年数の月で端数が出る場合は切り捨てに変更
            '2007/1/25社長の指示で変更
    End If
    Job_kin = Round(Moto_kin * JOB_nen * S_ritu, 0)
    Moto_kin = Int(Job_kin / 1000)
    JOB_gak = Moto_kin * 1000
    If (Job_kin - JOB_gak) > 0 Then Moto_kin = Moto_kin + 1
    退職金計算 = Moto_kin * 1000
End Function
'
Function 印刷用算出(Kihon1 As Long, Kihon2 As Long, nensu As Integer, tuki As Integer, S_ritu As Single, G_ritu As Single, max_nen As Integer)
Attribute 印刷用算出.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim Job_kin As Long, Moto_kin As Long, JOB_nen As Single
    Moto_kin = Round((Kihon1 + Kihon2) * G_ritu, 0)
    If nensu >= max_nen Then
            JOB_nen = max_nen
        Else
            JOB_nen = nensu + Application.RoundDown((tuki / 12), 2)
            '年数の月で端数が出る場合は切り捨てに変更
            '2007/1/25社長の指示で変更
    End If
    Job_kin = Round(Moto_kin * JOB_nen * S_ritu, 0)
    印刷用算出 = Job_kin
End Function
