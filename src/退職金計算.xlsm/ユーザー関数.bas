Attribute VB_Name = "���[�U�[�֐�"
Option Explicit

Function �ސE���v�Z(Kihon1 As Long, Kihon2 As Long, nensu As Integer, tuki As Integer, S_ritu As Single, G_ritu As Single, max_nen As Integer)
Attribute �ސE���v�Z.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim Job_kin As Long, Moto_kin As Long, JOB_nen As Single, JOB_gak As Long
    Moto_kin = Round((Kihon1 + Kihon2) * G_ritu, 0)
    If nensu >= max_nen Then
            JOB_nen = max_nen
        Else
            JOB_nen = nensu + Application.RoundDown((tuki / 12), 2)
            '�N���̌��Œ[�����o��ꍇ�͐؂�̂ĂɕύX
            '2007/1/25�В��̎w���ŕύX
    End If
    Job_kin = Round(Moto_kin * JOB_nen * S_ritu, 0)
    Moto_kin = Int(Job_kin / 1000)
    JOB_gak = Moto_kin * 1000
    If (Job_kin - JOB_gak) > 0 Then Moto_kin = Moto_kin + 1
    �ސE���v�Z = Moto_kin * 1000
End Function
'
Function ����p�Z�o(Kihon1 As Long, Kihon2 As Long, nensu As Integer, tuki As Integer, S_ritu As Single, G_ritu As Single, max_nen As Integer)
Attribute ����p�Z�o.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim Job_kin As Long, Moto_kin As Long, JOB_nen As Single
    Moto_kin = Round((Kihon1 + Kihon2) * G_ritu, 0)
    If nensu >= max_nen Then
            JOB_nen = max_nen
        Else
            JOB_nen = nensu + Application.RoundDown((tuki / 12), 2)
            '�N���̌��Œ[�����o��ꍇ�͐؂�̂ĂɕύX
            '2007/1/25�В��̎w���ŕύX
    End If
    Job_kin = Round(Moto_kin * JOB_nen * S_ritu, 0)
    ����p�Z�o = Job_kin
End Function
