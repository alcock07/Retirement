Attribute VB_Name = "M02_Calc"
Option Explicit

Sub �ސE���v�Z()
Attribute �ސE���v�Z.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim Kihon1   As Long    ' �{��
    Dim Kihon2   As Long    ' ����
    Dim S_ritu   As Single  ' �x����(���ȓs���̏ꍇ�N���ŕς�j
    Dim G_ritu   As Single  ' ��{��(0.8)
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
    
    lngKIN1 = Round((Kihon1 + Kihon2) * G_ritu, 0) '��{��(�{��+����) �~ ��{��
    lngKIN2 = Round(Kihon1 * G_ritu, 0) '��{��(�{��) �~ ��{��
    
    If lngY >= 35 Then
        lngY = 35 '�Α��N��������𒴂��Ă��������N�����
    End If
    
    '�ސE�ྯ�
    JOB_kin = Round(lngKIN1 * lngY * S_ritu, 0) '��L��{�z�ɔN�����|���Ďx�������|����
    '1000�~�����𐮗�
    lngKIN1 = Int(JOB_kin / 1000) '�x���z��1000�Ŋ���i�����_�ȉ��؂�̂āj
    JOB_gak = lngKIN1 * 1000 '�x���z��1000���|����
    If (JOB_kin - JOB_gak) > 0 Then lngKIN1 = lngKIN1 + 1 '���������z�����̎x���z��菭�Ȃ�������1000����
    Range("Q21") = lngKIN1 * 1000
    Range("Q23") = lngKIN1 * 1000
    Range("Q26") = JOB_kin
    
    '�@�̏ꍇ���
    JOB_kin = Round(lngKIN2 * lngY * S_ritu, 0)
    lngKIN2 = Int(JOB_kin / 1000) '�x���z��1000�Ŋ���i�����_�ȉ��؂�̂āj
    JOB_gak = lngKIN2 * 1000 '�x���z��1000���|����
    If (JOB_kin - JOB_gak) > 0 Then lngKIN2 = lngKIN2 + 1 '���������z�����̎x���z��菭�Ȃ�������1000����
    Range("Q22") = lngKIN2 * 1000
    Range("Q25") = JOB_kin
    
End Sub
