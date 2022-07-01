Attribute VB_Name = "M02_Calc"
Option Explicit

Sub ‘ŞE‹àŒvZ()
Attribute ‘ŞE‹àŒvZ.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim Kihon1   As Long    ' –{‹‹
    Dim Kihon2   As Long    ' ‰Á‹‹
    Dim S_ritu   As Single  ' x‹‹—¦(©ŒÈ“s‡‚Ìê‡”N”‚Å•Ï‚éj
    Dim G_ritu   As Single  ' Šî–{—¦(0.8)
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
    
    lngKIN1 = Round((Kihon1 + Kihon2) * G_ritu, 0) 'Šî–{‹‹(–{‹‹+‰Á‹‹) ~ Šî–{—¦
    lngKIN2 = Round(Kihon1 * G_ritu, 0) 'Šî–{‹‹(–{‹‹) ~ Šî–{—¦
    
    If lngY >= 35 Then
        lngY = 35 '‹Î‘±”N”‚ªãŒÀ‚ğ’´‚¦‚Ä‚¢‚½‚çãŒÀ”N”‚ğ¾¯Ä
    End If
    
    '‘ŞE‹à¾¯Ä
    JOB_kin = Round(lngKIN1 * lngY * S_ritu, 0) 'ã‹LŠî–{Šz‚É”N”‚ğŠ|‚¯‚Äx‹‹—¦‚ğŠ|‚¯‚é
    '1000‰~–¢–‚ğ®—
    lngKIN1 = Int(JOB_kin / 1000) 'x‹‹Šz‚ğ1000‚ÅŠ„‚éi¬”“_ˆÈ‰ºØ‚èÌ‚Äj
    JOB_gak = lngKIN1 * 1000 'x‹‹Šz‚É1000‚ğŠ|‚¯‚é
    If (JOB_kin - JOB_gak) > 0 Then lngKIN1 = lngKIN1 + 1 '®—‚µ‚½Šz‚ªŒ³‚Ìx‹‹Šz‚æ‚è­‚È‚©‚Á‚½‚ç1000‘«‚·
    Range("Q21") = lngKIN1 * 1000
    Range("Q23") = lngKIN1 * 1000
    Range("Q26") = JOB_kin
    
    '‡@‚Ìê‡¾¯Ä
    JOB_kin = Round(lngKIN2 * lngY * S_ritu, 0)
    lngKIN2 = Int(JOB_kin / 1000) 'x‹‹Šz‚ğ1000‚ÅŠ„‚éi¬”“_ˆÈ‰ºØ‚èÌ‚Äj
    JOB_gak = lngKIN2 * 1000 'x‹‹Šz‚É1000‚ğŠ|‚¯‚é
    If (JOB_kin - JOB_gak) > 0 Then lngKIN2 = lngKIN2 + 1 '®—‚µ‚½Šz‚ªŒ³‚Ìx‹‹Šz‚æ‚è­‚È‚©‚Á‚½‚ç1000‘«‚·
    Range("Q22") = lngKIN2 * 1000
    Range("Q25") = JOB_kin
    
End Sub
