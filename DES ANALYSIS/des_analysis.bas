Attribute VB_Name = "Module1"
Global BitsArr() As Boolean
Global IniByte As Byte
Global ResByte As Byte
Global t_a(0 To 64), t_b(0 To 64), PC1(0 To 56), PC2(0 To 48), yede(0 To 64), SHIFT_(0 To 256), P(0 To 32), IP_(0 To 64), EBiT(0 To 48), S1(0 To 64), S2(0 To 64), S3(0 To 64), S4(0 To 64), S5(0 To 64), S6(0 To 64), S7(0 To 64), S8(0 To 64), IP1(0 To 64) As Integer
Global lnrn(256), rnln(256), dchp(256), chp(256), pindex(256), IP_L(256), IP_R(256), pXoR(256), SB_Byte(256), C(257), D(257), KN(256), ER(256), KXoR(256), SBB(256), SB_Satir(256), SB_Sutun(256) As String
Global tbl As Tablolar
Global filename As String
Global gchp(256, 256), ind(256), adet

Type Tablolar    'LEN=1680
rec_pc1(1 To 56) As String * 2
rec_pc2(1 To 48) As String * 2
rec_shift(1 To 16) As String * 2
rec_ebit(1 To 48) As String * 2
rec_pbox(1 To 32) As String * 2
rec_ip(1 To 64) As String * 2
rec_ip1(1 To 64) As String * 2
rec_sb1(1 To 64) As String * 2
rec_sb2(1 To 64) As String * 2
rec_sb3(1 To 64) As String * 2
rec_sb4(1 To 64) As String * 2
rec_sb5(1 To 64) As String * 2
rec_sb6(1 To 64) As String * 2
rec_sb7(1 To 64) As String * 2
rec_sb8(1 To 64) As String * 2
End Type


