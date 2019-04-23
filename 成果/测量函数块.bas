Attribute VB_Name = "ģ��2"
Option Explicit


Function pi1()
pi1 = Application.WorksheetFunction.pi()   '����Excel�е�pi
End Function


Function dutodms(du) As String '��ת��Ϊ�ȷ���
Dim d#, m#, ms#, s#
If du > 0 Then
    d = Int(du)
Else
    d = -Int(Abs(du))
End If
ms = Abs(du - d) * 60
m = Int(ms)
s = Int((ms - m) * 60)
dutodms = d & "��" & m & "��" & s & "��"
End Function


Function dutohudu(du)                  '��ת��Ϊ����
dutohudu = du * pi1 / 180
End Function


Function hudutodu(hudu)                 '����ת��Ϊ��
hudutodu = hudu * 180 / pi1
End Function


Function dmstodu(dms)                          '�ȷ���ת��Ϊ��
Dim dms1#, d1#: Dim m1#: Dim s1#
dms1 = Abs(dms)
d1 = Int(dms1)
m1 = Round(((dms1 - d1) * 100), 0)             'Ϊʲô��ROUND��������Ϊ���ֹ�bug
s1 = (((dms1 - d1) * 100) - m1) * 100
If dms > 0 Then
    dmstodu = d1 + m1 / 60 + s1 / 3600
Else
    dmstodu = -(d1 + m1 / 60 + s1 / 3600)
End If
End Function


Function fangwei(XA, YA, XB, YB)       '���㷽λ�ǣ����޽ǵķ�����
Dim X As Double
Dim Y As Double
X = XB - XA
Y = YB - YA
If X = 0 Then
    If Y > 0 Then
        fangwei = 90
    Else
        fangwei = 270
    End If
Else
    If X > 0 Then
        If Y >= 0 Then
            fangwei = Atn((YB - YA) / (XB - XA)) * 180 / pi1
        Else
            fangwei = 360 - Atn(Abs((YB - YA) / (XB - XA))) * 180 / pi1
        End If
    Else
        If Y >= 0 Then
            fangwei = 180 - Atn(Abs((YB - YA) / (XB - XA))) * 180 / pi1
        Else
            fangwei = 180 + Atn(Abs((YB - YA) / (XB - XA))) * 180 / pi1
        End If
    End If
End If
End Function
Rem                 ����һ�ָ��򵥵ķ�����
Rem  fangwei = 180 - 90 * Abs(yb - ya + 10 ^ (-10)) / (yb - ya + 10 ^ (-10)) - Atn((xb - xa) / (yb - ya + 10 ^ (-10))) * 180 / pi1
Rem                 ����yb-ya ����Ϊ 0  ���Լ���10 ^ (-10��   һ����λ


Function juli(XA, YA, XB, YB)           '�������
juli = Sqr((XB - XA) ^ 2 + (YB - YA) ^ 2)
End Function


Function zhengsuan(XA, YA, juli, fangwei) As Variant         '��������(fangwei�Ƕȵĸ�ʽ)
Dim XB#, YB#                                        '�����ж������ֵ���������
XB = XA + juli * Cos(dutohudu(fangwei))
YB = YA + juli * Sin(dutohudu(fangwei))
zhengsuan = Array(XB, YB)
End Function    '������������ն������ֵ�ĺ���
