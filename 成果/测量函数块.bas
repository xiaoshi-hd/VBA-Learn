Attribute VB_Name = "模块2"
Option Explicit


Function pi1()
pi1 = Application.WorksheetFunction.pi()   '调用Excel中的pi
End Function


Function dutodms(du) As String '度转化为度分秒
Dim d#, m#, ms#, s#
If du > 0 Then
    d = Int(du)
Else
    d = -Int(Abs(du))
End If
ms = Abs(du - d) * 60
m = Int(ms)
s = Int((ms - m) * 60)
dutodms = d & "°" & m & "′" & s & "″"
End Function


Function dutohudu(du)                  '度转化为弧度
dutohudu = du * pi1 / 180
End Function


Function hudutodu(hudu)                 '弧度转化为度
hudutodu = hudu * 180 / pi1
End Function


Function dmstodu(dms)                          '度分秒转化为度
Dim dms1#, d1#: Dim m1#: Dim s1#
dms1 = Abs(dms)
d1 = Int(dms1)
m1 = Round(((dms1 - d1) * 100), 0)             '为什么用ROUND函数，因为出现过bug
s1 = (((dms1 - d1) * 100) - m1) * 100
If dms > 0 Then
    dmstodu = d1 + m1 / 60 + s1 / 3600
Else
    dmstodu = -(d1 + m1 / 60 + s1 / 3600)
End If
End Function


Function fangwei(XA, YA, XB, YB)       '计算方位角（象限角的方法）
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
Rem                 还有一种更简单的方法：
Rem  fangwei = 180 - 90 * Abs(yb - ya + 10 ^ (-10)) / (yb - ya + 10 ^ (-10)) - Atn((xb - xa) / (yb - ya + 10 ^ (-10))) * 180 / pi1
Rem                 这里yb-ya 不能为 0  所以加上10 ^ (-10）   一步到位


Function juli(XA, YA, XB, YB)           '计算距离
juli = Sqr((XB - XA) ^ 2 + (YB - YA) ^ 2)
End Function


Function zhengsuan(XA, YA, juli, fangwei) As Variant         '坐标正算(fangwei是度的格式)
Dim XB#, YB#                                        '函数有多个返回值用数组输出
XB = XA + juli * Cos(dutohudu(fangwei))
YB = YA + juli * Sin(dutohudu(fangwei))
zhengsuan = Array(XB, YB)
End Function    '可以用数组接收多个返回值的函数
