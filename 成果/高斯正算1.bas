Attribute VB_Name = "ģ��3"
Sub gaosizhengsuan1()                                '��˹����,һ���㷨
'ԭʼ�����Զȵĸ�ʽ����;Ϊʲô���öȷ���,��Ϊ�ȷ�����ķ�Χ����,��ز�����Ҫ���ߵľ���

Select Case ComboBox1.ListIndex         'ѡ���������
    Case 0
     a = 6378245                        '���򳤰���
     b = 6356863.0188                   '����̰���
     f = 1 / 298.3       '54            '�������
     e1 = 0.006693421622966             '�����һƫ���ʵ�ƽ��
     e2 = 0.006738525414683             '����ڶ�ƫ���ʵ�ƽ��
    Case 1
     a = 6378140
     b = 6356755.28815753
     f = 1 / 298.257         '80
     e1 = 0.00669438499959
     e2 = 0.00673950181947
    Case 2
     a = 6378139
     b = 6356752.3142
     f = 1 / 298.257223563    'wgs84
     e1 = 0.00669437999013
     e2 = 0.00673949674223
    Case Else
     a = 6378137
     b = 6356752.31414
     f = 1 / 298.257222101    '2000
     e1 = 0.0066943800229
     e2 = 0.00673949677548
    End Select



Dim BX As Double             'ԭʼ����B
Dim L As Double              'ԭʼ����L
Dim B1 As Double             '��ԭʼ����Bת��Ϊ����
Dim L1 As Double             '��ԭʼ����Lת��Ϊ��
L = TextBox6.Value           '125.35
BX = TextBox7.Value           '43.88
B1 = dutohudu(BX)
L1 = L



Dim L0 As Double             '���������߾���                'L0 = 123
Dim LX As Double             'vba����������ִ�Сд,���  LX = L1 - L0
Dim N As Double             'î��Ȧ���ʰ뾶
Dim t As Double
Dim yita As Double
Dim daihao As Integer       'ͶӰ����

If OptionButton1.Value = True Then              '�жϷִ�ѡ��,�������������߾��Ⱥ�LX
    If L1 Mod 6 = 0 Then
        L0 = 6 * Int(L1 / 6) - 3
        daihao = Int(L1 / 6)
    Else
        L0 = 6 * (Int(L1 / 6) + 1) - 3
        daihao = (Int(L1 / 6) + 1)
    End If
    LX = dutohudu(L1 - L0)
Else
    If (L1 - 1.5) Mod 3 = 0 Then
        L0 = 3 * Int((L1 - 1.5) / 3)
    Else
        L0 = 3 * (Int((L1 - 1.5) / 3) + 1)
    End If
    LX = dutohudu(L1 - L0)
End If

N = a / Sqr(1 - e1 * Sin(B1) ^ 2)
t = Tan(B1)
yita = Sqr(e2) * Cos(B1)



Dim A0 As Double
Dim A2 As Double
Dim A4 As Double
Dim A6 As Double
Dim A8 As Double

A0 = 1 + 3 * e1 / 4 + 45 * e1 ^ 2 / 64 + 350 * e1 ^ 3 / 512 + 11025 * e1 ^ 4 / 16384
A2 = -0.5 * (3 * e1 / 4 + 60 * e1 ^ 2 / 64 + 525 * e1 ^ 3 / 512 + 17640 * e1 ^ 4 / 16384)
A4 = 1 * (15 * e1 ^ 2 / 64 + 210 * e1 ^ 3 / 512 + 8820 * e1 ^ 4 / 16384) / 4
A6 = -1 * (35 * e1 ^ 3 / 512 + 2520 * e1 ^ 4 / 16384) / 6
A8 = 1 * (315 * e1 ^ 4 / 16384) / 8


Dim X0 As Double             'LX=0ʱ���ӳ������������߻���
Dim X As Double
Dim Y As Double
Dim YY As String                '���д��ŵ�Yֵ

X0 = a * (1 - e1) * (A0 * B1 + A2 * Sin(2 * B1) + A4 * Sin(4 * B1) + A6 * Sin(6 * B1) + A8 * Sin(8 * B1))

X = X0 + N / 2 * t * Cos(B1) ^ 2 * LX ^ 2 + N / 24 * t * Cos(B1) ^ 4 * (5 - t ^ 2 + 9 * yita ^ 2 + 4 * yita ^ 4) * LX ^ 4 _
  + N / 720 * t * Cos(B1) ^ 6 * (61 - 58 * t ^ 2 + t ^ 4 + 270 * yita ^ 2 - 330 * yita ^ 2 * t ^ 2) * LX ^ 6
  
Y = 500000 + N * Cos(B1) * LX + N / 6 * Cos(B1) ^ 3 * (1 - t ^ 2 + yita ^ 2) * LX ^ 3 _
  + N / 120 * Cos(B1) ^ 5 * (5 - 18 * t ^ 2 + t ^ 4 + 14 * yita ^ 2 - 58 * t ^ 2 * yita ^ 2) * LX ^ 5

YY = daihao & Y                 '����д��ŵ�Y����

TextBox8.Value = X
TextBox9.Value = YY

End Sub

