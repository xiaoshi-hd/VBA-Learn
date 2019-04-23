VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ��˹������ 
   Caption         =   "��˹������"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11655
   OleObjectBlob   =   "��˹������.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "��˹������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a#, b#, f#, e1, e2      '����ģ�����




Private Sub UserForm_Initialize()       'Ϊ��ѡ���趨ֵ
ComboBox1.AddItem "54����"
ComboBox1.AddItem "80����"
ComboBox1.AddItem "WGS84����"
ComboBox1.AddItem "2000����"
End Sub


Private Sub ComboBox1_Change()          '����ѡ���ѡ����ʾ���ı�����
Select Case ComboBox1.ListIndex
    Case 0
     a = 6378245
     b = 6356863.0188
     f = 1 / 298.3       '54
     e1 = 0.006693421622966
     e2 = 0.006738525414683
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

TextBox1.Value = a
TextBox2.Value = b
TextBox3.Value = f
TextBox4.Value = e1
TextBox5.Value = e2
End Sub
Private Sub CommandButton3_Click()
TextBox6.Value = ""
TextBox7.Value = ""
TextBox8.Value = ""
TextBox9.Value = ""
TextBox10.Value = ""
TextBox11.Value = ""
TextBox14.Value = ""
TextBox16.Value = ""
End Sub

Private Sub CommandButton4_Click()

TextBox12.Value = ""
TextBox13.Value = ""
TextBox15.Value = ""
TextBox17.Value = ""
End Sub


Private Sub CommandButton5_Click()
TextBox10.Value = dutodms(TextBox6.Value)
TextBox11.Value = dutodms(TextBox7.Value)
End Sub


Private Sub CommandButton1_Click()              '���и�˹����������

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
    
If OptionButton3.Value = True Then
    Call gaosizhengsuan
Else
    Call gaosifansuan
End If
End Sub

Sub gaosizhengsuan()                                '��˹����

Dim BY As Double             'ԭʼ����B                     '�Զȵĸ�ʽ����
Dim LY As Double              'ԭʼ����L
Dim B1 As Double             '��ԭʼ����Bת��Ϊ����

LY = TextBox6.Value           '125.35
BY = TextBox7.Value           '43.88
B1 = dutohudu(BY)




Dim L0 As Double             '���������߾���                'L0 = 123
Dim L As Double             'vba����������ִ�Сд,���  L = LY - L0
Dim N As Double             'î��Ȧ���ʰ뾶
Dim t As Double
Dim yita As Double
Dim daihao As Integer       'ͶӰ����

If OptionButton1.Value = True Then              '�жϷִ�ѡ��,�������������߾��Ⱥ�LX
    If LY Mod 6 = 0 Then
        L0 = 6 * Int(LY / 6) - 3
        daihao = Int(LY / 6)
    Else
        L0 = 6 * (Int(LY / 6) + 1) - 3
        daihao = (Int(LY / 6) + 1)
    End If
    L = dutohudu(LY - L0)
Else
    If (LY - 1.5) Mod 3 = 0 Then
        L0 = 3 * Int((LY - 1.5) / 3)
        daihao = Int((LY - 1.5) / 3)
    Else
        L0 = 3 * (Int((LY - 1.5) / 3) + 1)
        daihao = (Int((LY - 1.5) / 3) + 1)
    End If
    L = dutohudu(LY - L0)
End If

N = a / Sqr(1 - e1 * Sin(B1) ^ 2)
t = Tan(B1)
yita = Sqr(e2) * Cos(B1)


Dim A0 As Double
Dim A2 As Double
Dim A4 As Double
Dim A6 As Double
Dim A8 As Double
Dim M0#, M2#, M4#, M6#, M8#

M0 = a * (1 - e1)
M2 = 3 * e1 * M0 / 2
M4 = 5 * e1 * M2
M6 = 7 * e1 * M4 / 6
M8 = 9 * e1 * M6 / 8

A0 = M0 + M2 / 2 + 3 * M4 / 8 + 5 * M6 / 16 + 35 * M8 / 128
A2 = M2 / 2 + M4 / 2 + 15 * M6 / 32 + 7 * M8 / 16
A4 = M4 / 8 + 3 * M6 / 16 + 7 * M8 / 32
A6 = M6 / 32 + M8 / 16
A8 = M8 / 128


Dim X0 As Double             'LX=0ʱ���ӳ������������߻���
Dim X As Double
Dim Y As Double
Dim ABC As Double           '����ʽ���м�������Է����ʽ̫���ӱ���

ABC = (A2 - A4 + A6) + (2 * A4 - 16 * A6 / 3) * Sin(B1) ^ 2 + 16 * A6 * Sin(B1) ^ 4 / 3
X0 = A0 * B1 - Sin(B1) * Cos(B1) * ABC

X = X0 + N * Sin(B1) * Cos(B1) * L ^ 2 / 2 + N * Sin(B1) * Cos(B1) ^ 3 * (5 - t ^ 2 + 9 * yita ^ 2 + 4 * yita ^ 4) * L ^ 4 / 24 _
  + N * Sin(B1) * Cos(B1) ^ 5 * (61 - 58 * t ^ 2 + t ^ 4 + 270 * yita ^ 2 - 330 * yita ^ 2 * t ^ 2) * L ^ 6 / 720
  
Y = 500000 + N * Cos(B1) * L + N * Cos(B1) ^ 3 * (1 - t ^ 2 + yita ^ 2) * L ^ 3 / 6 _
  + N * Cos(B1) ^ 5 * (5 - 18 * t ^ 2 + t ^ 4 + 14 * yita ^ 2 - 58 * t ^ 2 * yita ^ 2) * L ^ 5 / 120

'Dim YY As String
'YY = daihao & Y                 '����д��ŵ�Y����

TextBox8.Value = X
TextBox9.Value = Y
TextBox14.Value = daihao
TextBox16.Value = L0

End Sub

Sub gaosifansuan()                     '��˹����

Dim X As Double
Dim Y As Double
Dim Y1 As Double
Dim daihao As Variant
X = TextBox8.Value
Y = TextBox9.Value
'daihao = Left(CStr(Y), 2)
'daihao = CDbl(daihao)
'Y = Y - daihao * 1000000
Y1 = Y - 500000                 '��˹ͶӰ��y����Ҫ��ȥ500000



Dim A0 As Double
Dim A2 As Double
Dim A4 As Double
Dim A6 As Double
Dim A8 As Double
Dim M0#, M2#, M4#, M6#, M8#

M0 = a * (1 - e1)
M2 = 3 * e1 * M0 / 2
M4 = 5 * e1 * M2
M6 = 7 * e1 * M4 / 6
M8 = 9 * e1 * M6 / 8

A0 = M0 + M2 / 2 + 3 * M4 / 8 + 5 * M6 / 16 + 35 * M8 / 128
A2 = M2 / 2 + M4 / 2 + 15 * M6 / 32 + 7 * M8 / 16
A4 = M4 / 8 + 3 * M6 / 16 + 7 * M8 / 32
A6 = M6 / 32 + M8 / 16
A8 = M8 / 128


Dim BF As Double                        '��������BF
Dim BF0 As Double
Dim BFI As Double
Dim FBF As Double
Dim I As Integer
BF = X / A0                         '�趨��ֵ
BF0 = BF
Do
    I = 0
    FBF = -A2 * Sin(2 * BF0) / 2 + A4 / Sin(4 * BF0) / 4 - A6 * Sin(6 * BF0) / 6 + A8 * Sin(8 * BF0) / 8
    BFI = (X - FBF) / A0
    If Abs(BFI - BF0) >= pi1 * 10 ^ (-8) / (36 * 18) Then
        BF0 = BFI
        I = 1
    End If
Loop While I = 1
BF = BFI


Dim TF As Double
Dim xitaf As Double
Dim NF As Double
Dim MF As Double


TF = Tan(BF)
xitaf = Sqr(e2) * Cos(BF)
NF = a / Sqr(1 - e1 * Sin(BF) ^ 2)
MF = NF / (1 + e2 * Cos(BF) ^ 2)



Dim BX As Double
Dim BY As Double                    '��ʽ�������γ�ȣ�Ϊ���ȵ�λ
Dim L As Double
Dim L1 As Double                    '��ʽ����ľ���
Dim L0 As Double                    '���������߾���

'L0 = CDbl(InputBox(PROMPT:="���������������߾���", Title:="������ʾ", Default:="123"))  'ת��Ϊdouble����
L0 = TextBox16.Value

BY = BF - TF * Y1 ^ 2 / (2 * MF * NF) + TF * Y1 ^ 4 * (5 + 3 * TF ^ 2 + xitaf ^ 2 - 9 * xitaf ^ 2 * TF ^ 2) / (24 * MF * NF ^ 3) _
 - TF * Y1 ^ 6 * (61 + 90 * TF ^ 2 + 45 * TF ^ 4) / (720 * MF * NF ^ 5)

L1 = Y1 / (NF * Cos(BF)) - Y1 ^ 3 * (1 + 2 * TF ^ 2 + xitaf ^ 2) / (6 * NF ^ 3 * Cos(BF)) _
+ Y1 ^ 5 * (5 + 28 * TF ^ 2 + 24 * TF ^ 4 + 6 * xitaf ^ 2 + 8 * xitaf ^ 2 * TF ^ 2) / (120 * NF ^ 5 * Cos(BF))


L = hudutodu(L1) + L0
BX = hudutodu(BY)                         '����5λС��

TextBox6.Value = L
TextBox7.Value = BX


End Sub


Private Sub CommandButton2_Click()          '��������

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

Call gaosifansuan           '��˹����

Dim L As Double             'vba����������ִ�Сд,���  L = LY - L0
Dim L0 As Double             '���������߾���                'L0 = 123
Dim daihao As Integer       'ͶӰ����

Dim BY As Double             'ԭʼ����B                     '�Զȵĸ�ʽ����
Dim LY As Double              'ԭʼ����L
Dim B1 As Double             '��ԭʼ����Bת��Ϊ����

LY = TextBox6.Value           '125.35
BY = TextBox7.Value           '43.88
B1 = dutohudu(BY)


If OptionButton5.Value = True Then              '�жϷִ�ѡ��,�������������߾��Ⱥ�LX
    daihao = TextBox15.Value
    L0 = 6 * daihao - 3
ElseIf OptionButton6.Value = True Then
    daihao = TextBox15.Value
    L0 = 3 * daihao
    ElseIf OptionButton7.Value = True Then
        L0 = 6 * Int(LY / 6) - 3
        daihao = Int(LY / 6)
        ElseIf OptionButton8.Value = True Then
            L0 = 3 * Int(LY / 6)
            daihao = Int(LY / 6)
Else
    L0 = TextBox16.Value
End If


L = dutohudu(LY - L0)

Dim N As Double             'î��Ȧ���ʰ뾶
Dim t As Double
Dim yita As Double

N = a / Sqr(1 - e1 * Sin(B1) ^ 2)
t = Tan(B1)
yita = Sqr(e2) * Cos(B1)

Dim A0 As Double
Dim A2 As Double
Dim A4 As Double
Dim A6 As Double
Dim A8 As Double
Dim M0#, M2#, M4#, M6#, M8#

M0 = a * (1 - e1)
M2 = 3 * e1 * M0 / 2
M4 = 5 * e1 * M2
M6 = 7 * e1 * M4 / 6
M8 = 9 * e1 * M6 / 8

A0 = M0 + M2 / 2 + 3 * M4 / 8 + 5 * M6 / 16 + 35 * M8 / 128
A2 = M2 / 2 + M4 / 2 + 15 * M6 / 32 + 7 * M8 / 16
A4 = M4 / 8 + 3 * M6 / 16 + 7 * M8 / 32
A6 = M6 / 32 + M8 / 16
A8 = M8 / 128


Dim X0 As Double             'LX=0ʱ���ӳ������������߻���
Dim X As Double
Dim Y As Double
Dim ABC As Double           '����ʽ���м�������Է����ʽ̫���ӱ���

ABC = (A2 - A4 + A6) + (2 * A4 - 16 * A6 / 3) * Sin(B1) ^ 2 + 16 * A6 * Sin(B1) ^ 4 / 3
X0 = A0 * B1 - Sin(B1) * Cos(B1) * ABC

X = X0 + N * Sin(B1) * Cos(B1) * L ^ 2 / 2 + N * Sin(B1) * Cos(B1) ^ 3 * (5 - t ^ 2 + 9 * yita ^ 2 + 4 * yita ^ 4) * L ^ 4 / 24 _
  + N * Sin(B1) * Cos(B1) ^ 5 * (61 - 58 * t ^ 2 + t ^ 4 + 270 * yita ^ 2 - 330 * yita ^ 2 * t ^ 2) * L ^ 6 / 720
  
Y = 500000 + N * Cos(B1) * L + N * Cos(B1) ^ 3 * (1 - t ^ 2 + yita ^ 2) * L ^ 3 / 6 _
  + N * Cos(B1) ^ 5 * (5 - 18 * t ^ 2 + t ^ 4 + 14 * yita ^ 2 - 58 * t ^ 2 * yita ^ 2) * L ^ 5 / 120

TextBox12.Value = X
TextBox13.Value = Y

End Sub


Function pi1()
pi1 = Application.WorksheetFunction.pi()   '����Excel�е�pi
End Function


Function dutohudu(du)                  '��ת��Ϊ����
dutohudu = du * pi1 / 180
End Function


Function hudutodu(hudu)                 '����ת��Ϊ��
hudutodu = hudu * 180 / pi1
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
s = (ms - m) * 60
dutodms = d & "��" & m & "��" & s & "��"
End Function
