VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "伤害计算器"
   ClientHeight    =   6165
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   9870
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton delfile 
      Caption         =   "删除"
      Height          =   300
      Left            =   6120
      TabIndex        =   64
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton writefile 
      Caption         =   "存储"
      Height          =   300
      Left            =   4200
      TabIndex        =   63
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton readfile 
      Caption         =   "读取"
      Height          =   300
      Left            =   5160
      TabIndex        =   62
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox filename 
      Height          =   270
      Left            =   4800
      TabIndex        =   60
      Top             =   5040
      Width           =   2055
   End
   Begin VB.ComboBox rate 
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   5040
      List            =   "Form1.frx":0002
      TabIndex        =   59
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox dmg3 
      Height          =   375
      Left            =   7680
      TabIndex        =   57
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox dmg2 
      Height          =   375
      Left            =   7680
      TabIndex        =   55
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox critrate 
      Height          =   270
      Left            =   1680
      TabIndex        =   53
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox ans5 
      Height          =   270
      Left            =   8280
      TabIndex        =   48
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox ans6 
      Height          =   270
      Left            =   8280
      TabIndex        =   47
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox ans7 
      Height          =   270
      Left            =   8280
      TabIndex        =   46
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox ans8 
      Height          =   270
      Left            =   8280
      TabIndex        =   45
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox ans1 
      Height          =   270
      Left            =   8280
      TabIndex        =   40
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox ans2 
      Height          =   270
      Left            =   8280
      TabIndex        =   39
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox ans3 
      Height          =   270
      Left            =   8280
      TabIndex        =   38
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox ans4 
      Height          =   270
      Left            =   8280
      TabIndex        =   37
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox eleb 
      Height          =   270
      Left            =   5040
      TabIndex        =   34
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox lvf 
      Height          =   270
      Left            =   5040
      TabIndex        =   32
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox mas 
      Height          =   270
      Left            =   5040
      TabIndex        =   29
      Top             =   1800
      Width           =   1215
   End
   Begin VB.OptionButton Option0 
      Caption         =   "无反应"
      Height          =   495
      Left            =   4680
      TabIndex        =   28
      Top             =   360
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "增幅反应"
      Height          =   495
      Left            =   4680
      TabIndex        =   27
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "激化反应"
      Height          =   495
      Left            =   4680
      TabIndex        =   26
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox res 
      Height          =   270
      Left            =   1680
      TabIndex        =   24
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox nodef 
      Height          =   270
      Left            =   1680
      TabIndex        =   22
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox redef 
      Height          =   270
      Left            =   1680
      TabIndex        =   20
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox mlv 
      Height          =   270
      Left            =   1680
      TabIndex        =   18
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox plv 
      Height          =   270
      Left            =   1680
      TabIndex        =   16
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox crit 
      Height          =   270
      Left            =   1680
      TabIndex        =   14
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox bon 
      Height          =   270
      Left            =   1680
      TabIndex        =   12
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox att5 
      Height          =   270
      Left            =   1680
      TabIndex        =   10
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox dmg1 
      Height          =   375
      Left            =   7680
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox att4 
      Height          =   270
      Left            =   1680
      TabIndex        =   8
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox att3 
      Height          =   270
      Left            =   1680
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox att2 
      Height          =   270
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox att1 
      Height          =   270
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton ans 
      Caption         =   "计算"
      Height          =   615
      Left            =   4560
      TabIndex        =   0
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label22 
      Caption         =   "文件名"
      Height          =   255
      Left            =   4200
      TabIndex        =   61
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label21 
      Caption         =   "最终伤害(未暴击)"
      Height          =   255
      Left            =   7560
      TabIndex        =   58
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label20 
      Caption         =   "最终伤害(暴击)"
      Height          =   255
      Left            =   7680
      TabIndex        =   56
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "暴击率"
      Height          =   255
      Left            =   1080
      TabIndex        =   54
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lable114 
      Caption         =   "暴击乘区"
      Height          =   255
      Left            =   7560
      TabIndex        =   52
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label anse2 
      Caption         =   "增幅乘区"
      Height          =   255
      Left            =   7560
      TabIndex        =   51
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label19 
      Caption         =   "防御乘区"
      Height          =   255
      Left            =   7560
      TabIndex        =   50
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "抗性乘区"
      Height          =   255
      Left            =   7560
      TabIndex        =   49
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "总攻击力"
      Height          =   255
      Left            =   7560
      TabIndex        =   44
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "基础伤害"
      Height          =   255
      Left            =   7560
      TabIndex        =   43
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label anse1 
      Caption         =   "激化加成"
      Height          =   255
      Left            =   7560
      TabIndex        =   42
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "增伤乘区"
      Height          =   255
      Left            =   7560
      TabIndex        =   41
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "最终伤害(期望)"
      Height          =   255
      Left            =   7680
      TabIndex        =   36
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label e4 
      Caption         =   "反应伤害提高"
      Height          =   255
      Left            =   3960
      TabIndex        =   35
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label e3 
      Caption         =   "等级系数"
      Height          =   255
      Left            =   4320
      TabIndex        =   33
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label e2 
      Caption         =   "反应倍率"
      Height          =   255
      Left            =   4320
      TabIndex        =   31
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label e1 
      Caption         =   "元素精通"
      Height          =   255
      Left            =   4320
      TabIndex        =   30
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "敌人抗性"
      Height          =   255
      Left            =   960
      TabIndex        =   25
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "无视防御"
      Height          =   255
      Left            =   960
      TabIndex        =   23
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "减防"
      Height          =   255
      Left            =   1320
      TabIndex        =   21
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "敌人等级"
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "角色等级"
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "暴击伤害"
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "各类伤害加成"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "技能倍率"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "百分比加成"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "固定数值加成"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "攻击力蓝值"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "攻击力白值"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mode As Integer
Dim flv(1 To 90) As Double
Sub lvinit()
    flv(1) = 17.23
    flv(2) = 18.67
    flv(3) = 20.1
    flv(4) = 21.51
    flv(5) = 22.14
    flv(6) = 25.04
    flv(7) = 26.4
    flv(8) = 28.5
    flv(9) = 31.31
    flv(10) = 34.09
    flv(11) = 37.58
    flv(12) = 40.3
    flv(13) = 44.45
    flv(14) = 48.56
    flv(15) = 53.34
    flv(16) = 59.5
    flv(17) = 64.18
    flv(18) = 69.52
    flv(19) = 74.8
    flv(20) = 80.74
    flv(21) = 85.92
    flv(22) = 91.74
    flv(23) = 96.82
    flv(24) = 103.23
    flv(25) = 108.21
    flv(26) = 113.14
    flv(27) = 118.03
    flv(28) = 122.88
    flv(29) = 130.38
    flv(30) = 136.47
    flv(31) = 143.18
    flv(32) = 149.16
    flv(33) = 155.76
    flv(34) = 161.65
    flv(35) = 169.45
    flv(36) = 176.54
    flv(37) = 184.22
    flv(38) = 191.84
    flv(39) = 199.41
    flv(40) = 207.55
    flv(41) = 215.64
    flv(42) = 224.3
    flv(43) = 233.52
    flv(44) = 243.31
    flv(45) = 256.17
    flv(46) = 268.31
    flv(47) = 281.61
    flv(48) = 294.81
    flv(49) = 309.16
    flv(50) = 323.4
    flv(51) = 336.93
    flv(52) = 350.37
    flv(53) = 364.33
    flv(54) = 378.2
    flv(55) = 398.63
    flv(56) = 416.51
    flv(57) = 434.27
    flv(58) = 452.52
    flv(59) = 472.44
    flv(60) = 492.83
    flv(61) = 513.68
    flv(62) = 539.13
    flv(63) = 565.6
    flv(64) = 592.48
    flv(65) = 624.47
    flv(66) = 651.6
    flv(67) = 679.72
    flv(68) = 707.69
    flv(69) = 737.22
    flv(70) = 765.42
    flv(71) = 794.61
    flv(72) = 824.79
    flv(73) = 851.37
    flv(74) = 877.8
    flv(75) = 914.29
    flv(76) = 946.63
    flv(77) = 979.35
    flv(78) = 1010.78
    flv(79) = 1044.84
    flv(80) = 1077.04
    flv(81) = 1109.64
    flv(82) = 1143.17
    flv(83) = 1176.54
    flv(84) = 1209.74
    flv(85) = 1253.8
    flv(86) = 1288.84
    flv(87) = 1325.35
    flv(88) = 1363.33
    flv(89) = 1405.48
    flv(90) = 1446.88
End Sub
Function cal1(T As ComboBox) As Double
Dim str As String
str = T.Text
cal1 = calculate(str)
End Function
Function cal(T As TextBox) As Double
Dim str As String
str = T.Text
cal = calculate(str)
End Function
Function calculate(T As String) As Double
    s = T
    If Len(s) = 0 Then
        calculate = 0
        Exit Function
    End If
    
    Dim nums() As Double
    Dim ops() As String
    Dim tempNum As String
    tempNum = ""
    
    Dim i As Integer
    Dim c As String
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If IsDigitOrDotOrSign(c, tempNum, nums, ops) Then
            tempNum = tempNum & c
        Else
            If tempNum <> "" Then
                AddToNums nums, CDbl(tempNum)
                tempNum = ""
            End If
            
            If IsArrayInitialized(nums) Then
                Dim numsCount As Integer
                numsCount = GetNumsCount(nums)
                Dim opsCount As Integer
                opsCount = GetOpsCount(ops)
                
                If numsCount = opsCount Then
                    tempNum = c
                Else
                    AddToOps ops, c
                End If
            Else
                tempNum = c
            End If
        End If
    Next i
    
    If tempNum <> "" Then
        AddToNums nums, CDbl(tempNum)
    End If
    
    If GetNumsCount(nums) = 0 Then
        calculate = 0
        Exit Function
    End If
    
    Dim j As Integer
    j = 0
    While j < GetOpsCount(ops)
        If ops(j) = "*" Then
            nums(j) = nums(j) * nums(j + 1)
            RemoveDoubleElement nums, j + 1
            RemoveStringElement ops, j
        Else
            j = j + 1
        End If
    Wend
    
    Dim result As Double
    result = nums(0)
    For j = 0 To GetOpsCount(ops) - 1
        If ops(j) = "+" Then
            result = result + nums(j + 1)
        Else
            result = result - nums(j + 1)
        End If
    Next j
    
    calculate = result
End Function

Private Function IsDigitOrDotOrSign(c As String, tempNum As String, nums() As Double, ops() As String) As Boolean
    If IsNumeric(c) Or c = "." Then
        IsDigitOrDotOrSign = True
    ElseIf c = "+" Or c = "-" Then
        If tempNum = "" Then
            Dim numsCount As Integer
            numsCount = GetNumsCount(nums)
            Dim opsCount As Integer
            opsCount = GetOpsCount(ops)
            IsDigitOrDotOrSign = (numsCount = opsCount)
        Else
            IsDigitOrDotOrSign = False
        End If
    Else
        IsDigitOrDotOrSign = False
    End If
End Function

Private Sub AddToNums(arr() As Double, value As Double)
    If Not IsArrayInitialized(arr) Then
        ReDim arr(0)
        arr(0) = value
    Else
        ReDim Preserve arr(UBound(arr) + 1)
        arr(UBound(arr)) = value
    End If
End Sub

Private Sub AddToOps(arr() As String, value As String)
    If Not IsArrayInitialized(arr) Then
        ReDim arr(0)
        arr(0) = value
    Else
        ReDim Preserve arr(UBound(arr) + 1)
        arr(UBound(arr)) = value
    End If
End Sub

Private Function GetNumsCount(arr() As Double) As Integer
    If IsArrayInitialized(arr) Then
        GetNumsCount = UBound(arr) + 1
    Else
        GetNumsCount = 0
    End If
End Function

Private Function GetOpsCount(arr() As String) As Integer
    If IsArrayInitialized(arr) Then
        GetOpsCount = UBound(arr) + 1
    Else
        GetOpsCount = 0
    End If
End Function

Private Function IsArrayInitialized(arr As Variant) As Boolean
    On Error Resume Next
    Dim uboundCheck As Integer
    uboundCheck = UBound(arr)
    If Err.Number = 0 Then
        IsArrayInitialized = True
    Else
        IsArrayInitialized = False
    End If
    On Error GoTo 0
End Function

Private Sub RemoveDoubleElement(arr() As Double, ByVal index As Integer)
    Dim i As Integer
    For i = index To UBound(arr) - 1
        arr(i) = arr(i + 1)
    Next i
    If UBound(arr) > 0 Then
        ReDim Preserve arr(UBound(arr) - 1)
    Else
        Erase arr
    End If
End Sub

Private Sub RemoveStringElement(arr() As String, ByVal index As Integer)
    Dim i As Integer
    For i = index To UBound(arr) - 1
        arr(i) = arr(i + 1)
    Next i
    If UBound(arr) > 0 Then
        ReDim Preserve arr(UBound(arr) - 1)
    Else
        Erase arr
    End If
End Sub
Sub init()
att1.Text = "808"
att2.Text = "3147"
att3.Text = ""
att4.Text = ""
att5.Text = "428.9"
bon.Text = "0.31*400+86.6+56"
critrate.Text = "63.5"
crit.Text = "157.1"
plv.Text = "90"
mlv.Text = "60"
res.Text = "-140-30"
lvf.Text = "1446.88"
rate.Text = "1.25"
mas.Text = "110"
End Sub


Private Sub ans_Click()
Dim BaseDMG As Double
BaseDMG = cal(att1) * (1 + cal(att4) / 100) + cal(att2) + cal(att3)
ans1.Text = BaseDMG
BaseDMG = BaseDMG * cal(att5) / 100
ans2.Text = BaseDMG

If mode = 1 Then
    Dim temp As Double
    temp = cal(lvf) * cal1(rate) * (1 + (5 * cal(mas)) / (cal(mas) + 1200) + cal(eleb) / 100)
    BaseDMG = BaseDMG + temp
    ans3.Text = temp
Else
    ans3.Text = ""
End If

Dim BonusDMG As Double
BonusDMG = 1 + cal(bon) / 100
ans4.Text = BonusDMG

Dim CritDMG As Double
CritDMG = 1 + cal(crit) / 100
ans5.Text = CritDMG

Dim ele As Double
If mode = 2 Then
    ele = cal1(rate) * (1 + 2.78 * cal(mas) / (cal(mas) + 1400) + cal(eleb) / 100)
    ans6.Text = ele
Else
    ele = 1
    ans6.Text = ""
End If

Dim def As Double
def = (cal(plv) + 100) / ((cal(plv) + 100) + (cal(mlv) + 100) * (1 - cal(nodef) / 100) * (1 - cal(redef) / 100))
ans7.Text = def

Dim resis As Double
If cal(res) > 75 Then
    resis = 1 / (1 + 4 * cal(res) / 100)
Else
    If cal(res) >= 0 Then
        resis = 1 - cal(res) / 100
    Else
        resis = 1 - cal(res) / 100 / 2
    End If
End If
ans8.Text = resis

Dim critf As Double
If cal(critrate) > 100 Then
    critf = CritDMG
Else
    critf = (1 + cal(critrate) / 100 * cal(crit) / 100)
End If
dmg1.Text = BaseDMG * BonusDMG * critf * def * resis * ele
dmg2.Text = BaseDMG * BonusDMG * CritDMG * def * resis * ele
dmg3.Text = BaseDMG * BonusDMG * def * resis * ele
End Sub



Private Sub Form_Load()
Call lvinit
rate.List(0) = "超激化1.15"
rate.List(1) = "蔓激化1.25"
rate.List(2) = "水蒸发2.0"
rate.List(3) = "火蒸发1.5"
rate.List(4) = "火融化2.0"
rate.List(5) = "冰融化1.5"
Call Option0_Click
Option0.value = True
Call init
End Sub

Private Sub Option2_Click()
mode = 2
e1.Enabled = True
e2.Enabled = True
e3.Enabled = False
e4.Enabled = True
anse1.Enabled = False
anse2.Enabled = True
End Sub
Private Sub Option1_Click()
mode = 1
e1.Enabled = True
e2.Enabled = True
e3.Enabled = True
e4.Enabled = True
anse1.Enabled = True
anse2.Enabled = False
End Sub
Private Sub Option0_Click()
mode = 0
e1.Enabled = False
e2.Enabled = False
e3.Enabled = False
e4.Enabled = False
anse1.Enabled = False
anse2.Enabled = False

End Sub

Private Sub plv_Change()
If cal(plv) <= 90 And cal(plv) >= 1 Then
    lvf.Text = flv(Int(cal(plv)))
Else
    lvf.Text = "1446.88"
End If
End Sub
Private Sub rate_Change()
Select Case rate.Text
    Case "超激化1.15"
        rate.Text = "1.15"
    Case "蔓激化1.25"
        rate.Text = "1.25"
    Case "水蒸发2.0", "火融化2.0"
        rate.Text = "2.0"
    Case "火蒸发1.5", "冰融化1.5"
        rate.Text = "1.5"
End Select
End Sub

Private Sub rate_Click()
Select Case rate.Text
    Case "超激化1.15"
        rate.Text = "1.15"
    Case "蔓激化1.25"
        rate.Text = "1.25"
    Case "水蒸发2.0", "火融化2.0"
        rate.Text = "2.0"
    Case "火蒸发1.5", "冰融化1.5"
        rate.Text = "1.5"
End Select
End Sub

Private Sub writefile_Click()
Dim File As String
Dim response As Integer
File = filename.Text
If Trim(File) = "" Then
        MsgBox "文件名不能为空！", vbExclamation
        Exit Sub
    End If
If Dir(File) <> "" Then
    response = MsgBox("将覆盖文件：" & vbCrLf & File, vbQuestion + vbYesNo, "确认覆盖")
    If response = vbNo Then
        Exit Sub
    End If
 End If
Open File For Output As #1
Print #1, att1.Text
Print #1, att2.Text
Print #1, att3.Text
Print #1, att4.Text
Print #1, att5.Text
Print #1, bon.Text
Print #1, critrate.Text
Print #1, crit.Text
Print #1, plv.Text
Print #1, mlv.Text
Print #1, redef.Text
Print #1, nodef.Text
Print #1, res.Text
Print #1, mode
Print #1, mas.Text
Print #1, rate.Text
Print #1, lvf.Text
Print #1, eleb.Text
Close #1
MsgBox "文件保存成功！", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "保存文件时出错：" & Err.Description, vbCritical
    If Not EOF(1) Then Close #1 ' 确保文件关闭
End Sub
Private Sub readfile_Click()
    Dim File As String
    Dim temp As String  ' 用于临时存储每行数据
    Dim response As Integer
    File = filename.Text
    If Trim(File) = "" Then
        MsgBox "文件名不能为空！", vbExclamation
        Exit Sub
    End If
    ' 检查文件是否存在
    If Dir(File) = "" Then
        MsgBox "文件不存在!", vbExclamation
        Exit Sub
    End If
    response = MsgBox("将覆盖控制台内容：" & vbCrLf & File, vbQuestion + vbYesNo, "确认覆盖")
    If response = vbNo Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    Open File For Input As #1
    
    ' 按顺序读取每一行并赋值
    Line Input #1, temp
    att1.Text = temp
    Line Input #1, temp
    att2.Text = temp
    Line Input #1, temp
    att3.Text = temp
    Line Input #1, temp
    att4.Text = temp
    Line Input #1, temp
    att5.Text = temp
    Line Input #1, temp
    bon.Text = temp
    Line Input #1, temp
    critrate.Text = temp
    Line Input #1, temp
    crit.Text = temp
    Line Input #1, temp
    plv.Text = temp
    Line Input #1, temp
    mlv.Text = temp
    Line Input #1, temp
    redef.Text = temp
    Line Input #1, temp
    nodef.Text = temp
    Line Input #1, temp
    res.Text = temp
    Line Input #1, temp
    temp = CInt(temp)
    Line Input #1, temp
    mas.Text = temp
    Line Input #1, temp
    rate.Text = temp
    Line Input #1, temp
    lvf.Text = temp
    Line Input #1, temp
    eleb.Text = temp
    Close #1
    Exit Sub
'    If mode = 0 Then
'        Call Option0_Click
'    End If
'    If mode = 1 Then
'        Call Option1_Click
'    End If
'    If mode = 2 Then
'        Call Option2_Click
'    End If
ErrorHandler:
    MsgBox "读取文件时出错: " & Err.Description, vbExclamation
    Close #1
End Sub
Private Sub delfile_Click()
    Dim File As String
    Dim response As Integer
    
    ' 获取文件名
    File = filename.Text
    If Trim(File) = "" Then
        MsgBox "文件名不能为空！", vbExclamation
        Exit Sub
    End If
    ' 检查文件是否存在
    If Dir(File) = "" Then
        MsgBox "文件不存在！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 弹出确认对话框
    response = MsgBox("确认删除文件：" & vbCrLf & File, vbQuestion + vbYesNo, "确认删除")
    
    ' 如果用户选择"是"
    If response = vbYes Then
        On Error GoTo ErrorHandler
        
        ' 删除文件
        Kill File
        MsgBox "文件删除成功！", vbInformation, "完成"
        
        ' 清空文件名显示（可选）
        'FileName.Text = ""
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "删除文件时出错：" & vbCrLf & Err.Description, vbCritical, "错误"
End Sub
