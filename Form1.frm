VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "点名程序优化版 v3.0"
   ClientHeight    =   6645
   ClientLeft      =   4725
   ClientTop       =   2880
   ClientWidth     =   9165
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   9165
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   15000
      Left            =   7200
      Top             =   240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "?"
      Height          =   375
      Left            =   8160
      TabIndex        =   11
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "普通点名"
      Height          =   735
      Left            =   6840
      TabIndex        =   10
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "加权点名指定人数"
      Height          =   735
      Left            =   5520
      TabIndex        =   8
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   6240
      TabIndex        =   6
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "不重复点名总人数"
      Height          =   735
      Left            =   3960
      TabIndex        =   5
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   510
      Left            =   5640
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   1560
      Width           =   615
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "华文隶书"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "总读取人数："
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   12
      Top             =   5520
      Width           =   8775
   End
   Begin VB.Label Label4 
      Caption         =   "by: DullJZ"
      Height          =   375
      Left            =   7920
      TabIndex        =   9
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "点名人数（缺省则为总人数）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   7
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "加权（%）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   1560
      Width           =   495
   End
   Begin VB.Menu yiyan 
      Caption         =   "一言"
      Begin VB.Menu API1 
         Caption         =   "API1（默认）"
      End
      Begin VB.Menu API2 
         Caption         =   "API2"
      End
      Begin VB.Menu API3 
         Caption         =   "API3（毒鸡汤）"
      End
      Begin VB.Menu lookupAPI 
         Caption         =   "查看API"
      End
   End
   Begin VB.Menu advanced 
      Caption         =   "高级"
      WindowList      =   -1  'True
      Begin VB.Menu help_use 
         Caption         =   "使用文档"
      End
      Begin VB.Menu LoadLocalText 
         Caption         =   "导入本地一言库（未完成）"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sum As Integer '总人数
Dim sum_url As Integer '总url数
Dim a As Integer, b As Integer
Dim c(1 To 1000) As String '存储学生姓名
Dim saying(1 To 20000) As String, saying_num As Integer '一言
Dim LocalText As Boolean
Dim APIurl(1 To 20) As String
Dim xmlHttp As Object

Private Function sentence(url As String)
Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
xmlHttp.Open "GET", url, True
xmlHttp.send (Null)
While xmlHttp.ReadyState <> 4
DoEvents
Wend
sentence = xmlHttp.responseText
If Mid(sentence, 1, 1) = "{" Then
    sentence = Mid(sentence, 50)
    sentence = Left(sentence, Len(sentence) - 3)
End If
End Function

Private Sub API1_Click()
API1.Checked = Not (API1.Checked)
If API1.Checked Then
    For i = 1 To 20
    If APIurl(i) = "" Then APIurl(i) = "https://v1.hitokoto.cn/?encode=text": Exit For
    Next i
Else
    For i = 1 To 20
    If APIurl(i) = "https://v1.hitokoto.cn/?encode=text" Then APIurl(i) = ""
    Next i
End If
End Sub

Private Sub API2_Click()
API2.Checked = Not (API2.Checked)
If API2.Checked Then
    For i = 1 To 20
    If APIurl(i) = "" Then APIurl(i) = "https://api.btstu.cn/yan/api.php": Exit For
    Next i
Else
    For i = 1 To 20
    If APIurl(i) = "https://api.btstu.cn/yan/api.php" Then APIurl(i) = ""
    Next i
End If
End Sub

Private Sub APi3_Click()
API3.Checked = Not (API3.Checked)
If API3.Checked Then
    For i = 1 To 20
    If APIurl(i) = "" Then APIurl(i) = "https://api.muxiaoguo.cn/api/dujitang": Exit For
    Next i
Else
    For i = 1 To 20
    If APIurl(i) = "https://api.muxiaoguo.cn/api/dujitang" Then APIurl(i) = ""
    Next i
End If
End Sub

Private Sub Command1_Click() '不重复点名总人数
List1.Clear
Dim flag(1000) As Boolean
Dim tmp(1000) As String
For i = 1 To 1000
    tmp(i) = c(i)
Next i
For i = 1 To sum
    t = Int(Rnd() * sum + 1)   '[1,sum]
    If Not flag(t) Then
        List1.AddItem tmp(t)
    Else
        i = i - 1
    End If
    flag(t) = True
Next i
End Sub

Private Sub Command2_Click()  '加权点名指定人数
List1.Clear
Dim tmp(10000) As String
Dim tmp1(10000) As String
For i = 1 To sum
    tmp1(i) = c(i)
Next
Dim members, num As Integer
Dim rate As Single
If Text3.Text <> "" Then members = Val(Text3.Text) Else members = sum
If Text2.Text <> "" Then rate = Val(Text2.Text)
If Text1.Text <> "" Then num = Val(Text1.Text)
If Not ((rate > 100 Or rate < 0) Or (num < 1 Or num > sum)) Then '排除输入错误情况
    If rate <> 0 And num <> 0 Then
        For i = 1 To Int(100 * rate)
            tmp(i) = tmp1(num)
        Next i
        For i = Int(100 * rate) + 1 To 10000
            t = Int(Rnd() * sum) + 1
            tmp(i) = tmp1(t)
        Next i
        For i = 1 To members
            t = Int(Rnd() * 10000 + 1)
            List1.AddItem tmp(t)
        Next i
    Else
        For i = 1 To members
            t = Int(Rnd() * sum + 1)
            List1.AddItem tmp1(t)
        Next i
    End If
Else
    MsgBox "输入错误！", 0, "error"
End If
End Sub

Private Sub Command3_Click()  '普通点名
List1.Clear
t = Int(Rnd() * sum + 1)
List1.AddItem c(t)
End Sub

Private Sub Command4_Click()
MsgBox "加权算法：以10000为基数，乘以权重，得到该同学的姓名重复次数(n)；其余(10000-n)个为随机生成的学生姓名；最后在10000个姓名中随机抽取", , "说明"
End Sub

Private Sub Form_Load()
Call API1_Click
List1.Clear
Randomize
APIurl(1) = "https://v1.hitokoto.cn/?encode=text"
'开始读取数据
On Error GoTo catch
try:
Open "students.txt" For Input Access Read As #1
i = 1
Do Until EOF(1) Or i > 1000
    Line Input #1, c(i)
    List1.AddItem c(i)
    i = i + 1
    sum = sum + 1
Loop
'读取数据完成
MsgBox "读取成功！共" & sum & "人", , "成功"
Label6.Caption = "总读取人数：" & sum & "人"
Close #1
GoTo finally
catch:
MsgBox "读取失败！", , "错误"
finally:
End Sub

Private Sub help_use_Click()
Shell "explorer https://dulljzblog.jz-home.top/2021/10/24/random-name-use/"
End Sub

Private Sub Label4_Click()
Shell "explorer https://dulljzblog.jz-home.top"
End Sub

Private Sub Label5_Click()
Call Timer1_Timer
End Sub

Private Sub Label6_Click()
List1.Clear
For i = 1 To sum
List1.AddItem c(i)
Next i
End Sub

Private Sub LoadLocalText_Click() '读取本地一言库

'开始读取数据
Me.CommonDialog1.ShowOpen
On Error GoTo catch
try:
Open Me.CommonDialog1.FileName For Input Access Read As #2
i = 1
Do Until EOF(2) Or i > 20000
    Line Input #2, saying(i)
    i = i + 1
Loop
'读取数据完成
MsgBox "读取成功！共" & i & "条", , "成功"
saying_num = i
LocalText = True
Close #2
GoTo finally

catch: MsgBox "错误！", , "error"

finally:
End Sub

Private Sub lookupAPI_Click()
Form2.Show
End Sub

Private Sub Timer1_Timer()
Dim APIurltmp(1 To 20) As String
sum_url = 0
For i = 1 To 20 '统计真实有效url数
If APIurl(i) <> "" Then sum_url = sum_url + 1: APIurltmp(sum_url) = APIurl(i)
Next i
For i = 1 To 20 '覆盖
APIurl(i) = APIurltmp(i)
Next i
Dim tmp As String
If Not (LocalText) Then
tmp = APIurl(Int(Rnd * sum_url) + 1)
Label5.Caption = sentence(tmp)
Else
Label5.Caption = saying(Int(Rnd * saying_num) + 1)
End If
End Sub
