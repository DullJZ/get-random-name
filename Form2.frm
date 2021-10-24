VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "ÊÕÂ¼API"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3555
   ScaleWidth      =   6600
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Text            =   "https://api.muxiaoguo.cn/api/dujitang"
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   2
      Left            =   1800
      TabIndex        =   3
      Text            =   "https://api.btstu.cn/yan/api.php"
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Text            =   "https://v1.hitokoto.cn/?encode=text"
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "API3"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "API2"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "API1"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
