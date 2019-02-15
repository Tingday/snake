VERSION 5.00
Begin VB.Form frmRule 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "贪吃蛇游戏规则"
   ClientHeight    =   2340
   ClientLeft      =   7455
   ClientTop       =   4110
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   3870
   StartUpPosition =   1  '所有者中心
   Begin VB.Label lbShow 
      BackStyle       =   0  'Transparent
      Caption         =   "游戏规则"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
lbShow.Caption = "1.方向盘控制蛇的行走方向。" & vbCrLf & _
                "2.蛇不可以碰到自己的身体任意部分，不然就会游戏失败。" & vbCrLf & _
                "3.蛇可以穿墙。"
                
End Sub
