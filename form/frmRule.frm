VERSION 5.00
Begin VB.Form frmRule 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "̰������Ϸ����"
   ClientHeight    =   2340
   ClientLeft      =   7455
   ClientTop       =   4110
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   3870
   StartUpPosition =   1  '����������
   Begin VB.Label lbShow 
      BackStyle       =   0  'Transparent
      Caption         =   "��Ϸ����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
lbShow.Caption = "1.�����̿����ߵ����߷���" & vbCrLf & _
                "2.�߲����������Լ����������ⲿ�֣���Ȼ�ͻ���Ϸʧ�ܡ�" & vbCrLf & _
                "3.�߿��Դ�ǽ��"
                
End Sub
