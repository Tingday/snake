VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����""̰����"""
   ClientHeight    =   4380
   ClientLeft      =   6945
   ClientTop       =   3705
   ClientWidth     =   6720
   LinkTopic       =   "MDIForm1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6720
   StartUpPosition =   1  '����������
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "̰����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lbOwner 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   360
      TabIndex        =   1
      Top             =   3000
      Width           =   6045
   End
   Begin VB.Label lbShow 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   6090
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   6360
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lbShow.Caption = "�ٶ�����0yufan0" & vbCrLf & _
                    "�汾:2.0(�ڲ��汾  20190215)" & vbCrLf & _
                    "�汾���� 2019 �ٶ�����0yufan0�������С�" & vbCrLf & _
                    "̰������Ϸ�����û�������������й�����������/������֪ʶ��Ȩ��������"
    lbOwner.Caption = "���ٶ�����0yufan0��ɣ�����Ʒ��Դ����������⿽��������������Ϣ��" & vbCrLf & _
                        "��ϵ��ʽ: woyufan@163.com "
End Sub

