VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "̰����"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   1  '����������
   Begin VB.Menu ��Ϸ 
      Caption         =   "��ʼ��Ϸ"
   End
   Begin VB.Menu �˳���Ϸ 
      Caption         =   "�˳���Ϸ"
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------
'�汾 2018.03.28 �޸� ���ߣ�0yufan0 vb��
'��ϵ woyufan@163.com ����������ϵ��
'��Դ��Ϊ���˱�д������ѧϰ�����˽�fps��Ϸ��д
'-------------------------------------------
'---------------------------------------------����Ϊ˽�б���----------------------------------------
Option Explicit
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function GetInputState Lib "user32.dll" () As Long

'��Ϸȫ��
Dim Game_Frame As RECTL '��Χǽ
Dim fps As Long '��Ϸˢ��Ƶ�ʣ���/�룩
Const Game_Wide = 15 '������
'�߱���
Dim Snakes() As Point '�߱���
Dim White_Food As Point '��ͨʳ��
Dim Red_Food As Point '����ʮ��
Dim Food_Eated As Point
Dim Snake_Speed As Long '�ߵ��ٶ�(����/Game_Wide)
'��Ϸ״̬
Dim mGame_State As Game_Status
Dim mSnake_Direction As Snake_Direction
Dim mUser_Direction As User_Direction
'��Ϸ�������
Dim score As Long '�ܵ÷� = ��ͨʳ��÷� + ����ʳ�� * 5
Dim N_White_Food As Long
Dim N_Red_Food As Long
'����
Private Type Point
        x As Long
        y As Long
End Type

Private Type RECTL
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'�б�����
Private Enum Game_Status
    Game_STATE_RUNNING = 0
    Game_STATE_PAUSE = 1
    Game_STATE_STOP = 2
End Enum
Private Enum Snake_Direction
    Snake_Direction_Up = 1
    Snake_direction_Down = 2
    Snake_Direction_Left = 3
    Snake_Direction_Right = 4
End Enum
Private Enum User_Direction
    User_Direction_Up = 5
    User_direction_Down = 6
    User_Direction_Left = 7
    User_Direction_Right = 8
End Enum
Private Enum Food_Color
    Food_Color_White = 0
    Food_Color_Red = 1
End Enum
'--------------------------------------------------------�������˽���---------------------------------------
'--------------------------------------------------------�� �� ��  ��---------------------------------------
Private Function MoveSnake() As Boolean '�ƶ������ж�
    Dim current_snake As Integer '������i
    Dim mSnake_Head As Point '��ͷ����
    Dim mSnake_length As Long '�߳��ȣ���λ1���ߣ�
    mSnake_length = UBound(Snakes)
    mSnake_Head = Snakes(0) '�����ͷ
    Select Case mUser_Direction '���ݷ����ƶ���ÿ��һ����λGame_Wide
        Case User_Direction_Up
            If mSnake_Direction <> Snake_direction_Down Then
                Snakes(0).y = Snakes(0).y - Game_Wide
                mSnake_Direction = Snake_Direction_Up
            Else
                mUser_Direction = User_direction_Down
                Snakes(0).y = Snakes(0).y + Game_Wide
            End If
        Case User_direction_Down
            If mSnake_Direction <> Snake_Direction_Up Then
                Snakes(0).y = Snakes(0).y + Game_Wide
                mSnake_Direction = Snake_direction_Down
            Else
                mUser_Direction = User_Direction_Up
                Snakes(0).y = Snakes(0).y - Game_Wide
            End If
        Case User_Direction_Left
            If mSnake_Direction <> Snake_Direction_Right Then
                Snakes(0).x = Snakes(0).x - Game_Wide
                mSnake_Direction = Snake_Direction_Left
            Else
                mUser_Direction = User_Direction_Right
                Snakes(0).x = Snakes(0).x + Game_Wide
            End If
        Case User_Direction_Right
            If mSnake_Direction <> Snake_Direction_Left Then
                Snakes(0).x = Snakes(0).x + Game_Wide
                mSnake_Direction = Snake_Direction_Right
            Else
                mUser_Direction = User_Direction_Left
                Snakes(0).x = Snakes(0).x - Game_Wide
            End If
    End Select
    '��ǽЧ��ʵ��
    If Snakes(0).y >= Game_Frame.Bottom Then Snakes(0).y = Game_Frame.Top
    If Snakes(0).y < Game_Frame.Top Then Snakes(0).y = Game_Frame.Bottom - Game_Wide
    If Snakes(0).x >= Game_Frame.Right Then Snakes(0).x = Game_Frame.Left
    If Snakes(0).x < Game_Frame.Left Then Snakes(0).x = Game_Frame.Right - Game_Wide
    '��ײ���
    For current_snake = 1 To mSnake_length
        If Snakes(0).x = Snakes(current_snake).x And Snakes(0).y = Snakes(current_snake).y Then
            mGame_State = Game_STATE_STOP
            ��Ϸ.Caption = "��ʼ��Ϸ"
            MsgBox "��Ϸ������"
            Exit Function
        End If
    Next current_snake
    'ʳ����ײ���
    If Snakes(0).x = White_Food.x And Snakes(0).y = White_Food.y Then
        N_White_Food = N_White_Food + 1
        Food_Eated = White_Food
        CreateFood Food_Color_White
        If N_White_Food Mod 8 = 0 Then CreateFood Food_Color_Red 'ÿ��8��ʳ��ͳ���һ����ɫʳ��
    End If
    '��ʳ����ײ���
    If Snakes(0).x = Red_Food.x And Snakes(0).y = Red_Food.y Then
        N_Red_Food = N_Red_Food + 1
        Food_Eated = Red_Food
        Red_Food.x = -1
        Red_Food.y = -1
    End If
    '��������
    score = N_White_Food + N_Red_Food * 5
     '�߳ɳ�
    If Food_Eated.x <> -1 And Food_Eated.y <> -1 Then
        Debug.Print Snakes(0).x & "," & Snakes(0).y
        Debug.Print Food_Eated.x & "," & Food_Eated.y
        mSnake_length = mSnake_length + 1
        ReDim Preserve Snakes(mSnake_length) As Point
        With Food_Eated
            .x = -1
            .y = -1
        End With
    End If
    '�ƶ�����
    For current_snake = mSnake_length To 1 Step -1
        If current_snake = 1 Then
            Snakes(current_snake) = mSnake_Head
        Else
            Snakes(current_snake) = Snakes(current_snake - 1)
        End If
    Next current_snake
    MoveSnake = True
End Function
'����ʳ��
Private Sub CreateFood(ByVal Color As Food_Color)
    Dim mFood As Point
    Do
        Randomize
        mFood.x = CInt(Rnd * (Game_Frame.Right - Game_Frame.Left - Game_Wide - Game_Wide) / Game_Wide) * Game_Wide + Game_Frame.Left + Game_Wide
        Randomize
        mFood.y = CInt(Rnd * (Game_Frame.Bottom - Game_Frame.Top - Game_Wide - Game_Wide) / Game_Wide) * Game_Wide + Game_Frame.Top + Game_Wide
    Loop Until FuInSnake(mFood) = False
    If Color = Food_Color_White Then
        White_Food = mFood
    ElseIf Color = Food_Color_Red Then
        Red_Food = mFood
    End If
End Sub

Private Function FuInSnake(ByRef White_Food As Point) As Boolean
    Dim i As Integer
    Dim n As Long
    n = UBound(Snakes)
    For i = 0 To n
        If White_Food.x = Snakes(i).x And White_Food.y = Snakes(i).y Then
            FuInSnake = True
            Exit Function
        End If
    Next i
End Function

Private Function HasRedim(ByRef x() As Point) As Boolean '�ж������Ƿ����
    On Error GoTo iEmpty
    Dim i As Long
    i = UBound(x)
    If i > 0 Then
        HasRedim = True
        Exit Function
    End If
iEmpty:
    HasRedim = False
    Err.Clear
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38
            mUser_Direction = User_Direction_Up
        Case 40
            mUser_Direction = User_direction_Down
        Case 37
            mUser_Direction = User_Direction_Left
        Case 39
            mUser_Direction = User_Direction_Right
    End Select
End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
    fps = 80
    Snake_Speed = 60 '���ٶ�
    Me.FontSize = 14
    Me.Font = "΢���ź�"
    With Red_Food
        .x = -1
        .y = -1
    End With
    mGame_State = Game_STATE_STOP
    With Game_Frame
        .Left = Game_Wide
        .Top = Game_Wide
        .Bottom = .Top + 30 * Game_Wide
        .Right = .Left + 32 * Game_Wide
    End With
    'ʳ��
    With Food_Eated
        .x = -1
        .y = -1
    End With
End Sub


Private Sub ��Ϸ_Click()
    Dim i As Integer
    '��ʼ��
    If ��Ϸ.Caption = "��ʼ��Ϸ" Then
        mGame_State = Game_STATE_RUNNING
        ��Ϸ.Caption = "��ͣ��Ϸ"
        ReDim Snakes(3) As Point
        '����С��
        Snakes(0).x = CLng(Game_Frame.Right / 2 / Game_Wide) * Game_Wide
        Snakes(0).y = CLng(Game_Frame.Bottom / 2 / Game_Wide) * Game_Wide
        For i = 1 To 3
            Snakes(i).x = Snakes(i - 1).x + Game_Wide
            Snakes(i).y = Snakes(i - 1).y
        Next i
        mUser_Direction = User_Direction_Left 'С��������
        mSnake_Direction = Snake_Direction_Left
        CreateFood Food_Color_White    '����ʳ��
        Call Game_Loop
    ElseIf ��Ϸ.Caption = "��ͣ��Ϸ" Then
        ��Ϸ.Caption = "������Ϸ"
        mGame_State = Game_STATE_PAUSE
    ElseIf ��Ϸ.Caption = "������Ϸ" Then
        mGame_State = Game_STATE_RUNNING
        ��Ϸ.Caption = "��ͣ��Ϸ"
    End If
End Sub
'��Ϸѭ��
Private Sub Game_Loop()
    Dim lsTime As Long
    Dim nwTime As Long
    Dim ltime_speed As Long
    Dim ntime_speed As Long
    While DoEvents
        If mGame_State = Game_STATE_RUNNING Then
            'UI�滭ˢ��
            nwTime = timeGetTime()
            If nwTime - lsTime >= 1000 / fps Then
                lsTime = nwTime
                Me.Cls
                Call Game_Draw
                Call Frame_Draw
                Me.Refresh
            End If
            '�߲�ˢ��
            ntime_speed = timeGetTime()
            If ntime_speed - ltime_speed >= Snake_Speed Then
                ltime_speed = ntime_speed
                Call MoveSnake
            End If
        End If
        Sleep 1   '�ӳ��Խ����ڴ�
    Wend
End Sub

Private Sub Frame_Draw()
    Me.FillColor = vbBlack
    Me.ForeColor = vbBlack
    Me.Line (Game_Frame.Left, Game_Frame.Top)-(Game_Frame.Right, Game_Frame.Bottom), , B '���߽�
    Me.CurrentX = Game_Frame.Right + Game_Wide
    Me.CurrentY = Game_Frame.Top + Game_Wide
    Me.Print "�ܷ֣�" & score
End Sub

Private Sub Game_Draw()
    Dim n As Integer
    Dim i As Integer
    n = UBound(Snakes) '����
    For i = 0 To n
        If i = 0 Then '����ͷ
            Me.FillColor = RGB(102, 205, 170)
            Me.ForeColor = RGB(102, 205, 170)
        Else
            Me.FillColor = RGB(0, 255, 255)
            Me.ForeColor = RGB(0, 255, 255)
        End If
        Me.Line (Snakes(i).x, Snakes(i).y)-(Snakes(i).x + Game_Wide, Snakes(i).y + Game_Wide), , BF
    Next i
    Me.FillColor = RGB(255, 215, 0)
    Me.ForeColor = RGB(255, 215, 0)
    Me.Line (White_Food.x, White_Food.y)-(White_Food.x + Game_Wide, White_Food.y + Game_Wide), , BF '����ʳ��
    If Red_Food.x <> -1 And Red_Food.y <> -1 Then '����ʳ��
        Me.FillColor = RGB(255, 0, 0)
        Me.ForeColor = RGB(255, 0, 0)
        Me.Line (Red_Food.x, Red_Food.y)-(Red_Food.x + Game_Wide, Red_Food.y + Game_Wide), , BF
    End If
End Sub

Private Sub �˳���Ϸ_Click()
    mGame_State = Game_STATE_STOP
    Unload Me
End Sub
