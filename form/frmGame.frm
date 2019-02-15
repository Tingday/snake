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
   Begin VB.Menu ��Ϸ���� 
      Caption         =   "��Ϸ����"
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
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

'��Ϸȫ��
Dim Game_Frame As RECTL '��Χǽ
Dim fps As Long '��Ϸˢ��Ƶ�ʣ���/�룩
Const Game_Wide = 15 '������
'�߱���
Dim Snakes() As Point '�߱���
Dim Food As Point '��ͨʳ��
Dim Food_Red As Point '����ʮ��
Dim Food_Eated As Point
Dim Snake_Speed As Long '�ߵ��ٶ�(����/Game_Wide)
'��Ϸ״̬
Dim mGame_State As Game_Status
Dim mSnake_Direction As Snake_Direction
Dim mUser_Direction As User_Direction
'��Ϸ�������
Dim score As Long '�ܵ÷� = ��ͨʳ��÷� + ����ʳ�� * 5
Dim N_Food_White As Long
Dim N_Food_Red As Long
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
    Dim Counter_i As Integer '������i
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
    For Counter_i = 1 To mSnake_length
        If Snakes(0).x = Snakes(Counter_i).x And Snakes(0).y = Snakes(Counter_i).y Then
            mGame_State = Game_STATE_STOP
            ��Ϸ.Caption = "��ʼ��Ϸ"
            MsgBox "��Ϸ������"
            Exit Function
        End If
    Next Counter_i
    If Snakes(0).x = Food.x And Snakes(0).y = Food.y Then 'ʳ����ײ���
        N_Food_White = N_Food_White + 1
        Food_Eated = Food
        CreateFood Food_Color_White
        If N_Food_White Mod 8 = 0 Then CreateFood Food_Color_Red 'ÿ��8��ʳ��ͳ���һ����ɫʳ��
    End If
    If Snakes(0).x = Food_Red.x And Snakes(0).y = Food_Red.y Then '��ʳ����ײ���
        N_Food_Red = N_Food_Red + 1
        Food_Eated = Food_Red
        Food_Red.x = -1
        Food_Red.y = -1
    End If
    score = N_Food_White + N_Food_Red * 5 '��������
    If Food_Eated.x <> -1 And Food_Eated.y <> -1 Then '�߳ɳ�
        For Counter_i = 0 To mSnake_length
            If Snakes(Counter_i).x = Food_Eated.x And Snakes(Counter_i).y = Food_Eated.y Then
                mSnake_length = mSnake_length + 1
                ReDim Preserve Snakes(mSnake_length) As Point
                With Food_Eated
                    .x = -1
                    .y = -1
                End With
            End If
        Next Counter_i
    End If
    For Counter_i = mSnake_length To 1 Step -1 'Ч��ʵ��
        If Counter_i = 1 Then
            Snakes(Counter_i) = mSnake_Head
        Else
            Snakes(Counter_i) = Snakes(Counter_i - 1)
        End If
    Next Counter_i
    MoveSnake = True
End Function

Private Sub CreateFood(ByVal Color As Food_Color)
    Dim mFood As Point
    Do
        Randomize
        mFood.x = CInt(Rnd * (Game_Frame.Right - Game_Frame.Left - Game_Wide - Game_Wide) / Game_Wide) * Game_Wide + Game_Frame.Left + Game_Wide
        Randomize
        mFood.y = CInt(Rnd * (Game_Frame.Bottom - Game_Frame.Top - Game_Wide - Game_Wide) / Game_Wide) * Game_Wide + Game_Frame.Top + Game_Wide
    Loop Until FuInSnake(mFood) = False
    If Color = Food_Color_White Then
        Food = mFood
    ElseIf Color = Food_Color_Red Then
        Food_Red = mFood
    End If
End Sub

Private Function FuInSnake(ByRef Food As Point) As Boolean
    Dim i As Integer
    Dim n As Long
    n = UBound(Snakes)
    For i = 0 To n
        If Food.x = Snakes(i).x And Food.y = Snakes(i).y Then
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
    Snake_Speed = 100 '���ٶ�
    With Food_Red
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
        Sleep 50   '�ӳ��Խ����ڴ�
    Wend
End Sub

Private Sub Frame_Draw()
    Me.FillColor = vbBlack
    Me.ForeColor = vbBlack
    Me.Line (Game_Frame.Left, Game_Frame.Top)-(Game_Frame.Right, Game_Frame.Bottom), , B '���߽�
    Me.CurrentX = Game_Frame.Right + Game_Wide
    Me.CurrentY = Game_Frame.Top + Game_Wide
    Me.Font = "΢���ź�"
    Me.FontSize = 14
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
    Me.FillColor = RGB(255, 255, 0)
    Me.ForeColor = RGB(255, 255, 0)
    Me.Line (Food.x, Food.y)-(Food.x + Game_Wide, Food.y + Game_Wide), , BF '����ʳ��
    If Food_Red.x <> -1 And Food_Red.y <> -1 Then '����ʳ��
        Me.FillColor = RGB(255, 0, 0)
        Me.ForeColor = RGB(255, 0, 0)
        Me.Line (Food_Red.x, Food_Red.y)-(Food_Red.x + Game_Wide, Food_Red.y + Game_Wide), , BF
    End If
End Sub

Private Sub �˳���Ϸ_Click()
    mGame_State = Game_STATE_STOP
    Unload Me
End Sub

Private Sub ��Ϸ����_Click()
    frmRule.Show , Me
End Sub

Private Sub ����_Click()
    frmAbout.Show , Me
End Sub
