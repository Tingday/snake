VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "贪吃蛇"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   1  '所有者中心
   Begin VB.Menu 游戏 
      Caption         =   "开始游戏"
   End
   Begin VB.Menu 退出游戏 
      Caption         =   "退出游戏"
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------
'版本 2018.03.28 修改 作者：0yufan0 vb吧
'联系 woyufan@163.com 有问题请联系我
'本源码为本人编写，认真学习可以了解fps游戏编写
'-------------------------------------------
'---------------------------------------------以下为私有变量----------------------------------------
Option Explicit
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function GetInputState Lib "user32.dll" () As Long

'游戏全局
Dim Game_Frame As RECTL '外围墙
Dim fps As Long '游戏刷新频率（次/秒）
Const Game_Wide = 15 '网格宽度
'蛇变量
Dim Snakes() As Point '蛇本体
Dim White_Food As Point '普通食物
Dim Red_Food As Point '奖励十五
Dim Food_Eated As Point
Dim Snake_Speed As Long '蛇的速度(毫秒/Game_Wide)
'游戏状态
Dim mGame_State As Game_Status
Dim mSnake_Direction As Snake_Direction
Dim mUser_Direction As User_Direction
'游戏分数相关
Dim score As Long '总得分 = 普通食物得分 + 奖励食物 * 5
Dim N_White_Food As Long
Dim N_Red_Food As Long
'类型
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
'列表数据
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
'--------------------------------------------------------变量到此结束---------------------------------------
'--------------------------------------------------------函 数 过  程---------------------------------------
Private Function MoveSnake() As Boolean '移动蛇与判断
    Dim current_snake As Integer '计数器i
    Dim mSnake_Head As Point '蛇头坐标
    Dim mSnake_length As Long '蛇长度（单位1节蛇）
    mSnake_length = UBound(Snakes)
    mSnake_Head = Snakes(0) '获得蛇头
    Select Case mUser_Direction '根据方向移动，每次一个单位Game_Wide
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
    '穿墙效果实现
    If Snakes(0).y >= Game_Frame.Bottom Then Snakes(0).y = Game_Frame.Top
    If Snakes(0).y < Game_Frame.Top Then Snakes(0).y = Game_Frame.Bottom - Game_Wide
    If Snakes(0).x >= Game_Frame.Right Then Snakes(0).x = Game_Frame.Left
    If Snakes(0).x < Game_Frame.Left Then Snakes(0).x = Game_Frame.Right - Game_Wide
    '碰撞检测
    For current_snake = 1 To mSnake_length
        If Snakes(0).x = Snakes(current_snake).x And Snakes(0).y = Snakes(current_snake).y Then
            mGame_State = Game_STATE_STOP
            游戏.Caption = "开始游戏"
            MsgBox "游戏结束！"
            Exit Function
        End If
    Next current_snake
    '食物碰撞检测
    If Snakes(0).x = White_Food.x And Snakes(0).y = White_Food.y Then
        N_White_Food = N_White_Food + 1
        Food_Eated = White_Food
        CreateFood Food_Color_White
        If N_White_Food Mod 8 = 0 Then CreateFood Food_Color_Red '每吃8个食物就出现一个红色食物
    End If
    '红食物碰撞检测
    If Snakes(0).x = Red_Food.x And Snakes(0).y = Red_Food.y Then
        N_Red_Food = N_Red_Food + 1
        Food_Eated = Red_Food
        Red_Food.x = -1
        Red_Food.y = -1
    End If
    '分数计算
    score = N_White_Food + N_Red_Food * 5
     '蛇成长
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
    '移动蛇体
    For current_snake = mSnake_length To 1 Step -1
        If current_snake = 1 Then
            Snakes(current_snake) = mSnake_Head
        Else
            Snakes(current_snake) = Snakes(current_snake - 1)
        End If
    Next current_snake
    MoveSnake = True
End Function
'创建食物
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

Private Function HasRedim(ByRef x() As Point) As Boolean '判断蛇体是否存在
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
    Snake_Speed = 60 '蛇速度
    Me.FontSize = 14
    Me.Font = "微软雅黑"
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
    '食物
    With Food_Eated
        .x = -1
        .y = -1
    End With
End Sub


Private Sub 游戏_Click()
    Dim i As Integer
    '初始化
    If 游戏.Caption = "开始游戏" Then
        mGame_State = Game_STATE_RUNNING
        游戏.Caption = "暂停游戏"
        ReDim Snakes(3) As Point
        '创建小蛇
        Snakes(0).x = CLng(Game_Frame.Right / 2 / Game_Wide) * Game_Wide
        Snakes(0).y = CLng(Game_Frame.Bottom / 2 / Game_Wide) * Game_Wide
        For i = 1 To 3
            Snakes(i).x = Snakes(i - 1).x + Game_Wide
            Snakes(i).y = Snakes(i - 1).y
        Next i
        mUser_Direction = User_Direction_Left '小蛇向左走
        mSnake_Direction = Snake_Direction_Left
        CreateFood Food_Color_White    '创建食物
        Call Game_Loop
    ElseIf 游戏.Caption = "暂停游戏" Then
        游戏.Caption = "继续游戏"
        mGame_State = Game_STATE_PAUSE
    ElseIf 游戏.Caption = "继续游戏" Then
        mGame_State = Game_STATE_RUNNING
        游戏.Caption = "暂停游戏"
    End If
End Sub
'游戏循环
Private Sub Game_Loop()
    Dim lsTime As Long
    Dim nwTime As Long
    Dim ltime_speed As Long
    Dim ntime_speed As Long
    While DoEvents
        If mGame_State = Game_STATE_RUNNING Then
            'UI绘画刷新
            nwTime = timeGetTime()
            If nwTime - lsTime >= 1000 / fps Then
                lsTime = nwTime
                Me.Cls
                Call Game_Draw
                Call Frame_Draw
                Me.Refresh
            End If
            '蛇步刷新
            ntime_speed = timeGetTime()
            If ntime_speed - ltime_speed >= Snake_Speed Then
                ltime_speed = ntime_speed
                Call MoveSnake
            End If
        End If
        Sleep 1   '延迟以降低内存
    Wend
End Sub

Private Sub Frame_Draw()
    Me.FillColor = vbBlack
    Me.ForeColor = vbBlack
    Me.Line (Game_Frame.Left, Game_Frame.Top)-(Game_Frame.Right, Game_Frame.Bottom), , B '画边界
    Me.CurrentX = Game_Frame.Right + Game_Wide
    Me.CurrentY = Game_Frame.Top + Game_Wide
    Me.Print "总分：" & score
End Sub

Private Sub Game_Draw()
    Dim n As Integer
    Dim i As Integer
    n = UBound(Snakes) '画蛇
    For i = 0 To n
        If i = 0 Then '画蛇头
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
    Me.Line (White_Food.x, White_Food.y)-(White_Food.x + Game_Wide, White_Food.y + Game_Wide), , BF '画白食物
    If Red_Food.x <> -1 And Red_Food.y <> -1 Then '画红食物
        Me.FillColor = RGB(255, 0, 0)
        Me.ForeColor = RGB(255, 0, 0)
        Me.Line (Red_Food.x, Red_Food.y)-(Red_Food.x + Game_Wide, Red_Food.y + Game_Wide), , BF
    End If
End Sub

Private Sub 退出游戏_Click()
    mGame_State = Game_STATE_STOP
    Unload Me
End Sub
