VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GamePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Private Type NewNote
        Name As String
        index As Integer
        time As Long
        Direction As Integer
    End Type
    Dim Notes() As NewNote
    Dim ShineTime As Long
    Public Title As String
'======================================
Public Sub Add(index As Integer, Name As String)
    If index = -1 Then Exit Sub
    
    ShineTime = GetTickCount
    
    If index = 99 Then Exit Sub
    If index = 98 Then Exit Sub
    
    ReDim Preserve Notes(UBound(Notes) + 1)
    With Notes(UBound(Notes))
        Randomize
        .Direction = Int(Rnd * 4)
        .index = index
        .Name = Name
        .time = GetTickCount
    End With
End Sub
Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    Dim spro As Single
    spro = (GetTickCount - ShineTime) / 300
    If spro > 1 Then spro = 1
    spro = 1 - Cubic(spro, 0, 0, 0, 1)
    
    Page.Clear argb(255, 120 * spro + Abs((GetTickCount Mod 4000) - 2000) / 2000 * 120 * spro, 120 * spro + Abs((GetTickCount Mod 8000) - 4000) / 4000 * 120 * spro, 120 * spro + Abs((GetTickCount Mod 2800) - 1400) / 1400 * 120 * spro) '清空画布

    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    Dim pro As Single, w As Long, h As Long
    Dim delMark As Boolean, mode As Integer, opro As Single
    
    For I = 1 To UBound(Notes)
        If I > UBound(Notes) Then Exit For
        pro = (GetTickCount - Notes(I).time) / 400
        delMark = False
        If pro > 1 Then delMark = True: pro = 1
        opro = pro
        If pro < 0.5 Then
            pro = pro / 0.5: mode = 0
        Else
            pro = 1 - (pro - 0.5) / 0.5: mode = 1
        End If
        pro = Cubic(pro, 0, 1, 1, 1)
        
        w = Page.Res.ImgSize(Notes(I).Name & ".gif", imgGetWidth)
        h = Page.Res.ImgSize(Notes(I).Name & ".gif", imgGetHeight)
        
        Page.Paint 0, Notes(I).index / 22 * GW, GH - GH * opro + 1, GW / 22, GH * opro, argb(255 - 255 * opro, 255, 255, 255)
        
        Select Case Notes(I).Direction
            Case 0: Page.DrawImage Notes(I).Name & ".gif", -w / 2 + w * 1 * pro, GH / 2, alpha:=pro, pos:=posOnCenter, Direction:=DirNormal
            Case 1: Page.DrawImage Notes(I).Name & ".gif", GW + w / 2 - w * 1 * pro, GH / 2, alpha:=pro, pos:=posOnCenter, Direction:=DirVertical
            Case 2: Page.DrawImage Notes(I).Name & ".gif", GW / 2, -h / 2 + h * 1 * pro, alpha:=pro, pos:=posOnCenter, Direction:=DirVertical
            Case 3: Page.DrawImage Notes(I).Name & ".gif", GW / 2, GH + h / 2 - h * 1 * pro, alpha:=pro, pos:=posOnCenter, Direction:=DirHorizontalVertical
        End Select
        
        If delMark Then
            Notes(I) = Notes(UBound(Notes))
            I = I - 1
            ReDim Preserve Notes(UBound(Notes) - 1)
        End If
    Next
    
    For I = 1 To UBound(Notes)
        opro = (GetTickCount - Notes(I).time) / 400
        If opro > 1 Then opro = 1
        Page.Paint 0, Notes(I).index / 22 * GW, GH - GH * opro + 1, GW / 22, GH * opro, argb(255 - 255 * opro, 255, 255, 255)
    Next
    
    Page.Writes Title, GW / 2, GH - 100, 36, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold, BorderColor:=argb(255, 0, 0, 0), BorderSize:=2
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    ReDim Notes(0)
    Title = "你的鬼畜音乐"
    
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.NewImages App.path & "\assets"
    '创建页面
    ECore.Add Page, "GamePage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
