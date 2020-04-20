VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Piafxxk"
   ClientHeight    =   5268
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5268
   ScaleWidth      =   9480
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command5 
      Caption         =   "SetTitle"
      Height          =   465
      Left            =   8016
      TabIndex        =   13
      Top             =   1800
      Width           =   1128
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ÊÓÆµÕ¹Ê¾"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6912
      TabIndex        =   12
      Top             =   4008
      Value           =   1  'Checked
      Width           =   2268
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "ÔÚÕâÀï°´¼üÅÌÖ±½ÓÑÝ×à"
      Top             =   840
      Width           =   8880
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   8136
      TabIndex        =   10
      Text            =   "250"
      Top             =   1272
      Width           =   984
   End
   Begin VB.CommandButton Command4 
      Caption         =   "-"
      Height          =   465
      Left            =   4464
      TabIndex        =   9
      Top             =   1464
      Width           =   456
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+"
      Height          =   465
      Left            =   3864
      TabIndex        =   8
      Top             =   1464
      Width           =   456
   End
   Begin VB.ListBox List1 
      Height          =   2688
      Left            =   264
      TabIndex        =   7
      Top             =   2352
      Width           =   6276
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   336
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1536
      Width           =   3492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Random"
      Height          =   465
      Left            =   6888
      TabIndex        =   2
      Top             =   4560
      Width           =   1128
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   465
      Left            =   8280
      TabIndex        =   1
      Top             =   4560
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Text            =   $"Form1.frx":0000
      Top             =   528
      Width           =   8880
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÒôÔ´"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   240
      TabIndex        =   6
      Top             =   1992
      Width           =   408
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÒôÔ´"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   240
      TabIndex        =   4
      Top             =   1152
      Width           =   408
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Òô·û´®"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   612
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type SoundSource
    Notes(23) As Long
    Name As String
End Type
Dim SS() As SoundSource
Dim TempS() As SoundSource
Dim p As Piano
Dim sMark As Boolean
Sub SetPlayRate(ByVal NRate As Single)
    BASS_ChannelSetAttribute SongHandle, BASS_ATTRIB_FREQ, 44100 * NRate
End Sub

Private Sub PlaySource(Name As String, index As Integer)
    For I = 1 To UBound(SS)
        If SS(I).Name = Name Then
            Dim Hwnd As Long
            'Hwnd = SS(i).Notes(index)
            ReDim Preserve TempS(UBound(TempS) + 1)
            Hwnd = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.path & "\source\" & SS(I).Name & "\source.wav"), 0, 0, 0)
            BASS_ChannelSetAttribute Hwnd, BASS_ATTRIB_FREQ, 44100 * (1 + Cubic(index / 23, 0, 1, 1, 1) * 2) * 0.2
            BASS_ChannelSetAttribute Hwnd, BASS_ATTRIB_VOL, 0.5 + Rnd * 0.4 - 0.2
            TempS(UBound(TempS)).Notes(index) = Hwnd
            TempS(UBound(TempS)).Name = SS(I).Name
            BASS_ChannelPlay Hwnd, True
            Exit For
        End If
    Next
End Sub
Private Sub StopSource(Name As String, index As Integer)
    For I = 1 To UBound(SS)
        If SS(I).Name = Name Then
            BASS_ChannelStop SS(I).Notes(index)
            For S = 1 To UBound(TempS)
                If S > UBound(TempS) Then Exit For
                If TempS(S).Name = Name And TempS(S).Notes(index) <> 0 Then
                    'BASS_ChannelStop TempS(s).Notes(index)
                    'BASS_MusicFree TempS(s).Notes(index)
                    'TempS(s) = TempS(UBound(TempS))
                    'ReDim Preserve TempS(UBound(TempS) - 1)
                    's = s - 1
                End If
            Next
            Exit For
        End If
    Next
End Sub
Public Sub PlayNote(Note As String)
    Dim temp As Integer
    Select Case Note
        Case "0": temp = 97
        Case "1": temp = 0
        Case "2": temp = 2
        Case "3": temp = 4
        Case "4": temp = 5
        Case "5": temp = 7
        Case "6": temp = 9
        Case "7": temp = 11
        Case "Q": temp = 12
        Case "W": temp = 14
        Case "E": temp = 16
        Case "R": temp = 17
        Case "T": temp = 19
        Case "Y": temp = 21
        Case "U": temp = 23
        Case "A": temp = 1
        Case "S": temp = 3
        Case "D": temp = 6
        Case "F": temp = 8
        Case "G": temp = 10
        Case "H": temp = 13
        Case "J": temp = 15
        Case "K": temp = 18
        Case "L": temp = 20
        Case ";": temp = 22
        Case " ": temp = 98
        Case "~": temp = 99
        Case "Z"
            List1.Clear
            If Combo1.ListCount > 1 Then List1.AddItem Combo1.List(0)
            temp = -1
        Case "X"
            List1.Clear
            If Combo1.ListCount > 2 Then List1.AddItem Combo1.List(1)
            temp = -1
        Case "C"
            List1.Clear
            If Combo1.ListCount > 3 Then List1.AddItem Combo1.List(2)
            temp = -1
        Case "V"
            List1.Clear
            If Combo1.ListCount > 4 Then List1.AddItem Combo1.List(3)
            temp = -1
        Case "B"
            List1.Clear
            If Combo1.ListCount > 5 Then List1.AddItem Combo1.List(4)
            temp = -1
        Case "N"
            List1.Clear
            If Combo1.ListCount > 6 Then List1.AddItem Combo1.List(5)
            temp = -1
        Case Else: temp = -1
    End Select
    If temp > -1 Then
        If p.lPlay <> -1 And temp <> 99 Then
            For S = 0 To List1.ListCount - 1
                StopSource List1.List(S), p.lPlay
            Next
        End If
        p.Play (temp)
        If temp = 98 Then
            For S = 0 To List1.ListCount - 1
                PlaySource List1.List(S), p.lPlay
            Next
        End If
        If temp <= 90 Then
            For S = 0 To List1.ListCount - 1
                PlaySource List1.List(S), temp
            Next
        End If
    End If
    
    
    For S = 0 To List1.ListCount - 1
        GameWindow.GamePage.Add temp, List1.List(S)
    Next
    
End Sub

Private Sub Check1_Click()
    GameWindow.Visible = (Check1.value = 1)
End Sub

Private Sub Command1_Click()
    If Command1.Caption = "Stop" Then sMark = True: Exit Sub
    
    sMark = False
    Command1.Caption = "Stop"
    Command2.Enabled = False
    Text1.Locked = True
    
    For I = 1 To UBound(TempS)
        For S = 0 To UBound(TempS(I).Notes)
            If TempS(I).Notes(S) <> 0 Then
                BASS_StreamFree TempS(I).Notes(S)
            End If
        Next
    Next
    ReDim TempS(0)
    
    Dim temp As Integer, per As Long
    per = Val(Text2.Text)
    Dim lt As Long, st As Long
    
    Text1.SetFocus
    
    For I = 1 To Len(Text1.Text)
        If sMark Then Exit For
        PlayNote (Mid(Text1.Text, I, 1))
        Text1.SelStart = I
        Text1.SelLength = 1
        lt = GetTickCount
        If temp = 98 Then
            st = per * 4
        Else
            st = per
        End If
        Do While GetTickCount - lt < st
            ECore.Display: DoEvents
        Loop
        DoEvents
    Next
    'Q154Q114QWQ54Q1654345EWWQ2Q654563554341Q654345EWWQ2Q65456355434Q454J45464566545Q454J4546456654
    'UT£ºUpon once a time
    
    Command1.Caption = "Play"
    Command2.Enabled = True
    Text1.Locked = False
End Sub

Private Sub Command2_Click()
    Dim R As String, n As String
    
    Randomize
    For I = 1 To Int(Rnd * 50) + 70
        Randomize
        n = Int(Rnd * 9)
        If n > 7 Then n = 0
        R = R & n
    Next
    
    Text1.Text = R
End Sub

Private Sub Command3_Click()
    List1.AddItem Combo1.List(Combo1.ListIndex)
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command5_Click()
    GameWindow.GamePage.Title = InputBox("")
End Sub

Private Sub Form_Load()
    BASS_Init -1, 44100, BASS_DEVICE_3D, Me.Hwnd, 0

    Text1.Text = "C5~~~Q~~~W~~~J~~~~~~~~~~~R~~~W~~~~~~~~~~~J~~~XQ~~~~~~~~~~~~~~CWQ~~~7~~~Q~~W~~RJ~~~W~~~J~~~R~~~T~~~~~~~T~~~~~~~T~~~~~~~~~~~~~~~~~~~~~~~R~~~T~RTL~~~~~~~~~~~~~~~W~~~~~~~J~~~R~JRT~~~~~~~~~~~~~~~XQ~~~~~~~Q~~~W~~~J~~~~~~~CJ~~~R~~~W~~~~~~~XW~~~CJ~~~Q~~~~~~~~~~~~~~~~~~~~~~~R~~~T~~~L~~~~~~~~~~~~~~~W~~~~~~~J~~~R~~~T~~~~~~~~~~~~~~~Q~~~~~~~Q~~~W~~~J~~~~~~~J~~~R~~~W~~~~~~~W~~~J~~~Q~~~~~~~~~~~~~~~~~~~~~~~"
    
    Set p = New Piano
    p.Init
    
    Dim f As String
    ReDim SS(0): ReDim TempS(0)
    f = Dir(App.path & "\source\", vbDirectory)
    Do While f <> ""
        If f <> "." And f <> ".." Then
            ReDim Preserve SS(UBound(SS) + 1)
            SS(UBound(SS)).Name = f
            For I = 0 To 23
                'Notes(i) = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & "\source\" & Combo1.List(Combo1.ListIndex) & "\" & i & ".wav"), 0, 0, 0)
                'Notes(i) = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & "\source\" & Combo1.List(Rnd * Combo1.ListCount) & "\source.wav"), 0, 0, 0)
                SS(UBound(SS)).Notes(I) = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.path & "\source\" & f & "\source.wav"), 0, 0, 0)
                BASS_ChannelSetAttribute SS(UBound(SS)).Notes(I), BASS_ATTRIB_FREQ, 44100 * (1 + Cubic(I / 23, 0, 1, 1, 1) * 2) * 0.2
            Next
            Combo1.AddItem f
        End If
        f = Dir(, vbDirectory)
        DoEvents
    Loop
    
    Combo1.ListIndex = 0
    Command3_Click
End Sub
Function Cubic(t As Single, arg0 As Single, arg1 As Single, arg2 As Single, arg3 As Single) As Single
    'Formula:B(t)=P_0(1-t)^3+3P_1t(1-t)^2+3P_2t^2(1-t)+P_3t^3
    'Attention:all the args must in this area (0~1)
    Cubic = (arg0 * ((1 - t) ^ 3)) + (3 * arg1 * t * ((1 - t) ^ 2)) + (3 * arg2 * (t ^ 2) * (1 - t)) + (arg3 * (t ^ 3))
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    BASS_Free
    p.Dispose
    On Error Resume Next
    Unload GameWindow
    End
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    PlayNote UCase(Chr(KeyAscii))
End Sub

Private Sub Text3_LostFocus()
    For I = 1 To UBound(TempS)
        For S = 0 To UBound(TempS(I).Notes)
            If TempS(I).Notes(S) <> 0 Then
                BASS_StreamFree TempS(I).Notes(S)
            End If
        Next
    Next
    ReDim TempS(0)
End Sub
