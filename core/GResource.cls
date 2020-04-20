VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As clsid) As Long
Public HotLoad As Boolean
Dim AssetsFolder() As String, FolderParam1() As Variant, FolderParam2() As Variant
Public SearchIndex As Long
Public NoLoadCrash As Boolean, NoLoadDirection As Boolean
Dim Mems() As GMem
Dim Effect As Long, effectw As Long, effecth As Long, effectimg As Long, effectindex As Integer
Dim DisposeMark As Boolean
Private Sub Class_Initialize()
    ReDim Mems(0)
    ReDim AssetsFolder(0): ReDim FolderParam1(0): ReDim FolderParam2(0)
    SearchIndex = 1
    NoLoadDirection = True
End Sub
Public Sub ClipCircle(n)
    Dim Index As Integer
    
    If TypeName(n) = "String" Then
        Index = GetImage(n)
    Else
        Index = n
    End If

    Dim path As Long, bmpGraph As Long, BMP As Long, BMP2 As Long, tempGraph As Long
    Dim W As Long, H As Long
    
    GdipGetImageWidth Mems(Index).ImgHwnd, W
    GdipGetImageHeight Mems(Index).ImgHwnd, H

    GdipCreateBitmapFromScan0 W, H, ByVal 0, PixelFormat32bppARGB, ByVal 0, BMP2
    GdipCreateBitmapFromScan0 W * 2, H * 2, ByVal 0, PixelFormat32bppARGB, ByVal 0, BMP

    GdipGetImageGraphicsContext BMP2, tempGraph
    GdipGetImageGraphicsContext BMP, bmpGraph
    GdipSetSmoothingMode tempGraph, SmoothingModeAntiAlias
    GdipSetSmoothingMode bmpGraph, SmoothingModeAntiAlias
    
    PoolCreatePath FillModeWinding, path
    GdipAddPathEllipseI path, 1, 1, W * 2 - 2, H * 2 - 2
    GdipSetClipPath bmpGraph, path, CombineModeReplace
    
    GdipDrawImageRectI bmpGraph, Mems(Index).ImgHwnd, 0, 0, W * 2, H * 2
    GdipResetClip bmpGraph
    
    GdipSetInterpolationMode tempGraph, InterpolationModeHighQualityBilinear
    
    GdipDrawImageRectI tempGraph, BMP, 0, 0, W - 1, H - 1

    GdipDeletePath path
    PoolDeleteGraphics bmpGraph
    PoolDisposeImage BMP
    PoolDeleteGraphics tempGraph
    PoolDisposeImage Mems(Index).ImgHwnd
    
    Mems(Index).ImgHwnd = BMP2
    
    Call ReDraw(Index)
End Sub
Private Sub ApplyEffect(n, Kind As GdipEffectType)
    If XPMode Then Exit Sub

    Dim Index As Integer
    
    If TypeName(n) = "String" Then
        Index = GetImage(n)
    Else
        Index = n
    End If
    
    effectimg = Mems(Index).ImgHwnd
    
    PoolCreateEffect2 Kind, Effect
    
    effectindex = Index
End Sub
Private Sub ApplyEffectStep2()
    If XPMode Then Exit Sub

    GdipGetImageWidth effectimg, effectw: GdipGetImageHeight effectimg, effecth
    GdipBitmapApplyEffect effectimg, Effect, NewRectL(0, 0, effectw, effecth), 0, 0, 0
    
    PoolDeleteEffect Effect '垃圾处理
    
    Call ReDraw(effectindex)
End Sub
Public Sub ApplyTintEffect(n, amount As Long, hue As Long)
    If XPMode Then Exit Sub

    Dim e As TintParams
    e.amount = amount
    e.hue = hue
    
    ApplyEffect n, Tint
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplySharpenEffect(n, amount As Single, Radius As Single)
    If XPMode Then Exit Sub

    Dim e As SharpenParams
    e.amount = amount
    e.Radius = Radius

    ApplyEffect n, Sharpen
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyBlurEffect(n, Radius As Long, edge As Integer)
    If XPMode Then Exit Sub

    Dim e As BlurParams
    e.Radius = Radius
    e.expandEdge = edge

    ApplyEffect n, Blur
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyBrightnessContrastEffect(n, brightnessLevel As Long, contrastLevel As Long)
    If XPMode Then Exit Sub

    Dim e As BrightnessContrastParams
    e.brightnessLevel = brightnessLevel
    e.contrastLevel = contrastLevel

    ApplyEffect n, BrightnessContrast
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyColorBalanceEffect(n, cyanRed As Long, magentaGreen As Long, yellowBlue As Long)
    If XPMode Then Exit Sub

    Dim e As ColorBalanceParams
    e.cyanRed = cyanRed
    e.magentaGreen = magentaGreen
    e.yellowBlue = YellowGreen

    ApplyEffect n, ColorBalance
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyColorCurveEffect(n, adjustment As CurveAdjustments, adjustValue As Long, channel As CurveChannel)
    If XPMode Then Exit Sub

    Dim e As ColorCurveParams
    e.adjustment = adjustment
    e.adjustValue = adjustValue
    e.channel = channel

    ApplyEffect n, ColorCurve
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyColorMatrixEffect(n, m() As Single)
    If XPMode Then Exit Sub

    Dim e As ColorMatrix
    For I = 0 To 4
        For S = 0 To 4
            e.m(I, S) = m(I, S)
        Next
    Next
    ApplyEffect n, ColorMatrix
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyHueSaturationLightnessEffect(n, hueLevel As Long, lightnessLevel As Long, saturationLevel As Long)
    If XPMode Then Exit Sub

    Dim e As HueSaturationLightnessParams
    e.hueLevel = hueLevel
    e.lightnessLevel = lightnessLevel
    e.saturationLevel = saturationLevel

    ApplyEffect n, HueSaturationLightness
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyLevelsEffect(n, highlight As Long, midtone As Long, Shadow As Long)
    If XPMode Then Exit Sub

    Dim e As LevelsParams
    e.highlight = highlight
    e.midtone = midtone
    e.Shadow = Shadow

    ApplyEffect n, levels
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Private Sub ReDraw(I As Integer)
    If Mems(I).GIF.Count > 0 Then Exit Sub
    
    Dim g As Long
    PoolCreateFromHdc Mems(I).Hwnd, g
    GdipSetSmoothingMode g, SmoothingModeAntiAlias
    GdipGraphicsClear g, 0
    GdipDrawImageRect g, Mems(I).ImgHwnd, 0, 0, Mems(I).W, Mems(I).H
    PoolDeleteGraphics g
    
    For S = 0 To 3
        PoolCreateFromHdc Mems(I).Imgs(S), g
        GdipSetSmoothingMode g, SmoothingModeAntiAlias
        GdipGraphicsClear g, 0
        GdipDrawImageRect g, Mems(I).ImgHwnd, IIf(S = 1 Or S = 3, Mems(I).W, 0), IIf(S = 2 Or S = 3, Mems(I).H, 0), Mems(I).W * IIf(S = 1 Or S = 3, -1, 1), Mems(I).H * IIf(S = 2 Or S = 3, -1, 1)
        PoolDeleteGraphics g
    Next
    
    Dim tempCo As Long, tCo(3) As Byte
    Dim iW As Long, iH As Long

    ReDim Preserve ColorLists(UBound(ColorLists) + 1)
    ReDim ColorLists(UBound(ColorLists)).IsAlpha(Mems(I).W + 1, Mems(I).H + 1)
    
    DC = Mems(I).Imgs(0)
    GdipGetImageWidth Mems(I).ImgHwnd, iW: GdipGetImageHeight Mems(I).ImgHwnd, iH
    For X = 0 To Mems(I).W - 1
        For Y = 0 To Mems(I).H - 1
            '取得像素点颜色
            GdipBitmapGetPixel Mems(I).ImgHwnd, X / (Mems(I).W - 1) * iW, Y / (Mems(I).H - 1) * iH, tempCo
            CopyMemory tCo(0), tempCo, 4
            
            If tCo(3) = 255 Then ColorLists(UBound(ColorLists)).IsAlpha(X, Y) = True
        Next
        DoEvents
    Next
    
    Mems(I).CrashIndex = UBound(ColorLists)
End Sub
Public Sub Add(mem As Long, Kind As Integer, Name As String, Optional W, Optional H, Optional Folder As String = "")
    ReDim Preserve Mems(UBound(Mems) + 1)
    With Mems(UBound(Mems))
        .Hwnd = mem
        .Kind = Kind
        .Name = Name
        .Folder = Folder
        If Not IsMissing(W) Then .W = W
        If Not IsMissing(H) Then .H = H
    End With
End Sub
Public Property Get ImgCount() As Long
    ImgCount = UBound(Mems)
End Property
Public Property Get ImgSize(n, sizeIndex As imgIndex)
    Dim Index As Integer
    
    If TypeName(n) = "String" Then
        Index = GetImage(n)
    Else
        Index = n
    End If

    Select Case sizeIndex
        Case 0
            ImgSize = Mems(Index).W
        Case 1
            ImgSize = Mems(Index).H
        Case 2
            ImgSize = Mems(Index).GIF.Count
    End Select
End Property
Public Sub Dispose()
    If DisposeMark Then Exit Sub
    Exit Sub
    For I = 1 To UBound(Mems)
        If Not Mems(I).copyed Then
            Select Case Mems(I).Kind
                Case 0 'DC
                    If Mems(I).Hwnd <> 0 Then DeleteObject Mems(I).Hwnd
                    If Mems(I).ImgHwnd <> 0 Then PoolDisposeImage Mems(I).ImgHwnd
                    For S = 0 To 3
                        If Mems(I).Imgs(S) <> 0 Then DeleteObject Mems(I).Imgs(S)
                    Next
                Case 1 'Graphics
                    PoolDeleteGraphics Mems(I).Hwnd
            End Select
            If Mems(I).GIF.Count > 0 Then
                For S = 0 To Mems(I).GIF.Count
                    DeleteObject Mems(I).GIF.frames(S)
                Next
            End If
        End If
    Next
    
    DisposeMark = True
End Sub
Public Property Get ResourceGIFHandle(Index As Integer, frame As Integer) As Long
    ResourceGIFHandle = Mems(Index).GIF.frames(frame)
End Property
Public Property Get ResourceGIFCurrentHandle(Index As Integer) As Long
    ResourceGIFCurrentHandle = Mems(Index).GIF.frames(Mems(Index).GIF.tick)
End Property
Public Property Get ResourceGIFTick(Index As Integer) As Long
    ResourceGIFTick = Mems(Index).GIF.tick
End Property
Public Sub GIFTick(Index As Integer)
    Mems(Index).GIF.Time = GetTickCount
    Mems(Index).GIF.tick = Mems(Index).GIF.tick + 1
    If Mems(Index).GIF.tick > Mems(Index).GIF.Count Then Mems(Index).GIF.tick = 0
End Sub
Public Property Get ResourceGIFCount(Index As Integer) As Long
    ResourceGIFCount = Mems(Index).GIF.Count
End Property
Public Property Get ResourceGIFTime(Index As Integer) As Long
    ResourceGIFTime = Mems(Index).GIF.Time
End Property
Public Property Get ResourceName(Index As Integer) As Long
    ResourceName = Mems(Index).Name
End Property
Public Property Get ResourceHandle(Index As Integer) As Long
    ResourceHandle = Mems(Index).Hwnd
End Property
Public Property Get ResourceCrashIndex(Index As Integer) As Long
    ResourceCrashIndex = Mems(Index).CrashIndex
End Property
Public Property Get ResourceHandle2(Index As Integer, Direction As ImgDirection) As Long
    ResourceHandle2 = Mems(Index).Imgs(Direction)
End Property
Public Property Get ResourceImageHandle(Index As Integer) As Long
    ResourceImageHandle = Mems(Index).ImgHwnd
End Property
Public Function GetImage(ByVal Name As String) As Integer
    For I = SearchIndex To UBound(Mems)
        If Mems(I).Kind = 0 Then
            If Mems(I).Name = Name Then GetImage = I: Exit For
        End If
    Next
    
    If HotLoad Then
        If GetImage = 0 Then
            For I = 1 To UBound(AssetsFolder)
                If IsExitAFile(AssetsFolder(I) & Name) Then
                    newImage AssetsFolder(I) & Name, FolderParam1(I), FolderParam2(I)
                    GetImage = UBound(Mems)
                    Exit For
                End If
            Next
        End If
    End If
End Function
Public Sub newImage(path As String, Optional arg1, Optional arg2, Optional Name As String = "")
    'when arg1 is missing ; i.w = i.w
    'when arg2 is missing but arg1 has content ; i.w = i.w * arg1
    'other ; i.w = arg1
    Dim si As Integer, I As Long, W As Long, H As Long, DC As Long, g As Long
    If IsMissing(arg1) Then si = 0
    If IsMissing(arg2) And (Not IsMissing(arg1)) Then si = 1
    If (Not IsMissing(arg2)) And (Not IsMissing(arg1)) Then si = 2
    
    PoolCreateBitmapFromFile StrPtr(path), I
    GdipGetImageWidth I, W: GdipGetImageHeight I, H
    
    If si = 1 Then W = Round(W * arg1): H = Round(H * arg1)
    If si = 2 Then W = arg1: H = arg2
    
    If Scales <> 1 Then
        W = W * Scales
        H = H * Scales
    End If
    
    Dim gC As clsid, gL As Long, Index As Long, temp() As String, CrashPath As Long
    CLSIDFromString StrPtr("{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"), gC
    GdipImageGetFrameCount I, gC, gL
    
    If gL > 0 Then  '动态GIF支持
        temp = Split(path, "\")
        Add 0, 0, temp(UBound(temp)), W, H, temp(UBound(temp) - 1)
        With Mems(UBound(Mems))
            .GIF.Count = gL
            ReDim .GIF.frames(gL)
            For Index = 0 To gL
                GdipImageSelectActiveFrame I, gC, Index
                DC = CreateCDC(W, H)
                PoolCreateFromHdc DC, g
                GdipDrawImageRect g, I, 0, 0, W, H
                PoolDeleteGraphics g
                .GIF.frames(Index) = DC
            Next
        End With
        
        Mems(UBound(Mems)).ImgHwnd = I
        Exit Sub
    End If
    
    temp = Split(path, "\")
    ReDim Preserve Mems(UBound(Mems) + 1)
    With Mems(UBound(Mems))
        .Name = IIf(Name = "", temp(UBound(temp)), Name)
        .Folder = temp(UBound(temp) - 1)
        .W = W
        .H = H
        .ImgHwnd = I
    End With
    
    For S = 0 To 3
        If NoLoadDirection And S > 0 Then Exit For
        DC = CreateCDC(W, H)
        PoolCreateFromHdc DC, g
        GdipDrawImageRect g, I, IIf(S = 1 Or S = 3, W, 0), IIf(S = 2 Or S = 3, H, 0), W * IIf(S = 1 Or S = 3, -1, 1), H * IIf(S = 2 Or S = 3, -1, 1)
        PoolDeleteGraphics g
        Mems(UBound(Mems)).Imgs(S) = DC
    Next
    
    If W > GW * 0.5 Or H > GH * 0.5 Or NoLoadCrash Then Exit Sub
    
    Dim tempCo As Long, tCo(3) As Byte
    Dim iW As Long, iH As Long

    ReDim Preserve ColorLists(UBound(ColorLists) + 1)
    ReDim ColorLists(UBound(ColorLists)).IsAlpha(W + 1, H + 1)
    
    DC = Mems(UBound(Mems)).Imgs(0)
    GdipGetImageWidth I, iW: GdipGetImageHeight I, iH
    For X = 0 To W - 1
        For Y = 0 To H - 1
            '取得像素点颜色
            GdipBitmapGetPixel I, X / (W - 1) * iW, Y / (H - 1) * iH, tempCo
            CopyMemory tCo(0), tempCo, 4

            If tCo(3) = 255 Then ColorLists(UBound(ColorLists)).IsAlpha(X, Y) = True
        Next
        DoEvents
    Next
    
    Mems(UBound(Mems)).CrashIndex = UBound(ColorLists)
End Sub
Public Sub NewImages(path As String, Optional arg1, Optional arg2)

    If HotLoad Then
        ReDim Preserve AssetsFolder(UBound(AssetsFolder) + 1)
        ReDim Preserve FolderParam1(UBound(FolderParam1) + 1)
        ReDim Preserve FolderParam2(UBound(FolderParam2) + 1)
        If Right(path, 1) <> "\" Then path = path & "\"
        AssetsFolder(UBound(AssetsFolder)) = path
        FolderParam1(UBound(FolderParam1)) = arg1
        FolderParam2(UBound(FolderParam2)) = arg2
        Exit Sub
    End If
    
    Dim f As String, C As Long
    If Right(path, 1) <> "\" Then path = path & "\"
    
    Dim Index As Integer, Tree As AssetsTree
    Tree.path = path: ReDim Tree.Files(0)
    Index = FindAssetsTree(path, arg1, arg2)
    
    f = Dir(path)
    Do While f <> ""
        C = C + 1
        ReDim Preserve Tree.Files(UBound(Tree.Files) + 1)
        f = Dir()
    Loop
    If Index = 0 Then
        PreLoadCount = PreLoadCount + C
    Else
        Tree = AssetsTrees(Index)
        For I = 1 To UBound(Tree.Files)
            ReDim Preserve Mems(UBound(Mems) + 1)
            Mems(UBound(Mems)) = Tree.Files(I)
            Mems(UBound(Mems)).copyed = True
        Next
        ReLoadCount = ReLoadCount + UBound(Tree.Files)
        Exit Sub
    End If
    
    f = Dir(path)
    I = 1
    Do While f <> ""
        If IsMissing(arg1) Then
            newImage path & f
        ElseIf IsMissing(arg2) Then
            newImage path & f, arg1
        Else
            newImage path & f, arg1, arg2
        End If
        Tree.Files(I) = Mems(UBound(Mems))
        I = I + 1
        f = Dir()
        DoEvents
        LoadedCount = LoadedCount + 1
    Loop
    
    If Index = 0 Then AddAssetsTree Tree, arg1, arg2
End Sub
