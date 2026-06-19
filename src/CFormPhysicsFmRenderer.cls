VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFormPhysicsFmRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare PtrSafe Function CreateDIBSection Lib "gdi32" (ByVal hdc As LongPtr, ByRef pbmi As BITMAPINFO, ByVal iUsage As Long, ByRef ppvBits As LongPtr, ByVal hSection As LongPtr, ByVal dwOffset As Long) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (ByRef lpPictDesc As PICTDESC, ByRef riid As GUID, ByVal fOwn As Long, ByRef lplpvObj As IPictureDisp) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc.dll" (ByVal IAccessible As Object, ByRef hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const GWL_STYLE As Long = -16, GWL_EXSTYLE As Long = -20
Private Const WS_EX_LAYERED = &H80000, WS_THICKFRAME = &H40000, WS_EX_DLGMODALFRAME = &H1&
Private Const LWA_COLORKEY = &H1
Private Const BI_RGB = 0, PICTYPE_BITMAP = 1
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type RGBA
    R As Byte
    G As Byte
    B As Byte
    A As Byte
End Type
Private Type POINTL
    X As Long
    Y As Long
End Type
Private Type LONGTYPE
    Value As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0) As RGBA
End Type
Private Type PICTDESC
    cbSizeofStruct As Long
    picType As Long
    hBitmap As LongPtr
    hPalette As LongPtr
    Reserved As LongPtr
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Implements ICFormPhysicsEx
Private WithEvents myCore As CFormPhysics
Attribute myCore.VB_VarHelpID = -1
Private IID As GUID, PIC As PICTDESC, BINFO As BITMAPINFO
Private ptL() As Long, ptR() As Long, ptT() As Long, ptB() As Long
Private myMask() As Long, myPX() As Byte, cl As RGBA, cl2 As RGBA
Private pw As Long, ph As Long, sw As Long, sh As Long, psize As Long
Private ubpX As Long
Private pdmg As Double
Private hWnd As LongPtr
Private Type CrackDir
    ox As Long
    oy As Long
    E  As Double
    lf As Double
    isFind As Boolean
End Type
Private Property Get Mask(ByRef X As Long, ByRef Y As Long) As Long
    If X < 0 Then X = 0
    If X > pw Then X = pw
    If Y < 0 Then Y = 0
    If Y > ph Then Y = ph
    Mask = myMask(X, Y)
End Property
Private Property Let Mask(ByRef X As Long, ByRef Y As Long, ByRef rhs As Long)
    If X < 0 Then X = 0
    If Y < 0 Then Y = 0
    If X > pw Then X = pw
    If Y > ph Then Y = ph
    myMask(X, Y) = rhs
End Property
Private Property Get PX(ByRef X As Long, ByRef Y As Long) As RGBA
    Dim i As Long: i = (pw * Y + X) * 4
    If i < 0 Then Exit Property
    If i > ubpX Then Exit Property
    With PX
        .B = myPX(i + 0)
        .G = myPX(i + 1)
        .R = myPX(i + 2)
    End With
End Property
Private Property Let PX(ByRef X As Long, ByRef Y As Long, ByRef tCL As RGBA)
    Dim i As Long: i = (pw * Y + X) * 4
    If i < 0 Then Exit Property
    If i > ubpX Then Exit Property
    With tCL
        myPX(i + 0) = .B
        myPX(i + 1) = .G
        myPX(i + 2) = .R
    End With
End Property
Private Sub UpdatePicture()
    Dim pict As IPictureDisp, pbits As LongPtr
    PIC.hBitmap = CreateDIBSection(0, BINFO, 0, pbits, 0, 0)
    CopyMemory pbits, VarPtr(myPX(0)), pw * ph * 4
    OleCreatePictureIndirect PIC, IID, True, pict
    Set myCore.mFrmRaw.Picture = pict
    myCore.mFrmRaw.Repaint
End Sub
Private Function FindMinEnergyDirection( _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal dx As Double, ByVal dy As Double) As CrackDir
    Dim ox As Long, oy As Long
    Dim nx As Long, ny As Long
    Dim bendE As Double, totalE As Double
    With FindMinEnergyDirection
        .E = 1E+30
        .isFind = False
    End With
    For ox = -1 To 1
        For oy = -1 To 1
            If Not (ox = 0 And oy = 0) Then
                nx = cx + ox
                ny = cy + oy
                If Mask(nx, ny) = 1 Then
                    bendE = (1# - (dx * ox + dy * oy))
                    With FindMinEnergyDirection
                        If PX(nx, ny).R > 5 Then
                            totalE = bendE
                            .lf = 1
                        Else
                            totalE = bendE + 0.5
                            .lf = 0.1
                        End If
                        If totalE < .E Then
                            .E = totalE
                            .ox = ox
                            .oy = oy
                            .isFind = True
                        End If
                    End With
                End If
            End If
        Next oy
    Next ox
End Function
Private Sub myCore_Crash(X As Double, Y As Double, dmg As Double, time As Long)
    Dim mx As Double, my As Double
    Dim sx As Long, sy As Long
    Dim cx As Long, cy As Long
    Dim dx As Double, dy As Double
    Dim life As Double, i As Long
    life = (dmg - pdmg) * 20 + 5
    If life < 0 Then
        With cl
            For i = 0 To ubpX Step 4
                myPX(i + 0) = .B
                myPX(i + 1) = .G
                myPX(i + 2) = .R
                myPX(i + 3) = 255
            Next i
        End With
    End If
    pdmg = dmg
    If X < Abs(X - sw) Then
        mx = X
        dx = 1
    Else
        mx = Abs(X - sw)
        dx = -1
    End If
    
    If Y < Abs(Y - sh) Then
        my = Y
        dy = 1
    Else
        my = Abs(Y - sh)
        dy = -1
    End If
    If mx < my Then 'X
        sy = Rnd * ph
        If mx = X Then
            sx = 1
        Else
            sx = pw - 1
        End If
        For i = 0 To 50
            If Mask(sx, sy + i) = 1 Then
                sy = sy + i
                Exit For
            ElseIf Mask(sx, sy - i) = 1 Then
                sy = sy - i
                Exit For
            End If
        Next i
    Else
        sx = Rnd * pw
        If my = Y Then
            sy = 1
        Else
            sy = ph - 1
        End If
        For i = 0 To 50
            If Mask(sx + i, sy) = 1 Then
                sx = sx + i
                Exit For
            ElseIf Mask(sx - i, sy) = 1 Then
                sx = sx - i
                Exit For
            End If
        Next i
    End If
    cx = sx
    cy = sy
    i = 0
    Do: i = i + 1
        If Mask(cx, cy) = 1 Then
            PX(cx, cy) = cl2
            With FindMinEnergyDirection(cx, cy, dx, dy)
                If .isFind Then
                    cx = cx + .ox
                    cy = cy + .oy
                    dx = .ox
                    dy = .oy
                    life = life - .lf
                Else
                    Exit Do
                End If
            End With
        End If
        If life < 0 Then Exit Do
    Loop Until i = psize
    UpdatePicture
End Sub
Private Sub InitPicture()
    Dim i As Long, bid As Long, bDist As Double, dist As Double
    Dim dx As Long, dy As Long
    Dim lx As Long, ly As Long
    Dim pL As Long, pU As Long
    Dim points() As POINTL
    ReDim points(CLng(Rnd * 30 + 70))
    pL = LBound(points)
    pU = UBound(points)
    With myCore
        cl = CL2RGBA(.tCL)
    End With
    With cl
        For i = 0 To ubpX Step 4
            myPX(i + 0) = .B
            myPX(i + 1) = .G
            myPX(i + 2) = .R
            myPX(i + 3) = 255
        Next i
    End With
    For i = pL To pU
        With points(i)
            .X = Rnd * pw
            .Y = Rnd * ph
        End With
    Next i
    ReDim region(pw, ph)
    ReDim myMask(pw, ph)
    For lx = 0 To pw
        dx = lx - points(pL).X
        For ly = 0 To ph
            bid = pL
            dy = ly - points(pL).Y
            bDist = dx * dx + dy * dy
            For i = pL + 1 To pU
                With points(i)
                    dx = lx - .X
                    dy = ly - .Y
                    dist = dx * dx + dy * dy
                    If dist < bDist Then
                        bDist = dist
                        bid = i
                    End If
                End With
            Next i
            region(lx, ly) = bid
        Next ly
    Next lx
    For lx = 0 To pw - 1
        For ly = 0 To ph - 1
            If region(lx, ly) <> region(lx, ly + 1) Then
                Mask(lx, ly) = 1
            ElseIf region(lx, ly) <> region(lx + 1, ly) Then
                Mask(lx, ly) = 1
            End If
        Next ly
    Next lx
End Sub
Private Sub ICFormPhysicsEx_init(core As CFormPhysics, Optional params As Variant = Empty)
    Dim ccl As Long, i As Long
    Set myCore = core
    With myCore
        pw = .mFrmRaw.width / .Px2Tw
        ph = .mFrmRaw.height / .Px2Tw
        psize = 2 * Sqr(pw * pw + ph * ph)
        sw = .ScrWidth * .Px2Tw
        sh = .ScrHeight * .Px2Tw
        .mFrmRaw.PictureAlignment = fmPictureAlignmentCenter
        .mFrmRaw.PictureSizeMode = fmPictureSizeModeStretch
    End With
    ReDim myPX(pw * ph * 4 - 1) As Byte
    ubpX = UBound(myPX)
    With BINFO.bmiHeader
        .biSize = Len(BINFO.bmiHeader)
        .biWidth = pw
        .biHeight = -ph
        .biPlanes = 1
        .biBitCount = 32
        .biSizeImage = pw * ph * 4
    End With
    With IID
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    With PIC
        .cbSizeofStruct = Len(PIC)
        .picType = PICTYPE_BITMAP
    End With
    ApplyMaskColor
    InitPicture
    UpdatePicture
End Sub
Private Sub ApplyMaskColor()
    WindowFromAccessibleObject myCore.mFrmRaw, hWnd
    Dim exStyle As LongPtr
    #If Win64 Then
        exStyle = GetWindowLongPtr(hWnd, GWL_EXSTYLE)
    #Else
        exStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    #End If
    #If Win64 Then
        SetWindowLongPtr hWnd, GWL_EXSTYLE, exStyle Or WS_EX_LAYERED
    #Else
        SetWindowLong hWnd, GWL_EXSTYLE, exStyle Or WS_EX_LAYERED
    #End If
    SetLayeredWindowAttributes hWnd, RGB(0, 254, 0), 0, LWA_COLORKEY
End Sub
Private Function CL2RGBA(ByRef cl As Long) As RGBA
    Dim L As LONGTYPE: L.Value = cl
    LSet CL2RGBA = L
End Function
Private Sub ICFormPhysicsEx_Terminate()
On Error GoTo err
    Set myCore.mFrmRaw.Picture = Nothing
err:
    Set myCore = Nothing
    hWnd = 0
End Sub
Private Property Get ICFormPhysicsEx_CreateInstance() As ICFormPhysicsEx
    Set ICFormPhysicsEx_CreateInstance = New CFormPhysicsFmRenderer
End Property
