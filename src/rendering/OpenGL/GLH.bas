Attribute VB_Name = "glh"
Option Explicit
Public Enum Glenum
    GLU_FALSE = &H0&
    GLU_TRUE = &H1&
    GLU_SMOOTH = &H186A0
    GLU_FLAT = &H186A1
    GLU_OUTSIDE = &H186B4
    GLU_INSIDE = &H186B5
    GL_AMBIENT = &H1200&
    GL_COLOR_BUFFER_BIT = &H4000&
    GL_COMPILE = &H1300&
    GL_COMPILE_AND_EXECUTE = &H1301&
    GL_DEPTH_BUFFER_BIT = &H100&
    GL_DEPTH_TEST = &HB71&
    GL_DIFFUSE = &H1201&
    GL_FALSE = &H0&
    GL_FASTEST = &H1101&
    GL_FLAT = &H1D00&
    GL_LESS = &H201&
    GL_LIGHT0 = &H4000&
    GL_LIGHT1 = &H4001&
    GL_LIGHT2 = &H4002&
    GL_LIGHT3 = &H4003&
    GL_LIGHTING = &HB50&
    GL_LINES = &H1&
    GL_LINE_LOOP = &H2&
    GL_LINE_SMOOTH = &HB20&
    GL_LINE_SMOOTH_HINT = &HC52&
    GL_LINE_STRIP = &H3&
    GL_MODELVIEW = &H1700&
    GL_NICEST = &H1102&
    GL_PERSPECTIVE_CORRECTION_HINT = &HC50&
    GL_POINTS = &H0&
    GL_POLYGON = &H9&
    GL_POSITION = &H1203&
    GL_PROJECTION = &H1701&
    GL_QUADS = &H7&
    GL_QUAD_STRIP = &H8&
    GL_SMOOTH = &H1D01&
    GL_SPECULAR = &H1202&
    GL_TRIANGLES = &H4&
    GL_TRIANGLE_FAN = &H6&
    GL_TRIANGLE_STRIP = &H5&
    GL_TRUE = &H1&
    GL_POINT_SMOOTH_HINT = &HC51&
    GL_FRONT = &H408&
    GL_BACK = &H409&
    GL_FRONT_AND_BACK = &H40A&
    GL_SHININESS = &H1601&
    GL_BLEND = &HBE2&
    GL_ZERO = &H0&
    GL_ONE = &H1&
    GL_SRC_COLOR = &H300&
    GL_ONE_MINUS_SRC_COLOR = &H301&
    GL_SRC_ALPHA = &H302&
    GL_ONE_MINUS_SRC_ALPHA = &H303&
    GL_DST_ALPHA = &H304&
    GL_ONE_MINUS_DST_ALPHA = &H305&
    GL_POINT_SMOOTH = &HB10&
    GL_SCISSOR_TEST = &HC11&
    GL_COLOR_MATERIAL = &HB57&
    GL_NORMALIZE = &HBA1&
    GL_RESCALE_NORMAL = &H803A&
    GL_POLYGON_OFFSET_FILL = &H8037&

    GL_VERTEX_ARRAY = &H8074&
    GL_NORMAL_ARRAY = &H8075&
    GL_COLOR_ARRAY = &H8076&
    GL_INDEX_ARRAY = &H8077&
    GL_TEXTURE_COORD_ARRAY = &H8078&
    GL_COLOR_INDEX = &H1900&
    GL_COLOR_INDEXES = &H1603&

    GL_MULTISAMPLE = &H809D&
    GL_SAMPLE_ALPHA_TO_COVERAGE = &H809E&
    GL_SAMPLE_ALPHA_TO_ONE = &H809F&
    GL_SAMPLE_COVERAGE = &H80A0&

    GL_BYTE = &H1400&
    GL_UNSIGNED_BYTE = &H1401&
    GL_SHORT = &H1402&
    GL_UNSIGNED_SHORT = &H1403&
    GL_INT = &H1404&
    GL_UNSIGNED_INT = &H1405&
    GL_FLOAT = &H1406&
    GL_DOUBLE = &H140A&
    GL_2_BYTES = &H1407&
    GL_3_BYTES = &H1408&
    GL_4_BYTES = &H1409&

    GL_INVALID_ENUM = &H500&
    GL_INVALID_VALUE = &H501&
    GL_INVALID_OPERATION = &H502&
    GL_STACK_OVERFLOW = &H503&
    GL_STACK_UNDERFLOW = &H504&
    GL_OUT_OF_MEMORY = &H505&
    GL_INVALID_FRAMEBUFFER_OPERATION = &H506&
    GL_CONTEXT_LOST = &H507&
    GL_TABLE_TOO_LARGE1 = &H8031&
    GL_ONE_MINUS_DST_COLOR = &H307&

    GL_MODELVIEW_MATRIX = &HBA6&
    GL_RGB = &H1907&
    GL_RGBA = &H1908&
    GL_CLIP_PLANE0 = &H3000&
    GL_CLIP_PLANE1 = &H3001&
    GL_AMBIENT_AND_DIFFUSE = &H1602&
    
    
    
End Enum
Public Const CS_VREDRAW = 1
Public Const CS_HREDRAW = 2
Public Const CS_OWNDC = 32
Public Const PFD_DOUBLEBUFFER = 1
Public Const PFD_DRAW_TO_WINDOW = 4
Public Const PFD_SUPPORT_OPENGL = 32
Public Const PFD_DRAW_TO_BITMAP = 8
Public Const PFD_SUPPORT_GDI = 16
Public Const DIB_PAL_COLORS = 1
Public Type PIXELFORMATDESCRIPTOR
    nSize As Long
    nVersion As Long
    dwFlags As Long
    iPixelType As Byte
    cColorBits As Byte
    cRedBits As Byte
    cRedShift As Byte
    cGreenBits As Byte
    cGreenShift As Byte
    cBlueBits As Byte
    cBlueShift As Byte
    cAlphaBits As Byte
    cAlphaShift As Byte
    cAccumBits As Byte
    cAccumRedBits As Byte
    cAccumGreenBits As Byte
    cAccumBlueBits As Byte
    cAccumAlphaBits As Byte
    cDepthBits As Byte
    cStencilBits As Byte
    cAuxBuffers As Byte
    iLayerType As Byte
    bReserved As Byte
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type
Public Type Vector2d
    X As Double
    Y As Double
End Type
Public Type Vector3d
    X As Double
    Y As Double
    z As Double
End Type
Public Type Color4
    R As Single
    G As Single
    B As Single
    a As Single
End Type
Public Type B4
    B(3) As Byte
End Type
Public Type S1
    s As Single
End Type
Public Type L1
    L As Long
End Type
Public Function B2Single(ByRef B1, ByRef B2, ByRef B3, ByRef B4) As Single
    Dim X As B4, Y As S1
    With X: .B(0) = B1: .B(1) = B2: .B(2) = B3: .B(3) = B4: End With
    LSet Y = X
    B2Single = Y.s
End Function
Public Function B2Long(ByRef B1, ByRef B2, ByRef B3, ByRef B4) As Long
    Dim X As B4, Y As L1
    With X: .B(0) = B1: .B(1) = B2: .B(2) = B3: .B(3) = B4: End With
    LSet Y = X
    B2Long = Y.L
End Function
Public Function Vector2d(ByVal X As Double, ByVal Y As Double) As Vector2d
    With Vector2d: .X = X: .Y = Y: End With
End Function
Public Function Vector3d(ByVal X As Double, ByVal Y As Double, ByVal z As Double) As Vector3d
    With Vector3d: .X = X: .Y = Y: .z = z: End With
End Function
Public Function Color4(ByVal R As Single, ByVal G As Single, ByVal B As Single, ByVal a As Single) As Color4
    With Color4: .a = a: .B = B: .G = G: .R = R: End With
End Function
Public Function FastSin(ByRef th As Double) As Double
    Static mySin(2000) As Double
    Static init As Boolean
    If Not init Then
        Dim i As Long
        For i = 0 To 2000
            mySin(i) = Sin(i * 0.01)
        Next i
        init = True
    End If
    FastSin = mySin(CLng(th * 100))
End Function
Public Function FastCos(ByRef th As Double) As Double
    Static mycos(2000) As Double
    Static init As Boolean
    If Not init Then
        Dim i As Long
        For i = 0 To 2000
            mycos(i) = Cos(i * 0.01)
        Next i
        init = True
    End If
    FastCos = mycos(CLng(th * 100))
End Function
