Attribute VB_Name = "Module1"
Option Explicit
 
'-----------------------------------------------------------------------------
' variables
'-----------------------------------------------------------------------------
Dim g_DX As New DirectX8
Dim g_D3D As Direct3D8              'Used to create the D3DDevice
Dim g_D3DDevice As Direct3DDevice8  'Our rendering device
Dim g_VB As Direct3DVertexBuffer8

' A structure for our custom vertex type
' representing a point on the screen
Type CUSTOMVERTEX
    x As Single         'x in screen space
    y As Single         'y in screen space
    z  As Single        'normalized z
    Color As Long       'vertex color
End Type

Enum Rending
    PointList = D3DPT_POINTLIST
    LineList = D3DPT_LINELIST
    LineStrip = D3DPT_LINESTRIP
    TriangleList = D3DPT_TRIANGLELIST
    TriangleStrip = D3DPT_TRIANGLESTRIP
    TriangleFan = D3DPT_TRIANGLEFAN
End Enum
' Our custom FVF, which describes our custom vertex structure
Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)

Const g_pi = 3.1415

Public VerticesN As Long
Public TrianglesN As Long
Public Vertices() As CUSTOMVERTEX

Public Rotate As String

'-----------------------------------------------------------------------------
' Name: InitD3D()
' Desc: Initializes Direct3D
'-----------------------------------------------------------------------------
Function InitD3D(hWnd As Long) As Boolean
    On Local Error Resume Next
    
    ' Create the D3D object
    Set g_D3D = g_DX.Direct3DCreate()
    If g_D3D Is Nothing Then Exit Function
    
    ' Get the current display mode
    Dim Mode As D3DDISPLAYMODE
    g_D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode
        
    ' Fill in the type structure used to create the device
    Dim d3dpp As D3DPRESENT_PARAMETERS
    d3dpp.Windowed = 1
    d3dpp.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    d3dpp.BackBufferFormat = Mode.Format
    
    ' Create the D3DDevice
    ' If you do not have hardware 3d acceleration. Enable the reference rasterizer
    ' using the DirectX control panel and change D3DDEVTYPE_HAL to D3DDEVTYPE_REF
    
    Set g_D3DDevice = g_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    If g_D3DDevice Is Nothing Then Exit Function
    
    ' Device state would normally be set here
    ' Turn off culling, so we see the front and back of the triangle
    g_D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

    ' Turn off D3D lighting, since we are providing our own vertex colors
    g_D3DDevice.SetRenderState D3DRS_LIGHTING, 0

    InitD3D = True
End Function

Function InitVertices()
    ReDim Vertices(VerticesN - 1) As CUSTOMVERTEX
End Function
'-----------------------------------------------------------------------------
' Name: InitGeometry()
' Desc: Creates a vertex buffer and fills it with our vertices.
'-----------------------------------------------------------------------------
Function InitGeometry() As Boolean
    ' Initialize three vertices for rendering a triangle
    Dim VertexSizeInBytes As Long
    
    VertexSizeInBytes = Len(Vertices(0))

    ' Create the vertex buffer.
    Set g_VB = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * VerticesN, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VB Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_VB, 0, VertexSizeInBytes * VerticesN, 0, Vertices(0)

    InitGeometry = True
End Function

'-----------------------------------------------------------------------------
' Name: Cleanup()
' Desc: Releases all previously initialized objects
'-----------------------------------------------------------------------------
Sub Cleanup()
    Set g_VB = Nothing
    Set g_D3DDevice = Nothing
    Set g_D3D = Nothing
End Sub

Sub Rotation()
    ' For our world matrix, we will just rotate the object.
    Dim matWorld As D3DMATRIX
    
    If Rotate = "y" Then
        D3DXMatrixRotationY matWorld, Timer
        g_D3DDevice.SetTransform D3DTS_WORLD, matWorld
    ElseIf Rotate = "x" Then
        D3DXMatrixRotationX matWorld, Timer
        g_D3DDevice.SetTransform D3DTS_WORLD, matWorld
    ElseIf Rotate = "z" Then
        D3DXMatrixRotationZ matWorld, Timer
        g_D3DDevice.SetTransform D3DTS_WORLD, matWorld
    End If
End Sub

Public Function CameraPos(ByVal posx As Long, ByVal posy As Long, ByVal posz As Long)
    ' The view matrix defines the position and orientation of the camera
    ' Set up our view matrix. A view matrix can be defined given an eye point,
    ' a point to lookat, and a direction for which way is up. Here, we set the
    ' eye five units back along the z-axis and up three units, look at the
    ' origin, and define "up" to be in the y-direction.
    
    Dim matView As D3DMATRIX
    D3DXMatrixLookAtLH matView, Vec3(Val(posx), Val(posy), Val(posz)), Vec3(0#, 0#, 0#), Vec3(0#, 1#, 0#)
    g_D3DDevice.SetTransform D3DTS_VIEW, matView
End Function
'-----------------------------------------------------------------------------
' Name: SetupMatrices()
' Desc: Sets up the world, view, and projection transform matrices.
'-----------------------------------------------------------------------------
Sub SetupMatrices()
    ' The projection matrix describes the camera's lenses
    ' For the projection matrix, we set up a perspective transform (which
    ' transforms geometry from 3D view space to 2D viewport space, with
    ' a perspective divide making objects smaller in the distance). To build
    ' a perpsective transform, we need the field of view (1/4 pi is common),
    ' the aspect ratio, and the near and far clipping planes (which define at
    ' what distances geometry should be no longer be rendered).
    Dim matProj As D3DMATRIX
    D3DXMatrixPerspectiveFovLH matProj, g_pi / 4, 1, 1, 1000
    g_D3DDevice.SetTransform D3DTS_PROJECTION, matProj
End Sub

'-----------------------------------------------------------------------------
' Name: Render()
' Desc: Draws the scene
'-----------------------------------------------------------------------------
Sub Render(BackColor As Long, Mode As Rending)
    Dim v As CUSTOMVERTEX
    Dim sizeOfVertex As Long
    
    If g_D3DDevice Is Nothing Then Exit Sub

    ' Clear the backbuffer to a blue color (ARGB = 000000ff)
    '
    ' To clear the entire back buffer we send down
    g_D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, BackColor, 1#, 0
     
    ' Begin the scene
    g_D3DDevice.BeginScene
    
    SetupMatrices
        
    'Draw the triangles in the vertex buffer
    sizeOfVertex = Len(v)
    g_D3DDevice.SetStreamSource 0, g_VB, sizeOfVertex
    g_D3DDevice.SetVertexShader D3DFVF_CUSTOMVERTEX
    g_D3DDevice.DrawPrimitive Mode, 0, TrianglesN
     
    ' End the scene
    g_D3DDevice.EndScene
    
    ' Present the backbuffer contents to the front buffer (screen)
    g_D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

'-----------------------------------------------------------------------------
' Name: vec3()
' Desc: helper function
'-----------------------------------------------------------------------------
Function Vec3(x As Single, y As Single, z As Single) As D3DVECTOR
    Vec3.x = x
    Vec3.y = y
    Vec3.z = z
End Function

Function CreateVertex(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal Color As ColorConstants) As CUSTOMVERTEX
    With CreateVertex
        .x = x
        .y = y
        .z = z
        .Color = Color
    End With
End Function

