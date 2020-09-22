Attribute VB_Name = "AgrDX"
'Agr3.0 DirectX
'Dated August 4, 2005
'Written by Nauful Aka Cade
's_u_cider@hotmail.com
'http://aura-blue.com/cade/vbsite/
'Thanks to Daniel Story, http://wannabegames.com
'Some code used from Jack Hoxley's tutorials
Option Explicit
Public DX As DirectX8
Public D3DX As D3DX8
Public D3D As Direct3D8
Public D3DDevice As Direct3DDevice8
Public D3DCaps As D3DCAPS8
Public DispMode As D3DDISPLAYMODE
Public DXGammaRamp As D3DGAMMARAMP
Public D3DWindow As D3DPRESENT_PARAMETERS
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Const Pi = 3.14159268
Public GDst!
Dim I As Long

Public Type RGBADesc
R As Single
G As Single
B As Single
A As Single
End Type

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Freq As Currency
Private StartTime As Currency
Private EndTime As Currency
Public TimeElapse As Single
Public TimeFactor As Double

Public Type D3DRect
X1 As Long
Y1  As Long
X2 As Long
Y2 As Long
End Type

Public Type Vector2DSng
Ix As Single
Iy As Single
Cx As Single
Cy As Single
Nx As Single
Ny As Single
Wx As Single
Wy As Single
End Type

Public Type UNLITVERTEX
    X As Single      ' Position
    Y As Single
    z As Single
    Nx As Single     ' Normal
    Ny As Single
    Nz As Single
    Diffuse As Long  ' diffuse color
    tu1 As Single    ' texture coordinates
    tv1 As Single
    tu2 As Single    ' texture coordinates
    tv2 As Single
End Type

Public Type LITVERTEX
    X As Single
    Y As Single
    z As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type
Public Const FVF_COLORVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
Public Const FVF_TLVERTEX = (D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR)
Public Const FVF_TLVERTEX2 = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2)
Public Const FVF_LVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1)
Public Const FVF_VERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)
Public Const LIT_FVF = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1)

Public Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Public Type TLVERTEX
    X As Single
    Y As Single
    z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Public Type TLVERTEX2
    X As Single
    Y As Single
    z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu1 As Single
    tv1 As Single
    tu2 As Single
    tv2 As Single
End Type

Public Const FVF2 = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2

Public Type TLVERTEX3
    p As D3DVECTOR4
    Color As Long
    Specular As Long
    TEX0 As D3DVECTOR2
    TEX1 As D3DVECTOR2
    TEX2 As D3DVECTOR2
End Type

Public Const FVF3 = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX3

Public Const FVF_PARTICLEVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
Public Type PARTICLEVERTEX
    v As D3DVECTOR
    Color As Long
    tu As Single
    tv As Single
End Type

Dim DXlngShaderArray() As Long
Dim DXlngShaderSize As Long
Dim DXBufferCode As D3DXBuffer
Public DXmatProj As D3DMATRIX
Public DXmatView As D3DMATRIX
Public DXmatWorld As D3DMATRIX
Public DXmatWorld2 As D3DMATRIX
Public DXmatTemp As D3DMATRIX
Public DXmatReflect As D3DMATRIX
Public DXmatShadow As D3DMATRIX
Public DXmatFinal As D3DMATRIX
Public DXMaterialBuffer As D3DXBuffer

Public DXMainFont As D3DXFont
Public DXMainFontDesc As IFont
Public DXTextRect As DxVBLibA.RECT
Public DXFnt As New StdFont

Public DXMultiSamples As Long

Public Type DXTexPoolObject
cTexture As Direct3DTexture8
cTextureMipMapLevel As Long
cTextureName As String
cTextureDesc() As D3DSURFACE_DESC
cLastUsedTick As Long
bInUse As Boolean
End Type
Public DXTexturePool() As DXTexPoolObject

Private TriStrip() As TLVERTEX
Private TriStripTemp() As TLVERTEX

Public DXRenderingDeviceName As String
Public DXRenderingDeviceInfo As D3DADAPTER_IDENTIFIER8

Public Type DXMeshTexture
cIndex As Long
cTexName As String
cTexAdv As Direct3DTexture8
End Type
Public Type DXT3DQuad
cTex As DXMeshTexture
bHasPlane As Boolean
cPlane As D3DPLANE
cVerts() As LITVERTEX
cVertCount As Long
cNormal As D3DVECTOR
End Type
Public Type DXT3DMesh
cMesh() As DXT3DQuad
cLights() As D3DLIGHT8
End Type

Public DXSceneLights() As D3DLIGHT8
Public Type DXMeshObject
cMaterials() As D3DMATERIAL8
cTextures() As DXMeshTexture
cMesh As D3DXMesh
cMaterialsNum As Long
cMatBuffer As D3DXBuffer
cMatrix As D3DMATRIX
cShadowMatrix As D3DMATRIX
cReflectMatrix As D3DMATRIX
cBoundingBoxMin As D3DVECTOR
cBoundingBoxCenter As D3DVECTOR
cBoundingBoxMax As D3DVECTOR
cBoundingSphereRadius As Single
cBoundingSphereCenter As D3DVECTOR
cBoundingSphereCenterUntransformed As D3DVECTOR4
cVerts() As D3DVERTEX
cAdjacency As D3DXBuffer
CPos As D3DVECTOR4
End Type
Public DXSceneMesh() As DXMeshObject
Private MeshBuffer As D3DXMesh

Private TempLV() As LITVERTEX

Public DXVertexBuffer As Direct3DVertexBuffer8

Public NullMatrix As D3DMATRIX

Public DX2DTransformMatrix As D3DMATRIX

Public VP As D3DVIEWPORT8

Public tStripTemp(3) As TLVERTEX
Public tStrip(3) As TLVERTEX

Public DXtempVec As D3DVECTOR
Public DXtempVec4 As D3DVECTOR4
Dim m_pDisplayTexture As Direct3DTexture8
Dim m_pDisplayTextureSurface As Direct3DSurface8
Dim m_pDisplayZSurface As Direct3DSurface8
Dim m_pBackBuffer As Direct3DSurface8
Dim m_pZBuffer As Direct3DSurface8

Public D3DXRts As D3DXRenderToSurface
Public D3DRenderTexture As Direct3DTexture8
Public D3DXRtsVP As D3DVIEWPORT8
Public D3DRenderSurface As Direct3DSurface8

Public Type DX2DMeshObjectFrame
TriVerts() As TLVERTEX
TextureName As String
End Type

Public Type DX2DMeshObject
Frames() As DX2DMeshObjectFrame
X As Single
Y As Single
End Type

Public v(4) As TLVERTEX
Public V2(4) As TLVERTEX2

Public vl(35) As LITVERTEX
Public vlt(1) As TLVERTEX
Public Sub DXLoad(Init_hWnd As Long, Optional MaxMultiSampling As Long, Optional MaxAnisotropy As Long)
On Error GoTo DXLoadErr
'Create blank objects in memory
Set DX = New DirectX8
Set D3D = DX.Direct3DCreate
Set D3DX = New D3DX8

'Create texture pool and scene mesh with blank objects as contents
ReDim DXTexturePool(0)
ReDim DXSceneMesh(0)

'Set up the display window
D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

D3DWindow.Windowed = 1
D3DWindow.SwapEffect = D3DSWAPEFFECT_DISCARD
D3DWindow.BackBufferFormat = DispMode.Format
D3DWindow.AutoDepthStencilFormat = DXGetBestStencilFormat
D3DWindow.EnableAutoDepthStencil = 1

'Retrieve device capabilities
D3D.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DCaps

'Use multisampling
If MaxMultiSampling > 0 Then
Dim CMS As Long
CMS = MaxMultiSampling
SetMultiSampling:
'Loop through the aa sampling valid modes until a valid mode has been found
If D3D.CheckDeviceMultiSampleType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DWindow.Windowed, CMS) = D3D_OK Then
D3DWindow.MultiSampleType = CMS
Else
If CMS > 0 Then
CMS = CMS - 1
GoTo SetMultiSampling
End If
End If
End If
If MaxAnisotropy > 0 Then
If MaxAnisotropy > D3DCaps.MaxAnisotropy Then
MaxAnisotropy = D3DCaps.MaxAnisotropy
End If
End If

Debug.Print "Max samples specified: " & MaxMultiSampling & " Actual Samples: " & CMS
DXMultiSamples = CMS

If Not D3D.CheckDeviceType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, DispMode.Format, 1) = D3D_OK Then
MsgBox "There was an error initializing DirectX", vbCritical, ""
DXUnLoad
End
End If

'Create the D3DDevice
Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Init_hWnd, D3DCREATE_MIXED_VERTEXPROCESSING, D3DWindow)
D3DDevice.SetVertexShader FVF_TLVERTEX
D3DDevice.SetRenderState D3DRS_LIGHTING, 0
D3DDevice.SetRenderState D3DRS_ZENABLE, 1
D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

Dim I As Long

'Enable ansiotropy texture filtering
If MaxAnisotropy > 0 Then
For I = 0 To 8
D3DDevice.SetTextureStageState I, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
D3DDevice.SetTextureStageState I, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
D3DDevice.SetTextureStageState I, D3DTSS_MIPFILTER, D3DTEXF_ANISOTROPIC
D3DDevice.SetTextureStageState I, D3DTSS_MAXANISOTROPY, MaxAnisotropy
Next
Debug.Print "Anisotropy texture filtering: " & MaxAnisotropy & "x"
Else
For I = 0 To 8
D3DDevice.SetTextureStageState I, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
D3DDevice.SetTextureStageState I, D3DTSS_MINFILTER, D3DTEXF_LINEAR
D3DDevice.SetTextureStageState I, D3DTSS_MIPFILTER, D3DTEXF_LINEAR
D3DDevice.SetTextureStageState I, D3DTSS_MAXANISOTROPY, MaxAnisotropy
Next
End If

'Set the zbuffer to render from back to front
D3DDevice.SetRenderState D3DRS_ZFUNC, D3DCMP_LESS

'Initialise the text renderer
DXFnt.Name = "Arial"
DXFnt.Size = 14
Set DXMainFontDesc = DXFnt
Set DXMainFont = D3DX.CreateFont(D3DDevice, DXMainFontDesc.hFont)
DXBeginScene
DXRenderText "Loading...", D3DColorRGBA(0, 0, 0, 255), 0, 0, 640, 480
DXEndScene

Dim TempLitVertex As LITVERTEX
Set DXVertexBuffer = D3DDevice.CreateVertexBuffer(Len(TempLitVertex) * 3, 0, LIT_FVF, D3DPOOL_DEFAULT)
Exit Sub
Dim tChr As String
D3D.GetAdapterIdentifier D3DADAPTER_DEFAULT, 0, DXRenderingDeviceInfo
For I = 0 To 512 - 1
tChr = Chr$(DXRenderingDeviceInfo.Description(I))
If IsNumeric(tChr) Or UCase(tChr) <> LCase(tChr) Or tChr = " " Then
DXRenderingDeviceName = DXRenderingDeviceName & tChr
End If
Next
Debug.Print "Rendering device is " & DXRenderingDeviceName
D3DXMatrixIdentity NullMatrix

'Set D3DRenderTexture = D3DX.CreateTexture(D3DDevice, 1024, 1024, 1&, 0&, D3DFMT_UNKNOWN, D3DPOOL_MANAGED)
'Set D3DXRts = D3DX.CreateRenderToSurface(D3DDevice, 1024, 1024, DispMode.Format, 0&, D3DFMT_UNKNOWN)

'With D3DXRtsVP
'.Height = 1024
'.Width = 1024
'.MaxZ = 1!
'End With

'Set D3DRenderSurface = D3DRenderTexture.GetSurfaceLevel(0&)

D3DDevice.SetRenderState D3DRS_ZENABLE, 0
D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
DXfvfTL
Exit Sub
DXLoadErr:
MsgBox D3DX.GetErrorString(Err.Number), vbCritical, "AgrDX:DXLoad"
End Sub
Public Sub DXUnLoad()
Dim I As Long
Dim I2 As Long
For I = 0 To UBound(DXTexturePool)
'Loop through and delete the texture pool's contents
DXRemoveTexture DXTexturePool(I).cTextureName
Next
'Remove the scene meshes
If UBound(DXSceneMesh) > 0 Then
For I = 0 To UBound(DXSceneMesh)
DXUnloadXMesh DXSceneMesh(I)
Next
End If
'Erase dimensioned objects
Erase DXTexturePool
Erase DXSceneMesh
Erase DXSceneLights
'Unload DirectX
Set DXVertexBuffer = Nothing
Set D3DDevice = Nothing
Set D3D = Nothing
Set DX = Nothing
Set D3DX = Nothing
Set DXFnt = Nothing
Set DXMainFontDesc = Nothing
Set DXMainFont = Nothing
End Sub
Public Sub DXBeginScene()
'&HCCCCFF
'Clear the render target and begin the scene
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, &HCCCCFF, 1#, 0
D3DDevice.BeginScene
End Sub
Public Sub DXEndScene(Optional OverrideDC As Long)
'End the scene
D3DDevice.EndScene
D3DDevice.Present ByVal 0, ByVal 0, OverrideDC, ByVal 0
If D3DDevice.GetRenderState(D3DRS_STENCILENABLE) = 1 Then
'If the stencil buffer is enabled, disable it and clear stencil
D3DDevice.SetRenderState D3DRS_STENCILENABLE, 0
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_STENCIL, 0, 0, 0
End If
If D3DDevice.GetRenderState(D3DRS_ZENABLE) = 1 Then
'If the zbuffer is enabled, clear zbuffer
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_ZBUFFER, 0, 0, 0
End If
End Sub
Public Sub DXEnumDevice()
On Local Error Resume Next
'Retrieve caps
With D3DCaps
If .MaxActiveLights = -1 Then
Debug.Print "Max Active Lights: Unlimited"
Else
Debug.Print "Max Active Lights: " & .MaxActiveLights
End If
Debug.Print "Maximum Point Vertex size: " & .MaxPointSize
Debug.Print "Maximum Texture Size: " & .MaxTextureWidth & "x" & .MaxTextureHeight
Debug.Print "Maximum Primatives in one call: " & .MaxPrimitiveCount
Debug.Print "Max Ansiotropy " & .MaxAnisotropy
Debug.Print "Max Pixel Shader Value " & .MaxPixelShaderValue
Debug.Print "Max Point Size " & .MaxPointSize
Debug.Print "Max Primitive Count " & .MaxPrimitiveCount
Debug.Print "Max Simultaneous Textures " & .MaxSimultaneousTextures
Debug.Print "Max Texture Aspect Ratio " & .MaxTextureAspectRatio
Debug.Print "Max Ubound Vertex " & .MaxVertexIndex
Debug.Print "Max Vertex Shader Public Const " & .MaxVertexShaderConst
Debug.Print "Max Volume Extent " & .MaxVolumeExtent
Debug.Print "Pixel shader version: " & .PixelShaderVersion
Debug.Print "Vertex shader version" & .VertexShaderVersion
If .TextureCaps And Not D3DPTEXTURECAPS_SQUAREONLY Then
    Debug.Print "Textures don't have to be square"
End If
If .TextureCaps And D3DPTEXTURECAPS_CUBEMAP Then
    Debug.Print "Device Supports Cube Mapping"
End If
If .TextureCaps And D3DPTEXTURECAPS_VOLUMEMAP Then
    Debug.Print "Device Supports Volume Mapping"
End If
If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, _
       D3DRTYPE_SURFACE, D3DFMT_D16) = D3D_OK Then
    Debug.Print "16 bit Z-Buffers are supported (D3DFMT_D16)"
End If
If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D16_LOCKABLE) = D3D_OK Then
    Debug.Print "Lockable 16 bit Z-Buffers are supported (D3DFMT_D16_LOCKABLE)"
End If
If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D24S8) = D3D_OK Then
    Debug.Print "32 bit divided between 24 bit Depth and 8 bit stencil are supported (D3DFMT_D24S8)"
End If
If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D24X4S4) = D3D_OK Then
    Debug.Print "32 bit divided between 24 bit depth, 4 bit stencil and 4 bit unused (D3DFMT_D24X4S4)"
End If
If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D24X8) = D3D_OK Then
    Debug.Print "24 bit Z-Buffer supported (D3DFMT_D24X8)"
End If
If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D32) = D3D_OK Then
    Debug.Print "Pure 32 bit Z-buffer supported (D3DFMT_D32)"
End If
End With
End Sub
Public Function DXGetBestStencilFormat() As Variant
Dim OutSF As Variant
'Retrieve the best format to use for stencil/depth buffer
D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
OutSF = D3DFMT_D16
If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, _
       D3DRTYPE_SURFACE, D3DFMT_D16) = D3D_OK Then
    OutSF = D3DFMT_D16
End If
If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D16_LOCKABLE) = D3D_OK Then
    OutSF = D3DFMT_D16_LOCKABLE
End If
If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D24X4S4) = D3D_OK Then
    OutSF = D3DFMT_D24X4S4
End If
If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D24S8) = D3D_OK Then
    OutSF = D3DFMT_D24S8
End If
If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D24X4S4) = D3D_OK Then
    'OutSF = D3DFMT_D24X4S4
End If
If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D24X8) = D3D_OK Then
    'OutSF = D3DFMT_D24X8
End If
If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, _
        D3DRTYPE_SURFACE, D3DFMT_D32) = D3D_OK Then
    'OutSF = D3DFMT_D32
End If
DXGetBestStencilFormat = OutSF
End Function
Public Function DXCreateLVertex(X As Single, Y As Single, z As Single, Color As Long, Specular As Long, tu As Single, tv As Single) As LITVERTEX
'Inline wrapper for a lit vertex
DXCreateLVertex.X = X
DXCreateLVertex.Y = Y
DXCreateLVertex.z = z
DXCreateLVertex.Color = Color
DXCreateLVertex.Specular = Specular
DXCreateLVertex.tu = tu
DXCreateLVertex.tv = tv
End Function
Public Function DXCreateTLVertex(X As Single, Y As Single, z As Single, rhw As Single, Color As Long, Specular As Long, tu As Single, tv As Single) As TLVERTEX
'Inline wrapper for a transformed and lit vertex, 2d
DXCreateTLVertex.X = X
DXCreateTLVertex.Y = Y
DXCreateTLVertex.z = z
DXCreateTLVertex.rhw = rhw
DXCreateTLVertex.Color = Color
DXCreateTLVertex.Specular = Specular
DXCreateTLVertex.tu = tu
DXCreateTLVertex.tv = tv
End Function
Public Function DXCreateTLVertex2(X As Single, Y As Single, z As Single, rhw As Single, Color As Long, Specular As Long, tu1 As Single, tv1 As Single, tu2 As Single, tv2 As Single) As TLVERTEX2
DXCreateTLVertex2.X = X
DXCreateTLVertex2.Y = Y
DXCreateTLVertex2.z = z
DXCreateTLVertex2.rhw = rhw
DXCreateTLVertex2.Color = Color
DXCreateTLVertex2.Specular = Specular
DXCreateTLVertex2.tu1 = tu1
DXCreateTLVertex2.tv1 = tv1
DXCreateTLVertex2.tu2 = tu2
DXCreateTLVertex2.tv2 = tv2
End Function
Public Function DXMakeVector(X As Single, Y As Single, z As Single) As D3DVECTOR
'Inline wrapper for a 3d vector
DXMakeVector.X = X
DXMakeVector.Y = Y
DXMakeVector.z = z
End Function
Public Function DXMakeVector4(X As Single, Y As Single, z As Single, w As Single) As D3DVECTOR4
'Inline wrapper for a 3d vector
DXMakeVector4.X = X
DXMakeVector4.Y = Y
DXMakeVector4.z = z
DXMakeVector4.w = w
End Function
Public Function DXMakePlane(A As Single, B As Single, C As Single, D As Single) As D3DPLANE
With DXMakePlane
.A = A
.B = B
.C = C
.D = D
End With
End Function
Public Function DXMakeVector2D(X As Single, Y As Single) As D3DVECTOR2
'Inline wrapper for a 2d vector
DXMakeVector2D.X = X
DXMakeVector2D.Y = Y
End Function
Public Function DXMakePixelShader(PSFileName As String)
'Assemble a pixel shader from a file, returning its handle
'Optimized
On Error GoTo PSErr
Set DXBufferCode = D3DX.AssembleShaderFromFile(PSFileName$, 0, "", Nothing)
DXlngShaderSize = DXBufferCode.GetBufferSize() / 4
ReDim DXlngShaderArray(DXlngShaderSize - 1)
Call D3DX.BufferGetData(DXBufferCode, 0&, 4&, DXlngShaderSize&, DXlngShaderArray(0))
DXMakePixelShader = D3DDevice.CreatePixelShader(DXlngShaderArray(0))
Set DXBufferCode = Nothing
Exit Function
PSErr:
MsgBox "Unable to create pixel shader", vbCritical, PSFileName
Set DXBufferCode = Nothing
End Function
Public Function DXMakePixelShaderFromMemory(PSContents As String)
'Assemble a pixel shader from a string, return its handle
'Optimized
On Error GoTo PSErr
Set DXBufferCode = D3DX.AssembleShader(PSContents$, 0, Nothing, "")
DXlngShaderSize = DXBufferCode.GetBufferSize() / 4
ReDim DXlngShaderArray(DXlngShaderSize - 1)
Call D3DX.BufferGetData(DXBufferCode, 0&, 4&, DXlngShaderSize&, DXlngShaderArray(0))
DXMakePixelShaderFromMemory = D3DDevice.CreatePixelShader(DXlngShaderArray(0))
Set DXBufferCode = Nothing
Exit Function
PSErr:
MsgBox "Unable to create pixel shader", vbCritical, ""
Set DXBufferCode = Nothing
End Function
Public Sub DXSetPixelShader(ByRef lngPixelShaderHandle As Long)
'Enables a pixel shader
'Optimized
D3DDevice.SetPixelShader lngPixelShaderHandle&
End Sub
Public Sub DXDeletePixelShader(ByRef lngPixelShaderHandle As Long)
'Deletes a pixel shader
'Optimized
D3DDevice.DeletePixelShader lngPixelShaderHandle&
End Sub
Public Sub DXSetPixelShaderCRegister(RegIndex As Long, ValR As Single, ValG As Single, ValB As Single, ValA As Single)
Dim SinArr(3)
SinArr(0) = ValR
SinArr(1) = ValG
SinArr(2) = ValB
SinArr(3) = ValA
D3DDevice.SetPixelShaderConstant RegIndex, SinArr(0), 4
End Sub
Public Function DXCreateLitVertex(X As Single, Y As Single, z As Single, Colour As Long, Specular As Long, tu As Single, tv As Single) As LITVERTEX
'Inline function for a different type of lit vertex
    DXCreateLitVertex.X = X
    DXCreateLitVertex.Y = Y
    DXCreateLitVertex.z = z
    DXCreateLitVertex.Color = Colour
    DXCreateLitVertex.Specular = Specular
    DXCreateLitVertex.tu = tu
    DXCreateLitVertex.tv = tv
End Function
Public Function DXftTOdw(flo As Single) As Long
'A helper function, converts from C++ Float to C++ DWord
    Dim Buf As D3DXBuffer
    Set Buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData Buf, 0, 4, 1, flo
    D3DX.BufferGetData Buf, 0, 4, 1, DXftTOdw
End Function
Public Function DXMakeColor(R!, G!, B!, A!) As D3DCOLORVALUE
'An inline wrapper to return a color but in non-DWord RGBA format
DXMakeColor.R = R
DXMakeColor.G = G
DXMakeColor.B = B
DXMakeColor.A = A
End Function
Public Sub DXSetupStencilBuffer()
'Create a stencil buffer, with overdraw functionality
With D3DDevice
.SetRenderState D3DRS_STENCILENABLE, 1
.SetRenderState D3DRS_STENCILFUNC, D3DCMP_ALWAYS
.SetRenderState D3DRS_STENCILREF, 0
.SetRenderState D3DRS_STENCILMASK, 0
.SetRenderState D3DRS_STENCILWRITEMASK, &HFFFFFFFF
.SetRenderState D3DRS_STENCILZFAIL, D3DSTENCILOP_INCRSAT
.SetRenderState D3DRS_STENCILFAIL, D3DSTENCILOP_ZERO
.SetRenderState D3DRS_STENCILPASS, D3DSTENCILOP_INCRSAT
End With
End Sub
Public Sub DXSetupStencilBufferStencil()
'Create a stencil buffer, with stencil functionality
With D3DDevice
.SetRenderState D3DRS_STENCILENABLE, 1
.SetRenderState D3DRS_STENCILFUNC, D3DCMP_ALWAYS
.SetRenderState D3DRS_STENCILREF, 0
.SetRenderState D3DRS_STENCILMASK, 0
.SetRenderState D3DRS_STENCILWRITEMASK, &HFFFFFFFF
.SetRenderState D3DRS_STENCILZFAIL, D3DSTENCILOP_ZERO
.SetRenderState D3DRS_STENCILFAIL, D3DSTENCILOP_ZERO
.SetRenderState D3DRS_STENCILPASS, D3DSTENCILOP_INCRSAT
End With
End Sub
Public Sub DXFreezeStencilBuffer()
'Create a stencil buffer, with stencil functionality
With D3DDevice
.SetRenderState D3DRS_STENCILZFAIL, D3DSTENCILOP_KEEP
.SetRenderState D3DRS_STENCILFAIL, D3DSTENCILOP_KEEP
.SetRenderState D3DRS_STENCILPASS, D3DSTENCILOP_KEEP
.SetRenderState D3DRS_ZFUNC, D3DCMP_NOTEQUAL
End With
End Sub
Public Sub DXRenderStencilBuffer()
'Display the stencil buffer on the render target
Dim TLVerts(0 To 3) As TLVERTEX
'//needed to size the quad properly.
D3DDevice.GetViewport VP

'//set up the default parameters for the TL Quad. The colour will
'//be altered for each level we draw.
TLVerts(0).rhw = 1: TLVerts(1).rhw = 1: TLVerts(2).rhw = 1: TLVerts(3).rhw = 1
TLVerts(0).Color = D3DColorXRGB(255, 0, 0)
TLVerts(1).Color = TLVerts(0).Color
TLVerts(2).Color = TLVerts(0).Color
TLVerts(3).Color = TLVerts(0).Color

TLVerts(0).X = 0: TLVerts(0).Y = 0
TLVerts(1).X = VP.Width: TLVerts(1).Y = 0
TLVerts(2).X = 0: TLVerts(2).Y = VP.Height
TLVerts(3).X = VP.Width: TLVerts(3).Y = VP.Height

D3DDevice.SetTexture 0, Nothing
D3DDevice.SetVertexShader FVF_TLVERTEX

'//This next line is necessary to clear any pixels with
'//a 0 overdraw value... as they wont be picked up in the next
'//stage.
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0

With D3DDevice
    '//we dont care whats in the z buffer, so ignore it.
    '.SetRenderState D3DRS_ZENABLE, 0

    .SetRenderState D3DRS_STENCILZFAIL, D3DSTENCILOP_KEEP
    .SetRenderState D3DRS_STENCILFAIL, D3DSTENCILOP_KEEP
    .SetRenderState D3DRS_STENCILPASS, D3DSTENCILOP_KEEP
    .SetRenderState D3DRS_STENCILFUNC, D3DCMP_NOTEQUAL
    .SetRenderState D3DRS_STENCILREF, 0

    'DRAW LEVEL 1 OVERDRAW
    .SetRenderState D3DRS_STENCILMASK, 1
    TLVerts(0).Color = D3DColorXRGB(0, 0, 255)
    TLVerts(1).Color = TLVerts(0).Color
    TLVerts(2).Color = TLVerts(0).Color
    TLVerts(3).Color = TLVerts(0).Color
    .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TLVerts(0), Len(TLVerts(0))

    'DRAW LEVEL 2 OVERDRAW
    .SetRenderState D3DRS_STENCILMASK, 2
    TLVerts(0).Color = D3DColorXRGB(0, 255, 0)
    TLVerts(1).Color = TLVerts(0).Color
    TLVerts(2).Color = TLVerts(0).Color
    TLVerts(3).Color = TLVerts(0).Color
    .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TLVerts(0), Len(TLVerts(0))

    'DRAW LEVEL 3 OVERDRAW
    .SetRenderState D3DRS_STENCILMASK, 4
    TLVerts(0).Color = D3DColorXRGB(255, 128, 0)
    TLVerts(1).Color = TLVerts(0).Color
    TLVerts(2).Color = TLVerts(0).Color
    TLVerts(3).Color = TLVerts(0).Color
    .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TLVerts(0), Len(TLVerts(0))

    'DRAW LEVEL 4 OVERDRAW
    .SetRenderState D3DRS_STENCILMASK, 8
    TLVerts(0).Color = D3DColorXRGB(255, 0, 0)
    TLVerts(1).Color = TLVerts(0).Color
    TLVerts(2).Color = TLVerts(0).Color
    TLVerts(3).Color = TLVerts(0).Color
    .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TLVerts(0), Len(TLVerts(0))

    'DRAW ALL REMAINING LEVELS OVERDRAW
    '(this level wont exist for a 4bit stencil buffer)
    .SetRenderState D3DRS_STENCILMASK, 256 - 16
    TLVerts(0).Color = D3DColorXRGB(255, 255, 0)
    TLVerts(1).Color = TLVerts(0).Color
    TLVerts(2).Color = TLVerts(0).Color
    TLVerts(3).Color = TLVerts(0).Color
    .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TLVerts(0), Len(TLVerts(0))

    'restore the various stage states...
    '.SetRenderState D3DRS_ZENABLE, 1
     .SetRenderState D3DRS_STENCILENABLE, 0
End With
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_STENCIL, 0, 0, 0
End Sub
Public Sub DXEnableAlpha()
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
End Sub
Public Sub DXDisableAlpha()
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End Sub
Public Sub DXSetBumpMapping(X!, Y!)
D3DDevice.SetRenderState D3DRS_TEXTUREFACTOR, DXVectorToRGBA(DXMakeVector(X, Y, 1), 16)
End Sub
Public Sub DXSetDot3BumpMappingState()
D3DDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2

        D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
        D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_DOTPRODUCT3
        D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_TFACTOR
        
        D3DDevice.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
        D3DDevice.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_CURRENT
        D3DDevice.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_TEXTURE
        
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
End Sub
Public Sub DXSetAlphaOneState()
'Set the states to use alpha
D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
End Sub
Public Sub DXSetAlphaState()
'Set the states to use alpha
D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub
Public Sub DXSetAlphaTextureState()
D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE
D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_CURRENT
End Sub
Public Sub DXSetTextureStateBlend()
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
End Sub
Public Sub DXSetTextureStateModulated()
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTCOLOR
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCCOLOR
End Sub
Public Sub DXSetTextureStateHighlighted()
D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub
Public Sub DXSetMultiTextureState()
'Set the states to use multitexturing
D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
D3DDevice.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
D3DDevice.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
D3DDevice.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
End Sub
Public Sub DXfvfTL()
'Set the state to use transformed and lit vertices
D3DDevice.SetVertexShader AgrDX.FVF_TLVERTEX
End Sub
Public Sub DXfvfTL2()
'Set the state to use transformed and lit vertices with 2 textures
D3DDevice.SetVertexShader AgrDX.FVF2
End Sub
Public Sub DXfvfL()
'Set the state to use lit vertices
D3DDevice.SetVertexShader AgrDX.FVF_LVERTEX
End Sub
Public Sub DXfvfV()
'Set the state to use lit vertices
D3DDevice.SetVertexShader AgrDX.FVF_VERTEX
End Sub
Public Sub DXEnableFog()
'Creates fog with no range defined
With D3DDevice
.SetRenderState D3DRS_FOGENABLE, 1
.SetRenderState D3DRS_FOGTABLEMODE, D3DFOG_NONE
.SetRenderState D3DRS_FOGVERTEXMODE, D3DFOG_LINEAR
.SetRenderState D3DRS_RANGEFOGENABLE, DXftTOdw(1)
End With
End Sub
Public Sub DXRenderText(TxtString As String, FontColor&, X!, Y!, w!, H!)
DXTextRect.Left = X
DXTextRect.Top = Y
DXTextRect.Right = DXTextRect.Left + w
DXTextRect.bottom = DXTextRect.Top + H
D3DX.DrawText DXMainFont, FontColor, TxtString, DXTextRect, DT_TOP Or DT_LEFT Or DT_WORDBREAK
End Sub
Public Function DXMakeMaterial(ByVal rReflect As Byte, ByVal gReflect As Byte, ByVal bReflect As Byte, ByVal rOwn As Byte, ByVal gOwn As Byte, ByVal bOwn As Byte, ByVal SpecPower As Single, ByVal Transparency As Byte) As D3DMATERIAL8
'Inline function that returns a material
    Dim retMat As D3DMATERIAL8
    Dim Reflect As D3DCOLORVALUE, Own As D3DCOLORVALUE

    Reflect.A = 1 - Transparency / 255
    Reflect.R = rReflect / 255
    Reflect.G = gReflect / 255
    Reflect.B = bReflect / 255
    Own.A = 1
    Own.R = rOwn / 255
    Own.G = gOwn / 255
    Own.B = bOwn / 255

    retMat.Ambient = Reflect
    retMat.Ambient.A = 1
    retMat.Diffuse = Reflect
    retMat.emissive = Own
    retMat.Specular.A = 1
    retMat.Specular.R = 1
    retMat.Specular.G = 1
    retMat.Specular.B = 1
    retMat.power = SpecPower

    DXMakeMaterial = retMat
End Function
Public Sub DXSetTexture(TexAddress As Long, TexName As String)
'D3DDevice.SetTexture TextureAddress, and the texture associated with TexName in the Texture Pool
'Optimized
If UCase(TexName) = "NULL" Then
D3DDevice.SetTexture TexAddress, Nothing
Exit Sub
End If
If IsNumeric(TexName) Then
D3DDevice.SetTexture TexAddress, DXTexturePool(CLng(TexName)).cTexture
Exit Sub
End If
Dim bFoundTex As Boolean
Dim cTexIndex As Integer
Dim I As Long
For I& = 0 To UBound(DXTexturePool)
If DXTexturePool(I).cTextureName$ = TexName$ And DXTexturePool(I).bInUse Then
bFoundTex = True
cTexIndex = I&
GoTo RenderTex:
End If
Next
If Not bFoundTex Then
DXAddTexture TexName$
Debug.Print TexName$ & " was not found in the texture pool so it has been loaded"
cTexIndex% = UBound(DXTexturePool)
End If
RenderTex:
DXTexturePool(cTexIndex).cLastUsedTick& = GetTickCount&
D3DDevice.SetTexture TexAddress&, DXTexturePool(cTexIndex).cTexture
End Sub
Public Function DXGetTextureIndex(TexName As String) As Long
'Retrieves the index in the texture pool of the associated texture
DXGetTextureIndex = -1
Dim I As Long
For I = 0 To UBound(DXTexturePool)
If UCase(DXTexturePool(I).cTextureName) = UCase(TexName) Then
DXGetTextureIndex = I
Exit Function
End If
Next
End Function
Public Sub DXRemoveTexture(TexName As String)
'Removes the texture associated with TexName from the texture pool
'Optimized
Dim I As Long
For I = 0 To UBound(DXTexturePool)
If DXTexturePool(I).cTextureName = TexName And DXTexturePool(I).bInUse Then
With DXTexturePool(I&)
.bInUse = False
Set .cTexture = Nothing
Debug.Print TexName$ & " in texture pool as " & I& & " has been removed"
End With
End If
Next
End Sub
Public Sub DXAddTexture(TexName As String, Optional TransparentColorKey As Long, Optional TextureFormat As CONST_D3DFORMAT, Optional TexWidth As Long, Optional TexHeight As Long, Optional MipMapLevels As Long = 4)
'Adds a texture into the texture pool
'Optimized
On Error GoTo DXTexErr
Dim cTexIndex As Integer
Dim bTexIsADuplicate As Boolean
Dim I As Long
bTexIsADuplicate = False
cTexIndex% = -1
If cTexIndex% = -1 Then
ReDim Preserve DXTexturePool(UBound(DXTexturePool) + 1&)
cTexIndex% = UBound(DXTexturePool)
End If
For I& = 0 To UBound(DXTexturePool)
If DXTexturePool(I).bInUse And DXTexturePool(I).cTextureName$ = TexName$ Then
Debug.Print TexName & " in texture pool already exists so it wont be added again"
ReDim Preserve DXTexturePool(UBound(DXTexturePool) - 1&)
Exit Sub
End If
If Not DXTexturePool(I).bInUse Then
cTexIndex% = I&
End If
Next
If TransparentColorKey& = 0& Then
TransparentColorKey& = D3DColorRGBA(255, 255, 255, 255)
End If
If TexWidth& = 0 Then TexWidth& = -1
If TexHeight& = 0 Then TexHeight& = -1
If TextureFormat = 0 Then
TextureFormat = D3DFMT_UNKNOWN
End If
With DXTexturePool(cTexIndex%)
Set .cTexture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\" & TexName$, TexWidth&, TexHeight&, MipMapLevels, 0&, TextureFormat, D3DPOOL_DEFAULT, D3DX_FILTER_TRIANGLE, D3DX_FILTER_TRIANGLE, TransparentColorKey&, ByVal 0&, ByVal 0&)
ReDim .cTextureDesc(MipMapLevels)
.cTexture.GetLevelDesc 0, .cTextureDesc(0)
For I = MipMapLevels - 1 To 0 Step -1
If I <> 0 Then .cTexture.GetLevelDesc I, .cTextureDesc(I)
If .cTextureDesc(I).Height <= D3DCaps.MaxTextureHeight And .cTextureDesc(I).Width <= D3DCaps.MaxTextureWidth Then
.cTextureMipMapLevel = I
.cTexture.SetLOD I
Else
End If
Next
.cTextureName$ = TexName$
.bInUse = Not bTexIsADuplicate
Debug.Print TexName$ & " in texture pool as " & cTexIndex% & " has been added with mipmap level " & .cTextureMipMapLevel
End With
Exit Sub
D3DDevice.SetTexture 0, Nothing
DXTexErr:
Debug.Print "Texture error! Texture name is " & TexName
End Sub
Public Sub DXTextureOptimizePool()
'Removes redundant entries in the texture pool
'Optimized
Dim I As Long
Dim I2 As Integer
Dim cUboundTexturePool As Integer
Dim bCanBeOptimized As Boolean
cUboundTexturePool% = UBound(DXTexturePool)

For I& = 0 To UBound(DXTexturePool)
For I2% = I + 1 To UBound(DXTexturePool)
If DXTexturePool(I).cTextureName$ = DXTexturePool(I2).cTextureName$ Then
'Set a duplicate texture to be removed
DXTexturePool(I2).bInUse = False
bCanBeOptimized = True
End If
Next
If TypeName(DXTexturePool(I).cTexture) = "Nothing" Then DXTexturePool(I).bInUse = False
If Not DXTexturePool(I).bInUse Then bCanBeOptimized = True
Next

If Not bCanBeOptimized Then
Debug.Print "The texture pool is currently at an optimal state"
Exit Sub
End If

StartOfOptimize:
For I& = 0 To UBound(DXTexturePool)
If DXTexturePool(I).bInUse = False Then

Debug.Print "Optimized texture pool by removing " & iIF(DXTexturePool(I).cTextureName$ = "", "a blank texture", DXTexturePool(I).cTextureName$) & " at index " & I
If DXTexturePool(I).cTextureName$ <> "" Then
Set DXTexturePool(I).cTexture = Nothing
DXTexturePool(I).cTextureName$ = ""
End If

'Remove a redundant texture
If I& = UBound(DXTexturePool) Then
Set DXTexturePool(UBound(DXTexturePool)).cTexture = Nothing
If UBound(DXTexturePool) > 0 Then ReDim Preserve DXTexturePool(UBound(DXTexturePool) - 1&)
Else
DXTexturePool(I).bInUse = DXTexturePool(UBound(DXTexturePool)).bInUse
DXTexturePool(I).cLastUsedTick& = DXTexturePool(UBound(DXTexturePool)).cLastUsedTick&
Set DXTexturePool(I).cTexture = DXTexturePool(UBound(DXTexturePool)).cTexture
DXTexturePool(I).cTextureName$ = DXTexturePool(UBound(DXTexturePool)).cTextureName$
Set DXTexturePool(UBound(DXTexturePool)).cTexture = Nothing
ReDim Preserve DXTexturePool(UBound(DXTexturePool) - 1&)
GoTo StartOfOptimize
End If

End If
DoEvents
Next
Debug.Print "Optimized texture pool, unoptimized size was " & cUboundTexturePool% + 1 & ", new optimized size is " & UBound(DXTexturePool) + 1&
End Sub
Public Sub DXRenderT3DMesh(cMesh As DXT3DMesh, WhiteTexIndex As Long)
'Transforms and renders renders a T3D format mesh
Dim I As Long
Dim iMeshLoop As Long
DXfvfL
D3DDevice.SetTexture 0, DXTexturePool(WhiteTexIndex).cTexture
DXSetAlphaState
D3DXMatrixIdentity DXmatTemp
D3DXMatrixRotationX DXmatTemp, -90 * (Pi / 180)
D3DDevice.SetTransform D3DTS_WORLD, DXmatTemp
DXEnableAlpha
If UBound(cMesh.cMesh) > 0 Then
For iMeshLoop = 1 To UBound(cMesh.cMesh)

If cMesh.cMesh(iMeshLoop).cTex.cTexName = "" Then
D3DDevice.SetTexture 0, DXTexturePool(WhiteTexIndex).cTexture
Else
If cMesh.cMesh(iMeshLoop).cTex.cIndex = -1 Then cMesh.cMesh(iMeshLoop).cTex.cIndex = DXGetTextureIndex(cMesh.cMesh(iMeshLoop).cTex.cTexName)
If cMesh.cMesh(iMeshLoop).cTex.cIndex > -1 Then D3DDevice.SetTexture 0, DXTexturePool(cMesh.cMesh(iMeshLoop).cTex.cIndex).cTexture
End If
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLEFAN, cMesh.cMesh(iMeshLoop).cVertCount, cMesh.cMesh(iMeshLoop).cVerts(0), Len(cMesh.cMesh(iMeshLoop).cVerts(1))
Next
End If
DXDisableAlpha
D3DXMatrixIdentity DXmatTemp
D3DDevice.SetTransform D3DTS_WORLD, DXmatTemp
End Sub
Public Sub DXCopyTriTL(ByRef TLSource As TLVERTEX, ByRef TLDest As TLVERTEX)
TLDest.Color = TLSource.Color
TLDest.Specular = TLSource.Specular
TLDest.tu = TLSource.tu
TLDest.tv = TLSource.tv
TLDest.rhw = TLSource.rhw
TLDest.X = TLSource.X
TLDest.Y = TLSource.Y
TLDest.z = TLSource.z
End Sub
Public Sub DXCopyTri(ByRef TLSource As LITVERTEX, ByRef TLDest As LITVERTEX)
TLDest.Color = TLSource.Color
TLDest.Specular = TLSource.Specular
TLDest.tu = TLSource.tu
TLDest.tv = TLSource.tv
TLDest.X = TLSource.X
TLDest.Y = TLSource.Y
TLDest.z = TLSource.z
End Sub
Public Sub DXSwapTri(ByRef Tri1 As LITVERTEX, ByRef Tri2 As LITVERTEX)
Dim TempV As LITVERTEX
DXCopyTri Tri1, TempV
DXCopyTri Tri2, Tri1
DXCopyTri TempV, Tri2
End Sub
Public Sub DXCopyTriU(ByRef TLSource As UNLITVERTEX, ByRef TLDest As UNLITVERTEX)
TLDest.Diffuse = TLSource.Diffuse
TLDest.tu1 = TLSource.tu2
TLDest.tv1 = TLSource.tv2
TLDest.tu2 = TLSource.tu2
TLDest.tv2 = TLSource.tv2
TLDest.X = TLSource.X
TLDest.Y = TLSource.Y
TLDest.z = TLSource.z
End Sub
Private Sub ArrayShiftLeft(ByRef Inp() As String)
Dim I&
For I = 0 To UBound(Inp) - 1
Inp(I) = Inp(I + 1)
Next
ReDim Preserve Inp(UBound(Inp) - 1)
End Sub
Public Function DXCreateT3DMesh(T3DStr As String) As DXT3DMesh
On Error Resume Next
Dim I As Long
Dim I2 As Long
Dim iMeshLoop As Long
Dim TArr() As String
Dim TLineArr() As String
Dim SplitTextureStr() As String
Dim SplitPolyStr() As String
Dim cTex As String
Dim cVertexColor As Long
Dim cVertexLocation() As String
Dim cVertexOrigin() As String
Dim cVertexPos() As String
Dim TempLoc() As String
Dim TempArr() As String
Dim TexU() As String
Dim TexV() As String
ReDim DXCreateT3DMesh.cMesh(0)
ReDim DXCreateT3DMesh.cLights(0)
TLineArr = Split(T3DStr$, vbCrLf)
For iMeshLoop = 0 To UBound(TLineArr)
If Trim$(TLineArr(iMeshLoop)) <> "" Then
TArr = Split(TLineArr(iMeshLoop), " ")

Do While TArr(0) = "" And UBound(TArr) > 0
ArrayShiftLeft TArr()
Loop

If Strings.Left(TArr(0), 8) = "Location" Then
ReDim Preserve TArr(1)
TArr(1) = Strings.Right(TArr(0), Len(TArr(0)) - 9)
TArr(1) = Replace(TArr(1), "(", "")
TArr(1) = Replace(TArr(1), ")", "")
TempLoc = Split(TArr(1), ",")
ReDim Preserve TempLoc(2)
Erase cVertexLocation
ReDim cVertexLocation(2)
For I = 0 To 2
Select Case Strings.Left(TempLoc(I), 1)
Case "X"
cVertexLocation(0) = TempLoc(I)
Case "Y"
cVertexLocation(1) = TempLoc(I)
Case "Z"
cVertexLocation(2) = TempLoc(I)
End Select
Next
For I = 0 To 2
If Len(cVertexLocation(I)) > 2 Then cVertexLocation(I) = Strings.Right(cVertexLocation(I), Len(cVertexLocation(I)) - 2)
Next
End If

If UBound(TArr) > 0 Then
If TArr(0) = "TextureU" Then
TexU = Split(TArr(UBound(TArr)), ",")
End If
If TArr(0) = "TextureV" Then
TexV = Split(TArr(UBound(TArr)), ",")
End If
If TArr(0) = "End" And TArr(1) = "PolyList" Then
cVertexLocation = Split("0,0,0", ",")
End If
If TArr(0) = "Begin" And TArr(1) = "Polygon" Then
If UBound(TArr) >= 2 Then
For I = 2 To UBound(TArr)
If Strings.Left(TArr(I), "7") = "Texture" Then
DXAddTexture Strings.Right(TArr(I), Len(TArr(I)) - 8), D3DColorRGBA(0, 0, 0, 0)
DXCreateT3DMesh.cMesh(UBound(DXCreateT3DMesh.cMesh)).cTex.cIndex = DXGetTextureIndex(Strings.Right(TArr(I), Len(TArr(I)) - 8))
DXCreateT3DMesh.cMesh(UBound(DXCreateT3DMesh.cMesh)).cTex.cTexName = Strings.Right(TArr(I), Len(TArr(I)) - 8)
End If
Next
End If
With DXCreateT3DMesh

ReDim Preserve .cMesh(UBound(.cMesh) + 1)
ReDim .cMesh(UBound(.cMesh)).cVerts(0)
.cMesh(UBound(.cMesh)).cTex.cIndex = -1
End With
End If
If TArr(0) = "End" And TArr(1) = "Polygon" Then
With DXCreateT3DMesh.cMesh(UBound(DXCreateT3DMesh.cMesh))
'ReDim Preserve .cVerts(UBound(.cVerts) - 1)
cVertexOrigin = Split("0,0,0", ",")
End With
End If
If TArr(0) = "Origin" Then
cVertexOrigin = Split(TArr(UBound(TArr)), ",")
End If
If TArr(0) = "Vertex" Then
With DXCreateT3DMesh.cMesh(UBound(DXCreateT3DMesh.cMesh)).cVerts(UBound(DXCreateT3DMesh.cMesh(UBound(DXCreateT3DMesh.cMesh)).cVerts))
cVertexPos = Split(TArr(UBound(TArr)), ",")
For I = 0 To 2
If cVertexLocation(I) <> "" Then cVertexPos(I) = CSng(cVertexPos(I)) + CSng(cVertexLocation(I))
Next
.X = cVertexPos(0) / 128
.Y = cVertexPos(1) / 128
.z = cVertexPos(2) / 128
'Select Case UBound(DXCreateT3DMesh.cMesh(UBound(DXCreateT3DMesh.cMesh)).cVerts)
'End Select
.Color = D3DColorRGBA(255, 255, 255, 255)
DXCreateT3DMesh.cMesh(UBound(DXCreateT3DMesh.cMesh)).cVerts(UBound(DXCreateT3DMesh.cMesh(UBound(DXCreateT3DMesh.cMesh)).cVerts)) = _
DXCreateLitVertex(.X, .Y, .z, .Color, 0, TexU(0) / 1, TexV(0) / 1)
End With
With DXCreateT3DMesh.cMesh(UBound(DXCreateT3DMesh.cMesh))
ReDim Preserve .cVerts(UBound(.cVerts) + 1)
End With
End If
End If

End If
Next
For I = 1 To UBound(DXCreateT3DMesh.cMesh)
With DXCreateT3DMesh.cMesh(I)
If UBound(.cVerts) = 4 Then
ReDim Preserve .cVerts(6)
End If
.cVertCount = UBound(.cVerts) / 3
End With
Next
Debug.Print "Mesh created, " & UBound(DXCreateT3DMesh.cMesh) & " faces, " & UBound(DXCreateT3DMesh.cLights) & " lights"
End Function
Public Function DXCreateT3DMeshMessed(T3DStr As String) As DXT3DMesh
'Creates a mesh from the input string of a T3D file
Dim I As Long
Dim I2 As Long
Dim iMeshLoop As Long
Dim TArr() As String
Dim TLineArr() As String
Dim SplitTextureStr() As String
Dim SplitPolyStr() As String
Dim cTex As String
Dim cVertex As Long
Dim cActualVertex As Long
Dim cVertexColor As Long
Dim cVertexBrushLocation(2) As Single
Dim cVertexOrigin(2) As Single
Dim cVertexRots(2) As Single
Dim TempRotAng As Single
Dim IsSheet As Boolean
Dim RotMatrix(2) As D3DMATRIX
Dim FinalRotMatrix As D3DMATRIX
Dim RotPoint As D3DVECTOR4
Dim RotPointTemp As D3DVECTOR
Dim cVertexNormal(2) As Single
Dim cTag As String
ReDim DXCreateT3DMeshMessed.cMesh(0)
ReDim DXCreateT3DMeshMessed.cLights(0)
TLineArr = Split(T3DStr$, vbCrLf)
For iMeshLoop = 0 To UBound(TLineArr)
'Dont parse empty lines
If Trim$(TLineArr(iMeshLoop)) <> "" Then
'Loop through the lines of the T3D mesh's string
TArr = Split(TLineArr(iMeshLoop), " ")

'Remove blank spaces
Do While TArr(0) = "" And UBound(TArr) > 0
ArrayShiftLeft TArr()
Loop

If UBound(TArr) >= 2 Then
'Check for texture for poly
If TArr(0) = "Begin" And TArr(1) = "Polygon" Then
SplitTextureStr = Split(TArr(2), "=")
If SplitTextureStr(0) = "Texture" Then
'Retrieve new texture for the new polygon
cTex = SplitTextureStr(1)
End If
End If
End If

If UBound(TArr) >= 1 Then
If TArr(0) = "End" And TArr(1) = "PolyList" Then
For I = 0 To UBound(cVertexBrushLocation)
cVertexBrushLocation(I) = 0
Next
For I = 0 To UBound(cVertexRots)
cVertexRots(I) = 0
Next
End If
End If

If TArr(0) = "Origin" Then
SplitPolyStr = Split(TArr(UBound(TArr)), ",")
For I = 0 To 2
cVertexOrigin(I) = SplitPolyStr(I)
Next
End If
If TArr(0) = "Normal" Then
SplitPolyStr = Split(TArr(UBound(TArr)), ",")
For I = 0 To 2
cVertexNormal(I) = SplitPolyStr(I)
Next
End If

If Strings.Left(TArr(0), 3) = "Tag" Then
cTag = Strings.Right(TArr(0), Len(TArr(0)) - 4)
End If

If UBound(TArr) >= 1 Then
If cTag = Chr(34) & "Light" & Chr(34) Then
With DXCreateT3DMeshMessed.cLights(UBound(DXCreateT3DMeshMessed.cLights))
.Position.X = cVertexBrushLocation(0) / 128
.Position.Y = cVertexBrushLocation(1) / 128
.Position.z = cVertexBrushLocation(2) / 128
.Diffuse.R = 255
.Diffuse.G = 255
.Diffuse.B = 255
.Range = 256
.Attenuation1 = 0.05
End With
ReDim Preserve DXCreateT3DMeshMessed.cLights(UBound(DXCreateT3DMeshMessed.cLights) + 1)
End If
End If

If Strings.Left(TArr(0), 8) = "Location" Then
TArr(0) = Replace(TArr(0), "(", "")
TArr(0) = Replace(TArr(0), ")", "")
TArr(0) = Strings.Right(TArr(0), Len(TArr(0)) - 9)
SplitTextureStr = Split(TArr(0), ",")
For I = 0 To 2
cVertexBrushLocation(I) = 0
Next
For I2 = 0 To UBound(SplitTextureStr)
SplitPolyStr = Split(SplitTextureStr(I2), "=")
Select Case SplitPolyStr(0)
Case "X"
cVertexBrushLocation(0) = SplitPolyStr(1)
Case "Y"
cVertexBrushLocation(1) = SplitPolyStr(1)
Case "Z"
cVertexBrushLocation(2) = SplitPolyStr(1)
End Select
Next
End If

If UBound(TArr) >= 1 Then
If TArr(0) = "End" And TArr(1) = "Actor" Then
For I = 0 To 2
cVertexRots(I) = 0
Next
End If
End If

If Strings.Left(TArr(0), 8) = "Rotation" Then
TArr(0) = Replace(TArr(0), "(", "")
TArr(0) = Replace(TArr(0), ")", "")
TArr(0) = Strings.Right(TArr(0), Len(TArr(0)) - 9)
SplitTextureStr = Split(TArr(0), ",")
For I2 = 0 To UBound(SplitTextureStr)
SplitPolyStr = Split(SplitTextureStr(I2), "=")
Select Case SplitPolyStr(0)
Case "Yaw"
cVertexRots(0) = SplitPolyStr(1) / 65536
Case "Pitch"
cVertexRots(1) = SplitPolyStr(1) / 65536
Case "Roll"
cVertexRots(2) = SplitPolyStr(1) / 65536
End Select
Next
End If

If UBound(TArr) >= 2 Then
If TArr(0) = "Begin" And TArr(1) = "Polygon" Then
'Initialize new polygon
DXCreateT3DMeshMessed.cMesh(UBound(DXCreateT3DMeshMessed.cMesh)).cVertCount = cVertex
DXCreateT3DMeshMessed.cMesh(UBound(DXCreateT3DMeshMessed.cMesh)).cNormal = DXMakeVector(cVertexNormal(0), cVertexNormal(1), cVertexNormal(2))
With DXCreateT3DMeshMessed.cMesh(UBound(DXCreateT3DMeshMessed.cMesh))
.cTex.cIndex = -1
End With
If cVertex = 3 Or cVertex = 4 And IsSheet Then
With DXCreateT3DMeshMessed.cMesh(UBound(DXCreateT3DMeshMessed.cMesh))
.bHasPlane = True
If .cVerts(1).Y = .cVerts(2).Y And .cVerts(2).Y = .cVerts(3).Y And .cVerts(3).Y = .cVerts(4).Y Then
.cPlane.B = 1
End If
End With
End If
For I = 0 To 2
cVertexOrigin(I) = 0
cVertexNormal(I) = 0
Next
IsSheet = False
If TArr(2) = "Item=Sheet" Then
IsSheet = True
End If
ReDim Preserve DXCreateT3DMeshMessed.cMesh(UBound(DXCreateT3DMeshMessed.cMesh) + 1)
With DXCreateT3DMeshMessed.cMesh(UBound(DXCreateT3DMeshMessed.cMesh))
.cTex.cIndex = -1
End With
DXCreateT3DMeshMessed.cMesh(UBound(DXCreateT3DMeshMessed.cMesh)).cTex.cTexName = cTex
For I = 0 To UBound(DXTexturePool)
If DXTexturePool(I).cTextureName = cTex Then
DXCreateT3DMeshMessed.cMesh(UBound(DXCreateT3DMeshMessed.cMesh)).cTex.cIndex = I
Exit For
End If
Next
cVertex = 0
End If
End If

cVertexColor = D3DColorRGBA(0, 0, 0, 255)

If UBound(TArr) >= 1 And cVertex < 4 Then
'Found a vertex
If TArr(0) = "Vertex" Then
'Retrieve coordinates
SplitPolyStr = Split(TArr(UBound(TArr)), ",")
'Find the actual vertex
cActualVertex = cVertex + 1
If cVertex = 0 Then ReDim DXCreateT3DMeshMessed.cMesh(UBound(DXCreateT3DMeshMessed.cMesh)).cVerts(32)
With DXCreateT3DMeshMessed.cMesh(UBound(DXCreateT3DMeshMessed.cMesh)).cVerts(cActualVertex)
'Create the vertex with the unreal modifications
RotPointTemp.X = SplitPolyStr(0)
RotPointTemp.Y = SplitPolyStr(1)
RotPointTemp.z = SplitPolyStr(2)
D3DXMatrixIdentity FinalRotMatrix
D3DXVec3Transform RotPoint, RotPointTemp, FinalRotMatrix
.X = (RotPoint.X + cVertexBrushLocation(0)) / 128
.Y = (RotPoint.Y + cVertexBrushLocation(1)) / 128
.z = (RotPoint.z + cVertexBrushLocation(2)) / 128
cVertex = cVertex + 1
.Color = cVertexColor
End With
End If
End If

End If
'Copy the vertexes to fill the surface
'If cActualVertex = 1 Then DXCopyTri DXCreateT3DMeshMessed.cMesh(UBound(DXCreateT3DMeshMessed.cMesh)).cVerts(cActualVertex), DXCreateT3DMeshMessed.cMesh(UBound(DXCreateT3DMeshMessed.cMesh)).cVerts(5) '5
'If cActualVertex = 2 Then DXCopyTri DXCreateT3DMeshMessed.cMesh(UBound(DXCreateT3DMeshMessed.cMesh)).cVerts(cActualVertex), DXCreateT3DMeshMessed.cMesh(UBound(DXCreateT3DMeshMessed.cMesh)).cVerts(4) '4
'If cActualVertex = 3 Then DXCopyTri DXCreateT3DMeshMessed.cMesh(UBound(DXCreateT3DMeshMessed.cMesh)).cVerts(cActualVertex), DXCreateT3DMeshMessed.cMesh(UBound(DXCreateT3DMeshMessed.cMesh)).cVerts(6) '6
Next
Dim Ix&
For Ix = 0 To UBound(DXCreateT3DMeshMessed.cMesh)
With DXCreateT3DMeshMessed
'ReDim Preserve .cMesh(Ix).cVerts(UBound(.cMesh(Ix).cVerts))
'.cMesh(Ix).cVertCount = .cMesh(Ix).cVertCount - 2
End With
Next
Debug.Print "Mesh created, " & UBound(DXCreateT3DMeshMessed.cMesh) & " faces"
End Function
Public Function DXCreateXMesh(XFilePath As String) As DXMeshObject
'Creates a mesh from the specified file
Dim I As Long
Dim I2 As Long
With DXCreateXMesh
'Create mesh
XFilePath = App.Path & "\" & XFilePath
Set .cMesh = D3DX.LoadMeshFromX(XFilePath, D3DXMESH_MANAGED, D3DDevice, .cAdjacency, .cMatBuffer, DXCreateXMesh.cMaterialsNum)
'Define memory
ReDim .cMaterials(.cMaterialsNum - 1)
ReDim .cTextures(.cMaterialsNum - 1)
'Reset matrix
D3DXMatrixIdentity .cMatrix

'Load textures
For I = 0 To .cMaterialsNum - 1
D3DX.BufferGetMaterial .cMatBuffer, I, .cMaterials(I)
.cMaterials(I).Ambient = .cMaterials(I).Diffuse
.cTextures(I).cTexName = D3DX.BufferGetTextureName(.cMatBuffer, I)
'Put texture into texture pool
If .cTextures(I).cTexName <> "" Then
DXAddTexture .cTextures(I).cTexName
Set .cTextures(I).cTexAdv = D3DX.CreateTextureFromFile(D3DDevice, App.Path & "\" & .cTextures(I).cTexName)
End If
Next

ReDim .cVerts(.cMesh.GetNumVertices) As D3DVERTEX
D3DXMeshVertexBuffer8GetData .cMesh, 0, Len(.cVerts(0)) * .cMesh.GetNumVertices, 0, .cVerts(0)

D3DX.ComputeBoundingBoxFromMesh .cMesh, .cBoundingBoxMin, .cBoundingBoxMax
.cBoundingBoxCenter.X = (.cBoundingBoxMin.X + .cBoundingBoxMax.X) / 2
.cBoundingBoxCenter.Y = (.cBoundingBoxMin.Y + .cBoundingBoxMax.Y) / 2
.cBoundingBoxCenter.z = (.cBoundingBoxMin.z + .cBoundingBoxMax.z) / 2
D3DX.ComputeBoundingSphereFromMesh .cMesh, .cBoundingSphereCenter, .cBoundingSphereRadius

Set .cMatBuffer = Nothing
End With
End Function
Public Sub DXUpdateXMeshBounds(ByRef cXMesh As DXMeshObject)
On Local Error Resume Next
With cXMesh
Dim TempVec As D3DVECTOR4
D3DXMeshVertexBuffer8GetData .cMesh, 0, Len(.cVerts(0)) * .cMesh.GetNumVertices, 0, .cVerts(0)
D3DX.ComputeBoundingBoxFromMesh .cMesh, .cBoundingBoxMin, .cBoundingBoxMax
D3DXVec3Transform TempVec, .cBoundingBoxMin, .cMatrix
.cBoundingBoxMin.X = TempVec.X
.cBoundingBoxMin.Y = TempVec.Y
.cBoundingBoxMin.z = TempVec.z
D3DXVec3Transform TempVec, .cBoundingBoxMax, .cMatrix
.cBoundingBoxMax.X = TempVec.X
.cBoundingBoxMax.Y = TempVec.Y
.cBoundingBoxMax.z = TempVec.z
'D3DXMeshVertexBuffer8SetData .cMesh, 0, Len(.cVerts(0)) * .cMesh.GetNumVertices, 0, .cVerts(0)
'D3DX.ComputeBoundingBoxFromMesh .cMesh, .cBoundingBoxMin, .cBoundingBoxMax
.cBoundingBoxCenter.X = (.cBoundingBoxMin.X + .cBoundingBoxMax.X) / 2
.cBoundingBoxCenter.Y = (.cBoundingBoxMin.Y + .cBoundingBoxMax.Y) / 2
.cBoundingBoxCenter.z = (.cBoundingBoxMin.z + .cBoundingBoxMax.z) / 2
D3DX.ComputeBoundingSphereFromMesh .cMesh, .cBoundingSphereCenter, .cBoundingSphereRadius
D3DXVec3Transform .cBoundingSphereCenterUntransformed, .cBoundingSphereCenter, .cMatrix
.cBoundingSphereCenter.X = .cBoundingSphereCenterUntransformed.X
.cBoundingSphereCenter.Y = .cBoundingSphereCenterUntransformed.Y
.cBoundingSphereCenter.z = .cBoundingSphereCenterUntransformed.z
End With
End Sub
Public Function DXRenderXMesh(ByRef cXMesh As DXMeshObject, Optional UseTexPoolIndex As Boolean, Optional UseFixTexPoolIndex As Long, Optional UseShadowMatrix As Boolean, Optional UseReflectionMatrix As Boolean, Optional UseShadowMaterial As Boolean)
'Renders a mesh onto the screen
Dim TempFVF As Long
Dim I As Long
With cXMesh
TempFVF = D3DDevice.GetVertexShader
D3DDevice.SetVertexShader .cMesh.GetFVF

Dim ShadowMaterial As D3DMATERIAL8

If UseShadowMatrix Then D3DDevice.SetTransform D3DTS_WORLD, .cShadowMatrix
If UseReflectionMatrix Then D3DDevice.SetTransform D3DTS_WORLD, .cReflectMatrix
If Not UseShadowMatrix And Not UseReflectionMatrix Then D3DDevice.SetTransform D3DTS_WORLD, .cMatrix
For I = 0 To .cMaterialsNum - 1

D3DDevice.SetMaterial .cMaterials(I)
If UseShadowMaterial Then D3DDevice.SetMaterial ShadowMaterial
If Not UseTexPoolIndex Then
If .cTextures(I).cTexName <> "" Then
D3DDevice.SetTexture 0, .cTextures(I).cTexAdv
Else
D3DDevice.SetTexture 0, Nothing
End If
Else
D3DDevice.SetTexture 0, DXTexturePool(UseFixTexPoolIndex).cTexture
End If

.cMesh.DrawSubset I
Next
D3DXMatrixIdentity NullMatrix
D3DDevice.SetTransform D3DTS_WORLD, NullMatrix
End With
D3DDevice.SetVertexShader TempFVF
End Function
Public Function DXUnloadXMesh(ByRef cXMesh As DXMeshObject)
'Deletes the contents of the specified mesh
Dim I As Long
On Error Resume Next
For I = 0 To UBound(cXMesh.cTextures)
Set cXMesh.cTextures(I).cTexAdv = Nothing
Next
Erase cXMesh.cMaterials
Erase cXMesh.cTextures
Set cXMesh.cMesh = Nothing
Set cXMesh.cMatBuffer = Nothing
cXMesh.cMaterialsNum = 0
End Function
Public Function DXIntersectTriangleCull(ByRef v0 As D3DVECTOR, ByRef V1 As D3DVECTOR, ByRef V2 As D3DVECTOR, vDir As D3DVECTOR, vOrig As D3DVECTOR, T As Single, U As Single, v As Single) As Boolean
'Returns true or false depending on if a triangle specified as v0,v1,v2 has been intersected by a ray from vOrig, vector vDir

    Dim edge1 As D3DVECTOR
    Dim edge2 As D3DVECTOR
    Dim pvec As D3DVECTOR
    Dim tVec As D3DVECTOR
    Dim qvec As D3DVECTOR
    Dim det As Single
    Dim fInvDet As Single
    
    'find vectors for the two edges sharing vert0
    D3DXVec3Subtract edge1, V1, v0
    D3DXVec3Subtract edge2, V2, v0
    
    'begin calculating the determinant - also used to caclulate u parameter
    D3DXVec3Cross pvec, vDir, edge2
    
    'if determinant is nearly zero, ray lies in plane of triangle
    det = D3DXVec3Dot(edge1, pvec)
    If (det < 0.0001) Then
        Exit Function
    End If
    
    'calculate distance from vert0 to ray origin
    D3DXVec3Subtract tVec, vOrig, v0

    'calculate u parameter and test bounds
    U = D3DXVec3Dot(tVec, pvec)
    If (U < 0 Or U > det) Then
        Exit Function
    End If
    
    'prepare to test v parameter
    D3DXVec3Cross qvec, tVec, edge1
    
    'calculate v parameter and test bounds
    v = D3DXVec3Dot(vDir, qvec)
    If (v < 0 Or (U + v > det)) Then
        Exit Function
    End If
    
    'calculate t, scale parameters, ray intersects triangle
    T = D3DXVec3Dot(edge2, qvec)
    fInvDet = 1 / det
    T = T * fInvDet
    U = U * fInvDet
    v = v * fInvDet
    If T = 0 Then Exit Function
    
    DXIntersectTriangleCull = True
    
End Function
Public Sub DXRender2DBox(X1!, Y1!, X2!, Y2!, cAlpha As Boolean, Filled As Boolean, Color As Long)
    
    If cAlpha Then
    DXSetAlphaState
    DXEnableAlpha
    End If
    D3DDevice.SetVertexShader FVF_TLVERTEX
    D3DDevice.SetTexture 0, Nothing
    v(0) = DXCreateTLVertex(X1, Y1, 0, 0, Color, 0, 0, 0)
    v(1) = DXCreateTLVertex(X2, Y1, 0, 0, Color, 0, 0, 0)
    If Filled Then
        v(2) = DXCreateTLVertex(X1, Y2, 0, 0, Color, 0, 0, 0)
        v(3) = DXCreateTLVertex(X2, Y2, 0, 0, Color, 0, 0, 0)
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, v(0), Len(v(0))
    Else
        v(2) = DXCreateTLVertex(X2, Y2, 0, 0, Color, 0, 0, 0)
        v(3) = DXCreateTLVertex(X1, Y2, 0, 0, Color, 0, 0, 0)
        v(4) = v(0)
        D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, 4, v(0), Len(v(0))
    End If
    If cAlpha Then
    DXDisableAlpha
    End If

End Sub
Public Sub DXRender2DLine(X1!, Y1!, X2!, Y2!, Color As Long)
    vlt(0) = DXCreateTLVertex(X1, Y1, 0, 0, Color, 0, 0, 0)
    vlt(1) = DXCreateTLVertex(X2, Y2, 0, 0, Color, 0, 0, 0)
    D3DDevice.SetTexture 0, Nothing
    D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 1, vlt(0), Len(vlt(0))
End Sub
Public Sub DXRender2DLineAdv(X1!, Y1!, X2!, Y2!, StartColor As Long, EndColor As Long)
    vlt(0) = DXCreateTLVertex(X1, Y1, 0, 0, StartColor, 0, 0, 0)
    vlt(1) = DXCreateTLVertex(X2, Y2, 0, 0, EndColor, 0, 0, 0)
    D3DDevice.SetTexture 0, Nothing
    D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 1, vlt(0), Len(vlt(0))
End Sub
Public Sub DXRender2DBoxTexturedFast(X1!, Y1!, X2!, Y2!, NOcAlpha As Boolean, Color As Long, Optional NOUOffset!, Optional NOVOffset!)

    v(0) = DXCreateTLVertex(X1, Y1, 0, 0, Color, 0, 0, 0)
    v(1) = DXCreateTLVertex(X2, Y1, 0, 0, Color, 0, 1, 0)
    v(2) = DXCreateTLVertex(X1, Y2, 0, 0, Color, 0, 0, 1)
    v(3) = DXCreateTLVertex(X2, Y2, 0, 0, Color, 0, 1, 1)
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, v(0), Len(v(0))

End Sub
Public Sub DXRender2DBoxTexturedWith2Tex(X1!, Y1!, X2!, Y2!, cAlpha As Boolean, Color As Long, Optional SkipVertexCreation As Boolean)
    
    If cAlpha Then
    DXSetAlphaState
    DXEnableAlpha
    End If
    
    If Not SkipVertexCreation Then
    V2(0) = DXCreateTLVertex2(X1, Y1, 0, 1, Color, 0, 0, 0, 0, 0)
    V2(1) = DXCreateTLVertex2(X2, Y1, 0, 1, Color, 0, 1, 0, 1, 0)
    V2(2) = DXCreateTLVertex2(X1, Y2, 0, 1, Color, 0, 0, 1, 0, 1)
    V2(3) = DXCreateTLVertex2(X2, Y2, 0, 1, Color, 0, 1, 1, 1, 1)
    End If
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, V2(0), Len(V2(0))
    
    If cAlpha Then
    DXDisableAlpha
    End If

End Sub
Public Sub DXRender2DBoxTexturedWith2TexRepeat(X1!, Y1!, X2!, Y2!, cAlpha As Boolean, Color As Long)

    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, V2(0), Len(V2(0))

End Sub
Public Sub DXRender2DBoxTextured(X1!, Y1!, X2!, Y2!, cAlpha As Boolean, Color As Long, Optional UOffset!, Optional VOffset!)
    
    If cAlpha Then
    DXSetAlphaState
    DXEnableAlpha
    End If
    D3DDevice.SetVertexShader FVF_TLVERTEX
    v(0) = DXCreateTLVertex(X1, Y1, 0, 0, Color, 0, 0 + UOffset, 0 + VOffset)
    v(1) = DXCreateTLVertex(X2, Y1, 0, 0, Color, 0, 1 + UOffset, 0 + VOffset)
    v(2) = DXCreateTLVertex(X1, Y2, 0, 0, Color, 0, 0 + UOffset, 1 + VOffset)
    v(3) = DXCreateTLVertex(X2, Y2, 0, 0, Color, 0, 1 + UOffset, 1 + VOffset)
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, v(0), Len(v(0))
    If cAlpha Then
    DXDisableAlpha
    End If

End Sub
Public Sub DXRender3DBox(X1!, Y1!, Z1!, X2!, Y2!, Z2!, cAlpha As Boolean, Filled As Boolean, Color As Long)

    If cAlpha Then
    DXSetAlphaState
    DXEnableAlpha
    End If
    D3DDevice.SetVertexShader FVF_LVERTEX
    D3DDevice.SetTexture 0, Nothing
vl(0) = DXCreateLVertex(X1, Y1, Z1, Color, 1, 0, 0)
vl(1) = DXCreateLVertex(X2, Y1, Z1, Color, 1, 0, 0)
vl(2) = DXCreateLVertex(X1, Y2, Z1, Color, 1, 0, 0)

vl(3) = DXCreateLVertex(X2, Y1, Z1, Color, 1, 0, 0)
vl(4) = DXCreateLVertex(X1, Y2, Z1, Color, 1, 0, 0)
vl(5) = DXCreateLVertex(X2, Y2, Z1, Color, 1, 0, 0)
'
vl(6) = DXCreateLVertex(X1, Y1, Z2, Color, 1, 0, 0)
vl(7) = DXCreateLVertex(X2, Y1, Z2, Color, 1, 0, 0)
vl(8) = DXCreateLVertex(X1, Y2, Z2, Color, 1, 0, 0)

vl(9) = DXCreateLVertex(X2, Y1, Z2, Color, 1, 0, 0)
vl(10) = DXCreateLVertex(X1, Y2, Z2, Color, 1, 0, 0)
vl(11) = DXCreateLVertex(X2, Y2, Z2, Color, 1, 0, 0)
'
vl(12) = DXCreateLVertex(X1, Y1, Z1, Color, 1, 0, 0)
vl(13) = DXCreateLVertex(X1, Y2, Z1, Color, 1, 0, 0)
vl(14) = DXCreateLVertex(X1, Y2, Z2, Color, 1, 0, 0)

vl(15) = DXCreateLVertex(X1, Y1, Z1, Color, 1, 0, 0)
vl(16) = DXCreateLVertex(X1, Y1, Z2, Color, 1, 0, 0)
vl(17) = DXCreateLVertex(X1, Y2, Z2, Color, 1, 0, 0)
'
vl(18) = DXCreateLVertex(X2, Y1, Z1, Color, 1, 0, 0)
vl(19) = DXCreateLVertex(X2, Y2, Z1, Color, 1, 0, 0)
vl(20) = DXCreateLVertex(X2, Y2, Z2, Color, 1, 0, 0)

vl(21) = DXCreateLVertex(X2, Y1, Z1, Color, 1, 0, 0)
vl(22) = DXCreateLVertex(X2, Y1, Z2, Color, 1, 0, 0)
vl(23) = DXCreateLVertex(X2, Y2, Z2, Color, 1, 0, 0)
'
vl(24) = DXCreateLVertex(X1, Y2, Z1, Color, 1, 0, 0)
vl(25) = DXCreateLVertex(X1, Y2, Z2, Color, 1, 0, 0)
vl(26) = DXCreateLVertex(X2, Y2, Z1, Color, 1, 0, 0)

vl(27) = DXCreateLVertex(X1, Y2, Z2, Color, 1, 0, 0)
vl(28) = DXCreateLVertex(X2, Y2, Z1, Color, 1, 0, 0)
vl(29) = DXCreateLVertex(X2, Y2, Z2, Color, 1, 0, 0)
'
vl(30) = DXCreateLVertex(X1, Y1, Z1, Color, 1, 0, 0)
vl(31) = DXCreateLVertex(X1, Y1, Z2, Color, 1, 0, 0)
vl(32) = DXCreateLVertex(X2, Y1, Z1, Color, 1, 0, 0)

vl(33) = DXCreateLVertex(X1, Y1, Z2, Color, 1, 0, 0)
vl(34) = DXCreateLVertex(X2, Y1, Z1, Color, 1, 0, 0)
vl(35) = DXCreateLVertex(X2, Y1, Z2, Color, 1, 0, 0)
'
    If Filled Then
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, (UBound(v) + 1) / 3, vl(0), Len(vl(0))
    Else
        D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, UBound(v), vl(0), Len(vl(0))
    End If
    If cAlpha Then
    DXDisableAlpha
    End If

End Sub
Public Function DXGenerateTriangleNormals(p0 As UNLITVERTEX, P1 As UNLITVERTEX, P2 As UNLITVERTEX) As D3DVECTOR


'//0. Variables required
    Dim v01 As D3DVECTOR 'Vector from points 0 to 1
    Dim v02 As D3DVECTOR 'Vector from points 0 to 2
    Dim vNorm As D3DVECTOR 'The final vector


'//1. Create the vectors from points 0 to 1 and 0 to 2
    D3DXVec3Subtract v01, DXMakeVector(P1.X, P1.Y, P1.z), DXMakeVector(p0.X, p0.Y, p0.z)
    D3DXVec3Subtract v02, DXMakeVector(P2.X, P2.Y, P2.z), DXMakeVector(p0.X, p0.Y, p0.z)


'//2. Get the cross product
    D3DXVec3Cross vNorm, v01, v02


'//3. Normalize this vector
    D3DXVec3Normalize vNorm, vNorm


'//4. Return the value
    DXGenerateTriangleNormals.X = vNorm.X
    DXGenerateTriangleNormals.Y = vNorm.Y
    DXGenerateTriangleNormals.z = vNorm.z


End Function
Public Function DXReturn2DTriStrip(TexWidth!, TexHeight!, TexRealWidth!, TexRealHeight!, TexAlpha!, IsShadow As Boolean) As TLVERTEX()
TexWidth = TexWidth / 2
TexHeight = TexHeight / 2
TexRealWidth = TexRealWidth / 2
TexRealHeight = TexRealHeight / 2
tStrip(0) = DXCreateTLVertex(-TexWidth, -TexHeight, 0, 1, D3DColorRGBA(255, 255, 255, TexAlpha), 0, 0, 0)
tStrip(1) = DXCreateTLVertex(TexWidth, -TexHeight, 0, 1, D3DColorRGBA(255, 255, 255, TexAlpha), 0, (TexWidth / TexRealWidth), 0)
tStrip(2) = DXCreateTLVertex(-TexWidth, TexHeight, 0, 1, D3DColorRGBA(255, 255, 255, TexAlpha), 0, 0, (TexHeight / TexRealHeight))
tStrip(3) = DXCreateTLVertex(TexWidth, TexHeight, 0, 1, D3DColorRGBA(255, 255, 255, TexAlpha), 0, (TexWidth / TexRealWidth), (TexHeight / TexRealHeight))
DXReturn2DTriStrip = tStrip
End Function
Public Sub DXRender2DTexture(TextureName As String, BlendTexture As String, X!, Y!, TexHeight!, TexWidth!, TexAlpha!, RotateAngle!, XScale!, YScale!, DrawShadow As Boolean, ShadowOffsetX!, ShadowOffsetY!, ShadowAlpha!, Optional UOffset!, Optional VOffset!, Optional AdditiveBlend As Boolean)
If TexAlpha > 255 Then TexAlpha = 255
TexWidth = TexWidth / 2
TexHeight = TexHeight / 2
Do While UOffset > 1
UOffset = UOffset - 1
Loop
Do While VOffset > 1
VOffset = VOffset - 1
Loop
Do While UOffset < -1
UOffset = UOffset + 1
Loop
Do While VOffset < -1
VOffset = VOffset + 1
Loop
tStrip(0) = DXCreateTLVertex(-TexWidth, -TexHeight, 0, 1, D3DColorRGBA(255, 255, 255, TexAlpha), 0, UOffset, VOffset)
tStrip(1) = DXCreateTLVertex(TexWidth, -TexHeight, 0, 1, D3DColorRGBA(255, 255, 255, TexAlpha), 0, 1 + UOffset, VOffset)
tStrip(2) = DXCreateTLVertex(-TexWidth, TexHeight, 0, 1, D3DColorRGBA(255, 255, 255, TexAlpha), 0, UOffset, 1 + VOffset)
tStrip(3) = DXCreateTLVertex(TexWidth, TexHeight, 0, 1, D3DColorRGBA(255, 255, 255, TexAlpha), 0, 1 + UOffset, 1 + VOffset)
DXTransform2DTriStrip tStrip, RotateAngle, RotateAngle, XScale, YScale, 0, 0, X, Y
DXSetTexture 0, TextureName
DXSetAlphaState
If AdditiveBlend Then
DXSetAlphaOneState
End If
If BlendTexture <> "" Then
DXSetTexture 1, BlendTexture
End If
DXEnableAlpha
If DrawShadow Then
TriStripTemp = tStrip
For I = 0 To 3
tStrip(I).X = tStrip(I).X + ShadowOffsetX
tStrip(I).Y = tStrip(I).Y + ShadowOffsetY
tStrip(I).Color = D3DColorRGBA(0, 0, 0, ShadowAlpha)
Next
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tStrip(0), Len(tStrip(0))
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriStripTemp(0), Len(TriStripTemp(0))
Else
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tStrip(0), Len(tStrip(0))
End If
DXDisableAlpha
End Sub
Public Sub DXRender2DTextureSprite(TextureName As String, X!, Y!, TexHeight!, TexWidth!, TexAlpha!, RotateAngle!, XScale!, YScale!, DrawShadow As Boolean, ShadowOffsetX!, ShadowOffsetY!, ShadowAlpha!, FrameWidth&, FrameHeight&, CurrentFrame!, FrameTexHeight&, FrameTexWidth&, UseLight As Boolean, LightR!, LightG!, LightB!)
Dim ctU!
Dim ctV!
Dim cntU!
Dim cntV!
ctV = 0
cntV = FrameHeight
Dim FramesPerRow As Single
FramesPerRow = FrameTexWidth / FrameWidth
Do While CurrentFrame > FramesPerRow
CurrentFrame = CurrentFrame - FramesPerRow
ctV = ctV + FrameHeight
cntV = cntV + FrameHeight
Loop
ctU = FrameWidth * CurrentFrame
cntU = FrameWidth * (CurrentFrame + 1)
ctU = ctU / FrameTexWidth
ctV = ctV / FrameTexHeight
cntU = cntU / FrameTexWidth
cntV = cntV / FrameTexHeight
If TexAlpha > 255 Then TexAlpha = 255
tStrip(0) = DXCreateTLVertex(-TexWidth / 2, -TexHeight / 2, 0, 1, D3DColorRGBA(255 * LightR, 255 * LightG, 255 * LightB, TexAlpha), 0, ctU, ctV)
tStrip(1) = DXCreateTLVertex(TexWidth / 2, -TexHeight / 2, 0, 1, D3DColorRGBA(255 * LightR, 255 * LightG, 255 * LightB, TexAlpha), 0, cntU, ctV)
tStrip(2) = DXCreateTLVertex(-TexWidth / 2, TexHeight / 2, 0, 1, D3DColorRGBA(255 * LightR, 255 * LightG, 255 * LightB, TexAlpha), 0, ctU, cntV)
tStrip(3) = DXCreateTLVertex(TexWidth / 2, TexHeight / 2, 0, 1, D3DColorRGBA(255 * LightR, 255 * LightG, 255 * LightB, TexAlpha), 0, cntU, cntV)
If UseLight Then
For I = 0 To 3
tStrip(I).Color = D3DColorRGBA(LightR * 255, LightG * 255, LightB * 255, TexAlpha)
Next
End If
DXTransform2DTriStrip tStrip, RotateAngle, RotateAngle, XScale, YScale, 0, 0, X, Y
DXSetTexture 0, TextureName
DXSetAlphaState
DXEnableAlpha
If DrawShadow Then
TriStripTemp = tStrip
For I = 0 To 3
tStrip(I).X = tStrip(I).X + ShadowOffsetX
tStrip(I).Y = tStrip(I).Y + ShadowOffsetY
tStrip(I).Color = D3DColorRGBA(0, 0, 0, ShadowAlpha)
Next
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, UBound(tStrip) - 1, tStrip(0), Len(tStrip(0))
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, UBound(TriStripTemp) - 1, TriStripTemp(0), Len(TriStripTemp(0))
Else
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, UBound(tStrip) - 1, tStrip(0), Len(tStrip(0))
End If
DXDisableAlpha
End Sub
Public Sub DXTransform2DTriStrip(ByRef TriStrip() As TLVERTEX, RotateX As Single, RotateY As Single, XScale As Single, YScale As Single, CurX As Single, CurY As Single, NewX As Single, NewY As Single)
Dim I As Long
D3DXMatrixIdentity DXmatTemp
D3DXMatrixRotationX DXmatTemp, RotateX * (Pi / 180)
D3DXMatrixRotationZ DXmatTemp, RotateY * (Pi / 180)
For I = 0 To UBound(TriStrip)
DXtempVec.X = TriStrip(I).X * XScale - CurX
DXtempVec.Y = TriStrip(I).Y * YScale - CurY
D3DXVec3Transform DXtempVec4, DXtempVec, DXmatTemp
TriStrip(I).X = DXtempVec4.X + NewX
TriStrip(I).Y = DXtempVec4.Y + NewY
Next
End Sub
Public Sub DXTransform2DVectorTriStrip(ByRef TriStrip() As D3DVECTOR, RotateX As Single, RotateY As Single, XScale As Single, YScale As Single, CurX As Single, CurY As Single, NewX As Single, NewY As Single)
Dim I As Long
D3DXMatrixIdentity DXmatTemp
D3DXMatrixRotationX DXmatTemp, RotateX * (Pi / 180)
D3DXMatrixRotationZ DXmatTemp, RotateY * (Pi / 180)
For I = 0 To UBound(TriStrip)
DXtempVec.X = TriStrip(I).X * XScale - CurX
DXtempVec.Y = TriStrip(I).Y * YScale - CurY
D3DXVec3Transform DXtempVec4, DXtempVec, DXmatTemp
TriStrip(I).X = DXtempVec4.X + NewX
TriStrip(I).Y = DXtempVec4.Y + NewY
'Debug.Assert DXtempVec4.X = DXtempVec.X And DXtempVec4.Y = DXtempVec.Y
Next
End Sub
Public Function DXRotateX(X As Single, Y As Single, Angle As Single) As Single
Dim Tx As Single, TAngle As Single
TAngle = Angle * (3.141592654 / 180)
Tx = X * Cos(TAngle) - Y * Sin(TAngle)
DXRotateX = Tx
End Function
Public Function DXRotateY(X As Single, Y As Single, Angle As Single) As Single
Dim Ty As Single, TAngle As Single
TAngle = Angle * (3.141592654 / 180)
Ty = Y * Cos(TAngle) + X * Sin(TAngle)
DXRotateY = Ty
End Function
Function Dist(X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant) As Single
Dist = Sqr(Abs(X1 - X2) * Abs(X1 - X2) + Abs(Y1 - Y2) * Abs(Y1 - Y2))
End Function
Function Dist3(X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant, Z1 As Variant, Z2 As Variant) As Single
Dist3 = Sqr(Abs(X1 - X2) * Abs(X1 - X2) + Abs(Y1 - Y2) * Abs(Y1 - Y2) + Abs(Z1 - Z2) * Abs(Z1 - Z2))
End Function
Public Function VectorTraceX(X1 As Single, X2 As Single, Dist As Single)
If Dist = 0 Then Dist = 1
VectorTraceX = ((X2 - X1) / Dist)
End Function
Public Function VectorTraceY(Y1 As Single, Y2 As Single, Dist As Single)
If Dist = 0 Then Dist = 1
VectorTraceY = ((Y2 - Y1) / Dist)
End Function
Public Function VectorTraceZ(Z1 As Single, Z2 As Single, Dist As Single)
If Dist = 0 Then Dist = 1
VectorTraceZ = ((Z2 - Z1) / Dist)
End Function
Public Function GetAngle(sX As Single, sY As Single, DX As Single, DY As Single) As Single
Dim slope!
Dim Angle!
    If sX = DX Then
        If sY < DY Then
            GetAngle = 90
        Else
            GetAngle = 270
        End If
        
        GetAngle = GetAngle + 90
        If GetAngle > 360 Then GetAngle = GetAngle - 360
        
        Exit Function
    Else
        slope = Abs(DY - sY) / Abs(DX - sX)
    End If
    
        Angle = Atn(slope)
        If DY < sY Then
            If DX < sX Then
                GetAngle = (Angle * 57.29578!) + 180
            Else
                GetAngle = -(Angle * 57.29578!) + 360
            End If
        Else
            If DX < sX Then
                GetAngle = -(Angle * 57.29578!) + 180
            Else
                GetAngle = Angle * 57.29578!
            End If
        End If
        
        GetAngle = GetAngle + 90
        If GetAngle > 360 Then GetAngle = GetAngle - 360
        
End Function
Public Function DXTiming(ByVal Action As Boolean)
On Error Resume Next 'When the timer has not been initialized, but the timer
                     'is stopped, a divide by zero error occurs
Select Case Action
    Case True 'Start
        QueryPerformanceFrequency Freq
        QueryPerformanceCounter StartTime
    Case False 'Stop
        QueryPerformanceCounter EndTime
        TimeElapse = (EndTime - StartTime) / (Freq / 1000)
End Select
End Function
Public Function MinMax(v, L, U)
If v < L Then v = L
If v > U Then v = U
MinMax = v
End Function
Public Sub DXEnablePointSprite()
DXfvfL
With D3DDevice
Dim DWFloat0 As Long
Dim DWFloat1 As Long
Dim DWFloatp08 As Long
DWFloat0 = DXftTOdw(0)
DWFloat1 = DXftTOdw(1)
DWFloatp08 = DXftTOdw(0.08)
.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
.SetRenderState D3DRS_POINTSCALE_ENABLE, 1
.SetRenderState D3DRS_POINTSIZE, DWFloatp08
.SetRenderState D3DRS_POINTSIZE_MIN, DWFloat0
.SetRenderState D3DRS_POINTSCALE_A, DWFloat0
.SetRenderState D3DRS_POINTSCALE_B, DWFloat0
.SetRenderState D3DRS_POINTSCALE_C, DWFloat1
End With
End Sub
Public Sub DXDisablePointSprite()
With D3DDevice
.SetRenderState D3DRS_POINTSPRITE_ENABLE, 0
.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
End With
End Sub
Public Function DXHitBoundingBox(BoxMin As D3DVECTOR, BoxMax As D3DVECTOR, Pos As D3DVECTOR)
If Pos.X > BoxMin.X And Pos.Y > BoxMin.Y And Pos.z > BoxMin.z And Pos.X < BoxMax.X And Pos.Y < BoxMax.Y And Pos.z < BoxMax.z Then
DXHitBoundingBox = True
Else
DXHitBoundingBox = False
End If
End Function
Public Function DXHitMesh(cMeshObj As DXMeshObject, CPos As D3DVECTOR, cDir As D3DVECTOR)
Dim bIsHit As Long
Dim cDist As Single
Dim tempVecPos As D3DVECTOR4
Dim tempVecDir As D3DVECTOR4
With cMeshObj
D3DX.Intersect .cMesh, CPos, cDir, bIsHit, 0, 0, 0, cDist, 0
End With
If bIsHit = 1 Then
DXHitMesh = True
End If
End Function
Public Function DXVectorToRGBA(Vec As D3DVECTOR, fHeight As Single) As Long
    Dim R As Integer, G As Integer, B As Integer, A As Integer
    R = 127 * Vec.X + 128
    G = 127 * Vec.Y + 128
    B = 127 * Vec.z + 128
    A = 255 * fHeight
    
    DXVectorToRGBA = D3DColorRGBA(R, G, B, A)
End Function
Public Function DXInterpolate2DMeshLinear(FromMesh As DX2DMeshObject, ToMesh As DX2DMeshObject, A As Single, ByRef OutMesh As DX2DMeshObject)
Dim V1 As D3DVECTOR2, V2 As D3DVECTOR2, V3 As D3DVECTOR2
Dim iI As Long
Dim iIF As Long
If A > 1 Then A = 1
If A < 0 Then A = 0
For iI = 1 To UBound(ToMesh.Frames)
For iIF = 0 To UBound(ToMesh.Frames(iI).TriVerts)
With FromMesh.Frames(iI).TriVerts(iIF)
V1.X = .X
V1.Y = .Y
End With
With ToMesh.Frames(iI).TriVerts(iIF)
V2.X = .X
V2.Y = .Y
End With
If A <> 0 And A <> 1 Then D3DXVec2Lerp V3, V1, V2, A
If A = 0 Then V3.X = V1.X: V3.Y = V1.Y
If A = 1 Then V3.X = V2.X: V3.Y = V2.Y
OutMesh.Frames(iI).TriVerts(iIF).X = V3.X
OutMesh.Frames(iI).TriVerts(iIF).Y = V3.Y
Next
Next
End Function
