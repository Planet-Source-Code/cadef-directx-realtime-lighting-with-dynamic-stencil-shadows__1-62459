VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AgrDX Tutorial 7 by Cade - Realtime Lighting & Stencil Volume Shadows - http://cade.sytes.net"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Stopped As Boolean
Dim CFps As Long
Dim LFps As Long
Dim Lt As Long
Dim LFpsTick As Long
Dim ObjTex As Long
Dim TriVerts(3) As TLVERTEX
Dim I As Long
Dim I2 As Long
Dim LR As Long
Dim tW!
Dim tH!
Dim vX!
Dim vY!
Dim LmX&
Dim LmY&
Dim LightAmount!
Dim GDst!
Dim ObjectsX() As Long
Dim ObjectsY() As Long
Dim ShadowVX(3) As Single
Dim ShadowVY(3) As Single
Dim ShadowProjectDistance As Long
Private Const NumObj As Long = 5 - 1
Dim TriLight() As TLVERTEX
Dim TriLightTemp() As TLVERTEX
Dim ShadowVolume() As TLVERTEX

Private Type LightObj
X As Single
Y As Single
Radius As Single
End Type

Dim SceneLights() As LightObj

Dim LightMapBuffer As Direct3DTexture8
Dim LightMapBufferSurface As Direct3DSurface8
Dim LightBuffer As Direct3DTexture8
Dim LightBufferSurface As Direct3DSurface8
Dim ShadowBuffer As Direct3DTexture8
Dim ShadowBufferSurface As Direct3DSurface8
Dim BackBuffer As Direct3DSurface8
Dim BackBufferCache As Direct3DTexture8
Dim BackBufferCacheSurface As Direct3DSurface8
Private Sub RenderTriLight(X!, Y!, Scalar!)
D3DDevice.SetTexture 0, Nothing
ReDim TriLightTemp(UBound(TriLight))
For I = 0 To UBound(TriLightTemp)
TriLightTemp(I) = DXCreateTLVertex((TriLight(I).X * (Scalar / 128)) + X, (TriLight(I).Y * (Scalar / 128)) + Y, 0, 0, TriLight(I).Color, 0, 1, 1)
Next
DXSetAlphaState
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLEFAN, UBound(TriLightTemp) - 1, TriLightTemp(0), Len(TriLightTemp(0))
End Sub
Private Sub MakeTriLight(Secs As Long)
ReDim TriLight(Secs + 1)
With TriLight(0)
.X = 0
.Y = 0
.Color = D3DColorRGBA(255, 255, 255, 255)
End With
For I = 1 To Secs + 1
With TriLight(I)
.X = DXRotateX(0, -128, ((360 / Secs) * I))
.Y = DXRotateY(0, -128, ((360 / Secs) * I))
.Color = D3DColorRGBA(255, 255, 255, 0)
End With
Next
End Sub
Public Function LightNormal(X1!, Y1!, X2!, Y2!, X3!, Y3!, X4!, Y4!) As Single
Ang1 = (GetAngle(X1, Y1, X2, Y2) - 180)
Ang2 = (GetAngle(X3, Y3, X4, Y4) - 180)
LightNormal = Abs(Ang1 - Ang2)
Debug.Print Round(LightNormal, 5)
End Function
Private Sub RenderObjectShadow(X As Long, Y As Long, LX As Single, LY As Single, BoxIndex As Long, ProjectDistance As Long)
'This doesnt render anything to the backbuffer, only to the stencil buffer
ShadowProjectDistance = ProjectDistance + 64
Dim CS As Single 'Current Shadow
For I = 0 To UBound(TriVerts)
With TriVerts(I)
Select Case I
Case 0
.X = X - tW '+ 1
.Y = Y - tH '+ 1
Case 1
.X = X + tW '- 1
.Y = Y - tH '+ 1
Case 3
.X = X - tW '+ 1
.Y = Y + tH '- 1
Case 2
.X = X + tW '- 1
.Y = Y + tH '- 1
End Select
GDst = Dist(LX, LY, .X, .Y)
If GDst = 0 Then GDst = 1
'Project the points away
ShadowVX(I) = (((.X - LX) / GDst) * ShadowProjectDistance) + .X
ShadowVY(I) = (((.Y - LY) / GDst) * ShadowProjectDistance) + .Y
.Color = D3DColorRGBA(0, 0, 0, 0)
End With
Next

'Create the shadow volume
ReDim ShadowVolume(0)
For I = 0 To UBound(TriVerts) + 1
CS = I
If I > UBound(TriVerts) Then CS = 0
With TriVerts(CS)
'If LightNormal(.X, .Y, LX, LY, X / 1, Y / 1, LX, LY) > 0.25 Then
ReDim Preserve ShadowVolume(UBound(ShadowVolume) + 2)
ShadowVolume(UBound(ShadowVolume)) = DXCreateTLVertex(.X, .Y, 0, 1, D3DColorRGBA(0, 0, 0, 0), 1, 0, 0)
ShadowVolume(UBound(ShadowVolume) - 1) = DXCreateTLVertex(ShadowVX(CS), ShadowVY(CS), 0, 1, D3DColorRGBA(0, 0, 0, 0), 1, 0, 0)
'End If
End With
Next

'Subtract the shadow volume from the stencil buffer
D3DDevice.SetRenderState D3DRS_STENCILPASS, D3DSTENCILOP_DECRSAT
If UBound(ShadowVolume) > 0 Then D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, UBound(ShadowVolume) - 2, ShadowVolume(1), Len(ShadowVolume(0))

'And add back the object
D3DDevice.SetRenderState D3DRS_STENCILPASS, D3DSTENCILOP_INCRSAT
D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 2, TriVerts(0), Len(TriVerts(0))
End Sub
Private Sub RenderObject(X As Long, Y As Long, BoxIndex As Long)
For I = 0 To UBound(TriVerts)
With TriVerts(I)
Select Case I
Case 0
.X = X - tW
.Y = Y - tH
Case 1
.X = X + tW
.Y = Y - tH
Case 2
.X = X - tW
.Y = Y + tH
Case 3
.X = X + tW
.Y = Y + tH
End Select
.Color = D3DColorRGBA(255, 255, 255, 255)
End With
Next

D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriVerts(0), Len(TriVerts(0))
End Sub
Private Sub Form_Load()
'Caption = ""
Visible = True
DXLoad hWnd, 0, 0
DXEnableAlpha
DXSetAlphaState
D3DDevice.SetRenderState D3DRS_ZENABLE, 0

ReDim SceneLights(1 To 3)
With SceneLights(1)
.X = 0
.Y = 0
.Radius = 512
End With
With SceneLights(2)
.X = 640
.Y = 480 / 2
.Radius = 512
End With
With SceneLights(3)
.X = 0
.Y = 0
.Radius = 256
End With

DXAddTexture "Object.bmp", D3DColorRGBA(0, 0, 0, 0)

ObjTex = DXGetTextureIndex("Object.bmp")

D3DDevice.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_ALWAYS

'Make a light out of a triangle fan made of 64 triangles
MakeTriLight 64

Randomize

'Create the buffers
'LightBuffer contains the fully computated lightmap
Set LightBuffer = D3DDevice.CreateTexture(640, 480, 1, 0, DispMode.Format, D3DPOOL_DEFAULT)
Set BackBufferCache = D3DDevice.CreateTexture(640, 480, 1, 0, DispMode.Format, D3DPOOL_DEFAULT)

'Use surfaces to render to the textures
Set LightBufferSurface = LightBuffer.GetSurfaceLevel(0)
Set BackBufferCacheSurface = BackBufferCache.GetSurfaceLevel(0)

'Pointer to the backbuffer
Set BackBuffer = D3DDevice.GetBackBuffer(ByVal 0, D3DBACKBUFFER_TYPE_MONO)

ReDim ObjectsX(NumObj)
ReDim ObjectsY(NumObj)
For I = 0 To NumObj
ObjectsX(I) = Rnd * ScaleWidth
ObjectsY(I) = Rnd * ScaleHeight
Next

For I = 0 To 3
TriVerts(I) = DXCreateTLVertex(0, 0, 0, 1, 0, 1, 0, 0)
With TriVerts(I)
Select Case I
Case 0
.tu = 0
.tv = 0
Case 1
.tu = 1
.tv = 0
Case 2
.tu = 0
.tv = 1
Case 3
.tu = 1
.tv = 1
End Select
End With
Next

D3DDevice.SetVertexShader FVF_TLVERTEX
'Scince we will only use 2 textures at one time (stage 0 and 1),
'disable texture stages 2 and above when rendering
D3DDevice.SetTextureStageState 2, D3DTSS_COLOROP, D3DTOP_DISABLE

Do Until Stopped
If LFps = 0 Then LFps = 1
CFps = CFps + 1
TimeElapse = TimeElapse / 10
DXTiming True

D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 255), 0, 0
D3DDevice.BeginScene

'Tex width, height
tW = DXTexturePool(ObjTex).cTextureDesc(0).Width / 2
tH = DXTexturePool(ObjTex).cTextureDesc(0).Height / 2

D3DDevice.SetTexture 0, Nothing

D3DDevice.SetVertexShader FVF_TLVERTEX

For I2 = 1 To UBound(SceneLights)
'For each light
'First set the stencil buffer to 2
'Now it has 2 values below it, 1 and 0
With D3DDevice
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_STENCIL, 0, 0, 2
.SetRenderState D3DRS_STENCILENABLE, 1
.SetRenderState D3DRS_STENCILFUNC, D3DCMP_ALWAYS
.SetRenderState D3DRS_STENCILMASK, 0
End With
For LR = 0 To NumObj
RenderObjectShadow ObjectsX(LR), ObjectsY(LR), SceneLights(I2).X, SceneLights(I2).Y, LR, SceneLights(I2).Radius / 1
Next
With D3DDevice
.SetRenderState D3DRS_STENCILZFAIL, D3DSTENCILOP_KEEP
.SetRenderState D3DRS_STENCILFAIL, D3DSTENCILOP_KEEP
.SetRenderState D3DRS_STENCILPASS, D3DSTENCILOP_KEEP
.SetRenderState D3DRS_STENCILFUNC, D3DCMP_NOTEQUAL
.SetRenderState D3DRS_STENCILMASK, 2
'Light only the areas where the stencil has not changed
RenderTriLight SceneLights(I2).X, SceneLights(I2).Y, SceneLights(I2).Radius
.SetRenderState D3DRS_STENCILENABLE, 0
End With

Next
D3DDevice.CopyRects BackBuffer, ByVal 0, 0, LightBufferSurface, 0

D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(255, 255, 255, 255), 0, 0

'The backbuffer is blank at this point, the lighting is stored in the lightmap buffer

DXSetTexture 0, "Object.bmp"
For I2 = 0 To NumObj
RenderObject ObjectsX(I2), ObjectsY(I2), I2
Next

D3DDevice.CopyRects BackBuffer, ByVal 0, 0, BackBufferCacheSurface, 0

'Multiply the backbuffer which contains the rendered scene with the lightmap
D3DDevice.SetTexture 0, LightBuffer
D3DDevice.SetTexture 1, BackBufferCache
D3DDevice.SetVertexShader FVF2
DXSetMultiTextureState
'Comment out the line under this to render the unlit scene
DXRender2DBoxTexturedWith2Tex 0, 0, 640, 480, False, D3DColorRGBA(255, 255, 255, 255)
D3DDevice.SetVertexShader FVF_TLVERTEX
D3DDevice.SetTexture 1, Nothing

With SceneLights(3)
.X = LmX
.Y = LmY
End With

DXRenderText "FPS: " & LFps & vbCrLf, D3DColorRGBA(128, 192, 255, 255), 0, 0, 512, 64

D3DDevice.EndScene

'Visible = True
'DoEvents
D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
'Exit Sub

DoEvents


If Abs(GetTickCount - LFpsTick) > 1000 Then
LFps = CFps
CFps = 0
LFpsTick = GetTickCount
IsFirstFrame = False
End If
DXTiming False
DoEvents
Loop

Set BackBuffer = Nothing
Set LightBufferSurface = Nothing
Set BackBufferCacheSurface = Nothing
Set LightBuffer = Nothing
Set BackBufferCache = Nothing
DXUnLoad
End
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LmX = X
LmY = Y
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Stopped = True
End Sub
