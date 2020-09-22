Attribute VB_Name = "MExplode"
Option Explicit

' Get Mouse PointAPI
Type PointAPI
   x As Long
   y As Long
End Type
Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public MousePoint As PointAPI
'----------------------------------------------------------------
Global Const ExplodePointCount        As Byte = 5
' Smoke Properties for ExplodePoint
Type ExplodePointProperties
    Active      As Boolean    ' Active ExplodePoint
    x           As Single     ' Position x, y
    y           As Single     '
    Angle       As Single     ' Direction ExplodePoint
    Speed       As Single     ' Speed ExplodePoint
    Turn        As Single     ' Turn ExplodePoint
    Kill        As Byte
End Type
Public ExplodePoint(ExplodePointCount)  As ExplodePointProperties
'----------------------------------------------------------------
' this to add create more smoke
Global Const SmokeCount As Byte = 250

' Smoke Properties for Smoke
Type SmokeProperties
    Active      As Boolean    ' Active Smoke
    x           As Single     ' Position x, y
    y           As Single     '
    Angle       As Single     ' Direction Smoke
    Speed       As Single     ' Speed Smoke
    Turn        As Single     ' Turn Smoke
    '--------------------------------------------------------------
    TimeKill    As Byte
    Length      As Byte
    '--------------------------------------------------------------
    ' Size Smoke Front and Back
    '--------------------------------------------------------------
    'SizeFront   As Byte
    'SizeBack    As Byte
End Type
Public Smoke(SmokeCount)      As SmokeProperties

'----------------------------------------------------------------
' Direct 3D Frame and Object
'----------------------------------------------------------------
' Frames
Public SmokeFrame(SmokeCount) As Direct3DRMFrame3

' Meshes (loaded 3D objects from a *.x file)
'
'                 +-> Use () because we don't how many
'                 |   Direct 3D Object want loading
'                 |
Public SmokeObject()         As Direct3DRMMeshBuilder3

' Texture
Private MatOverride          As D3DRMMATERIALOVERRIDE
Private TextureSmoke         As Direct3DRMTexture3

Public ExplodePointFollowMouse        As Boolean
Public KeyPress                     As Boolean
Public TimeRndExplode               As Byte

Sub LoadObject_Fire()
    Dim i As Byte
    
    ReDim SmokeObject(2)
    For i = 0 To 1
        Set SmokeObject(i) = D3D.CreateMeshBuilder()
    Next i
    
    ' Load Bitmap for smoke texture (256 color, because i use Depth Color 16 bit)
    ' If you want smoke texture more (16 bit color, set Depth Color more 16 bit)
    Set TextureSmoke = D3D.LoadTexture("Fire.bmp")
    With TextureSmoke
        .SetDecalTransparency D_TRUE
        .SetDecalTransparentColor RGB(0, 0, 0)
    End With
    ' Load Direct 3D Object
    With SmokeObject(0)
        .LoadFromFile App.Path & "\Fire.x", 0, 0, Nothing, Nothing
        .SetTexture TextureSmoke
        .ScaleMesh 1, 1, 1
    End With
    '--------------------------------------------------------------
    Set TextureSmoke = D3D.LoadTexture("FireSmoke.bmp")
    With TextureSmoke
        .SetDecalTransparency D_TRUE
        .SetDecalTransparentColor RGB(0, 0, 0)
    End With
    With SmokeObject(1)
        .LoadFromFile App.Path & "\Fire.x", 0, 0, Nothing, Nothing
        .SetTexture TextureSmoke
        .ScaleMesh 1, 1, 1
    End With
    '--------------------------------------------------------------
    ' Set Frame for Direct 3D Object
    For i = 0 To SmokeCount
        Set SmokeFrame(i) = D3D.CreateFrame(FrameRoot)
    Next i
End Sub

Sub ExplodePointCreate(x As Single, y As Single, Angle As Single, Speed As Single, Turn As Single)
    Dim i As Byte
    
    For i = 0 To ExplodePointCount
        If ExplodePoint(i).Active = False Then
            ExplodePoint(i).Active = True
            ExplodePoint(i).x = x
            ExplodePoint(i).y = y
            ExplodePoint(i).Angle = Angle
            ExplodePoint(i).Speed = Speed
            ExplodePoint(i).Turn = Turn
            ExplodePoint(i).Kill = 0
            Exit Sub
        End If
    Next i
End Sub

Sub ExplodePointMove()
    Dim i            As Byte
    Dim RndSpeed  As Integer
    Dim RndAngle  As Integer
    Dim xFollow   As Integer
    Dim yFollow   As Integer
    
    For i = 0 To ExplodePointCount
        If ExplodePoint(i).Active = True Then
        
            GetCursorPos MousePoint
            
            If ExplodePointFollowMouse = False Then
                ' Make smoke around, around
                xFollow = 800 / 2
                yFollow = 600 / 2
            Else
                ' Make smoke chase mouse
                xFollow = MousePoint.x
                yFollow = MousePoint.y
            End If
            
         '[-------------------------------------------------]
         '[ ENGINEAAK : Calculation moving ExplodePoint       ]
         '[-------------------------------------------------]
            'Engine ExplodePoint(i).Angle, ExplodePoint(i).x, ExplodePoint(i).y, xFollow, yFollow, ExplodePoint(i).Speed, ExplodePoint(i).Turn
         '[-------------------------------------------------]
         '[ ENGINEAAK : Don't forget replace with new value ]
         '[-------------------------------------------------]
            'ExplodePoint(i).x = EngineResult.x
            'ExplodePoint(i).y = EngineResult.y
            'ExplodePoint(i).Angle = EngineResult.Angle
         '[-------------------------------------------------]
            RndSpeed = Int(Rnd * 2 + 1)
            RndAngle = Int(Rnd * 360 + 1) '+ 90
            
            CreateSmoke ExplodePoint(i).x, -ExplodePoint(i).y, (RndAngle / 1), (RndSpeed / 1), (RndSpeed / 1), 20
            
            ExplodePoint(i).Kill = ExplodePoint(i).Kill + 1
            If ExplodePoint(i).Kill > 10 Then
                ExplodePoint(i).Active = False
            End If
        End If
    Next i
End Sub

Sub CreateSmoke(x As Single, y As Single, Angle As Single, Speed As Single, Turn As Single, Length As Byte)
    Dim i As Byte
    
    For i = 0 To SmokeCount
        If Smoke(i).Active = False Then
            Smoke(i).Active = True
            Smoke(i).x = x
            Smoke(i).y = y
            Smoke(i).Angle = Angle
            Smoke(i).Speed = Speed
            Smoke(i).Turn = Turn
            Smoke(i).TimeKill = 0
            Smoke(i).Length = Length
            SmokeFrame(i).SetPosition Nothing, Smoke(i).x, Smoke(i).y, 0
            SmokeFrame(i).AddScale D3DRMCOMBINE_AFTER, 0, 0, 0
            SmokeFrame(i).AddVisual SmokeObject(0)
             ' After Create get out sub
            Exit Sub
        End If
    Next i
End Sub

Sub SmokeMove()
    Dim i               As Byte
    Dim MatOverrideCacl As Single
    Dim SmokeKill       As Byte
    Dim ZoomSmokeSet    As Single
    Dim Speed           As Single
    
    SmokeKill = 50
    For i = 0 To SmokeCount
        If Smoke(i).Active = True Then
        
            If Smoke(i).TimeKill > Int(Smoke(i).Length / 2) Then
                Speed = Smoke(i).Speed - (Smoke(i).TimeKill / 5)
            Else
                Speed = Smoke(i).Speed '- (Smoke(i).TimeKill / 5)
            End If
         '[-------------------------------------------------]
         '[ ENGINEAAK : Calculation moving smoke            ]
         '[-------------------------------------------------]
            Engine Smoke(i).Angle, Smoke(i).x, Smoke(i).y, 0, 0, Speed, Smoke(i).Turn
         '[-------------------------------------------------]
         '[ ENGINEAAK : Don't forget replace with new value ]
         '[-------------------------------------------------]
            Smoke(i).x = EngineResult.x
            Smoke(i).y = EngineResult.y
            Smoke(i).Angle = EngineResult.Angle
            Smoke(i).Angle = Smoke(i).Angle + 1
         '[-------------------------------------------------]
                        
            Smoke(i).TimeKill = Smoke(i).TimeKill + 1
            ' Kill Object Fire change With Dark Fire
            If Smoke(i).TimeKill = Int(Smoke(i).Length / 2) Then
                SmokeFrame(i).DeleteVisual SmokeObject(0)
                SmokeFrame(i).AddVisual SmokeObject(1)
            End If
            
            If Smoke(i).TimeKill > Smoke(i).Length Then
                Smoke(i).Active = False
                SmokeFrame(i).DeleteVisual SmokeObject(1)
            End If
            
            If Smoke(i).TimeKill < Int(Smoke(i).Length / 3) Then
                ' Calc for density smoke
                MatOverrideCacl = 0.75 / Smoke(i).Length
                ' Fire
                With MatOverride
                    .lFlags = D3DRMMATERIALOVERRIDE_DIFFUSE_ALPHAONLY
                    .dcDiffuse.a = 0.5 - (Smoke(i).TimeKill * MatOverrideCacl) * 2
                End With
                ZoomSmokeSet = (Smoke(i).TimeKill / 10)    ' Zoom Tail smoke
            Else
                ' Calc for density smoke
                MatOverrideCacl = 0.5 / Smoke(i).Length
                ' Smoke
                With MatOverride
                    .lFlags = D3DRMMATERIALOVERRIDE_DIFFUSE_ALPHAONLY
                    .dcDiffuse.a = 0.5 - (Smoke(i).TimeKill * MatOverrideCacl)
                End With
                ZoomSmokeSet = 1 + (Smoke(i).TimeKill / 20)      ' Zoom Tail smoke
            End If
        
            SmokeFrame(i).SetMaterialOverride MatOverride
            
            ' Set Rotation, Zoom (Scale) and Position Direct 3D
            SmokeFrame(i).AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, (Pi / 2)
            SmokeFrame(i).AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, DegreeToRadian(-Smoke(i).Angle)
            SmokeFrame(i).AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, -DegreeToRadian(90)   ' 90=Position ship up
            SmokeFrame(i).AddScale D3DRMCOMBINE_AFTER, ZoomSmokeSet, ZoomSmokeSet, ZoomSmokeSet
            SmokeFrame(i).SetPosition Nothing, Smoke(i).x, Smoke(i).y, 0 'Smoke(i).TimeKill
        End If
    Next i
End Sub



