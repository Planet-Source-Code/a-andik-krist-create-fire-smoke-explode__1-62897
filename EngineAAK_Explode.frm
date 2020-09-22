VERSION 5.00
Begin VB.Form EngineAAK_Explode 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "EngineAAK_Explode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' [923342960150006  S T T D T T S  600051069243329]
' [===============================================]
' [        Introduction: EngineAAK Ver 1.1        ]
' [        -------------------------------        ]
' [      Title: Create Explode use Direct 3D      ]
' [===============================================]
' [     How make Explode use Direct 3D Object     ]
' [-----------------------------------------------]
' [ Not for sale or commercial without permission ]
' [-----------------------------------------------]
' [              By: A. Andik Krist.              ]
' [            -----------------------            ]
' [              JAKARTA - INDONESIA              ]
' [-----------------------------------------------]
' [                                               ]
' [        for Comments, Suggestions & Ideas      ]
' [          E-mails me: aakchat@yahoo.com        ]
' [               Date: 15-Okt-2005               ]
' [                                               ]
' [===============91923=29873=30006===============]
Option Explicit

Dim ProgramFinish As Boolean

Private Sub Form_Click()
    ProgramFinish = True
End Sub

Private Sub Form_Load()
    Dim i           As Byte
    Dim ScrWidth    As Long
    Dim ScrHeight   As Long
    Dim TxtIntro    As String
    Dim TxtMid      As Integer
        
    '--------------------------------------------------------------
    ScrWidth = 800   ' 640
    ScrHeight = 600  ' 480
    '--------------------------------------------------------------
    ' Init Direct3D
    D3DInit EngineAAK_Explode, ScrWidth, ScrHeight, 16
    '--------------------------------------------------------------
    ' Initialize Frame Direct3D like Root, Camera, Light
    FrameD3DInit ScrWidth, ScrHeight
    '--------------------------------------------------------------
    ' Initialize Object and Frame for Fire
    LoadObject_Fire
    '--------------------------------------------------------------
    FrameRoot.SetSceneBackground RGB(0, 0, 0)
    
    Do                                      ' Loop main until ProgramFinish=True
        On Local Error Resume Next
        DoEvents
        D3D_ViewPort.Clear D3DRMCLEAR_TARGET Or D3DRMCLEAR_ZBUFFER 'ClS Viewport.
        D3D_Device.Update                   ' Update the Direct3D Device.
        
        ' Just Text
        '--------------------------------------------------------------
        TxtIntro = "Introduction EngineAAK Ver 1.1 (from my first program HomingMissile)"
        TxtMid = (ScrWidth / 2) - (Len(TxtIntro) * 3.5)
        BackBuffer.DrawText TxtMid, 0, TxtIntro, False
        '--------------------------------------------------------------
        TxtIntro = "EngineAAK use for part Create Explode"
        TxtMid = (ScrWidth / 2) - (Len(TxtIntro) * 3.5)
        BackBuffer.DrawText TxtMid, 15, TxtIntro, False
        '--------------------------------------------------------------
        TxtIntro = "EngineAAK for : ARCADE SHOTTER / RTS / RPG / RACE (..test is working..)"
        TxtMid = (ScrWidth / 2) - (Len(TxtIntro) * 3.5)
        BackBuffer.DrawText TxtMid, 30, TxtIntro, False
        '--------------------------------------------------------------

        ExplodePointMove
        
        SmokeMove
        
        RandomExplode
        
        'DelayGame 21                        ' Set 41 FPS
        
        GetCursorPos MousePoint
        
        D3D_ViewPort.Render FrameRoot       ' Render the 3D Objects
        
        '--------------------------------------------------------------
        TxtIntro = "Click form to EXIT"
        TxtMid = (ScrWidth / 2) - (Len(TxtIntro) * 3.5)
        BackBuffer.DrawText TxtMid, ScrHeight - 20, TxtIntro, False
        
        Primary.Flip Nothing, DDFLIP_WAIT   ' Flip the BackBuffer with the FrontBuffer.
    Loop Until ProgramFinish = True
    End
    
End Sub

Sub RandomExplode()
    Dim RndX As Single
    Dim RndY As Single
    
    TimeRndExplode = TimeRndExplode + 1
    If TimeRndExplode = 5 Then
        TimeRndExplode = 0
        RndX = Int(Rnd * 800 + 1)
        RndY = Int(Rnd * 600 + 1)
        ExplodePointCreate RndX, RndY, 0, 5, 2.5
    End If
End Sub


