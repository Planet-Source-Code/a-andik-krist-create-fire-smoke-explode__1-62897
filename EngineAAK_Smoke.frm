VERSION 5.00
Begin VB.Form EngineAAK_Smoke 
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
Attribute VB_Name = "EngineAAK_Smoke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' [923342960150006  S T T D T T S  600051069243329]
' [===============================================]
' [        Introduction: EngineAAK Ver 1.1        ]
' [        -------------------------------        ]
' [       Title: Create Smoke use Direct 3D       ]
' [===============================================]
' [      How make smoke use Direct 3D Object      ]
' [  3 smoke will create (Red, Green, Blue), and  ]
' [  3 smoke will move a round, press M for make  ]
' [               smoke chase mouse               ]
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
    D3DInit EngineAAK_Smoke, ScrWidth, ScrHeight, 16
    '--------------------------------------------------------------
    ' Initialize Frame Direct3D like Root, Camera, Light
    FrameD3DInit ScrWidth, ScrHeight
    '--------------------------------------------------------------
    ' Initialize Object and Frame for Smoke
    LoadObject_Smoke
    '--------------------------------------------------------------
    FrameRoot.SetSceneBackground RGB(50, 50, 50)
    
    ' Create 3 Point for create smoke
    SmokePointCreate ScrWidth / 2, ScrHeight / 2, 0, 5, 2.5             ' Red Smoke
    SmokePointCreate ScrWidth / 2, (ScrHeight / 2) - 1, 120, 5, 2.5, 1  ' Blue Smoke  add -1 for make smoke same around
    SmokePointCreate ScrWidth / 2, (ScrHeight / 2) - 1, 240, 5, 2.5, 2  ' Green Smoke add -1 for make smoke same around
    
    Do                                      ' Loop main until ProgramFinish=True
        On Local Error Resume Next
        DoEvents
        D3D_ViewPort.Clear D3DRMCLEAR_TARGET Or D3DRMCLEAR_ZBUFFER 'ClS Viewport.
        D3D_Device.Update                   ' Update the Direct3D Device.
        
        ' Just Text
        '--------------------------------------------------------------
        TxtIntro = "Introduction EngineAAK Ver 1.1 (from my first program HomingMissile)"
        TxtMid = (ScrWidth / 2) - (Len(TxtIntro) * 3.5)
        BackBuffer.DrawText TxtMid, 300 - 20 - 15, TxtIntro, False
        '--------------------------------------------------------------
        TxtIntro = "EngineAAK use for moving and create smoke trail"
        TxtMid = (ScrWidth / 2) - (Len(TxtIntro) * 3.5)
        BackBuffer.DrawText TxtMid, 300 - 20, TxtIntro, False
        '--------------------------------------------------------------
        TxtIntro = "EngineAAK for : ARCADE SHOTTER / RTS / RPG / RACE (..test is working..)"
        TxtMid = (ScrWidth / 2) - (Len(TxtIntro) * 3.5)
        BackBuffer.DrawText TxtMid, 300 - 20 + 15, TxtIntro, False
        '--------------------------------------------------------------

        SmokePointMove
        
        SmokeMove
        
        DelayGame 21                        ' Set 41 FPS
        
        GetCursorPos MousePoint
        
        D3D_ViewPort.Render FrameRoot       ' Render the 3D Objects
        
        '--------------------------------------------------------------
        TxtIntro = "Click form to EXIT, Press M for smoke follow and not mouse"
        TxtMid = (ScrWidth / 2) - (Len(TxtIntro) * 3.5)
        BackBuffer.DrawText TxtMid, ScrHeight - 20, TxtIntro, False
        
        ' KeyBoard Press M for smoke chase mouse
        DI_Device.GetDeviceStateKeyboard DI_State
        If DI_State.Key(DIK_M) <> 0 Then
            If KeyPress = False Then
                If SmokePointFollowMouse = False Then
                    SmokePointFollowMouse = True
                Else
                    SmokePointFollowMouse = False
                End If
                KeyPress = True
            End If
        Else
            KeyPress = False
        End If
        
        Primary.Flip Nothing, DDFLIP_WAIT   ' Flip the BackBuffer with the FrontBuffer.
    Loop Until ProgramFinish = True
    End
    
End Sub

