VERSION 5.00
Begin VB.Form Form1 
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This Direct3D example was written by Vadim, I tried
' to explain everything as easy as possible, but if you
' Have any further questions you can e-mail me (vadimcoder@mail.ru)

' Main DirectX components - These are needed for directX to function
Dim DXMain As New DirectX7          ' The DirectX core file
Dim DDMain As DirectDraw4           ' The directdraw layer
Dim RMMain As Direct3DRM3           ' The direct3D layer

' Directinput components - Used to log various input devices
Dim DIMain As DirectInput           ' The DirectInput core
Dim DIDevice As DirectInputDevice   ' The directinputdevice
Dim DIState As DIKEYBOARDSTATE      ' An array holding the state of all the keys

' Directdraw surfaces - This is where the screen is drawn
Dim DSPrim As DirectDrawSurface4    ' The frontbuffer, this is what you see on your screen
Dim DSBack As DirectDrawSurface4    ' The backbuffer, here everything is drawn
Dim SDPrim As DDSURFACEDESC2        ' The surfacedescription
Dim DDBack As DDSCAPS2              ' Surface info

' Direct3D - Main objects - The viewport and direct3D device
Dim RMDevice As Direct3DRMDevice3   ' The retained mode device
Dim RMView As Direct3DRMViewport2   ' The direct3D viewport (the screen)

' Direct3D - Frames & Meshes - Frames are "containers" for the 3D meshes
Dim FRRoot As Direct3DRMFrame3      ' The main frame (is drawn on the backbuffer)
Dim FRLight As Direct3DRMFrame3     ' Frame containing the spotlight
Dim FRCam As Direct3DRMFrame3       ' Frame containing the camera
Dim FRShip As Direct3DRMFrame3      ' Frame containing the ship
Dim MSShip As Direct3DRMMeshBuilder3 ' Mesh containing the ship

' Direct3D - Lights - There are two lights in this program
Dim LTMain As Direct3DRMLight       ' The ambient lighting
Dim LTSpot As Direct3DRMLight       ' The spotlight (used to add more realistic lighting on the ship

' Ship data - Various variables containing the ships position and rotation
Dim dXAng As Single                 ' X Angle
Dim dYAng As Single                 ' Y Angle
Dim iXPos As Integer                ' X Position
Dim iYPos As Integer                ' Y Position

Sub DX_Init()

  ' In this sub all the main directX components are initialized
  ' and configured

  Set DDMain = DXMain.DirectDraw4Create("")
  ' First we create the directdraw object
    
  DDMain.SetCooperativeLevel Form1.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
  DDMain.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT
  ' We tell directdraw we want to go fullscreen and enter
  ' the resolution and bitdepth
    
  SDPrim.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
  SDPrim.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_3DDEVICE Or DDSCAPS_COMPLEX Or DDSCAPS_FLIP
  SDPrim.lBackBufferCount = 1
  Set DSPrim = DDMain.CreateSurface(SDPrim)
  ' Now we create the screen and add a single backbuffer
           
  DDBack.lCaps = DDSCAPS_BACKBUFFER
  Set DSBack = DSPrim.GetAttachedSurface(DDBack)
  DSBack.SetForeColor RGB(255, 255, 255)
  ' The backbuffer is initailized and the fore(text)color is
  ' set to white
    
  Set RMMain = DXMain.Direct3DRMCreate()
  ' Next we create the direct3D device. In this progran we will
  ' use the retained mode of direct3D. Direct3D originally only
  ' consisted of the immediate mode in which you had to hard-code
  ' every single polygon of an object. Later on Microsoft made
  ' a front-end of the immediate mode: the retained mode, in which
  ' objects can be loaded straigt from the harddisk.

  Set RMDevice = RMMain.CreateDeviceFromSurface("IID_IDirect3DRGBDevice", DDMain, DSBack, D3DRMDEVICE_DEFAULT)
  RMDevice.SetBufferCount 2
  RMDevice.SetQuality D3DRMRENDER_GOURAUD
  RMDevice.SetTextureQuality D3DRMTEXTURE_NEAREST
  RMDevice.SetRenderMode D3DRMRENDERMODE_BLENDEDTRANSPARENCY
  ' This part needs some explanation. Fist we tell the direct3D
  ' device we'll be using the RGB device. You can choose between
  ' the RGB device which uses only software rendering or the HAL
  ' device with which directX checks out which functions your
  ' video card can handle, the rest is emulated. The HAL device
  ' is usually much faster, but it takes longer to load.
  ' After that there are some variables containing the picture
  ' quality. The last two lines should not be changed, but the
  ' SetQuality line can be D3DRMRENDER_GOURAUD for maximum quality
  ' D3DRMRENDER_FLAT for medium quality or D3DRMRENDER_WIREFRAME
  ' for minimal quality.
  
  Set DIMain = DXMain.DirectInputCreate()
  Set DIDevice = DIMain.CreateDevice("GUID_SysKeyboard")
  DIDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
  DIDevice.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
  DIDevice.Acquire
  ' Finally we tell directinput we'll be using the keyboard

End Sub

Sub DX_Scene()
    
  ' In this sub the scene objects are loaded and direct3D
  ' is set up

  Set FRRoot = RMMain.CreateFrame(Nothing)
  Set FRLight = RMMain.CreateFrame(FRRoot)
  Set FRShip = RMMain.CreateFrame(FRRoot)
  Set FRCam = RMMain.CreateFrame(FRRoot)
  ' We set up the hierarchy in which the root frame is the
  ' parent of all of the other frames. This is needed to
  ' render all of the objects
  
  FRRoot.SetSceneBackgroundRGB 0, 0, 0
  ' The background color is set to black. Please note that
  ' In directX, RGB-colors are not defined using the numbers
  ' 0 to 255 but by a value for 0 to 1
  
  FRCam.SetPosition Nothing, 0, 50, 0
  Set RMView = RMMain.CreateViewport(RMDevice, FRCam, 0, 0, 640, 480)
  RMView.SetBack 300
  ' We set the position of the camera and use it as an viewport
  ' also we set the draw-depth, which is usually 100, but in this
  ' program we need a bigger value to prevent the ship from disappearing
    
  FRLight.SetPosition Nothing, 0, 25, 0
  Set LTSpot = RMMain.CreateLightRGB(D3DRMLIGHT_POINT, 1, 1, 1)
  FRLight.AddLight LTSpot
  ' Next we create a spotlight and set up the position of it.
  ' Note that we link the light to the frame. That way, a single
  ' Light definition can be used in multiple frames
    
  Set LTMain = RMMain.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.5, 0.5, 0.5)
  FRRoot.AddLight LTMain
  ' Also an ambient light is created, to prevent the dark
  ' sides from the ship from becoming totally black.

  Set MSShip = RMMain.CreateMeshBuilder()
  MSShip.LoadFromFile App.Path & "\ship.x", 0, 0, Nothing, Nothing
  MSShip.ScaleMesh 0.5, 0.5, 0.5
  FRShip.AddVisual MSShip
  ' Now we can load the ship mesh. Which is a file on the
  ' Harddisk with the ship I created in 3D Studio MAX. you
  ' can replace the file by any object or ship.
  
  dXAng = 0
  dYAng = 0
  iXPos = 0
  iYPos = 1
  ' To make sure the camera starts at the right angle, the
  ' Y-position of the ship is'nt 0, you can set it to 0 and
  ' find out what I mean

End Sub

Sub DX_Render()
    
  ' In this sub, the objects are rendered, an infinite
  ' do...loop is used for speed reasons. You can use a timer,
  ' but that would be much slower, and besides, we haven't
  ' used any controls so far, so why start now...

  On Local Error Resume Next
  ' Debug handler, if everything is alright, this is obsolete
    
  Do
    DoEvents
    ' Tell the computer to give the system some time to do
    ' all the things it needs to do
    DX_Input
    ' Call the directinput sub (below)
    RMView.Clear D3DRMCLEAR_TARGET Or D3DRMCLEAR_ZBUFFER
    ' Clear the viewport
    RMDevice.Update
    ' Update the direct3D device
    DSBack.DrawText 200, 0, "Direct3D Retained Mode By Vadim", False
    DSBack.DrawText 140, 460, "Press [Esc] to exit or use the arrow keys to control the ship", False
    ' Draw the text messages on the screen
    RMView.Render FRRoot
    ' Render the scene
    DSPrim.Flip Nothing, DDFLIP_WAIT
    ' Swap the backbuffer and the frontbuffer
  Loop

End Sub

Sub DX_Input()

  ' In this sub, the keyboard handlings will be processed

  DIDevice.GetDeviceStateKeyboard DIState
  ' Get the keyboardstate and store it in the array

  If DIState.Key(DIK_ESCAPE) <> 0 Then Call DX_Exit
  ' Terminate the program if the escape key is pressed

  If DIState.Key(DIK_LEFT) <> 0 Then
    Let dXAng = dXAng + 0.2
    Let dYAng = dYAng - 0.2
    If dXAng > 1 Then Let dXAng = 1
  End If
  ' If the left key is pressed, then the angle of the ship
  ' is increased, I also tilt the ship a bit to create a more
  ' realistic turn, the if...then makes sure it doesn't tilt to much.
  
  If DIState.Key(DIK_RIGHT) <> 0 Then
    Let dXAng = dXAng - 0.2
    Let dYAng = dYAng + 0.2
    If dXAng < -1 Then Let dXAng = -1
  End If
  ' Same as the left key, only inverted

  If DIState.Key(DIK_RIGHT) = 0 And DIState.Key(DIK_LEFT) = 0 Then
    If dXAng < 0 Then Let dXAng = dXAng + 0.2
    If dXAng > 0 Then Let dXAng = dXAng - 0.2
  End If
  ' If neither left or right key is pressed, the ship should
  ' tilt back to a horizontal position (which is 0) so if the
  ' tilt is larger than 0 it should be reduced and vice versa
  
  If DIState.Key(DIK_UP) <> 0 Then
    Let iXPos = iXPos + (Sin(dYAng) * 2)
    Let iYPos = iYPos + (Cos(dYAng) * 2)
  End If
  ' If the up key is pressed, the ship should move. Here some
  ' mathematics show up, but these are merely to make sure
  ' the ship moves according to it's angle.

  FRShip.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, dYAng
  FRShip.AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, dXAng
  ' Here we actually rotate the ship according to the new angles
  
  FRShip.SetPosition Nothing, iXPos, 2.5, iYPos
  FRCam.LookAt FRShip, Nothing, D3DRMCONSTRAIN_Z
  ' Finally we move the ship and tell the camera it should
  ' look at the ship.
    
End Sub

Sub DX_Exit()

  ' This sub is called when the program is terminated.

  Call DDMain.RestoreDisplayMode
  Call DDMain.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)
  Call DIDevice.Unacquire
  ' All devices are restored and the resloution is set back
  ' to it's original state.
  
  End
  ' End the program

End Sub

Private Sub Form_Load()
    
  Me.Show
  ' Show the form, it sounds stupid. But if you don't show
  ' the form, the screen starts doing stange things on some
  ' computers
    
  DoEvents
  ' Let the computer do it's things
    
  DX_Init    ' Initialize DirectX
  DX_Scene   ' Initialize the scene
  DX_Render  ' Start the rendering loop

End Sub

' For any questions or hints, mail me at: vadimcoder@mail.ru
