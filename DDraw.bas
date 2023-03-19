Attribute VB_Name = "DDraw"
'*****************************************************************
'
'IMPORTANT TO NOTE:
'
'When using resource files and the CreateSurfaceFromResource
'command in DirectX 7.0, you will not be able to run your program
'by simply pressing F5 or selecting the run menu item, you will
'get an error. dX7 will only recognize your resource file if you
'compile your program and RUN THE COMPILED VERSION.
'
' - Lucky
'
'*****************************************************************


Public Type BITMAP '14 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Public Type BITMAPINFOHEADER '40 bytes
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
Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Global Const DIB_RGB_COLORS = 0&
Global Const BI_RGB = 0&

Global Const pixR As Integer = 3
Global Const pixG As Integer = 2
Global Const pixB As Integer = 1
'mudanca , pegar cores

Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

'Major DX Objects
Global DX As DirectX7
Global dd As DirectDraw7

Global Primary As DirectDrawSurface7       'Primary surface
Global BackBuffer As DirectDrawSurface7    'Backbuffer surface
Global ddsdPrimary As DDSURFACEDESC2       'Primary surface description
Global ddsdBackBuffer As DDSURFACEDESC2    'Backbuffer surface description

Global Running As Boolean               'Determines if the program should still be running
Global Facing As Integer                'Determines the direction the ship is currently facing
Global ZoomX As Long                   'How far are we zoomed?
Global ZoomY As Long
Global ColourFill As Integer            'Variable that determines the colour of the background

Global SpriteWidth As Long                  'Width of the sprite we're displaying
Global SpriteHeight As Long                   'Height of the sprite we're displaying

Type Sprites
x As Long
y As Long
Width As Long
Height As Long
imagem As DirectDrawSurface7
End Type

'           cars,angles,upDown,pneus
Global explosao(1 To 19) As Sprites
Global sombra As Sprites
Global cars() As Sprites    'Array of surfaces that will contain our ship
Global pistas() As Sprites


Global pista_horizontal_amarela As Sprites
Global pista_vertical_amarela As Sprites
Global Mapa As Sprites

Global nivel As Long
Global OrigemdoPulo As Long
Global LastRampaStatus As Boolean
Global LastRampa As Long
Global LastLadeiraStatus As Boolean
Global EspereSetas As Boolean
Global RetornandoAPista As Boolean
'0 = reta branca
'1 = reta amarela
'2 = pulando rampa
'3 = descendo rampa

'Global pontinho As Sprites
    
Public Sub Initialize()

    'This routine initializes the display mode, and the primary/backbuffer complex
    
    'Handles errors
    'On Local Error GoTo ErrOut
    
    'Creates the directdraw object


    Set DX = New DirectX7
    Set dd = DX.DirectDrawCreate("")

    'Set the cooperative level and displaymode...
    'Call dd.SetCooperativeLevel(FrmDirectX.hWnd, DDSCL_NORMAL)
    Call dd.SetCooperativeLevel(FrmDirectX.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE Or DDSCL_ALLOWREBOOT)
    Call dd.SetDisplayMode(640, 480, 32, 0, DDSDM_DEFAULT)
    
    'Create the primary complex surface with one backbuffer
    ddsdPrimary.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsdPrimary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    ddsdPrimary.lBackBufferCount = 1

    Set Primary = dd.CreateSurface(ddsdPrimary)
    
    'Get the backbuffer from the primary surface
    Dim caps As DDSCAPS2
    caps.lCaps = DDSCAPS_BACKBUFFER
    Set BackBuffer = Primary.GetAttachedSurface(caps)
    
    'Set the colour (for text output) of the backbuffer
    BackBuffer.SetForeColor vbWhite

    'Initialize the global variables
    Facing = 0
    Zoom = 1
    ColourFill = 0
    Running = True
    'Load the sprites
    
    LoadSurfaces
         
    'Clears the buffer
    ClearBuffer
    
    Exit Sub
    
ErrOut:
    Running = False             'If there's an error, exit the program
    
End Sub

Private Sub LoadSurfaces()
Dim s As OLE_COLOR
Dim i As Integer
Dim cor As Sprites
    Dim lockrect As RECT
Dim SrcDesc         As DDSURFACEDESC2
Dim u As Long
Dim x As Long
Dim y As Long
Preto = 0
    'This routine loads all of the surfaces we're going to be using
    SpriteWidth = 123
    SpriteHeight = 72
    'For i = 0 To 23
    'cars(0, i, 0, 0).height = SpriteHeight
    'cars(0, i, 0, 0).width = SpriteWidth
    'Next i
        Dim a As Long
        Dim rEmpty As RECT, rEmpty2 As RECT
        Dim ddsdOrigine As DDSURFACEDESC2
        Dim CorRGB As RGBColour
        Dim rC As Byte
        Dim bC As Byte
        Dim gC As Byte
        

    
        
    'carrega as outras cores
   
    SortearTintas
   
    
    AlterarCorCarro (RGB(CoresTinta(0).vermelho, CoresTinta(0).verde, CoresTinta(0).azul))
    'CarregarDemaisCores
    
        For a = 1 To 19
            LoadSprite explosao(a), App.Path & "\graficos\animacoes\explosao\explosao" & a & ".bmp", 450, 376, CLng(Preto)
        Next a
        'trechos retos
        LoadSprite pistas(0), App.Path & "\graficos\pistas\pista1.bmp", 476, 120, CLng(Preto)
        
        'pinta os 10 checkpoints verticais
        
        
        
        'rampas
        LoadSprite pistas(2), App.Path & "\graficos\pistas\pista_rampa.bmp", 620, 224, CLng(Preto)
        
        
        LoadSprite pistas(4), App.Path & "\graficos\pistas\pista2_rampa.bmp", 500, 225, CLng(Preto)
        
        
        
        LoadSprite pistas(6), App.Path & "\graficos\pistas\pista3_rampa.bmp", 684, 164, CLng(Preto)
        
        
        LoadSprite pistas(8), App.Path & "\graficos\pistas\pista_curva_ALTA_ESQ.bmp", 934, 302, CLng(Preto)
        
        
        LoadSprite pistas(10), App.Path & "\graficos\pistas\pista1_horizontal.bmp", 417, 139, CLng(Preto)
        
        
        LoadSprite pistas(12), App.Path & "\graficos\pistas\pista_rampa_horizontal.bmp", 623, 179, CLng(Preto)
        
        
        LoadSprite pistas(14), App.Path & "\graficos\pistas\pista2_horizontal.bmp", 607, 199, CLng(Preto)
        
        
        LoadSprite pistas(16), App.Path & "\graficos\pistas\pista_curva_ALTA_DIR.bmp", 612, 270, CLng(Preto)
        
       
        LoadSprite pistas(18), App.Path & "\graficos\pistas\pista_rampa_vertical_direita.bmp", 634, 171, CLng(Preto)
        
       
        LoadSprite pistas(20), App.Path & "\graficos\pistas\pista_curva_BAIXA_ESQ.bmp", 601, 273, CLng(Preto)
        
       
        LoadSprite pistas(22), App.Path & "\graficos\pistas\rampinha_horizontal.bmp", 674, 300, CLng(Preto)
        
              
       LoadSprite pistas(24), App.Path & "\graficos\pistas\pista_curva_BAIXA_DIR.bmp", 837, 180, CLng(Preto)
        
        
        LoadSprite pistas(26), App.Path & "\graficos\pistas\rampinha_vertical.bmp", 610, 173, CLng(Preto)
        
        LoadSprite pistas(27), App.Path & "\graficos\pistas\pista2ladeirashorizontal.bmp", 569, 220, CLng(Preto)
        
        LoadSprite pistas(28), App.Path & "\graficos\pistas\rampinha_horizontal2.bmp", 492, 141, CLng(Preto)
        
        LoadSprite pistas(29), App.Path & "\graficos\pistas\pista4_rampa.bmp", 586, 134, CLng(Preto)
        
        LoadSprite pistas(30), App.Path & "\graficos\pistas\rampinha_vertical2.bmp", 449, 174, CLng(Preto)
        
        LoadSprite pistas(31), App.Path & "\graficos\pistas\cruz_direita.bmp", 812, 248, CLng(Preto)
        
        LoadSprite pistas(32), App.Path & "\graficos\pistas\rampastep_horizontal.bmp", 456, 134, CLng(Preto)
        'sombra
        LoadSprite sombra, App.Path & "\graficos\sombra.bmp", 58, 29, branco
        LoadSprite Oleo, App.Path & "\graficos\equipamentos\oleo.bmp", 33, 20, RGB(4, 4, 4)
        
        'fumacas
        LoadSprite Fumaca(0), App.Path & "\graficos\animacoes\fumaca4.bmp", 10, 7, CLng(Preto)
        LoadSprite Fumaca(1), App.Path & "\graficos\animacoes\fumaca1.bmp", 24, 29, CLng(Preto)
        LoadSprite Fumaca(2), App.Path & "\graficos\animacoes\fumaca2.bmp", 30, 27, CLng(Preto)
        LoadSprite Fumaca(3), App.Path & "\graficos\animacoes\fumaca3.bmp", 37, 35, CLng(Preto)
        'LoadSprite pontinho, "c:\point.bmp", 22, 14, clng(Preto)
     '   pistas(0).width = 472
      '  pistas(0).height = 120
        'lasers
      
        Dim AngleToRotate As Long
        LoadSprite laser(21), App.Path & "\graficos\equipamentos\laser.bmp", 34, 24, CLng(Preto)
        For x = 0 To 23
        If x <> 21 Then
           LoadSprite laser(x), App.Path & "\graficos\paleta\preto.bmp", 144, 134, CLng(Preto)
            AngleToRotate = GetAngleToMoveObject(x)
            RotateSprite laser(21), laser(x), AngleToRotate, 0, 0
        End If
        Next x
        
        'pilotos
        LoadSprite pilotos(0), App.Path & "\graficos\pilotos\ivanzypher.bmp", 136, 149, CLng(Preto)
        LoadSprite pilotos(1), App.Path & "\graficos\pilotos\jake badlands.bmp", 136, 149, CLng(Preto)
        LoadSprite pilotos(2), App.Path & "\graficos\pilotos\katarina lyons.bmp", 136, 149, CLng(Preto)
        LoadSprite pilotos(3), App.Path & "\graficos\pilotos\snake synders.bmp", 136, 149, CLng(Preto)
        LoadSprite pilotos(4), App.Path & "\graficos\pilotos\tarquin.bmp", 136, 149, CLng(Preto)
        LoadSprite pilotos(5), App.Path & "\graficos\pilotos\cyberhawks.bmp", 136, 149, CLng(Preto)
        
        LoadSprite SelectScreen, App.Path & "\graficos\animacoes\tela1.bmp", 486, 46, CLng(Preto)
        'LoadSprite Mapa.imagem, App.Path & "\graficos\cars\sequencia1206_.bmp", 1000, 1000, clng(Preto)
        'Mapa.width = 1000
        'Mapa.height = 1000
End Sub

Public Sub ClearBuffer()

Dim DestRect As RECT

    'This routine clears the backbuffer and displays the text
    With DestRect           'This rectangle must be defined the same as the screen (which we set as 640x480)
        .Bottom = (480 / ZoomY)
        .Left = 0
        .Right = (640 / ZoomX)
        .Top = 0
    End With
    
    'Fill the entire backbuffer with the colour dictated by "ColourFill"
    BackBuffer.BltColorFill DestRect, ColourFill
    
    'Draw the text on the backbuffer the describe the keys
    
End Sub

Public Sub LoadSprite(ByRef Sprite As Sprites, File As String, ByVal bWidth As Long, ByVal bHeight As Long, ColourKey As Long)

'Zoom = 1
bWidth = bWidth / ZoomX
bHeight = bHeight / ZoomY
Dim CKey As DDCOLORKEY
Dim ddsdNewSprite As DDSURFACEDESC2
    'Dim Sprite As DirectDrawSurface7
    'This routine loads sprites in the FObject file and sets their colour keys
    ddsdNewSprite.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT           'Set the surface description to include the Capabilities, Width and Height
    ddsdNewSprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN                    'Set the surface's capabilities to be an offscreen surface
    ddsdNewSprite.lWidth = bWidth                                           'Set the width of the surface
    ddsdNewSprite.lHeight = bHeight                                         'Set the height of the surface
    'Set Sprite = dd.CreateSurfaceFromResource("", File, ddsdNewSprite)      'Load the bitmap from the resource file into the surface using the surface description
    Set Sprite.imagem = dd.CreateSurfaceFromFile(File, ddsdNewSprite)
    
    CKey.low = ColourKey                                                    'Set the low value of the colour key
    CKey.high = ColourKey                                                   'and the high value (in this case they're the same because we're not using a range)
    Sprite.imagem.SetColorKey DDCKEY_SRCBLT, CKey                                  'Set the sprites colourkey using the key just created
    Sprite.Height = bHeight
    Sprite.Width = bWidth
End Sub

Public Sub DisplaySprite(tSprite As Sprites, Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, Optional ByVal tZoom As Long = 100, Optional ByVal NoCamera As Boolean)
If UseBots = True Then Exit Sub
Dim SrcRect As RECT
Dim DestRect As RECT
'#imagem enquadrada dentro da tela
'camera visao
'
'  - /  -
'   /
'0,0
'pra direita , camera diminue em y
'//camera acompanha o carro
'para cima camera diminui em x
If NoCamera = False Then
    If CameraSeguirOutroPlayer = 0 Then
    Camera.x = Player(0).position2D.x - (80 / ZoomX)
    Camera.y = Player(0).position2D.y - ((250 / ZoomY) / ZoomY)
    Else
    Camera.x = OtherPlayers(CameraSeguirOutroPlayer).position2D.x - (80 / ZoomX)
    Camera.y = OtherPlayers(CameraSeguirOutroPlayer).position2D.y - ((250 / ZoomY) / ZoomY)
    End If
'//camera fixa
'Camera.x = 900
'Camera.y = -600

  x = x - Camera.x + CamLeft + ControlCameraX
  y = y - Camera.y + CamUp - ControlCameraY
End If

    With SrcRect
        .Bottom = tSprite.Height
        .Left = 0
        .Right = tSprite.Width
        .Top = 0
    End With
  
  If gl_noclipping = False Then
    If x < 0 Then SrcRect.Left = Abs(x)
    If y < 0 Then SrcRect.Top = Abs(y)
    If x + tSprite.Width > (640 / ZoomX) Then SrcRect.Right = tSprite.Width - (x + tSprite.Width - (640 / ZoomX))
    If y + tSprite.Height > (480 / ZoomY) Then SrcRect.Bottom = tSprite.Height - (y + tSprite.Height - (480 / ZoomY))
  
  'posiciona
  
  Dim PosX As Long
  Dim PosY As Long
  
  If x < 0 Then PosX = 0 Else PosX = x
  If y < 0 Then PosY = 0 Else PosY = y
  Else
  PosX = x
  PosY = y
  End If
  BackBuffer.BltFast PosX, PosY, tSprite.imagem, SrcRect, gl_flag
  Exit Sub
  
  '#se toda a imagem cabe na tela
  If x >= 0 And y >= 0 And x + tSprite.Width <= (640 / ZoomX) And y + tSprite.Height <= (480 / ZoomY) Then
    'Set up the source rectangle
    With SrcRect
        .Bottom = tSprite.Height
        .Left = 0
        .Right = tSprite.Width
        .Top = 0
    End With
    'Set up the destination rectangle (taking zoom into account)
    'BackBuffer.Blt DestRect, tSprite.imagem, SrcRect, DDBLT_KEYSRC Or DDBLT_WAIT
    
    BackBuffer.BltFast x, y, tSprite.imagem, SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
 
    Exit Sub
   End If
   'Exit Sub

'#ponta esquerda alta , porem parte da imagem aparece
    If x < 0 And y < 0 And x + tSprite.Width > 0 And y + tSprite.Height > 0 Then
        With SrcRect
        .Bottom = tSprite.Height
        .Left = Abs(x)
        .Right = tSprite.Width
        .Top = Abs(y)
    End With
    'Set up the destination rectangle (taking zoom into account)
        'Blit the surface on to the backbuffer at the specified location, using the colour key of the source
        BackBuffer.BltFast 0, y, tSprite.imagem, SrcRect, DDBLTFAST_SRCCOLORKEY
        Exit Sub
    End If


Exit Sub
'#ponta esquerda sai da tela , porem parte da imagem aparece
    If x < 0 And x + tSprite.Width >= 0 And y >= 0 And y < (480 / ZoomY) And y + tSprite.Height < (480 / ZoomY) Then
        With SrcRect
        .Bottom = tSprite.Height
        .Left = 0 + Abs(x)
        .Right = tSprite.Width
        .Top = 0
    End With
    'Set up the destination rectangle (taking zoom into account)
        'Blit the surface on to the backbuffer at the specified location, using the colour key of the source
        BackBuffer.BltFast 0, y, tSprite.imagem, SrcRect, DDBLTFAST_SRCCOLORKEY
        Exit Sub
    End If
    
'#ponta direita sai da tela , porem parte da imagem aparece
    If x + tSprite.Width >= (640 / ZoomX) And x >= 0 And x < (640 / ZoomX) And y >= 0 And y < (480 / ZoomY) And y + tSprite.Height < (480 / ZoomY) Then
        With SrcRect
        .Bottom = tSprite.Height
        .Left = 0
        .Right = tSprite.Width - (x + tSprite.Width - (640 / ZoomX))
        .Top = 0
    End With
    'Set up the destination rectangle (taking zoom into account)
        'Blit the surface on to the backbuffer at the specified location, using the colour key of the source
        BackBuffer.BltFast x, y, tSprite.imagem, SrcRect, DDBLTFAST_SRCCOLORKEY
        Exit Sub
    End If
    
    '#ponta alta sai da tela , porem parte da imagem aparece
    If y < 0 And y + tSprite.Height >= 0 And x <= 0 Then
        With SrcRect
        .Bottom = tSprite.Height
        .Left = 0
        .Right = tSprite.Width
        .Top = 0 + Abs(y)
    End With
    'Set up the destination rectangle (taking zoom into account)
        'Blit the surface on to the backbuffer at the specified location, using the colour key of the source
        BackBuffer.BltFast 0, 0, tSprite.imagem, SrcRect, DDBLTFAST_SRCCOLORKEY
        Exit Sub
    End If
    
    If y < 0 And y + tSprite.Height >= 0 And x > 0 And x < (640 / ZoomX) Then
        With SrcRect
        .Bottom = tSprite.Height
        .Left = 0
        .Right = tSprite.Width
        .Top = 0 + Abs(y)
    End With
    'Set up the destination rectangle (taking zoom into account)
        'Blit the surface on to the backbuffer at the specified location, using the colour key of the source
        BackBuffer.BltFast x, 0, tSprite.imagem, SrcRect, DDBLTFAST_SRCCOLORKEY
        Exit Sub
    End If


'#ponta baixa sai da tela , porem parte da imagem aparece
    
    If y + tSprite.Height >= (480 / ZoomY) And y >= 0 And y < (480 / ZoomY) And x <= 0 Then
        'Do
        'Loop
        With SrcRect
        .Bottom = tSprite.Height - (y + tSprite.Height - (480 / ZoomY))
        .Left = 0 + Abs(x)
        .Right = tSprite.Width
        .Top = 0
    End With
    'Set up the destination rectangle (taking zoom into account)
        'Blit the surface on to the backbuffer at the specified location, using the colour key of the source
        BackBuffer.BltFast 0, y, tSprite.imagem, SrcRect, DDBLTFAST_SRCCOLORKEY
        Exit Sub
    End If
    
    If y + tSprite.Height >= (480 / ZoomY) And y >= 0 And y < (480 / ZoomY) And x > 0 Then
        'Do
        'Loop
        With SrcRect
        .Bottom = tSprite.Height - (y + tSprite.Height - (480 / ZoomY))
        .Left = 0 + Abs(x)
        .Right = tSprite.Width
        .Top = 0
    End With
    'Set up the destination rectangle (taking zoom into account)
        'Blit the surface on to the backbuffer at the specified location, using the colour key of the source
        BackBuffer.BltFast x, y, tSprite.imagem, SrcRect, DDBLTFAST_SRCCOLORKEY
        Exit Sub
    End If

End Sub

Public Sub Flip()
On Error Resume Next
    CountFrames = CountFrames + 1
    'Flip the attached surface (the backbuffer) to the screen
If NoVSync = True Then
   Primary.Flip Nothing, DDFLIP_WAIT
Else
    Primary.Flip Nothing, DDFLIP_NOVSYNC
End If
 ''Primary.Flip Nothing,
    
End Sub

Public Sub Terminate()
End
Dim i As Integer

    'This routine must destroy all surfaces and restore display mode
    Set Primary = Nothing
    Set BackBuffer = Nothing
    For i = 0 To UBound(cars)
        Set cars(0, 0, i, 0, 0).imagem = Nothing
    Next
    'For i = 0 To 23
     '   Set laser(i).imagem = Nothing
    'Next
    
    Call dd.RestoreDisplayMode
    Call dd.SetCooperativeLevel(FrmDirectX.hWnd, DDSCL_NORMAL)

End Sub



