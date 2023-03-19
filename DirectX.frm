VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmDirectX 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "DirectX Demo"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "DirectX.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAlterarCorCarro 
      Index           =   0
      Left            =   240
      Top             =   1080
   End
   Begin VB.Timer TmrChangeColor 
      Left            =   2040
      Top             =   360
   End
   Begin VB.Timer tmrcorridacompletada 
      Left            =   2640
      Top             =   240
   End
   Begin VB.Timer tmrControlCamera 
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Timer tmrStartGame 
      Left            =   3480
      Top             =   1920
   End
   Begin VB.Timer tmrattackBonus 
      Left            =   3600
      Top             =   1200
   End
   Begin VB.Timer tmrSayIt 
      Left            =   4200
      Top             =   2760
   End
   Begin VB.Timer tmrConnectError 
      Left            =   2640
      Top             =   1320
   End
   Begin VB.Timer tmrRefresh 
      Left            =   3600
      Top             =   480
   End
   Begin VB.Timer tmrProcessGetList 
      Left            =   4080
      Top             =   960
   End
   Begin MSWinsockLib.Winsock GetServerInfo 
      Index           =   0
      Left            =   1560
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrwaitDikX 
      Left            =   2400
      Top             =   840
   End
   Begin VB.Timer tmrShowAtack 
      Left            =   3240
      Top             =   2280
   End
   Begin MSWinsockLib.Winsock StringStream 
      Left            =   1080
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrStart 
      Left            =   3120
      Top             =   960
   End
   Begin VB.Timer tmrDikX 
      Left            =   840
      Top             =   2160
   End
   Begin MSWinsockLib.Winsock DataStream 
      Left            =   2040
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrProccessSaidText 
      Left            =   2280
      Top             =   2040
   End
   Begin VB.Timer tmrjoined 
      Left            =   120
      Top             =   2640
   End
   Begin VB.Timer tmrShowString 
      Left            =   240
      Top             =   1800
   End
   Begin VB.Timer TimerConnect 
      Left            =   2040
      Top             =   2640
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1560
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer naoderrapar 
      Left            =   4200
      Top             =   2160
   End
   Begin VB.Timer PararDerrapagem 
      Left            =   3840
      Top             =   1680
   End
   Begin VB.Timer tmrexplosaowait2 
      Left            =   3600
      Top             =   2520
   End
   Begin VB.Timer tmrFrames 
      Left            =   3000
      Top             =   1680
   End
   Begin VB.Timer tmrRetorno 
      Left            =   1440
      Top             =   480
   End
   Begin VB.Timer tmrexplosaowait 
      Left            =   840
      Top             =   360
   End
   Begin VB.Timer tmrexplosao 
      Left            =   2760
      Top             =   2520
   End
   Begin VB.Timer tmrSetas 
      Left            =   1080
      Top             =   2520
   End
End
Attribute VB_Name = "FrmDirectX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'
'IMPORTANT TO NOTE:
'
'When using resource files and the CreateSurfaceFromResource
'command in DirectX 7.0, you will not be able to run your program
'by simply pressing F5 or selecting the run menu item, you will
'get an error. dX7 will only recognize your resource file if you
'compile your program and RUN THE COMPILED VERSION.
'player(0
' - Lucky
'
'*****************************************************************

'...............
'regras importantes
'poligono(0) do jogo deve ser nivel 0 (nem subida e nem descida)
Option Explicit
Public showInfo As Boolean
'Public WithEvents winsock1 As TCPIP
Enum PolignType
vertical = 0
horizontal = 1
curva_esquerda_baixa = 2
curva_esquerda_alta = 3
curva_direita_baixa = 4
curva_direita_alta = 5
buraco = 6
Rampa = 7
checkpoint = 8
largada = 9
RampaH = 10
sObjeto = 11

End Enum

Enum lados
 nenhum = 0
 esquerda = 1
 direita = 2
 cima = 3
 baixo = 4
End Enum


Enum Mundos
chem_vi = 0
drakonis = 1
bogmire = 2
new_mojave = 3
nho = 4
inferno = 5
End Enum




Private Sub DataStream_Connect()
DataStream.Tag = "connected"

End Sub

Private Sub DataStream_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
 'If PintandoCarro = True Then Exit Sub
 Dim buffer() As Byte
 Dim command As command
 Dim OtherTemp As otherPlayersData
 Dim x As Long
 Dim y As Long
 ReDim buffer(0 To bytesTotal - 1)
 Dim oSize As Long
  
  DataStream.GetData buffer
    'primeiro descobre o comando
   oSize = Int(bytesTotal / LenB(OtherTemp)) - 1
    
    If bytesTotal >= LenB(OtherTemp) Then
    
    For y = 0 To oSize
     'localiza quem possui o id
     
     If bytesTotal - (y * LenB(OtherTemp)) >= LenB(OtherTemp) Then
     CopyMemory ByVal VarPtr(OtherTemp), buffer(y * LenB(OtherTemp)), LenB(OtherTemp)
     
        For x = 1 To 100
        'SaidText((x Mod 4) + 1) = OtherPlayers(x).data.Id & "  " & OtherTemp.Id
        If OtherPlayers(x).Data.id = OtherTemp.id Then
         Form2.Caption = OtherTemp.blow
        If bytesTotal >= LenB(OtherPlayers(x).Data) Then
            CopyMemory ByVal VarPtr(OtherPlayers(x).Data), ByVal VarPtr(OtherTemp), LenB(OtherTemp)
            If GameStarted = True And OtherTemp.CarroExplosao = 0 Then Crash = CarCrash
            If Crash = True Then
                Sound.WavPlay sons.eCarCrash, EffectsVolume
            End If
                
                
            OtherPlayers(x).ImageIndex = x
            
            If OtherPlayers(x).Active = False Then
                
                'PintandoCarro = True
                tmrAlterarCorCarro(x).Tag = CStr(OtherPlayers(x).Data.color)
                tmrAlterarCorCarro(x).Interval = 10
                'AlterarCorCarro OtherPlayers(x).Data.color, True, x
                'PintandoCarro = False
            End If
            
            OtherPlayers(x).Active = True
            OtherPlayers(x).AcaboudeReceber = True
            Exit For
        End If
            
        End If
        'DoEvents
       Next x
     If OtherTemp.NewObject.Active = True Then
        
        For x = 0 To 2000
            If ObjectsFromNet(x).Active = False Then
                'CopyMemory ByVal VarPtr(ObjectsFromNet(x)), ByVal VarPtr(OtherTemp.NewObject), lenb(ObjectsFromNet(x))
                ObjectsFromNet(x) = OtherTemp.NewObject
                If ObjectsFromNet(x).tipo = slaser Then ObjectsFromNet(x).PolignToUse = CreateObjectPolign(60, 100, OtherTemp.positionX, OtherTemp.positionY, slaser)
                If ObjectsFromNet(x).tipo = sOil Then ObjectsFromNet(x).PolignToUse = CreateObjectPolign(20, 50, OtherTemp.positionX, OtherTemp.positionY, sOil)
                Exit For
            End If
        Next x
    End If
    End If
Next y
End If
'SaidText(4) = "fim"
    
End Sub

Private Sub DataStream_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
If Number = 10050 Or Number = 10051 Or Number = 10052 Or Number = 10053 Or Number = 10054 Or Number = 10057 Or Number = 10058 Or Number = 10060 Or Number = 10061 Or Number = 10064 Or Number = 10065 Then
DataStream.Tag = "error"
DataStream.Close
DataStream.Connect server.Address, 20778
End If

End Sub

Private Sub Form_Click()

 'seleciona o texto na list
 On Error Resume Next
'ReDim GetServerInfo(0 To 20000)

Dim FontInfo As New StdFont
 Dim x As Long
 Dim k As Long
 If GameStatus = 4 Then
    If MouseX > 30 And MouseX < 490 Then
    'y=ShiftServerPage + 90 + x * 12
    'y-12x=ShiftServerPage + 90
   ' -12x=ShiftServerPage + 90-y
   '-x=(ShiftServerPage + 90-y)/12
    'x=-(ShiftServerPage + 90-y)/12
    'detetca o x texto
 
    x = -(ShiftServerPage + 90 - MouseY) / 12
       
    If x >= LBound(serverInfo) And x <= UBound(serverInfo) Then
        For k = 0 To UBound(serverInfo)
            serverInfo(k).Selected = False
        Next k
        serverInfo(x).Selected = True
        SelectedList = x
        End If
    End If
 
 'conectar
 'BackBuffer.DrawText 410, 310, "Conectar", False
 If MouseX > 410 And MouseX < 530 And MouseY > 310 And MouseY < 335 Then
        If serverInfo(SelectedList).noIp <> Empty Then
            BackBuffer.SetFontTransparency True
            BackBuffer.SetForeColor vbWhite
            FontInfo.Bold = True
            FontInfo.Size = 12
            FontInfo.name = "Verdana"
            FontInfo.Italic = True
            BackBuffer.SetFont FontInfo
            GameStatus = 5
            Sound.WavPlay sons.aceito, EffectsVolume
            If serverInfo(SelectedList).name <> Empty Then ConnectStatus = "conectando-se com " & serverInfo(SelectedList).name Else ConnectStatus = "conectando-se com " & serverInfo(SelectedList).noIp
            If serverInfo(SelectedList).noIp <> Empty Then
                TimerConnect.Interval = 1
            End If
        Else
            Sound.WavPlay sons.naoaceito, EffectsVolume
        End If
    End If
 
 'atualizar
            If MouseX > 281 And MouseX < 399 And MouseY > 300 And MouseY < 335 Then
                For k = 0 To UBound(serverInfo)
                 '   serverInfo(k).Selected = False
                  '  serverInfo(k).name = Empty
                   ' serverInfo(k).noIp = Empty
                    serverInfo(k).PlayersInInfo = Empty
                    'serverInfo(k).ping = Empty
                    
                Next k
                ShiftServerPage = 0
                RefreshingServer = True
                tmrRefresh.Interval = 3000
                
            End If
            
 'voltar
            
            If MouseX > 161 And MouseX < 280 And MouseY > 300 And MouseY < 335 Then
                GameStatus = 3
                
            End If
    End If
 'BackBuffer.DrawText 335, ShiftServerPage + 90 + x * 12, serverInfo(x).PlayersInInfo, False
           
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 96 Then
 
Exit Sub
End If

If ConsoleVisible = True Then
    If KeyAscii = 13 Then 'enter pressionado
        ConsoleVisible = False = False
        'GlobalString.command = SendingText
        'DeletePasswordIntoName MyName
        'ps = Trim(MyName) & ": " & Trim(TextToSend)
        'PutStringInArray GlobalString.parametroStr, ps
        'GlobalString.StringLength = Len(ps)
          'SaidText(1) = "teste" & TextToSend
          'SaidText(2) = "teste" & MyName
     'SendString StringStream, ps
     ExecuteCommands ConsoleTxt
     ConsoleTxt = Empty
    'TextToSend = Empty
    Exit Sub
    End If
    
    If KeyAscii = 8 Then
    ConsoleTxt = Left(ConsoleTxt, Len(ConsoleTxt) - 1)
    Exit Sub
    End If
   'If KeyAscii <> 13 Then TextToSend = TextToSend & Chr$(KeyAscii)

If KeyAscii <> 13 And KeyAscii <> 8 And Chr(KeyAscii) <> "'" Then ConsoleTxt = ConsoleTxt & Chr(KeyAscii)

Exit Sub

''''sa
End If

Dim ps As String
Dim ext As Long
'teclas no game

     'If GameStatus = 6 Then
        'If KeyAscii = 8 Then
           ' MyName = Left(MyName, lenb(MyName) - 1)
          '  End
         '   Exit Sub
        'End If
    
       ' MyName = MyName & Chr$(KeyAscii)
      '  Exit Sub
     'End If

Dim command As command
tmrProccessSaidText.Interval = 4000
Dim sCont As Long
Dim res As Long
If EscrevendoTexto = False And ConectadoAoServidor = True And UCase(Chr(KeyAscii)) = "Y" Then EscrevendoTexto = True: TextToSend = Empty: Exit Sub
If EscrevendoTexto = True Then
    If KeyAscii = 13 Then 'enter pressionado
        EscrevendoTexto = False
        GlobalString.command = SendingText
        DeletePasswordIntoName MyName
        ps = Trim(MyName) & ": " & Trim(TextToSend)
        PutStringInArray GlobalString.parametroStr, ps
        'GlobalString.StringLength = Len(ps)
          'SaidText(1) = "teste" & TextToSend
          'SaidText(2) = "teste" & MyName
     SendString StringStream, ps
    TextToSend = Empty
    Exit Sub
    End If
    If KeyAscii = 8 Then
    TextToSend = Left(TextToSend, Len(TextToSend) - 1)
    Exit Sub
    End If
    If KeyAscii <> 13 Then TextToSend = TextToSend & Chr$(KeyAscii)
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 192 Then ' tecla "'"
  ConsoleVisible = Not (ConsoleVisible)

End If

End Sub

Private Sub Form_Load()
'Sleep 80000
'End
'Dim k As otherPlayersData

'End
 
 PistaAtual = 0
'' DDBLTFAST_DESTCOLORKEY = 2
''DDBLTFAST_DONOTWAIT = 32
''DDBLTFAST_NOCOLORKEY = 0
 ''DDBLTFAST_SRCCOLORKEY = 1
'' DDBLTFAST_WAIT = 16
VSync = True
ConsoleRoll = -250
AudioVolume = 10000
LarryVolume = 10000
MusicVolume = 10000
EffectsVolume = 10000
CamSpeed = 12
cl_message = True
gl_flag = DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
cl_gaitestimation = True

UseBots = False
If UseBots = True Then Form2.Visible = True
BotConnectAt = PegarTextoArquivo(App.Path & "/servidor.txt")
Dim a As Long
 Dim b As Long
 Dim d As Long
 Dim meuindice As Long
 Dim indicedele As Long
 Dim x As Long
For x = 1 To 3000
Load GetServerInfo(x)
Next x
For x = 1 To 20
    Load tmrAlterarCorCarro(x)
Next x
 'meuindice = 21
 'indicedele = 21
 'a = indicedele - 12
  '                     a = validateIndex(a)
   '                     b = meuindice - a
    '                    '-12
     '                   d = indicedele + b
      '                  d = validateIndex(d)
       
        '                End
'End
 'On Error Resume Next
Dim c As otherPlayersData
'Open "c:\colors.txt" For Append As #122
Dim s As OLE_COLOR

Dim rC As Byte
        Dim bC As Byte
        Dim gC As Byte
   
   Set Sound = New clsDxSound
   
SelectedList = -1
'MyName = InputBox("digite o seu apelido para ser identificado no server", "Rock N'Roll Racing Remix")
PodedesenharoCarro = False
MyName = PegarValor("software\\RockNRollRacing", "name")
If UseBots = True Then MyName = "bot"

DeletePasswordIntoName MyName
MenorNivel = -10000
'ShowGrides = False
ShowGeometry = False
diff = 0.00001
tantoaAbaixar = 115
zAxis = 75
    yAxis = 41
    
tmrFrames.Interval = 1000
VelocidadeMaxima = 7
ZoomX = 1
ZoomY = 1
ReDim poligono(0)
ReDim ServerAddress(0)
ReDim RecPos(0 To 20) As pos
ReDim Player(0)
'Player(0).color = CLng(PegarValor("software\\RockNRollRacing", "cor"))

ReDim ArmasTraseirasNaPista(0)
ReDim ArmasFrontaisNaPista(0)
ReDim SprayPixadas(0)
Player(0).armas.Traseiras.tipo = sOil
Player(0).armas.Traseiras.quantidade = 7
Player(0).armas.Traseiras.QuantasTenho = 7

Player(0).armas.Frontais.QuantasTenho = 7
Player(0).armas.Frontais.quantidade = 7
Player(0).armas.Frontais.tipo = slaser
ReDim CheckPoints(0 To 999)
    '           cars,angles,upDown,pneus
    'player -  updown ->0= nivel normal , 1 =pra cima,1 nivel , 2=pra cima, nivel2,3=pra baixo nivel, 4=pra baixo nivel 2
    ReDim cars(0 To 19, 0 To 4, 0 To 23, 0 To 4, 0 To 2)
    '           até 20 carros na pista
    ReDim pistas(0 To 50)
   
   Sound.Initialize FrmDirectX 'initialize audio

   
   If UseBots = False Then DDraw.Initialize          'Initialize DirectDraw
    If UseBots = True Then Running = True
    DInput.Initialize         'Initialize DirectInput
    
    If UseBots = False Then
   nPlayWav(sons.paranoid) = Sound.WavLoad(App.Path & "\audio\paranoid.wav", True)
  nPlayWav(sons.badtobone) = Sound.WavLoad(App.Path & "\audio\bad to the bone.wav", True)
  nPlayWav(sons.pettergun) = Sound.WavLoad(App.Path & "\audio\peter gun.wav", True)
  nPlayWav(sons.borntobewild) = Sound.WavLoad(App.Path & "\audio\born to be wild.wav", True)
  nPlayWav(sons.highwayStar) = Sound.WavLoad(App.Path & "\audio\highway star.wav", True)
  nPlayWav(sons.derrapagem) = Sound.WavLoad(App.Path & "\audio\derrapagem.wav")
  nPlayWav(sons.cornercrash) = Sound.WavLoad(App.Path & "\audio\cornercrash.wav")
  nPlayWav(sons.getobjects) = Sound.WavLoad(App.Path & "\audio\getobjects.wav")
  nPlayWav(sons.salto) = Sound.WavLoad(App.Path & "\audio\salto.wav")
  nPlayWav(sons.eLaser) = Sound.WavLoad(App.Path & "\audio\laser.wav")
  nPlayWav(sons.eOleo) = Sound.WavLoad(App.Path & "\audio\oleo.wav")
  nPlayWav(sons.aceito) = Sound.WavLoad(App.Path & "\audio\menu2.wav")
  nPlayWav(sons.moving) = Sound.WavLoad(App.Path & "\audio\menu1.wav")
  nPlayWav(sons.naoaceito) = Sound.WavLoad(App.Path & "\audio\menu3.wav")
  nPlayWav(sons.eExplosao) = Sound.WavLoad(App.Path & "\audio\explosao.wav")
  nPlayWav(sons.eLaser2) = Sound.WavLoad(App.Path & "\audio\laser.wav")
  nPlayWav(sons.eLaser3) = Sound.WavLoad(App.Path & "\audio\laser.wav")
  nPlayWav(sons.eLaser4) = Sound.WavLoad(App.Path & "\audio\laser.wav")
  nPlayWav(sons.eLaser5) = Sound.WavLoad(App.Path & "\audio\laser.wav")
  nPlayWav(sons.eLaser6) = Sound.WavLoad(App.Path & "\audio\laser.wav")
  nPlayWav(sons.eLaser7) = Sound.WavLoad(App.Path & "\audio\laser.wav")
  nPlayWav(sons.eOleo2) = Sound.WavLoad(App.Path & "\audio\oleo.wav")
  nPlayWav(sons.eOleo3) = Sound.WavLoad(App.Path & "\audio\oleo.wav")
  nPlayWav(sons.eOleo4) = Sound.WavLoad(App.Path & "\audio\oleo.wav")
  nPlayWav(sons.eOleo5) = Sound.WavLoad(App.Path & "\audio\oleo.wav")
  nPlayWav(sons.eOleo6) = Sound.WavLoad(App.Path & "\audio\oleo.wav")
  nPlayWav(sons.eOleo7) = Sound.WavLoad(App.Path & "\audio\oleo.wav")
  nPlayWav(sons.eCarCrash) = Sound.WavLoad(App.Path & "\audio\carcrash.wav")
  nPlayWav(sons.stageset) = Sound.WavLoad(App.Path & "\audio\larry\1.wav")
  nPlayWav(sons.abouttoblow) = Sound.WavLoad(App.Path & "\audio\larry\2.wav")
  nPlayWav(sons.always) = Sound.WavLoad(App.Path & "\audio\larry\3.wav")
  nPlayWav(sons.dollar) = Sound.WavLoad(App.Path & "\audio\larry\4.wav")
  nPlayWav(sons.fadethelast) = Sound.WavLoad(App.Path & "\audio\larry\5.wav")
  nPlayWav(sons.jaminthefirst) = Sound.WavLoad(App.Path & "\audio\larry\6.wav")
  nPlayWav(sons.LastLap) = Sound.WavLoad(App.Path & "\audio\larry\7.wav")
  nPlayWav(sons.hurrysup) = Sound.WavLoad(App.Path & "\audio\larry\8.wav")
  nPlayWav(sons.ouch) = Sound.WavLoad(App.Path & "\audio\larry\9.wav")
  nPlayWav(sons.UaiPaud) = Sound.WavLoad(App.Path & "\audio\larry\10.wav")
  nPlayWav(sons.wow) = Sound.WavLoad(App.Path & "\audio\larry\11.wav")
  nPlayWav(sons.First) = Sound.WavLoad(App.Path & "\audio\larry\12.wav")
  nPlayWav(sons.Second) = Sound.WavLoad(App.Path & "\audio\larry\13.wav")
  nPlayWav(sons.third) = Sound.WavLoad(App.Path & "\audio\larry\14.wav")
  nPlayWav(sons.notcomplete) = Sound.WavLoad(App.Path & "\audio\larry\15.wav")
   nPlayWav(sons.ecyber) = Sound.WavLoad(App.Path & "\audio\larry\23.wav")
  nPlayWav(sons.eIvan) = Sound.WavLoad(App.Path & "\audio\larry\16.wav")
  nPlayWav(sons.eJake) = Sound.WavLoad(App.Path & "\audio\larry\17.wav")
  nPlayWav(sons.eKatarina) = Sound.WavLoad(App.Path & "\audio\larry\18.wav")
  nPlayWav(sons.eRip) = Sound.WavLoad(App.Path & "\audio\larry\19.wav")
  nPlayWav(sons.eSnake) = Sound.WavLoad(App.Path & "\audio\larry\20.wav")
  nPlayWav(sons.eTarquin) = Sound.WavLoad(App.Path & "\audio\larry\21.wav")
  nPlayWav(sons.carneage) = Sound.WavLoad(App.Path & "\audio\larry\22.wav")
 'nPlayWav(sons.apresentacao) = Sound.WavLoad(App.Path & "\audio\1.wav", True)
     'DesenhePista chem_vi, 0, 0, True
     'DesenhePista2 chem_vi, 0, 0, True
     'DesenhePista3 chem_vi, 0, 0, True
    Me.Show                     'Show the form
End If

     'Sound.WavPlay sons.carneage
     'Exit Sub
       'pilotos
    'ecyber = 43
    'eIvan = 44
    'eJake = 45
    'eKatarina = 46
    'eRip = 47
    'eSnake = 48
    'eTarquin = 49
    
    'carneage = 50
    'eHawk = 51
       
    If UseBots = False Then
    LoadSprite planetaNaTela, App.Path & "\graficos\paleta\preto.bmp", 300, 300, CLng(Preto)
    
    LoadSprite Logo(0), App.Path & "\graficos\logotipo\logo.bmp", 439, 152, CLng(Preto)
    LoadSprite Logo(1), App.Path & "\graficos\logotipo\logo2.bmp", 640, 480, CLng(Preto)
    LoadSprite Logo(2), App.Path & "\graficos\logotipo\logo4.bmp", 496, 182, RGB(255, 255, 255)
    'IvanZypher = 0
'JakeBlanders = 1
'KatarinaLyons = 2
'SnakeSanders = 3
'Tarquin = 4
'CyberHawks = 5

    LoadSprite PilotChoose(0), App.Path & "\graficos\logotipo\ivan_screen.bmp", 505, 432, CLng(Preto)
    LoadSprite PilotChoose(1), App.Path & "\graficos\logotipo\jake_screen.bmp", 505, 432, CLng(Preto)
    LoadSprite PilotChoose(2), App.Path & "\graficos\logotipo\katarina_screen.bmp", 505, 432, CLng(Preto)
    LoadSprite PilotChoose(3), App.Path & "\graficos\logotipo\snake_screen.bmp", 505, 432, CLng(Preto)
    LoadSprite PilotChoose(4), App.Path & "\graficos\logotipo\tarquin_screen.bmp", 505, 432, CLng(Preto)
    LoadSprite PilotChoose(5), App.Path & "\graficos\logotipo\cyberhawk_screen.bmp", 505, 432, CLng(Preto)
    
    LoadSprite planetas(0), App.Path & "\graficos\planetas\ivan.bmp", 189, 191, CLng(Preto)
    LoadSprite planetas(1), App.Path & "\graficos\planetas\jake.bmp", 189, 191, CLng(Preto)
    LoadSprite planetas(2), App.Path & "\graficos\planetas\katarina.bmp", 189, 191, CLng(Preto)
    LoadSprite planetas(3), App.Path & "\graficos\planetas\snake.bmp", 189, 191, CLng(Preto)
    LoadSprite planetas(4), App.Path & "\graficos\planetas\tarquin.bmp", 189, 191, CLng(Preto)
    LoadSprite planetas(5), App.Path & "\graficos\planetas\cyber.bmp", 189, 191, CLng(Preto)
    
    LoadSprite estrelas(0), App.Path & "\graficos\estrelas\estrela1.bmp", 27, 18, CLng(Preto)
    LoadSprite estrelas(1), App.Path & "\graficos\estrelas\estrela2.bmp", 11, 13, CLng(Preto)
    LoadSprite estrelas(2), App.Path & "\graficos\estrelas\estrela3.bmp", 16, 17, CLng(Preto)
    LoadSprite estrelas(3), App.Path & "\graficos\estrelas\estrela4.bmp", 13, 17, CLng(Preto)
    LoadSprite estrelas(4), App.Path & "\graficos\estrelas\estrela5.bmp", 13, 17, CLng(Preto)
    
    ReDim botao(0 To 3)
    LoadSprite botao(0), App.Path & "\graficos\b\b1.bmp", 120, 35, CLng(Preto)
    LoadSprite botao(1), App.Path & "\graficos\b\b1.bmp", 120, 35, CLng(Preto)
    LoadSprite botao(2), App.Path & "\graficos\b\b1.bmp", 120, 35, CLng(Preto)
    LoadSprite Tela1, App.Path & "\graficos\b\b2.bmp", 500, 300, CLng(Preto)
    LoadSprite Roll, App.Path & "\graficos\b\roll.bmp", 27, 220, CLng(Preto)
    
    LoadSprite MainScreen, App.Path & "\graficos\logotipo\mainscreen.bmp", 509, 445, CLng(Preto)
    LoadSprite buycarScreen, App.Path & "\graficos\logotipo\buycar.bmp", 510, 443, CLng(Preto)
    LoadSprite interrogacao, App.Path & "\graficos\logotipo\interrogacao.bmp", 75, 40, CLng(Preto)
    'LoadSprite spray, "c:\spray\spray.bmp", 75, 40, CLng(Preto)
    LoadSprite ConsoleScr, App.Path & "\graficos\b\b1.bmp", 640, 280, CLng(Preto)
    LoadSprite ConsoleTxtscr, App.Path & "\graficos\paleta\cinza.bmp", 560, 40, CLng(Preto)
    LoadSprite ConsoleScr, App.Path & "\graficos\b\b1.bmp", 640, 280, CLng(Preto)
    Dim mp(0) As pixelCoord
    For a = 0 To 4
        LoadSprite Tinta(a), App.Path & "\graficos\logotipo\tinta.bmp", 66, 32, CLng(Preto)
        ChangeColors Tinta(a), RGB(255, 0, 0), RGB(CoresTinta(a).vermelho, CoresTinta(a).verde, CoresTinta(a).azul), mp
        ChangeColors Tinta(a), RGB(0, 0, 255), Luminosidade(RGB(CoresTinta(a).azul, CoresTinta(a).verde, CoresTinta(a).vermelho), 3), mp
        ChangeColors Tinta(a), RGB(0, 255, 0), Luminosidade(RGB(CoresTinta(a).azul, CoresTinta(a).verde, CoresTinta(a).vermelho), 5), mp
    Next a
    End If
    
    If UseBots = True Then
    Player(0).piloto = CyberHawks
    Player(0).TopSpeed = 8
    Player(0).TopAcceleration = 14.285
    Player(0).TopCorner = 5  'quanto maior , pior a velocidade
    Player(0).TopJumping = -100
    
            GameStatus = 5
            Sound.WavPlay sons.aceito, EffectsVolume
            serverInfo(SelectedList).name = BotConnectAt
            serverInfo(SelectedList).noIp = BotConnectAt
            If serverInfo(SelectedList).name <> Empty Then ConnectStatus = "conectando-se com " & serverInfo(SelectedList).name Else ConnectStatus = "conectando-se com " & serverInfo(SelectedList).noIp
            If serverInfo(SelectedList).noIp <> Empty Then
                TimerConnect.Interval = 1
                
            End If
    
    End If
    
    MainLoop                    'Run the main loop
    
End Sub

Private Sub MainLoop()
On Error Resume Next
Dim x As Long
Dim CurrRed As Integer
Dim CurrGreen As Integer
Dim CurrBlue As Integer
Dim FontInfo As New StdFont
Dim PreviousFont As New StdFont
Dim processLeft As Long
Dim processRight As Long
Dim processPlanet As Long
Dim StarPos(0 To 100) As pos
Dim DestRect As RECT
Dim emptyRect As RECT
Dim res As Long
Dim r As Byte
Dim g As Byte
Dim b As Byte
Dim buffer() As Byte
PlayerNumber = 3
ShiftImageY = -1700

'Player(0).velocidade = 1
   'On Error GoTo ErrOut
    Dim p As Long
    Dim k As Long
'    Player(0).position.x = 200
 '   Player(0).position.y = 0
    Player(0).position.x = 0 / ZoomX
    Player(0).position.y = 0 / ZoomY
    Player(0).position.Z = 0
    Camera.x = 100
    Camera.y = 350
    Player(0).Car = Maraudercar
    Player(0).Car_Image_Index = 21
    
    
'sorteia a posicao da estrelas
'Randomize Timer
Dim ULimit As Long
Dim LLimit As Long
Dim ShiftStar As Long
Dim rotatePlanet As Long
Dim mp(0) As pixelCoord
For x = 0 To 100
Randomize Timer
ULimit = 240
LLimit = 80
StarPos(x).y = Int((ULimit - LLimit) * Rnd) + LLimit
ULimit = 2000
LLimit = 0

StarPos(x).x = Int((ULimit - LLimit) * Rnd) + LLimit
ULimit = 5
LLimit = 0
StarPos(x).altura = Int((ULimit - LLimit) * Rnd) + LLimit
Next x

    'if gamestatus=ShowTela1
    'PodedesenharoCarro = True
     '   GameStatus = 6
      '  GameStarted = True
       ' ConectadoAoServidor = True
'Player(0).piloto = IvanZypher
'Player(0).TopSpeed = 8
 '   Player(0).TopAcceleration = 14.285
  ' Player(0).TopCorner = 5  'quanto maior , pior a velocidade
   ' Player(0).TopJumping = -100
'PistaAtual = 1
    'If Player(0).piloto = CyberHawks Then
     '   Player(0).TopAcceleration = 16
      '  Player(0).TopJumping = -110
    'End If

    'If Player(0).piloto = JakeBlanders Then
     '   Player(0).TopAcceleration = 16
      '  Player(0).TopCorner = 2
    'End If

    'If Player(0).piloto = KatarinaLyons Then
     '   Player(0).TopJumping = -110
      '  Player(0).TopCorner = 2
    'End If

    'If Player(0).piloto = SnakeSanders Then
     '   Player(0).TopAcceleration = 16
      '  Player(0).TopSpeed = 9
    'End If

    'If Player(0).piloto = Tarquin Then
     '   Player(0).TopSpeed = 9
      '  Player(0).TopJumping = -110
    'End If

    'If Player(0).piloto = IvanZypher Then
     '   Player(0).TopAcceleration = 16
      '  Player(0).TopCorner = 2
    'End If
    
'-///////////////////////////////////////////
    Do While Running
    
    DInput.CheckKeys                                            'Get the current state of the keyboard
      '
    ''If DInput.aKeys(96) = False Then KeyPressedUp(96) = False
    ''If DInput.aKeys(96) And KeyPressedUp(96) = False Then
      ''  End
    ''End If
    'Player(0).color = 0
        'DDraw.ClearBuffer
       'DisplaySprite cars(player(0).color, 0, 1, 0, 0), 0, 0, , True
       'DDraw.Flip
       'DoEvents
       'GoTo fim
       Select Case GameStatus
        Case 0
            
            'mostra logotipo
            'GameStatus = ShowTela1
            MusicaPrincipal = sons.badtobone
            Sound.WavPlay MusicaPrincipal, MusicVolume
            DDraw.ClearBuffer
            DisplaySprite Logo(0), 60, 120, , True
            DDraw.Flip
            Sleep 4000
            'fade
            'FadeScreenToBlack
            GameStatus = 1
            
        Case 1
            
            'Sound.WavPlay sons.apresentacao
        BackBuffer.SetFontTransparency True
        BackBuffer.SetForeColor vbWhite
        FontInfo.Bold = True
        FontInfo.Size = 12
        FontInfo.name = "Verdana"
        FontInfo.Italic = True
        BackBuffer.SetFont FontInfo
        
'BackBuffer.SetFont FontInfo

           If ImageRoll = 0 Then ShiftImageY = ShiftImageY + 20
           If ImageRoll = 1 Then ShiftImageY = ShiftImageY - 20
                DDraw.ClearBuffer
                DisplaySprite Logo(1), 0, 0, , True
                DisplaySprite Logo(2), 60, ShiftImageY + 270, , True
                 BackBuffer.SetForeColor vbYellow
                 BackBuffer.DrawText 10, 240, "Visual Basic 6 Game Remake.Author: Pedro de Lima Freire - 2023 Brasil", False
                 BackBuffer.SetForeColor vbWhite
                 BackBuffer.DrawText 11, 239, "Visual Basic 6 Game Remake.Author: Pedro de Lima Freire - 2023 Brasil", False
                 BackBuffer.SetForeColor vbWhite
                 
                 BackBuffer.SetForeColor vbYellow
                 BackBuffer.DrawText 10, 270, "Projeto para faculdade. Apenas Uso Educacional . Não Comercial", False
                 BackBuffer.SetForeColor vbWhite
                 BackBuffer.DrawText 11, 269, "Projeto para faculdade. Apenas Uso Educacional . Não Comercial", False
                 BackBuffer.SetForeColor vbWhite
                 
                 BackBuffer.SetForeColor vbYellow
                 BackBuffer.DrawText 10, 300, "All Rights to Blizzard Entertainment.", False
                 BackBuffer.SetForeColor vbWhite
                 BackBuffer.DrawText 11, 299, "All Rights to Blizzard Entertainment.", False
                 BackBuffer.SetForeColor vbWhite
                 
                EscrevaNaTela

                DDraw.Flip
            Sleep 10
            If ShiftImageY >= 50 And ImageRoll = 0 Then ImageRoll = 1
            If ImageRoll = 1 And ShiftImageY <= -250 Then ImageRoll = 2
           'GameStatus = ShowTela2
           
        
        Case 2 'escolha dos pilotos
            
            DDraw.ClearBuffer
            'Select Case PlayerNumber
                DisplaySprite PilotChoose(PlayerNumber), 40, 0, , True
                'mostra as estrelas
                If processLeft > 0 And processPlanet > 0 Then
                    'processLeft = processLeft - 1
                    ShiftStar = ShiftStar + 1
                End If
                If processRight And processPlanet > 0 Then
                    'processRight = processRight - 1
                    ShiftStar = ShiftStar - 1
                End If
                
                For x = 0 To 100
                    If StarPos(x).x + ShiftStar > 0 And StarPos(x).x + ShiftStar < 440 Then DisplaySprite estrelas(StarPos(x).altura), StarPos(x).x + ShiftStar + 60, -StarPos(x).y + 260, , True
                Next x
               'mostra o planeta
                RotateSprite planetas(PlayerNumber), planetaNaTela, rotatePlanet Mod 360, 30, 0
                DisplaySprite planetaNaTela, processPlanet + 170, 10, , True
                rotatePlanet = rotatePlanet + 3
                processPlanet = processPlanet - 10
                If processPlanet <= 0 Then
                    processPlanet = 0
                    processRight = 0
                    processLeft = 0
                End If
                If rotatePlanet > 50000 Then rotatePlanet = 0
            'End Select
EscrevaNaTela
            
            DDraw.Flip
            planetaNaTela.imagem.BltColorFill emptyRect, RGB(0, 0, 0)
            Sleep 20
        Case 3
        
        DDraw.ClearBuffer
        DisplaySprite MainScreen, 50, 0, , True
        BackBuffer.SetFontTransparency True
        BackBuffer.SetForeColor vbWhite
        
        FontInfo.Size = 15
        FontInfo.name = "Comic Sans Ms"
        
        BackBuffer.SetFont FontInfo
        
        If MoveDrawBoxLeft < 0 Then MoveDrawBoxLeft = 0
        If MoveDrawBoxLeft > 256 Then MoveDrawBoxLeft = 256
        If MoveDrawBoxLeft = 0 Then BackBuffer.DrawText 100, 400, "Start a Race", False
        If MoveDrawBoxLeft = 64 Then BackBuffer.DrawText 100, 400, "indisponivel", False
        If MoveDrawBoxLeft = 128 Then BackBuffer.DrawText 100, 400, "indisponivel", False
        If MoveDrawBoxLeft = 192 Then BackBuffer.DrawText 100, 400, "Escolher carro / cor", False
        If MoveDrawBoxLeft = 256 Then BackBuffer.DrawText 100, 400, "sair do jogo", False
        
        BackBuffer.DrawText 230, 30, "nome: " & MyName, False
        DisplaySprite pilotos(PlayerNumber), 85, 10, , True
        'sp = sp Mod 24
        DisplaySprite cars(0, 0, FlashDrawBox Mod 24, 0, 0), 80, 190, , True
        BackBuffer.setDrawWidth 3
        BackBuffer.SetForeColor RGB(239, 244, 23)
        res = FlashDrawBox Mod 10
        If Not (res >= 0 And res <= 5) Then BackBuffer.DrawBox 80 + MoveDrawBoxLeft, 320, 142 + MoveDrawBoxLeft, 363
EscrevaNaTela
        
        DDraw.Flip
        FlashDrawBox = FlashDrawBox + 1
        Sleep 50
        
        Case 7
        BackBuffer.SetFontTransparency True
        BackBuffer.SetForeColor vbWhite
        FontInfo.Bold = False
        FontInfo.Size = 12
        FontInfo.name = "Comic Sans Ms"
        FontInfo.Italic = False
        BackBuffer.SetFont FontInfo
        
        DDraw.ClearBuffer
        DisplaySprite buycarScreen, 50, 0, , True
        DisplaySprite cars(0, 0, FlashDrawBox Mod 24, 0, 0), 298, 309, , True
        DisplaySprite interrogacao, 198, 300, , True
        DisplaySprite interrogacao, 430, 300, , True
        Dim a As Long
        For a = 0 To 4
            DisplaySprite Tinta(a), 90, 250 + (34 * a), , True
        Next a
    
        BackBuffer.DrawText 270, 70, "item: Marauder " & Player(0).color, False
        BackBuffer.setDrawWidth 3
        BackBuffer.SetForeColor RGB(239, 244, 23)
        res = FlashDrawBox Mod 10
        If Not (res >= 0 And res <= 5) Then
            'If MoveDrawBoxTop Mod 6 = 0 Then
                'BackBuffer.DrawBox 90, 208, 152, 241
            'Else
                If MoveDrawBoxTop = 0 Then MoveDrawBoxTop = 1
                BackBuffer.DrawBox 90, 216 + (MoveDrawBoxTop Mod 6) * 34, 152, 245 + (MoveDrawBoxTop Mod 6) * 34
            'End If
        End If
EscrevaNaTela
        
        DDraw.Flip
        FlashDrawBox = FlashDrawBox + 1
        Sleep 50
        
        Case 4
        'mostrar servidores disponiveis
         
         DDraw.ClearBuffer
        FontInfo.Size = 12
        FontInfo.Italic = True
        BackBuffer.SetFont FontInfo
         
         BackBuffer.SetForeColor vbWhite
         DisplaySprite Tela1, 30, 50, , True
         DisplaySprite Roll, 480, 80, , True
         DisplaySprite botao(0), 400, 300, , True
         BackBuffer.DrawText 410, 310, "Conectar", False
         
         
         DisplaySprite botao(1), 280, 300, , True
         BackBuffer.DrawText 290, 310, "Atualizar", False
         
         
         DisplaySprite botao(2), 160, 300, , True
         BackBuffer.DrawText 170, 310, "<<", False
         
          FontInfo.Size = 8
            
            FontInfo.Italic = False
            BackBuffer.SetFont FontInfo
         BackBuffer.DrawText 60, 57, "servidor", False
         BackBuffer.DrawText 330, 57, "players", False
         BackBuffer.DrawText 420, 57, "latencia", False
         
         BackBuffer.SetForeColor RGB(0, 740, 10)
         BackBuffer.DrawBox 50, 80, 480, 300
         ShowServerList
EscrevaNaTela
         
         DDraw.Flip
        
        
        Case 5
        DDraw.ClearBuffer
        FontInfo.Size = 8
        FontInfo.Italic = False
        BackBuffer.SetFont FontInfo
         
         BackBuffer.SetForeColor vbWhite
         
            DDraw.ClearBuffer
            DisplaySprite Tela1, 30, 50, , True
            BackBuffer.DrawText 156, 195, ConnectStatus, False
            'DisplaySprite botao(0), 400, 300, , True
            'BackBuffer.DrawText 410, 310, "Conectar", False
EscrevaNaTela
            
            DDraw.Flip
        
        
        Case 6
            ''GameFps = Int(CDbl(1) / CDbl((Timer - TimerPassed)))
            If Timer - TimerPassed > 0.015 Or TimerPassed = 0 Then '66 Fps
                
                TimerPassed = Timer
                
                HandleKeys              'Check the DirectInput data for significant keypresses
            End If
        
        
        
        Case 8
        DDraw.ClearBuffer
        DisplaySprite cars(0, 0, 9, 1, 0), 1, 0, , True
        res = GetColorAtMousePos
        BackBuffer.SetForeColor res
        UnRGB res, r, g, b
        BackBuffer.DrawText 350, 400, "RGB( " & r & " " & g & " " & b & " ) ", False
EscrevaNaTela
        
        DDraw.Flip
        End Select
        DoEvents                'Give windows its chance to do things
        
    
    
        
    'verifica o mouse
        If GameStatus = 4 Then
            If MouseX > 480 And MouseX < 507 And MouseY > 80 And MouseY < 120 Then ShiftServerPage = ShiftServerPage + 1
            If MouseX > 480 And MouseX < 507 And MouseY > 260 And MouseY < 300 Then ShiftServerPage = ShiftServerPage - 1
            
            
        End If
    'processa teclas
'If DInput.aKeys(DIK_ESCAPE) Then End
If DInput.aKeys(DIK_ESCAPE) Then
    If GameStatus = 5 Then
    Winsock1.Close
    DataStream.Close
    StringStream.Close
    ConectadoAoServidor = False
     GameStatus = 4
    End If
End If



If GameStatus = 7 Then
If DInput.aKeys(DIK_DOWN) = False Then KeyPressedUp(DIK_DOWN) = False
If DInput.aKeys(DIK_DOWN) And KeyPressedUp(DIK_DOWN) = False Then
    Player(0).color = Player(0).color + 1
    'If Player(0).color > 7 Then Player(0).color = 7: Sound.WavPlay sons.naoaceito, EffectsVolume Else Sound.WavPlay sons.moving, EffectsVolume
    If Player(0).color > 254 Then
        Player(0).color = 254: Sound.WavPlay sons.naoaceito, EffectsVolume
    Else
        Sound.WavPlay sons.moving, EffectsVolume
        MoveDrawBoxTop = MoveDrawBoxTop + 1
        If MoveDrawBoxTop > 5 Then
        For a = 0 To 4
            LoadSprite Tinta(a), App.Path & "\graficos\logotipo\tinta.bmp", 66, 32, CLng(Preto)
            ChangeColors Tinta(a), RGB(255, 0, 0), RGB(CoresTinta(a + Player(0).color).vermelho, CoresTinta(a + Player(0).color).verde, CoresTinta(a + Player(0).color).azul), mp
            ChangeColors Tinta(a), RGB(0, 0, 255), Luminosidade(RGB(CoresTinta(a + Player(0).color).azul, CoresTinta(a + Player(0).color).verde, CoresTinta(a + Player(0).color).vermelho), 3), mp
            ChangeColors Tinta(a), RGB(0, 255, 0), Luminosidade(RGB(CoresTinta(a + Player(0).color).azul, CoresTinta(a + Player(0).color).verde, CoresTinta(a + Player(0).color).vermelho), 5), mp
        Next a
        MoveDrawBoxTop = 0
        End If
    End If
    'BackBuffer.DrawBox 90, 208 + MoveDrawBoxTop * 50, 152, 241
    'DDraw.Flip
    
    AlterarCorCarro (RGB(CoresTinta(Player(0).color).vermelho, CoresTinta(Player(0).color).verde, CoresTinta(Player(0).color).azul)), False
    KeyPressedUp(DIK_DOWN) = True
End If

    If DInput.aKeys(DIK_UP) = False Then KeyPressedUp(DIK_UP) = False
    If DInput.aKeys(DIK_UP) Then
        Player(0).color = Player(0).color - 1
        If Player(0).color < 0 Then
            Player(0).color = 0: Sound.WavPlay sons.naoaceito, EffectsVolume
        Else
            Sound.WavPlay sons.moving, EffectsVolume
            MoveDrawBoxTop = MoveDrawBoxTop - 1
        'If MoveDrawBoxTop < 0 Then MoveDrawBoxTop = 0
            If MoveDrawBoxTop <= 0 Then
                For a = 0 To 4
                    LoadSprite Tinta(a), App.Path & "\graficos\logotipo\tinta.bmp", 66, 32, CLng(Preto)
                    ChangeColors Tinta(a), RGB(255, 0, 0), RGB(CoresTinta(a - 4 + Player(0).color).vermelho, CoresTinta(a - 4 + Player(0).color).verde, CoresTinta(a - 4 + Player(0).color).azul), mp
                    ChangeColors Tinta(a), RGB(0, 0, 255), Luminosidade(RGB(CoresTinta(a - 4 + Player(0).color).azul, CoresTinta(a - 4 + Player(0).color).verde, CoresTinta(a - 4 + Player(0).color).vermelho), 3), mp
                    ChangeColors Tinta(a), RGB(0, 255, 0), Luminosidade(RGB(CoresTinta(a - 4 + Player(0).color).azul, CoresTinta(a - 4 + Player(0).color).verde, CoresTinta(a - 4 + Player(0).color).vermelho), 5), mp
                Next a
                MoveDrawBoxTop = 5
            End If
        End If
        'BackBuffer.DrawBox 90, 208 + MoveDrawBoxTop * 50, 152, 241
        'DDraw.Flip
        AlterarCorCarro (RGB(CoresTinta(Player(0).color).vermelho, CoresTinta(Player(0).color).verde, CoresTinta(Player(0).color).azul)), False
        KeyPressedUp(DIK_UP) = True
        
    End If
End If

If GameStatus = 3 Then
If DInput.aKeys(DIK_RIGHT) = False Then KeyPressedUp(DIK_RIGHT) = False
If DInput.aKeys(DIK_RIGHT) And KeyPressedUp(DIK_RIGHT) = False Then
    MoveDrawBoxLeft = MoveDrawBoxLeft + 64
    Sound.WavPlay sons.moving, EffectsVolume
    KeyPressedUp(DIK_RIGHT) = True
End If

    If DInput.aKeys(DIK_LEFT) = False Then KeyPressedUp(DIK_LEFT) = False
    If DInput.aKeys(DIK_LEFT) Then
        MoveDrawBoxLeft = MoveDrawBoxLeft - 64
        Sound.WavPlay sons.moving, EffectsVolume
        KeyPressedUp(DIK_LEFT) = True
    End If
End If


    If DInput.aKeys(DIK_RETURN) = False Then KeyPressedUp(DIK_RETURN) = False
    If KeyPressedUp(DIK_RETURN) = False And DInput.aKeys(DIK_RETURN) Then
        KeyPressedUp(DIK_RETURN) = True
    If GameStatus = 7 Then GameStatus = 3: Sound.WavPlay sons.aceito, EffectsVolume: AlterarCorCarro (RGB(CoresTinta(Player(0).color).vermelho, CoresTinta(Player(0).color).verde, CoresTinta(Player(0).color).azul)):  GoTo fim
    If GameStatus = 3 Then
      'ProcessGetList
              'mostrar servidores disponiveis
        
        If MoveDrawBoxLeft = 256 Then Running = False: Sound.WavPlay sons.aceito, EffectsVolume: Exit Do
        If MoveDrawBoxLeft = 0 Then
        Sound.WavPlay sons.aceito, EffectsVolume
        tmrProcessGetList.Interval = 1
         DDraw.ClearBuffer
        FontInfo.Size = 12
        FontInfo.Italic = True
        BackBuffer.SetFont FontInfo
         
         BackBuffer.SetForeColor vbWhite
         DisplaySprite Tela1, 30, 50, , True
         DisplaySprite Roll, 530, 50, , True
         DisplaySprite botao(0), 400, 300, , True
         BackBuffer.DrawText 410, 310, "Conectar", False
         
         
         DisplaySprite botao(1), 280, 300, , True
         BackBuffer.DrawText 290, 310, "Atualizar", False
         
         
         DisplaySprite botao(2), 160, 300, , True
         BackBuffer.DrawText 170, 310, "<<", False
         
          FontInfo.Size = 8
            
            FontInfo.Italic = False
            BackBuffer.SetFont FontInfo
         BackBuffer.DrawText 60, 57, "servidor", False
         BackBuffer.DrawText 330, 57, "players", False
         BackBuffer.DrawText 420, 57, "latencia", False
         
         BackBuffer.SetForeColor RGB(0, 740, 10)
         BackBuffer.DrawBox 50, 80, 480, 300
         DDraw.Flip
         RefreshingServer = True
         If RefreshingServer = True Then BackBuffer.DrawText 300, 60, "atualizando lista de servidores", False
         GameStatus = 4
         End If
         
         If MoveDrawBoxLeft = 192 Then
         Sound.WavPlay sons.aceito, EffectsVolume
         GameStatus = 7
         End If
         
         If MoveDrawBoxLeft <> 192 And MoveDrawBoxLeft <> 0 And MoveDrawBoxLeft <> 256 Then Sound.WavPlay sons.naoaceito, EffectsVolume
     'GameStatus = 4: GoTo fim
End If

If GameStatus = 2 Then
    Player(0).piloto = PlayerNumber
    Player(0).TopSpeed = 8
    Player(0).TopAcceleration = 14.285
    Player(0).TopCorner = 5  'quanto maior , pior a velocidade
    Player(0).TopJumping = -100

    If Player(0).piloto = CyberHawks Then
        Player(0).TopAcceleration = 16
        Player(0).TopJumping = -110
    End If

    If Player(0).piloto = JakeBlanders Then
        Player(0).TopAcceleration = 16
        Player(0).TopCorner = 2
    End If

    If Player(0).piloto = KatarinaLyons Then
        Player(0).TopJumping = -110
        Player(0).TopCorner = 2
    End If

    If Player(0).piloto = SnakeSanders Then
        Player(0).TopAcceleration = 16
        Player(0).TopSpeed = 9
    End If

    If Player(0).piloto = Tarquin Then
        Player(0).TopSpeed = 9
        Player(0).TopJumping = -110
    End If

    If Player(0).piloto = IvanZypher Then
        Player(0).TopAcceleration = 16
        Player(0).TopCorner = 2
    End If
        GameStatus = 7
            Sound.WavPlay sons.aceito, EffectsVolume
         
        End If
        If GameStatus = 1 Then GameStatus = 2: Sound.WavPlay sons.aceito, EffectsVolume
        KeyPressedUp(DIK_RETURN) = True
    End If
    
    If DInput.aKeys(DIK_LEFT) = False Then KeyPressedUp(DIK_LEFT) = False
    If KeyPressedUp(DIK_LEFT) = False And DInput.aKeys(DIK_LEFT) And GameStatus = 2 Then
        PlayerNumber = PlayerNumber - 1
        KeyPressedUp(DIK_LEFT) = True
        If PlayerNumber < 0 Then
            PlayerNumber = 0
            Sound.WavPlay sons.naoaceito, EffectsVolume
        Else
           Sound.WavPlay sons.moving, EffectsVolume
              processLeft = 100
            processRight = 0
            processPlanet = 200
        End If
    End If
    
    If DInput.aKeys(DIK_RIGHT) = False Then KeyPressedUp(DIK_RIGHT) = False
    If KeyPressedUp(DIK_RIGHT) = False And DInput.aKeys(DIK_RIGHT) And GameStatus = 2 Then
        PlayerNumber = PlayerNumber + 1
        KeyPressedUp(DIK_RIGHT) = True
        If PlayerNumber > 5 Then
            PlayerNumber = 5
            Sound.WavPlay sons.naoaceito, EffectsVolume
        Else
           Sound.WavPlay sons.moving, EffectsVolume
            processLeft = 0
            processRight = 100
            processPlanet = 200
        End If
    End If
fim:
    Loop
   
ErrOut:

    ExitProgram                 'If an error occurs, leave the program
    
End Sub

Public Sub ExitProgram()

    DDraw.Terminate             'Unload the DirectDraw variables
    DInput.Terminate            'Unload the DirectInput variables
    End                         'End the program
    
End Sub

Private Sub HandleKeys()

On Error Resume Next

Dim pneu As Long
Dim c As Long
Dim ext As Long
Dim s As Long
Dim r As Long
Dim tempObject As Objetos
Dim mydata As otherPlayersData
Dim buffer() As Byte
Dim x As Long
                    Dim sCont As Long
                    Dim res As Long
                    
Dim command As command
If UseBots = True Then DInput.aKeys(DIK_Z) = True: ProntoPraCorrer = True

''Sleep 1000 / SaltoReal
'primeiro desenha a sombra

''DInput.CheckKeys                                            'Get the current state of the keyboard
'If Camera.x < Player(0).position2D.x - (80 / ZoomX) Then
 '   Camera.x = Camera.x + 1
'End If
'If Camera.y < Player(0).position2D.y - ((250 / ZoomY) / ZoomY) Then
 '   Camera.y = Camera.y + 1
'End If

If DInput.aKeys(DIK_Y) And ConectadoAoServidor = True Then EscrevendoTexto = True
'If DInput.aKeys(DIK_C) And ConectadoAoServidor = False And EscrevendoTexto = False Then TimerConnect.Interval = 5000: Conectar = True
If DInput.aKeys(DIK_SPACE) Then Sair = False
If Sair = True Then
    
    If Timer - sairTimer >= 7 Then
        command.command = Quit
        res = 0
        Do
        res = res + 1
        If res > 100 Then Exit Do
        Loop Until send(Winsock1.SocketHandle, ByVal VarPtr(command), LenB(command.command), 0) <> SOCKET_ERROR
        Winsock1.Close
        DataStream.Close
        StringStream.Close
        ConectadoAoServidor = False
        If GameStatus = 5 Then GameStatus = 4
        GameStatus = 4          'Has the escape key been pressed?
        Sound.WavPlay sons.abouttoblow, LarryVolume
        Sair = False
         Player(0).VoltasDadas = 0
    Player(0).armas.Traseiras.quantidade = Player(0).armas.Traseiras.QuantasTenho
    Player(0).armas.Frontais.quantidade = Player(0).armas.Frontais.QuantasTenho
    GameStarted = False
    Player(0).velocidade = 0
    jump = 0
    Player(0).blow = 0
      Player(0).position.x = 0
      Player(0).position.y = 0
      'Player(0).position.Z = -20
                    naoelevarAgora = False
                    jump = 0
                    RampaSaltada = False
                    processealtura_caindo = False
                    processeAltura = False
                    OrigemdoPulo = 0
                    Player(0).CarroSaltouPelaRampa = False
                    Player(0).CarroNaRampadaFrente = False
                    Player(0).CarroNaRampadeTras = False
                    Player(0).AlturaReal = Player(0).position.Z
                    alturaAnterior = Player(0).AlturaReal
                    BotaoSaltoPressionado = False
                    AplicandoGravidade = False
                    'verifica se quando tocar o chao o carro explode
                    ExtendRampaCont = 0
                    ExtendRampa = False
                    LastRampa = 0
                    RetornadoAPista = False
                    ProcesseSuspensao = True
                    SubaDescaCarroAposSalto = False
                    For x = 0 To 2000
                        ObjectsFromNet(x).Active = False
                    Next x
                    For x = 0 To UBound(ArmasTraseirasNaPista)
                        ArmasTraseirasNaPista(x).Active = False
                    Next x
    
                    For x = 0 To UBound(ArmasFrontaisNaPista)
                        ArmasFrontaisNaPista(x).Active = False
                    Next x
    
                    For x = 0 To UBound(SprayPixadas)
                        SprayPixadas(x).Status = 0
                    Next x
                    Exit Sub
        
    End If
    DDraw.ClearBuffer
    BackBuffer.DrawText 0, 0, "saindo do servidor em " & Int(7 - (Timer - sairTimer)) & " segundos ", False
    BackBuffer.DrawText 0, 15, "pressione espaco para cancelar ", False
    DDraw.Flip
    Exit Sub
End If
If DInput.aKeys(DIK_ESCAPE) Then
    If Sair = False Then sairTimer = Timer
    Sair = True
End If
If DInput.aKeys(DIK_TAB) Then
    command.command = ShowTab
    res = 0
    If showingTab = False Then
        Do
            res = res + 1
            If res > 50 Then Exit Do 'falha
        Loop Until send(Winsock1.SocketHandle, ByVal VarPtr(command), GetSizeToSend(command.command), 0) <> SOCKET_ERROR
        'ReDim buffer(0 To GetSizeToSend(command.command) - 1)
        'CopyMemory buffer(0), command, GetSizeToSend(command.command)
        'Winsock1.SendData buffer
        
    End If
    showingTab = True
    
End If

'If DInput.aKeys(DIK_RETURN) And ConectadoAoServidor And SelectingPlayer Then
'player escolhido
'Player(0).piloto = PlayerNumber
'SelectingPlayer = False
'envia o jogador
'command.command = PlayAndCarReady
'command.parametro1 = Player(0).piloto
'padroes

'res = 0
'Do
'res = res + 1
'If res > 100 Then Exit Do 'falha
'Loop Until send(Winsock1.SocketHandle, ByVal VarPtr(command), GetSizeToSend(command.command), 0) <> SOCKET_ERROR

'End If

If Conectar = True Then
TimerConnect.Interval = 5000
'If CrashDerrapagem = True Then
'End If
'DDraw.ClearBuffer
BackBuffer.DrawText 26, 185, "pressione ESC pra sair", False
BackBuffer.DrawText 56, 205, ConnectStatus, False
'DDraw.Flip

'Exit Sub
End If


    'DInput.CheckKeys                                            'Get the current state of the keyboard
If CrashDerrapagem = True And FrmDirectX.PararDerrapagem.Interval = 0 Then
    derrapar = True
    If OtherTime = 0 Then FrmDirectX.PararDerrapagem.Interval = 2000 Else FrmDirectX.PararDerrapagem.Interval = OtherTime
End If
'AnguloDerrapagem = GetAngleFromCarImage(0)
'Player(0).velocidade = Player(0).velocidade + 0.5
    ''If DInput.aKeys(DIK_UP) Then Player(0).position.Z = Player(0).position.Z + 1
    ''If DInput.aKeys(DIK_DOWN) Then Player(0).position.Z = Player(0).position.Z - 1
    'If DInput.aKeys(DIK_1) Then xAxis = xAxis - 1
    'If DInput.aKeys(DIK_2) Then xAxis = xAxis + 1
    'If DInput.aKeys(DIK_3) Then yAxis = yAxis - 1
    'If DInput.aKeys(DIK_4) Then yAxis = yAxis + 1
    'If DInput.aKeys(DIK_5) Then zAxis = zAxis - 1
    'If DInput.aKeys(DIK_6) Then zAxis = zAxis + 1
    GetAngle = GetAngleFromCarImage(GetLastValidImageIndex)

    If RetornandoAPista = True Then
    tmrexplosaowait.Interval = 0
    tmrexplosao.Interval = 0
    SetStatus
    MoverCarro Player(0), -5, GetAngle
    'DrawPoligono 1
    DDraw.Flip              'Flip the backbuffer to the screen
        DDraw.ClearBuffer       'Erase the backbuffer so that it can be drawn on again
        
    Exit Sub
    End If
    If BotaoSaltoPressionado = True And CrashDerrapagem = False Then derrapar = False ': Sound.WavStop sons.derrapagem
    If Player(0).velocidade <= 4.285 Then LastRampaStatus = False: LastLadeiraStatus = False: Player(0).CarroSaltouPelaRampa = False
       'If Player(0).velocidade > 4 Then
       If LastRampaStatus = True And Player(0).CarroNaRampadaFrente = False Then Player(0).CarroSaltouPelaRampa = True: LastRampaStatus = False
       If LastLadeiraStatus = True And Player(0).CarroNaRampadeTras = False Then Player(0).CarroSaltouPelaRampa = True: LastLadeiraStatus = False
    
     If (Attack3 = True And GameStarted = True) Or (DInput.aKeys(DIK_A) = False) Then KeyPressedUp(DIK_A) = False
    Attack3 = False
    If processeAltura = False And ProntoPraCorrer = True And GameStarted = True And Player(0).CarroExplodiu = False And Player(0).CarroVaiExplodir = False And EscrevendoTexto = False And ConsoleVisible = False And DInput.aKeys(DIK_A) And KeyPressedUp(DIK_A) = False Or (Player(0).CarroSaltouPelaRampa = True And processeAltura = False) Then
    If processeAltura = False And processealtura_caindo = False And carrofora = False Then processeAltura = True: KeyPressedUp(DIK_A) = True: Sound.WavPlay sons.salto, EffectsVolume: mydata.JumpHit = 1
    If ProntoPraCorrer = True And (GameStarted = True And DInput.aKeys(DIK_A) And EscrevendoTexto = False And ConsoleVisible = False And carrofora = False) Then BotaoSaltoPressionado = True
    
    End If
    
    
    If processeAltura = True Or ProcesseSuspensao Then
        
        If processealtura_caindo = False Then
            If Player(0).CarroSaltouPelaRampa = False And OrigemdoPulo = 0 Then SetarOrigem
            ''If Player(0).CarroSaltouPelaRampa = True And OrigemdoPulo = 0 Then OrigemdoPulo = 3
                ''Player(0).position.altura = Player(0).position.altura + (0.7142 * SaltoReal)
                jump = jump - (500.857 / SaltoReal / 2)
                
            
            If Player(0).CarroSaltouPelaRampa = False Then
                If jump < -100 And ProcesseSuspensao = False Then processealtura_caindo = True
                If jump < -40 And ProcesseSuspensao = True Then processealtura_caindo = True: ProcesseSuspensao = False: naoelevarAgora = True: elevacao = 3
                
            End If
            
            
        Else
                If naoelevarAgora = False Then Player(0).velocidade = Player(0).velocidade - (0.015 / SaltoReal) Else Player(0).velocidade = Player(0).velocidade - (0.15 / SaltoReal)
                If Player(0).velocidade < 0 Then Player(0).velocidade = 0
                If SaltoReal <> 0 Then
                    jump = jump + (120 / SaltoReal)
                Else
                    jump = jump + (120 / 0.1)
                End If
                'Player(0).position.y = Player(0).position.y + (0.4285 * SaltoReal)
                
                If carrofora = False Then
                If Player(0).AlturaReal > Player(0).position.Z Then
                    If AplicandoGravidade And poligono(CarroEstaNoPoligono).tangente <> 0 Then Player(0).velocidade = Player(0).velocidade * 0.7
                    naoelevarAgora = False
                    jump = 0
                    RampaSaltada = False
                    processealtura_caindo = False
                    processeAltura = False
                    OrigemdoPulo = 0
                    Player(0).CarroSaltouPelaRampa = False
                    Player(0).CarroNaRampadaFrente = False
                    Player(0).CarroNaRampadeTras = False
                    Player(0).AlturaReal = Player(0).position.Z
                    alturaAnterior = Player(0).AlturaReal
                    BotaoSaltoPressionado = False
                    AplicandoGravidade = False
                    'verifica se quando tocar o chao o carro explode
                    ExtendRampaCont = 0
                    ExtendRampa = False
                    LastRampa = 0
                    RetornadoAPista = False
                    If SubaDescaCarroAposSalto Then ProcesseSuspensao = True:: Sound.WavPlay sons.cornercrash, EffectsVolume
                    SubaDescaCarroAposSalto = False
                End If
            Else
                If Player(0).AlturaReal > MenorNivel And Player(0).CarroExplodiu = False Then FrmDirectX.tmrexplosaowait.Interval = 1
            End If
            
            End If
                    If processeAltura = False Or Player(0).CarroSaltouPelaRampa = True Then
            ''If Player(0).CarroNaRampadaFrente = False And Player(0).CarroNaRampadeTras = False Then Player(0).position.angulo = GetAngleFromCarImage(Player(0).Car_Image_Index)
            ''If Player(0).CarroNaRampadaFrente = True Then Player(0).position.angulo = GetAngleFromCarImage(Player(0).Car_Image_Index) - 11
            ''If Player(0).CarroNaRampadeTras = True Then Player(0).position.angulo = GetAngleFromCarImage(Player(0).Car_Image_Index) + 16
        Else
            Player(0).position.angulo = GetAngleFromCarImage(Player(0).Car_Image_Index)
        End If
        If BotaoSaltoPressionado = True Then
            If processealtura_caindo = False Then elevacao = 1: Player(0).declive = subindo
        Else
  ''          If processealtura_caindo = False Then elevacao = 3
      ''  If processealtura_caindo = True And (jump < -35) Then elevacao = 0
            If ExtendRampa = False Then
                If naoelevarAgora = False And (processealtura_caindo = True And (jump >= -90 And jump <= -70)) Then elevacao = 3
                If naoelevarAgora = False And processealtura_caindo = True And (jump > -70) Then elevacao = 4
            End If
        End If
        If BotaoSaltoPressionado = True Then
            If naoelevarAgora = False And processealtura_caindo = True And (jump < -99) Then elevacao = 0: Player(0).declive = reto
            If naoelevarAgora = False And (processealtura_caindo = True And (jump >= -90 And jump <= -50)) Then elevacao = 3
            If naoelevarAgora = False And processealtura_caindo = True And (jump > -50) Then elevacao = 4
        Else
        End If
        
        
    ''Player(0).position.Z = jump
    
    End If
    

atualizar2D:
    
    Atualize2D
    
    
    If Player(0).AlturaReal < Player(0).position.Z Then processeAltura = True
    'If DInput.aKeys(DIK_ESCAPE) Then Running = False            'Has the escape key been pressed?
    If GameStarted = True And DInput.aKeys(DIK_DOWN) And EscrevendoTexto = False And ConsoleVisible = False And carrofora = False And processeAltura = False And ProntoPraCorrer = True Then
        If Player(0).velocidade > 0 Then
            Player(0).velocidade = Player(0).velocidade - (14.285 / SaltoReal)
        End If
        If Player(0).velocidade < 0 Then
            Player(0).velocidade = Player(0).velocidade + (14.285 / SaltoReal)
        End If
    End If
       '/RENDER
       
       'armas frente
       If DInput.aKeys(DIK_S) = False Then KeyPressedUp(DIK_S) = False
       
       If (Attack = True And GameStarted = True) Or (GameStarted = True And DInput.aKeys(DIK_S) And KeyPressedUp(DIK_S) = False And EscrevendoTexto = False) And ConsoleVisible = False And ProntoPraCorrer = True Then
       Attack = False
       If Player(0).armas.Frontais.quantidade <= 0 Then Sound.WavPlay sons.naoaceito, EffectsVolume
       If Player(0).armas.Frontais.quantidade > 0 Then
            mydata.laserHit = 1
            If Player(0).armas.Frontais.quantidade = 7 Then Sound.WavPlay sons.eLaser7, EffectsVolume
            If Player(0).armas.Frontais.quantidade = 6 Then Sound.WavPlay sons.eLaser6, EffectsVolume
            If Player(0).armas.Frontais.quantidade = 5 Then Sound.WavPlay sons.eLaser5, EffectsVolume
            If Player(0).armas.Frontais.quantidade = 4 Then Sound.WavPlay sons.eLaser4, EffectsVolume
            If Player(0).armas.Frontais.quantidade = 3 Then Sound.WavPlay sons.eLaser3, EffectsVolume
            If Player(0).armas.Frontais.quantidade = 2 Then Sound.WavPlay sons.eLaser2, EffectsVolume
            If Player(0).armas.Frontais.quantidade = 1 Then Sound.WavPlay sons.eLaser, EffectsVolume
            KeyPressedUp(DIK_S) = True
            ext = UBound(ArmasFrontaisNaPista)
            
            Select Case Player(0).armas.Frontais.tipo
            Case slaser
                ArmasFrontaisNaPista(ext).positionX = CLng(Player(0).position.x) + 60
                ArmasFrontaisNaPista(ext).positionY = CLng(Player(0).position.y)
                ArmasFrontaisNaPista(ext).positionZ = CLng(Player(0).AlturaReal)
                ArmasFrontaisNaPista(ext).VideoPos.x = CLng(Player(0).position2D.x)
                ArmasFrontaisNaPista(ext).VideoPos.y = CLng(Player(0).position2D.y)
                ArmasFrontaisNaPista(ext).tipo = Player(0).armas.Frontais.tipo
                ArmasFrontaisNaPista(ext).extra = Player(0).Car_Image_Index
                ArmasFrontaisNaPista(ext).Active = True
                ArmasFrontaisNaPista(ext).id = Player(0).id
                ArmasFrontaisNaPista(ext).PolignToUse = CreateObjectPolign(30, 50, Player(0).position.x, Player(0).position.y, slaser)
                Randomize Timer
                ArmasFrontaisNaPista(ext).handle = Int((65530 - 0) * Rnd) + 0
                ReDim Preserve ArmasFrontaisNaPista(0 To ext + 1)
                Player(0).armas.Frontais.quantidade = Player(0).armas.Frontais.quantidade - 1
                
                    If DataStream.Tag = "connected" And Player(0).receivedID = True Then
                        
                        tempObject = ArmasFrontaisNaPista(ext)
                        
                   GoTo naoatire:
                        
                        
                    End If

                    
            End Select
        End If
        End If
       
If DInput.aKeys(DIK_X) = False Then KeyPressedUp(DIK_X) = False
       If (Attack2 = True And GameStarted = True) Or (GameStarted = True And DInput.aKeys(DIK_X) And KeyPressedUp(DIK_X) = False And EscrevendoTexto = False) And ConsoleVisible = False And ProntoPraCorrer = True Then
       'solta arma traseira
'            Traseiras
       Attack2 = False
            KeyPressedUp(DIK_X) = True
            If Player(0).armas.Traseiras.quantidade <= 0 Then Sound.WavPlay sons.naoaceito, EffectsVolume
            If Player(0).armas.Traseiras.quantidade > 0 Then
            'solta a arma de tras
            mydata.oleoHit = 1
            ext = UBound(ArmasTraseirasNaPista)
            Select Case Player(0).armas.Traseiras.tipo
            Case sOil
                'nao solta no ar
                If Player(0).AlturaReal = Player(0).position.Z Then
                    If Player(0).position2D.x + 50 <> 0 And Player(0).position2D.y + 45 <> 0 Then
                        ArmasTraseirasNaPista(ext).VideoPos.x = CLng(Player(0).position2D.x) + 50
                        ArmasTraseirasNaPista(ext).VideoPos.y = CLng(Player(0).position2D.y) + 45
                        ArmasTraseirasNaPista(ext).positionX = Player(0).position.x
                        ArmasTraseirasNaPista(ext).positionY = Player(0).position.y
                        ArmasTraseirasNaPista(ext).positionZ = CLng(Player(0).AlturaReal)
                        ArmasTraseirasNaPista(ext).tipo = Player(0).armas.Traseiras.tipo
                        ArmasTraseirasNaPista(ext).Active = True
                        ArmasTraseirasNaPista(ext).PolignToUse = CreateObjectPolign(20, 40, Player(0).position.x, Player(0).position.y, sOil)
                        ArmasTraseirasNaPista(ext).id = Player(0).id
                        Randomize Timer
                        waitDikX = True: tmrwaitDikX.Interval = 1000
                        Do
                            s = Int(9 * Rnd) + 1
                        Loop Until s <> 0
                    
                        ArmasTraseirasNaPista(ext).extra = s
                        ReDim Preserve ArmasTraseirasNaPista(0 To ext + 1)
                        Player(0).armas.Traseiras.quantidade = Player(0).armas.Traseiras.quantidade - 1
                        'envia
                        If DataStream.Tag = "connected" And Player(0).receivedID = True Then
                            tempObject = ArmasTraseirasNaPista(ext)
                        End If
                        If Player(0).armas.Traseiras.quantidade = 7 Then Sound.WavPlay sons.eOleo7, EffectsVolume
                        If Player(0).armas.Traseiras.quantidade = 6 Then Sound.WavPlay sons.eOleo6, EffectsVolume
                        If Player(0).armas.Traseiras.quantidade = 5 Then Sound.WavPlay sons.eOleo5, EffectsVolume
                        If Player(0).armas.Traseiras.quantidade = 4 Then Sound.WavPlay sons.eOleo4, EffectsVolume
                        If Player(0).armas.Traseiras.quantidade = 3 Then Sound.WavPlay sons.eOleo3, EffectsVolume
                        If Player(0).armas.Traseiras.quantidade = 2 Then Sound.WavPlay sons.eOleo2, EffectsVolume
                        If Player(0).armas.Traseiras.quantidade = 1 Then Sound.WavPlay sons.eOleo, EffectsVolume
                    End If
                'DisplaySprite Oleo, Player(0).position2D.x + 50, Player(0).position.y + 45
                End If
            End Select
            
            End If
        End If
            
naoatire:
       If DInput.aKeys(DIK_N) Then ShowGeometry = Not ShowGeometry

       If ShowGeometry = True Then
            ShowGrides = True
            
        Else
            ShowGrides = False
        End If
       
       If GameStarted = True And DInput.aKeys(DIK_Z) And EscrevendoTexto = False And ConsoleVisible = False And carrofora = False And Player(0).CarroExplodiu = False And Player(0).CarroVaiExplodir = False And Player(0).CarroSeChocouQuinaExterna = False And Player(0).CarroSeChocouQuinaInterna = False And Player(0).AlturaReal >= Player(0).position.Z And ProntoPraCorrer = True Then
            If processeAltura = False Then
                '14.285= normal
                Player(0).velocidade = Player(0).velocidade + (Player(0).TopAcceleration / SaltoReal)
                If Player(0).velocidade > Player(0).TopSpeed Then Player(0).velocidade = Player(0).TopSpeed
                If Player(0).declive = subindo Then
                    Player(0).velocidade = Player(0).velocidade - ((Player(0).TopAcceleration - 1) / SaltoReal)
                    If Player(0).velocidade < -Player(0).TopSpeed Then Player(0).velocidade = -Player(0).TopSpeed
                End If
        
            If Player(0).declive = descendo Then
                Player(0).velocidade = Player(0).velocidade + (Player(0).TopAcceleration / SaltoReal)
                If Player(0).velocidade > Player(0).TopSpeed Then Player(0).velocidade = Player(0).TopSpeed
            End If
        End If
       Else
            If Player(0).declive = reto Then
                If Player(0).velocidade > 0 Then
                    If Player(0).AlturaReal >= Player(0).position.Z Then
                        Player(0).velocidade = Player(0).velocidade - (4.285 / SaltoReal)
                    Else
                        ''carro no salto nao perde velocidade demais
                        Player(0).velocidade = Player(0).velocidade - (0.54285 / SaltoReal)
                    End If
                    If Player(0).velocidade <= 0 And derrapar = False Then Player(0).velocidade = 0
                    
                End If
                If Player(0).velocidade < 0 Then
                    Player(0).velocidade = Player(0).velocidade + (14.285 / SaltoReal)
                    If Player(0).velocidade >= 0 And derrapar = False Then Player(0).velocidade = 0
                End If
            End If
            If Player(0).declive = subindo Then
            
               ''essa aqui
                If processeAltura = False Then
                        Player(0).velocidade = Player(0).velocidade - (11.42 / SaltoReal)
                     '   If Player(0).velocidade < -7 Then Player(0).velocidade = -7
                    If poligono(CarroEstaNoPoligono).piso = Rampa Or poligono(CarroEstaNoPoligono).piso = RampaH Then Player(0).velocidade = Player(0).velocidade - Abs((poligono(CarroEstaNoPoligono).tangente / SaltoReal))
                End If
            End If
        
           If Player(0).declive = descendo Then
             'If processeAltura = False Then
                'essa aqui no retorno
                If RetornadoAPista = False Then
                    If processeAltura = True Or ProcesseSuspensao = True Then
                        Player(0).velocidade = Player(0).velocidade - (0.785 / SaltoReal)
                    Else
                        Player(0).velocidade = Player(0).velocidade + Abs((11.42 / SaltoReal))
                    End If
                    'If Player(0).velocidade > 7 Then Player(0).velocidade = 7
                End If
             'End If
        End If
        
       End If
       
       
       
       'Player(0).CarroSaltouPelaRampa = False
       If processeAltura = False Then
        If Player(0).CarroNaRampadaFrente = True Then
            If (Player(0).Car_Image_Index > 18 And Player(0).Car_Image_Index <= 23) Or Player(0).Car_Image_Index = 0 Or Player(0).Car_Image_Index = 1 Or Player(0).Car_Image_Index = 2 Or Player(0).Car_Image_Index = 3 Or Player(0).Car_Image_Index = 4 Then LastRampaStatus = True
        End If
         If Player(0).CarroNaRampadeTras = True Then
            If Not ((Player(0).Car_Image_Index > 18 And Player(0).Car_Image_Index <= 23) Or Player(0).Car_Image_Index = 0 Or Player(0).Car_Image_Index = 1 Or Player(0).Car_Image_Index = 2 Or Player(0).Car_Image_Index = 3 Or Player(0).Car_Image_Index = 4) Then LastLadeiraStatus = True
        End If
       End If
       If processeAltura = False Then
            If Player(0).CarroNaRampadaFrente = False And Player(0).CarroNaRampadeTras = False Then Player(0).position.angulo = GetAngleFromCarImage(Player(0).Car_Image_Index)
            ''If Player(0).CarroNaRampadaFrente = True Then Player(0).position.angulo = GetAngleFromCarImage(Player(0).Car_Image_Index) - 11
            If Player(0).CarroNaRampadeTras = True Then Player(0).position.angulo = GetAngleFromCarImage(Player(0).Car_Image_Index) + 16
        Else
            Player(0).position.angulo = GetAngleFromCarImage(Player(0).Car_Image_Index)
        End If
        
              
            If derrapar = False Then
                MoverCarro Player(0), Player(0).velocidade, Player(0).position.angulo
            Else
                MoverCarro Player(0), Player(0).velocidade, AnguloDerrapagem
                'Player(0).velocidade = Player(0).velocidade * 0.95
                
            End If
            For x = 0 To UBound(SprayPixadas)
    If SprayPixadas(x).Status <> 0 Then
        DisplaySprite spray, SprayPixadas(x).position.x, SprayPixadas(x).position.y
    End If
Next x
        'If Player(0).position.angulo >= 0 And Player(0).position.angulo <= 45 Then pneu = (Player(0).position2D.x Mod 15) Mod 3
        'If Player(0).position.angulo > 45 And Player(0).position.angulo <= 90 Then pneu = (Player(0).position2D.y Mod 15) Mod 3
        'If Player(0).position.angulo > 90 And Player(0).position.angulo <= 135 Then pneu = (Player(0).position2D.x Mod 15) Mod 3
        'If Player(0).position.angulo > 135 And Player(0).position.angulo < 180 Then pneu = (Player(0).position2D.y Mod 15) Mod 3
        'If Player(0).position.angulo >= 180 And Player(0).position.angulo <= 225 Then pneu = (Player(0).position2D.x Mod 15) Mod 3
        'If Player(0).position.angulo > 225 And Player(0).position.angulo <= 270 Then pneu = (Player(0).position2D.y Mod 15) Mod 3
        'If Player(0).position.angulo > 270 And Player(0).position.angulo <= 315 Then pneu = (Player(0).position2D.x Mod 15) Mod 3
        'If Player(0).position.angulo > 315 And Player(0).position.angulo <= 359 Then pneu = (Player(0).position2D.y Mod 15) Mod 3
        pneu = (Player(0).position2D.x Mod 15) Mod 3
        pneu = Abs(pneu)
        'frmScreen.Caption = Player(0).position.angulo & "  " & pneu & " " & Player(0).position.y & " " & Player(0).position.y Mod 9
        'sprite Marauder.picMarauder(pneu * 24 + Player(0).Car_Image_Index), Marauder_Masks.picMarauder(Player(0).Car_Image_Index), 200, 350
         'DesenhePista chem_vi, 0, 100

'cars(0, Player(0).Car_Image_Index, 0, pneu).height = 72
        ''Atualize2D
        'If ExtendRampa = True Then elevacao = 1
    DesenheTodosObjetos
        ProcesseFumacaOthers = False
        DrawOtherPlayers
        
        If processealtura_caindo = True Then
            DDraw.DisplaySprite sombra, CLng(Player(0).position2D.x) + 30, CLng(Player(0).position2D.y) + 60
        End If

        If Player(0).CarroVaiExplodir = True Then elevacao = 4
        If Player(0).CarroExplodiu = False Then
            If Player(0).CarroNaRampadaFrente = True Then
                If processeAltura = False Then
                    If (Player(0).Car_Image_Index > 18 And Player(0).Car_Image_Index <= 23) Or Player(0).Car_Image_Index = 0 Or Player(0).Car_Image_Index = 1 Or Player(0).Car_Image_Index = 2 Or Player(0).Car_Image_Index = 3 Or Player(0).Car_Image_Index = 4 Then
                         If PodedesenharoCarro = True Then DDraw.DisplaySprite cars(0, 0, Player(0).Car_Image_Index, 1, pneu), CLng(Player(0).position2D.x), CLng(Player(0).position2D.y)    'Display the appropriate frame of the sprite
                    
                    Else
                        If PodedesenharoCarro = True Then DDraw.DisplaySprite cars(0, 0, Player(0).Car_Image_Index, 3, pneu), CLng(Player(0).position2D.x), CLng(Player(0).position2D.y)
                    'DDraw.DisplaySprite sombra, CLng(Player(0).position.x) + 30, CLng(Player(0).position.y) + 60
                    End If
                Else
                'DDraw.DisplaySprite cars(player(0).color, Player(0).Car_Image_Index, elevacao, pneu), CLng(Player(0).position.x), CLng(Player(0).position.y) - Player(0).position.altura      'Display the appropriate frame of the sprite
                    If PodedesenharoCarro = True Then DDraw.DisplaySprite cars(0, 0, Player(0).Car_Image_Index, elevacao, pneu), CLng(Player(0).position2D.x), CLng(Player(0).position2D.y)
                'DDraw.DisplaySprite sombra, CLng(Player(0).position.x) + 30, CLng(Player(0).position.y) + 30
                End If
            Else
                If Player(0).CarroNaRampadeTras <> True And PodedesenharoCarro = True Then DDraw.DisplaySprite cars(0, 0, Player(0).Car_Image_Index, elevacao, pneu), CLng(Player(0).position2D.x), CLng(Player(0).position2D.y)
            End If
        
           If Player(0).CarroNaRampadeTras = True Then
            If processeAltura = False Then
                If (Player(0).Car_Image_Index > 18 And Player(0).Car_Image_Index <= 23) Or Player(0).Car_Image_Index = 0 Or Player(0).Car_Image_Index = 1 Or Player(0).Car_Image_Index = 2 Or Player(0).Car_Image_Index = 3 Or Player(0).Car_Image_Index = 4 Then
                    If PodedesenharoCarro = True Then DDraw.DisplaySprite cars(0, 0, Player(0).Car_Image_Index, 3, pneu), CLng(Player(0).position2D.x), CLng(Player(0).position2D.y)
                    
                Else
                    If PodedesenharoCarro = True Then DDraw.DisplaySprite cars(0, 0, Player(0).Car_Image_Index, 1, pneu), CLng(Player(0).position2D.x), CLng(Player(0).position2D.y)
                    
                End If
            Else
                If PodedesenharoCarro = True Then DDraw.DisplaySprite cars(0, 0, Player(0).Car_Image_Index, elevacao, pneu), CLng(Player(0).position2D.x), CLng(Player(0).position2D.y)
                'DDraw.DisplaySprite sombra, CLng(Player(0).position.x) + 30, CLng(Player(0).position.y) + 60
            End If
          Else
            If PodedesenharoCarro = True Then If Player(0).CarroNaRampadaFrente <> True Then DDraw.DisplaySprite cars(0, 0, Player(0).Car_Image_Index, elevacao, pneu), CLng(Player(0).position2D.x), CLng(Player(0).position2D.y)
            
          End If
        End If
        'Dim cadaponto As Long
        'BackBuffer.SetForeColor RGB(0, 0, 0)
        'For cadaponto = 0 To 3
         
        'Next cadaponto
        'BackBuffer.SetForeColor RGB(255, 255, 255)
        If Player(0).CarroExplodiu = True Then
            Player(0).velocidade = 0
            If Int(contexplosao) > 0 Then DDraw.DisplaySprite explosao(Int(contexplosao)), CLng(Player(0).position2D.x) - 250, CLng(Player(0).position2D.y) - 175
            tmrexplosao.Interval = 100
        End If
        
            
'sprays

'processa blow
If Player(0).blow > 0 Then
    CriarFumacaDerrapagem Player(0).blow
    ProcesseFumaca = True
Else
    ProcesseFumaca = False
End If
If (ProcesseFumaca = True Or ProcesseFumacaOthers = True) And UseBots = False Then
    
    For x = 0 To 1000
    If FumacaPos(x).Status <> 0 Then
        DisplaySprite Fumaca(Int(FumacaPos(x).Status - 1)), CLng(FumacaPos(x).position.x), CLng(FumacaPos(x).position.y)
        FumacaPos(x).Status = FumacaPos(x).Status - 0.2
        FumacaPos(x).position.y = FumacaPos(x).position.y - 1
        
        If FumacaPos(x).Status <= 0 Then
            FumacaPos(x).Status = 0
            FumacaPos(x).position.x = 0
            FumacaPos(x).position.y = 0
        End If
    End If
    Next x
    
    'verifica se não há mais fumaca a ser desenhada
    'For x = 0 To 299
     '   If FumacaPos(x).Status <> 0 Then GoTo aindahafumaca:
    'Next x
    'ProcesseFumaca = False
'aindahafumaca:
End If
   
If ConectadoAoServidor = False And Conectar = False Then
'BackBuffer.DrawText 0, 15, "Pressione <ESC> pra sair", False
'BackBuffer.DrawText 0, 25, "Pressione C para se conectar com o servidor de teste (bora jogar multiplayer!)", False
'BackBuffer.DrawText 0, 36, "'rock n' roll racing' reprogramming (09/03/2011)  by Snes Fan Remix ", False
'BackBuffer.DrawText 0, 50, "procure pela comunidade SNES Fan Remix no orkut, e participe", False
Else
'If ConectadoAoServidor = True Then BackBuffer.DrawText 0, 65, "conectado no servidor", False
End If
'BackBuffer.DrawText 350, 312, "velocidade: " & Player(0).velocidade, False
'BackBuffer.DrawText 350, 320, "FPS: " & FPS, False
'BackBuffer.DrawText 250, 352, "diferente da versao anterior , que usava deteccao por pixel ", False
'BackBuffer.DrawText 250, 362, "esta versao usa deteccao por poligono ", False
''BackBuffer.DrawText 350, 332, "velocidade real: " & SaltoReal, False
''BackBuffer.DrawText 0, 0, Player(0).position2D.x & "   " & Player(0).position2D.y, False
''BackBuffer.DrawText 30, 20, Player(0).position.x & "   " & Player(0).position.y, False
'DisplaySprite fumaca(3), CLng(10), CLng(30)
'For x = 0 To 3
'BackBuffer.DrawText Player(0).ChockPoints(x).x - Camera.x, Player(0).ChockPoints(x).y - Camera.y, x, False
'Next x

DrawAllPoligonos CLng(Player(0).position.x), CLng(Player(0).position.y)
If DInput.aKeys(DIK_TAB) = False Then
    If showingTab = True Then
        showingTab = False
        For x = 1 To 20
            TabScreen(x).name = Empty
            TabScreen(x).Ping = 0
            TabScreen(x).pontos = 0
            TabScreen(x).deads = 0
            TabScreen(x).kills = 0
        Next x
    End If
End If

If DInput.aKeys(DIK_T) = False Then KeyPressedUp(DIK_T) = False
If DInput.aKeys(DIK_T) = True And KeyPressedUp(DIK_T) = False And ProntoPraCorrer = True Then
     KeyPressedUp(DIK_T) = True
    ext = UBound(SprayPixadas)
    ReDim Preserve SprayPixadas(0 To ext + 1)
    SprayPixadas(ext).position.x = Player(0).position2D.x
    SprayPixadas(ext).position.y = Player(0).position2D.y
    SprayPixadas(ext).Status = 1
    
End If
If DInput.aKeys(DIK_I) Then showInfo = Not (showInfo)
If showInfo Then
BackBuffer.DrawText 30, (0 * 12) + 20, "Altura: " & Player(0).AlturaReal, False
            BackBuffer.DrawText 30, (1 * 12) + 20, "tiros: " & Player(0).armas.Frontais.QuantasTenho, False
            BackBuffer.DrawText 30, (2 * 12) + 20, "explodiu? " & Player(0).CarroExplodiu, False
            BackBuffer.DrawText 30, (3 * 12) + 20, "Todo na Pista " & Player(0).CarroTodoNaPista, False
            BackBuffer.DrawText 30, (4 * 12) + 20, "Vai explodir " & Player(0).CarroVaiExplodir, False
            BackBuffer.DrawText 30, (6 * 12) + 20, "posicao do carro 3D (na memoria) x,y,z , anguloº  " & Int(Player(0).position.x) & " , " & Int(Player(0).position.y) & " , " & Int(Player(0).position.Z) & " , " & Int(Player(0).position.angulo) & "º", False
            BackBuffer.DrawText 30, (7 * 12) + 20, "Posicao do carro 2D (na tela) x,y  " & Int(Player(0).position2D.x - Camera.x + CamLeft + ControlCameraX) & " , " & Int(Player(0).position2D.y - Camera.y + CamUp + ControlCameraY), False
            BackBuffer.DrawText 30, (9 * 12) + 20, "Velocidade " & Int(Player(0).velocidade), False
            
End If
If SelectingPlayer = True Then SelectPlayer



        DDraw.Flip              'Flip the backbuffer to the screen
        DDraw.ClearBuffer       'Erase the backbuffer so that it can be drawn on again
        SetStatus
'envia os dados para o servidor

If DataStream.Tag = "connected" And Player(0).receivedID = True And PodedesenharoCarro = True Then
mydata.commando = dataCarStream
mydata.Car = Player(0).Car
mydata.Car_Image_Index = Player(0).Car_Image_Index
mydata.elevacao = elevacao
mydata.id = Player(0).id
mydata.positionX = Player(0).position.x
mydata.positionY = Player(0).position.y
mydata.positionZ = Player(0).AlturaReal

mydata.ShowSombra = CByte(processealtura_caindo)
mydata.velocidade = Player(0).velocidade
mydata.commando = dataCarStream
mydata.CarroExplosao = contexplosao
mydata.NewObject = tempObject
mydata.color = RGB(CoresTinta(Player(0).color).vermelho, CoresTinta(Player(0).color).verde, CoresTinta(Player(0).color).azul)
'Player(0).blow = 3
mydata.blow = Player(0).blow
If UseBots = True Then mydata.color = RGB(Int((255 - 0) * Rnd) + 0, Int((255 - 0) * Rnd) + 0, Int((255 - 0) * Rnd) + 0)

If CrashDerrapagem = False Then mydata.FumacaDerrapagem = CByte(derrapar)


 'mydata.blow = 10
  'Dim ps As otherPlayersData
  'CopyMemory ps, mydata, LenB(mydata)
  
  'Form2.Label4 = "blow = " & ps.blow
    'ReDim buffer(0 To LenB(mydata) - 1)
    'CopyMemory buffer(0), ByVal VarPtr(mydata), Len(mydata)
    'DataStream.SendData buffer
    res = 0
    Dim e As Long
    Dim res2 As Long
   

    Do
    res = res + 1
    If res > 25 Then Exit Do
    res2 = send(DataStream.SocketHandle, mydata, LenB(mydata), 0)
    Loop Until res2 <> SOCKET_ERROR
       ' ReDim b(0 To LenB(mydata) - 1)
      '  CopyMemory b(0), mydata, LenB(mydata)
     '   DataStream.SendData b
    'Form2.Caption = MyName
    'If res2 = SOCKET_ERROR Then DataStream.Connect server.Address, 20778: SaidText(1) = "falha em datastream"
       'res = send(DataStream.SocketHandle, ByVal VarPtr(mydata), Len(mydata), 0)
If UseBots = True Then Form2.List2.AddItem res
If res2 = 10050 Or res2 = 10051 Or res2 = 10052 Or res2 = 10053 Or res2 = 10054 Or res2 = 10057 Or res2 = 10058 Or res2 = 10060 Or res2 = 10061 Or res2 = 10064 Or res2 = 10065 Then
    DataStream.Tag = "error"
    DataStream.Close
    DataStream.Connect server.Address, 20778
    SaidText(1) = "falha fluxo de dados"
End If


End If
    'End If
If UseBots = True Then Form2.Label1 = res2
If UseBots = True Then Form2.Label2 = mydata.positionX & "  " & mydata.positionY
        If CrashDerrapagem = True Or (GameStarted = True And derrapar = True And DInput.aKeys(DIK_Z) And EscrevendoTexto = False And ConsoleVisible = False) And ProntoPraCorrer = True Then
            Player(0).velocidade = Player(0).velocidade - (16.285 / SaltoReal)
            If CrashDerrapagem = False And Player(0).velocidade < 0 Then Player(0).velocidade = 0: PararDerrapagem.Interval = 100
            Sound.WavPlay sons.derrapagem, EffectsVolume
        End If
       If derrapar = True And DInput.aKeys(DIK_A) And GameStarted = True And ProntoPraCorrer = True Then
            Player(0).velocidade = Player(0).velocidade - (0.285 / SaltoReal)
            If Player(0).velocidade < 0 Then Player(0).velocidade = 0: PararDerrapagem.Interval = 100
            'Sound.WavPlay sons.derrapagem
        End If
       
       If DInput.aKeys(DIK_LEFT) = False Then KeyPressedUp(DIK_LEFT) = False
       If DInput.aKeys(DIK_LEFT) And SelectingPlayer = True Then
        If KeyPressedUp(DIK_LEFT) = False Then
            PlayerNumber = PlayerNumber - 1
            If PlayerNumber < 0 Then PlayerNumber = 5
            KeyPressedUp(DIK_LEFT) = True
         End If
        End If
       
       If DInput.aKeys(DIK_RIGHT) = False Then KeyPressedUp(DIK_RIGHT) = False
       If DInput.aKeys(DIK_RIGHT) And SelectingPlayer = True Then
        If KeyPressedUp(DIK_RIGHT) = False Then
            PlayerNumber = PlayerNumber + 1
            If PlayerNumber > 5 Then PlayerNumber = 0
            KeyPressedUp(DIK_RIGHT) = True
         End If
        End If
       
       If GameStarted = True And SelectingPlayer = False And DInput.aKeys(DIK_LEFT) And Player(0).velocidade <> 0 And EscrevendoTexto = False And ConsoleVisible = False And carrofora = False And RetornandoAPista = False And Player(0).CarroExplodiu = False And Player(0).CarroVaiExplodir = False And Player(0).AlturaReal >= Player(0).position.Z And ProntoPraCorrer = True Then                              'If the left arrow is pressed, rotate the ship left
        
          If EspereSetas = False Then
            If derrapar = False And CrashDerrapagem = False Then AnguloDerrapagem = Player(0).position.angulo - (10 / SaltoReal)
            'If AnguloDerrapagem <= -36 Then AnguloDerrapagem = -36
             If processeAltura = False And derrapar = False Then Player(0).Car_Image_Index = Player(0).Car_Image_Index - 1
             If processeAltura = False And derrapar = True And dik_left_time_press Mod 100 = 0 Then Player(0).Car_Image_Index = Player(0).Car_Image_Index - 1: Sound.WavPlay sons.derrapagem, EffectsVolume
             
            If Player(0).Car_Image_Index < 0 Then Player(0).Car_Image_Index = 23
            If dik_left_time_press >= 100 And DontProcessDerrapar = False Then derrapar = True: CriarFumacaDerrapagem 3: ProcesseFumaca = True: Sound.WavPlay sons.derrapagem, EffectsVolume
            dik_left_time_press = dik_left_time_press + 50
            EspereSetas = True: tmrSetas.Interval = 50
          End If
          If Player(0).velocidade <= 4.285 And CrashDerrapagem = False Then PararDerrapagem.Interval = 100
          Exit Sub
        End If
    
    If GameStarted = True And DInput.aKeys(DIK_RIGHT) = False Then KeyPressedUp(DIK_RIGHT) = False
    If SelectingPlayer = False And DInput.aKeys(DIK_RIGHT) And Player(0).velocidade <> 0 And EscrevendoTexto = False And ConsoleVisible = False And carrofora = False And RetornandoAPista = False And Player(0).CarroExplodiu = False And Player(0).CarroVaiExplodir = False And Player(0).AlturaReal >= Player(0).position.Z And ProntoPraCorrer = True Then                                   'If the right arrow is pressed, rotate the ship right
      If EspereSetas = False Then
         If derrapar = False And CrashDerrapagem = False Then AnguloDerrapagem = Player(0).position.angulo + (10 / SaltoReal)
         'If AnguloDerrapagem >= 36 Then AnguloDerrapagem = 36
         If processeAltura = False And derrapar = False Then Player(0).Car_Image_Index = Player(0).Car_Image_Index + 1
        If processeAltura = False And derrapar = True And dik_right_time_press Mod 100 = 0 Then Player(0).Car_Image_Index = Player(0).Car_Image_Index + 1: Sound.WavPlay sons.derrapagem, EffectsVolume
        If Player(0).Car_Image_Index > 23 Then Player(0).Car_Image_Index = 0
        If dik_right_time_press >= 100 And DontProcessDerrapar = False Then derrapar = True: CriarFumacaDerrapagem 3: ProcesseFumaca = True: Sound.WavPlay sons.derrapagem, EffectsVolume
        dik_right_time_press = dik_right_time_press + 50
        EspereSetas = True: tmrSetas.Interval = 50
     End If
     If Player(0).velocidade <= 4.285 And CrashDerrapagem = False Then PararDerrapagem.Interval = 100
     Exit Sub
    End If
 
 If CrashDerrapagem = True Then Exit Sub
If Not (DInput.aKeys(DIK_LEFT)) Then dik_left_time_press = 0: PararDerrapagem.Interval = 100
If Not (DInput.aKeys(DIK_RIGHT)) Then dik_right_time_press = 0: PararDerrapagem.Interval = 100
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim res As Long
If CameraSeguirOutroPlayer <> 0 Then
If Button = 1 Then
CameraSeguirOutroPlayer = CameraSeguirOutroPlayer + 1
If CameraSeguirOutroPlayer > server.PlayersIn Then CameraSeguirOutroPlayer = 1
Do
    If OtherPlayers(CameraSeguirOutroPlayer).Data.id <> 0 Then Exit Sub
    res = res + 1
    If res > 101 Then Exit Sub
    CameraSeguirOutroPlayer = CameraSeguirOutroPlayer + 1
    If CameraSeguirOutroPlayer > server.PlayersIn Then CameraSeguirOutroPlayer = 1
    DoEvents
Loop
End If

If Button = 2 Then
CameraSeguirOutroPlayer = CameraSeguirOutroPlayer - 1
If CameraSeguirOutroPlayer < 1 Then CameraSeguirOutroPlayer = server.PlayersIn
If OtherPlayers(CameraSeguirOutroPlayer).Data.id <> 0 Then Exit Sub
    res = res + 1
    If res > 101 Then Exit Sub
    
    CameraSeguirOutroPlayer = CameraSeguirOutroPlayer - 1
    If CameraSeguirOutroPlayer < 1 Then CameraSeguirOutroPlayer = server.PlayersIn
    DoEvents
End If
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim x As Long
  
  Sound.Terminate
    Set Sound = Nothing

    
End Sub



Private Sub GetServerInfo_DataArrival(index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim s As String
Dim sB() As String
GetServerInfo(index).GetData s
's = StrConv(s, vbUnicode)
'MsgBox s
sB = Split(s, "*")
serverInfo(index).PlayersInInfo = sB(1)
serverInfo(index).name = Trim(sB(0))

'ServerStartedTime = Timer
If s <> Empty Then serverInfo(index).Ping = CStr(CLng((Timer - ServerStartedTime(index)) * 1000))
GetServerInfo(index).Close
End Sub


Private Sub GetServerInfo_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'    serverInfo(Index).PlayersInInfo = "OFF"

End Sub

Private Sub naoderrapar_Timer()
DontProcessDerrapar = False
End Sub

Private Sub PararDerrapagem_Timer()
derrapar = False
Sound.WavStop sons.derrapagem
OtherTime = 0
If CrashDerrapagem = True Then AnguloDerrapagem = 0
CrashDerrapagem = False
PararDerrapagem.Interval = 0
DontProcessDerrapar = True
naoderrapar.Interval = 100
End Sub



Private Sub StringStream_Connect()
StringStream.Tag = "connected"
End Sub

Private Sub StringStream_DataArrival(ByVal bytesTotal As Long)
If cl_message = False Then Exit Sub
Dim str As String
Dim x As Long
Dim command As String * 1
StringStream.GetData str
'str = StrConv(str, vbUnicode)
'SaidText(3) = str
command = Left(str, 1)
str = Right(str, Len(str) - 1)

Select Case command
Case "N"
MyName = str
SaidText(1) = "name = " & MyName

Case "J"
ShowJoinedPlayer str

Case "A"
ShowAtack str

Case "Y"
    For x = 1 To 5
        If x = 5 And SaidText(5) <> Empty Then
        SaidText(5) = SaidText(4)
        SaidText(4) = SaidText(3)
        SaidText(3) = SaidText(2)
        SaidText(2) = SaidText(1)
        SaidText(1) = str
        Exit Sub
        End If
        
        If SaidText(x) = Empty Then SaidText(x) = str: Exit Sub
        'SaidText(1) = Str
        
        
        Exit Sub
    Next x
    
End Select

End Sub

Private Sub StringStream_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
If Number = 10050 Or Number = 10051 Or Number = 10052 Or Number = 10053 Or Number = 10054 Or Number = 10057 Or Number = 10058 Or Number = 10060 Or Number = 10061 Or Number = 10064 Or Number = 10065 Then
StringStream.Tag = "error"
StringStream.Close
StringStream.Connect server.Address, 20780
End If
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub TimerConnect_Timer()
On Error Resume Next
' OtherPlayers(1).data.positionX = 0 / ZoomX
 '   OtherPlayers(1).data.positionY = 0 / ZoomY
  '  OtherPlayers(1).data.positionZ = 0
   ' OtherPlayers(1).data.Car = Maraudercar
    'OtherPlayers(1).data.Car_Image_Index = 21
    'OtherPlayers(1).data.elevacao = 0
    'OtherPlayers(1).data.ID = 10
 '   Dim command As command
'command.command = SendHelloServer
'If UseBots = False Then
    If UseBots = True Then serverInfo(SelectedList).noIp = BotConnectAt
    If UseBots = True Then server.Address = BotConnectAt
    If serverInfo(SelectedList).noIp <> Empty Then
        server.Address = serverInfo(SelectedList).noIp
        If UseBots = True Then server.Address = BotConnectAt
        Winsock1.Connect server.Address, 20777
        DataStream.Connect server.Address, 20778
        StringStream.Connect server.Address, 20780
    Else
    GameStatus = 4
    
    End If
TimerConnect.Interval = 0

'Else
 '   server.Address = BotConnectAt
  '  Winsock1.Connect server.Address, 20777
   ' DataStream.Connect server.Address, 20778
   ' StringStream.Connect server.Address, 20780
   ' TimerConnect.Interval = 0

'End If

End Sub


Private Sub tmrAlterarCorCarro_Timer(index As Integer)
    AlterarCorCarro Val(tmrAlterarCorCarro(index).Tag), True, CLng(index)
    tmrAlterarCorCarro(index).Interval = 0
End Sub

Private Sub tmrattackBonus_Timer()
AttackString = Empty
tmrattackBonus.Interval = 0
End Sub

Private Sub TmrChangeColor_Timer()
 Dim lockrect As RECT
Dim SrcDesc         As DDSURFACEDESC2
Dim pixel(0 To 3) As Long
Dim x As Long
Dim altura As Long
Dim largura As Long
    
    
    
    ImagemAmudar(ImagemdaVez).Sprite.imagem.GetSurfaceDesc SrcDesc
    lockrect.Right = SrcDesc.lWidth
    lockrect.Bottom = SrcDesc.lHeight
    ImagemAmudar(ImagemdaVez).Sprite.imagem.Lock lockrect, SrcDesc, DDLOCK_WAIT Or DDLOCK_NOSYSLOCK, 0
    
    Dim u As Long
    altura = ImagemAmudar(ImagemdaVez).AlturaAtual
    largura = ImagemAmudar(ImagemdaVez).LarguraAtual
        'For altura = 0 To ImagemAmudar.Height - 1
        'ImagemAmudar.imagem.Lock lockrect, SrcDesc, DDLOCK_WAIT Or DDLOCK_NOSYSLOCK, 0
     '   For largura = 0 To ImagemAmudar.Width - 1
            'DoEvents
            
            u = ImagemAmudar(ImagemdaVez).Sprite.imagem.GetLockedPixel(largura, altura)
            If u = ImagemAmudar(ImagemdaVez).corSource Then ImagemAmudar(ImagemdaVez).Sprite.imagem.SetLockedPixel largura, altura, CorDest
            'ImagemAmudar.imagem.SetLockedPixel largura, altura, 234565
      '  Next largura
        'ImagemAmudar.imagem.Unlock lockrect
    'Next altura
    ImagemAmudar(ImagemdaVez).LarguraAtual = ImagemAmudar(ImagemdaVez).LarguraAtual + 1
    
    If ImagemAmudar(ImagemdaVez).LarguraAtual > ImagemAmudar(ImagemdaVez).Sprite.Height Then
        ImagemAmudar(ImagemdaVez).LarguraAtual = 0
        ImagemAmudar(ImagemdaVez).AlturaAtual = ImagemAmudar(ImagemdaVez).AlturaAtual + 1
        If ImagemAmudar(ImagemdaVez).AlturaAtual > ImagemAmudar(ImagemdaVez).Sprite.Width Then
            ImagemAmudar(ImagemdaVez).id = 0
            ImagemAmudar(ImagemdaVez).Sprite.imagem.Unlock lockrect
            ImagemdaVez = ProxImagem
            Exit Sub
        End If
    End If
    ImagemAmudar(ImagemdaVez).Sprite.imagem.Unlock lockrect
    ImagemdaVez = ProxImagem
    
'TmrChangeColor.Interval = 0
End Sub

Private Sub tmrConnectError_Timer()
  If GameStatus = 5 Then
        GameStatus = 4
  End If
    tmrConnectError.Interval = 0
End Sub

Private Sub tmrControlCamera_Timer()
If FixeCam = True Then Exit Sub
Select Case Player(0).Car_Image_Index
Case 3
If ControlCameraY > 0 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < 0 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 0 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 0 Then ControlCameraX = ControlCameraX - 1

Case 4
If ControlCameraY > 30 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < 30 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 30 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 30 Then ControlCameraX = ControlCameraX - 1

Case 5
If ControlCameraY > 60 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < 60 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 60 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 60 Then ControlCameraX = ControlCameraX - 1

Case 6
If ControlCameraY > 90 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < 90 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 90 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 90 Then ControlCameraX = ControlCameraX - 1

Case 7
If ControlCameraY > 120 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < 120 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 120 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 120 Then ControlCameraX = ControlCameraX - 1

Case 8
If ControlCameraY > 150 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < 150 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 150 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 150 Then ControlCameraX = ControlCameraX - 1

Case 9
If ControlCameraY > 180 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < 180 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 180 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 180 Then ControlCameraX = ControlCameraX - 1

Case 10
If ControlCameraY > 150 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < 150 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 210 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 210 Then ControlCameraX = ControlCameraX - 1

Case 11
If ControlCameraY > 120 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < 120 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 240 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 240 Then ControlCameraX = ControlCameraX - 1

Case 12
If ControlCameraY > 90 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < 90 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 270 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 270 Then ControlCameraX = ControlCameraX - 1

Case 13
If ControlCameraY > 60 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < 60 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 300 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 300 Then ControlCameraX = ControlCameraX - 1

Case 14
If ControlCameraY > 30 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < 30 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 330 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 330 Then ControlCameraX = ControlCameraX - 1

Case 15
If ControlCameraY > 0 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < 0 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 360 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 360 Then ControlCameraX = ControlCameraX - 1

Case 16
If ControlCameraY > -30 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < -30 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 170 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 170 Then ControlCameraX = ControlCameraX - 1

Case 17
If ControlCameraY > -60 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < -60 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 140 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 140 Then ControlCameraX = ControlCameraX - 1

Case 18
If ControlCameraY > -90 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < -90 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 110 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 110 Then ControlCameraX = ControlCameraX - 1

Case 19
If ControlCameraY > -120 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < -120 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 80 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 80 Then ControlCameraX = ControlCameraX - 1

Case 20
If ControlCameraY > -150 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < -150 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 50 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 50 Then ControlCameraX = ControlCameraX - 1

Case 21
If ControlCameraY > -150 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < -150 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 20 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 20 Then ControlCameraX = ControlCameraX - 1

Case 22
If ControlCameraY > -120 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < -120 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 120 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 120 Then ControlCameraX = ControlCameraX - 1

Case 23
If ControlCameraY > -90 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < -90 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 120 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 120 Then ControlCameraX = ControlCameraX - 1

Case 0
If ControlCameraY > -60 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < -60 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 90 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 90 Then ControlCameraX = ControlCameraX - 1

Case 1
If ControlCameraY > -30 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < -30 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 60 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 60 Then ControlCameraX = ControlCameraX - 1

Case 2
If ControlCameraY > 0 Then ControlCameraY = ControlCameraY - 1
If ControlCameraY < 0 Then ControlCameraY = ControlCameraY + 1
If ControlCameraX < 30 Then ControlCameraX = ControlCameraX + 1
If ControlCameraX > 30 Then ControlCameraX = ControlCameraX - 1

End Select
End Sub

Private Sub tmrcorridacompletada_Timer()
 
 
 
 If ProntoPraCorrer = False And GameStarted = False And server.PlayersIn > 0 Then CameraSeguirOutroPlayer = Int((server.PlayersIn - 1) * Rnd) + 1
   If GameStarted = False Then PodedesenharoCarro = False
 tmrcorridacompletada.Interval = 0
  'cl_gaitestimation = ex_gaitestimation
End Sub

Private Sub tmrExplosao_Timer()
contexplosao = contexplosao + 1

            If contexplosao > 19 Then
                Player(0).blow = 0
                contexplosao = 0
                Player(0).CarroExplodiu = False
                tmrexplosao.Interval = 0
                tmrexplosaowait.Interval = 0
                Player(0).CarroVaiExplodir = False
                RetornandoAPista = True
                tmrRetorno.Interval = 1
                    
            End If
End Sub

Private Sub tmrexplosaowait_Timer()

Dim s As Long
tmrexplosaowait.Interval = 0
If Player(0).CarroExplodiu = True Then Exit Sub
If carrofora = True Then
                     Sound.WavPlay sons.eExplosao, EffectsVolume
                    Randomize Timer
                    s = Int((5 - 1) * Rnd) + 1
                    If s = 1 Then Sound.WavPlay sons.always, LarryVolume
                    If s = 2 Then Sound.WavPlay sons.hurrysup, LarryVolume
                    If s = 3 Then Sound.WavPlay sons.ouch, LarryVolume
                    If s = 4 Then Sound.WavPlay sons.UaiPaud, LarryVolume
                    If s = 5 Then Sound.WavPlay sons.wow, LarryVolume
                    
                    Player(0).CarroExplodiu = True
                    processealtura_caindo = False
                    processeAltura = False
                    Player(0).CarroSaltouPelaRampa = False
                    Player(0).CarroNaRampadaFrente = False
                    Player(0).CarroNaRampadeTras = False
                    'Player(0).AlturaReal = Player(0).position.Z
                    'alturaAnterior = Player(0).AlturaReal
                    BotaoSaltoPressionado = False
                    'AplicandoGravidade = False
                    'verifica se quando tocar o chao o carro explode
                    ExtendRampaCont = 0
                    ExtendRampa = False
                    LastRampa = 0
  tmrexplosao.Interval = 100
Else
            FrmDirectX.tmrexplosao.Interval = 0
            FrmDirectX.tmrexplosaowait.Interval = 0
            ContpixelAposRetornar = 0
            RetornandoAPista = False
            Player(0).CarroExplodiu = False
            Player(0).CarroVaiExplodir = False
            
End If
End Sub
Private Sub tmrexplosaowait2_Timer()

Dim s As Long
tmrexplosaowait2.Interval = 0
If Player(0).CarroExplodiu = True Then Exit Sub
If carrofora = True Then
                Sound.WavPlay sons.eExplosao, EffectsVolume
                    Randomize Timer
                    s = Int((5 - 1) * Rnd) + 1
                    If s = 1 Then Sound.WavPlay sons.always, LarryVolume
                    If s = 2 Then Sound.WavPlay sons.hurrysup, LarryVolume
                    If s = 3 Then Sound.WavPlay sons.ouch, LarryVolume
                    If s = 4 Then Sound.WavPlay sons.UaiPaud, LarryVolume
                    If s = 5 Then Sound.WavPlay sons.wow, LarryVolume
                    
                    Player(0).CarroExplodiu = True
                    processealtura_caindo = False
                    processeAltura = False
                    Player(0).CarroSaltouPelaRampa = False
                    Player(0).CarroNaRampadaFrente = False
                    Player(0).CarroNaRampadeTras = False
                    'Player(0).AlturaReal = Player(0).position.Z
                    'alturaAnterior = Player(0).AlturaReal
                    BotaoSaltoPressionado = False
                    'AplicandoGravidade = False
                    'verifica se quando tocar o chao o carro explode
                    ExtendRampaCont = 0
                    ExtendRampa = False
                    LastRampa = 0
  tmrexplosao.Interval = 100
Else
            FrmDirectX.tmrexplosao.Interval = 0
            FrmDirectX.tmrexplosaowait.Interval = 0
            ContpixelAposRetornar = 0
            RetornandoAPista = False
            Player(0).CarroExplodiu = False
            Player(0).CarroVaiExplodir = False
            
End If
End Sub


Private Sub tmrFrames_Timer()
GameFps = CountFrames
If CountFrames = 0 Then Exit Sub
FPS = 7 'reduz aumenta velocidade do carro
'SaltoReal = FPS * 6
SaltoReal = 40
VelocidadeVirtual = VelocidadeMaxima / FPS
CountFrames = 0
End Sub

Private Sub tmrjoined_Timer()
JString = Empty
tmrjoined.Interval = 0
End Sub


Private Sub tmrProccessSaidText_Timer()
'abaixa o texto
        SaidText(5) = SaidText(4)
        SaidText(4) = SaidText(3)
        SaidText(3) = SaidText(2)
        SaidText(2) = SaidText(1)
        SaidText(1) = Empty
        
End Sub

Private Sub tmrProcessGetList_Timer()
ProcessGetList
tmrProcessGetList.Interval = 0
End Sub

Private Sub tmrRefresh_Timer()
RefreshServers
tmrRefresh.Interval = 0
End Sub

Private Sub tmrRetorno_Timer()
'carro nao retornou,posicionar agora

If RetornandoAPista = False Then tmrRetorno.Interval = 0: Exit Sub

    'CheckPoints(PoligonoDono).pos.x = (Poligono(PoligonoDono).pos(0).x + (Poligono(PoligonoDono).pos(1).x - Poligono(PoligonoDono).pos(0).x) / 2) - 50
    'CheckPoints(PoligonoDono).pos.y = Poligono(PoligonoDono).pos(0).y + (Poligono(PoligonoDono).pos(3).y - Poligono(PoligonoDono).pos(0).y) / 2

    Player(0).position.x = CarLastPositionX
    Player(0).position.y = CarLastPositionY
    
    Player(0).position.Z = LastPolBefore
    
    If CheckPoints(LastCheckPoint).ForceToindex <> -1 Then Player(0).Car_Image_Index = CheckPoints(LastCheckPoint).ForceToindex
    Player(0).velocidade = 0
    'jump = 30
    alturaAnterior = -90
    processealtura_caindo = True
    processeAltura = True
    AplicandoGravidade = True
    RetornandoAPista = False
    SubaDescaCarroAposSalto = False
    tmrRetorno.Interval = 0
    RetornadoAPista = True
'    Exit Sub
'End If
'Next
'se nenhum desses , retornar a largada
 '   Player(0).position.x = LinhadeChegada(1).centerPos.x
 '   Player(0).position.y = LinhadeChegada(1).centerPos.y
 '   Player(0).position.Z = LinhadeChegada(1).centerPos.Z
 '   Player(0).position.Z = CheckPoints(x).centerPos.Z
 '   Player(0).velocidade = 0
 '   SubaDescaCarroAposSalto = False
 '   alturaAnterior = -90
    'Player(0).AlturaReal = Player(0).position.Z + jump
 '   processealtura_caindo = True
 '   processeAltura = True
 '   AplicandoGravidade = True
 '   RetornandoAPista = False
 '   RetornadoAPista = True
 '   tmrRetorno.Interval = 0
 '   Exit Sub
End Sub

Private Sub tmrSayIt_Timer()
Sound.WavPlay dizerOque, LarryVolume
tmrSayIt.Interval = 0
End Sub

Private Sub tmrSetas_Timer()
EspereSetas = False
' derrapar = False
End Sub


Private Sub tmrShowAtack_Timer()
FrmDirectX.tmrShowAtack.Interval = 0
Atack = Empty
End Sub

Private Sub tmrShowString_Timer()
ShString = Empty
tmrShowString.Interval = 0
End Sub

Private Sub tmrStart_Timer()
GameStarted = True
tmrStart.Interval = 0
End Sub

Private Sub tmrStartGame_Timer()
'paranoid = 0
    'badtobone = 1
    'pettergun = 2
    'borntobewild = 3
    'highwayStar = 4
GameStarted = True
PodedesenharoCarro = True
        GameStatus = 6
ConectadoAoServidor = True
tmrStartGame.Interval = 0
        'processeAltura = True
End Sub

Private Sub tmrwaitDikX_Timer()
waitDikX = False
tmrwaitDikX.Interval = 0
End Sub

Private Sub Winsock1_Connect()
TimerConnect.Interval = 0
Winsock1.Tag = "conectado"
'ShowString "conectado ao servidor"
'envia um ola ao servidor
ConnectStatus = "conectado , esperando resposta"
Dim command As command
command.command = SendHelloServer
 Dim res As Long
 
 Do
 res = res + 1
 If res > 500 Then Exit Sub
 Loop Until send(Winsock1.SocketHandle, command, GetSizeToSend(command.command), 0) <> SOCKET_ERROR
Sound.WavPlay MusicaPrincipal, MusicVolume
'ShowString "mandando hello"
   
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
  'Dim datacars As Other_Player_Stats
  'Sound.WavStop sons.apresentacao
  On Error Resume Next
 Dim b() As Byte
 Dim tStr As String
 Dim tpiloto As pilot
 Dim buffer() As Byte
 Dim tempbuffer() As Byte
 Dim command As command
 Dim OtherTemp As otherPlayersData
 Dim ObjetctReceived As SendObjects
 Dim SeekBuffer As Long
 Dim x As Long
 Dim oSize As Long
 Dim y As Long
 Dim s As Long
 Dim res As Long
   Dim res2 As Long
   Dim cont As Long
 ReDim buffer(0 To bytesTotal - 1)
  Winsock1.GetData buffer
         
    
    'primeiro descobre o comando
      oSize = Int(bytesTotal / LenB(command)) - 1
    Dim mComand As comandos
    
    Do
        
     If SeekBuffer >= bytesTotal Then
            'envia seu nome
            Exit Sub
     End If
    'If bytesTotal - SeekBuffer >= lenb(command) Then
       
            'For x = 0 To UBound(buffer)
            'Form2.List3.AddItem x & " = " & buffer(x)
            'Next x
           
        CopyMemory mComand, buffer(SeekBuffer), LenB(mComand)
        CopyMemory command, buffer(SeekBuffer), GetSizeToSend(mComand)
        
        'CopyMemory command.parametro1, buffer(SeekBuffer + GetSizeToSend(mComand)), lenb(command.parametro1)
    'Else
     '   CopyMemory ByVal VarPtr(command), buffer(SeekBuffer), bytesTotal - SeekBuffer
    'End If
  
    ''command.command =corridacompleta
 'If UseBots = True Then Form2.List2.AddItem mComand & "  " & SeekBuffer & " " & command.parametro1 & "  " & command.command
 'Form2.List2.AddItem "(0) = " & buffer(0)
 'Form2.List2.AddItem "(100) = " & buffer(100)
 'Form2.List2.AddItem "(200) = " & buffer(200)
    Select Case mComand
    
    
    Case rename
    
    If command.parametro1 = Player(0).id Then
    
    'If IsUnicode(command.parametroStr) Then
    MyName = GetStringInArray(command.parametroStr)
    'Else
    'MyName = Left(command.parametroStr, command.StringLength)
    GoTo NextStep
    End If
    
    
    
    
    
    'If IsUnicode(command.parametroStr) Then
    OtherPlayers(GetOtherPlayersFromID(command.parametro1)).name = GetStringInArray(command.parametroStr)
    'Else
  
    Case Destruirobjeto
    'procura qual objeto tem o id e o desativa
        
        For x = 0 To UBound(ArmasFrontaisNaPista)
            If ArmasFrontaisNaPista(x).handle = command.parametro1 Then
                ArmasFrontaisNaPista(x).Active = False
                ArmasFrontaisNaPista(x).id = 0
            End If
        Next x
        For x = 0 To UBound(ObjectsFromNet)
            If ObjectsFromNet(x).handle = command.parametro1 Then
                ObjectsFromNet(x).Active = False
                ObjectsFromNet(x).id = 0
            End If
        Next x
    Case RePlayersIn
    server.PlayersIn = command.parametro1
    
    
    Case corridacompleta
        
        GameStarted = False
        ProntoPraCorrer = False
        ex_gaitestimation = cl_gaitestimation
        cl_gaitestimation = False
        tmrcorridacompletada.Interval = 4000
   
    Case spectate
    
    ZereTudo
    SaidText(1) = "voce entrou como espectador, favor esperar a corrida terminar"
        PistaAtual = command.parametro1
        tmrControlCamera.Interval = 25
        GameStarted = True
        'GameStarted = False
        PodedesenharoCarro = False
        ProntoPraCorrer = False
        
        GameStatus = 6
        ConectadoAoServidor = True

        AllChocksCreated = False
        
            
            Sound.WavStop MusicaPrincipal
            MusicaPrincipal = command.parametro2
            
            'If MusicaPrincipal > 4 Then MusicaPrincipal = 0
            Sound.WavPlay MusicaPrincipal, MusicVolume
          s = Int((10 - 1) * Rnd) + 1
        If s <= 5 Then Sound.WavPlay sons.stageset, LarryVolume Else Sound.WavPlay sons.carneage, LarryVolume
      
    CameraSeguirOutroPlayer = Int((server.PlayersIn - 1) * Rnd) + 1
    
    Case isLeading
        If command.parametro1 = CyberHawks Then Sound.WavPlay sons.ecyber, LarryVolume
        If command.parametro1 = IvanZypher Then Sound.WavPlay sons.eIvan, LarryVolume
        If command.parametro1 = JakeBlanders Then Sound.WavPlay sons.eJake, LarryVolume
        If command.parametro1 = KatarinaLyons Then Sound.WavPlay sons.eKatarina, LarryVolume
        If command.parametro1 = SnakeSanders Then Sound.WavPlay sons.eSnake, LarryVolume
        If command.parametro1 = Tarquin Then Sound.WavPlay sons.eTarquin, LarryVolume
        dizerOque = sons.jaminthefirst
        FrmDirectX.tmrSayIt.Interval = 500
            ShowJoinedPlayer GetStringInArray(command.parametroStr) & " está liderando"
        
        
    Case LastLapp
    AttackString = "Last Lap"
    tmrattackBonus.Interval = 4000
    Sound.WavPlay sons.LastLap, LarryVolume
    
    Case AttackBonnus
    AttackString = "Attack Bonus"
    tmrattackBonus.Interval = 4000
    
    Case WhoPlacedFirst
    tpiloto = command.parametro1
    If tpiloto = CyberHawks Then Sound.WavPlay sons.ecyber, LarryVolume
    If tpiloto = IvanZypher Then Sound.WavPlay sons.eIvan, LarryVolume
    If tpiloto = JakeBlanders Then Sound.WavPlay sons.eJake, LarryVolume
    If tpiloto = KatarinaLyons Then Sound.WavPlay sons.eKatarina, LarryVolume
    If tpiloto = SnakeSanders Then Sound.WavPlay sons.eSnake, LarryVolume
    If tpiloto = Tarquin Then Sound.WavPlay sons.eTarquin, LarryVolume
    Sleep 500
    Sound.WavPlay sons.First, LarryVolume
    
    Case WhoPlacedSecond
    tpiloto = command.parametro1
    If tpiloto = CyberHawks Then Sound.WavPlay sons.ecyber, LarryVolume
    If tpiloto = IvanZypher Then Sound.WavPlay sons.eIvan, LarryVolume
    If tpiloto = JakeBlanders Then Sound.WavPlay sons.eJake, LarryVolume
    If tpiloto = KatarinaLyons Then Sound.WavPlay sons.eKatarina, LarryVolume
    If tpiloto = SnakeSanders Then Sound.WavPlay sons.eSnake, LarryVolume
    If tpiloto = Tarquin Then Sound.WavPlay sons.eTarquin, LarryVolume
    Sleep 500
    Sound.WavPlay sons.Second, LarryVolume
    
    Case WhoPlacedThird
    tpiloto = command.parametro1
    If tpiloto = CyberHawks Then Sound.WavPlay sons.ecyber, LarryVolume
    If tpiloto = IvanZypher Then Sound.WavPlay sons.eIvan, LarryVolume
    If tpiloto = JakeBlanders Then Sound.WavPlay sons.eJake, LarryVolume
    If tpiloto = KatarinaLyons Then Sound.WavPlay sons.eKatarina, LarryVolume
    If tpiloto = SnakeSanders Then Sound.WavPlay sons.eSnake, LarryVolume
    If tpiloto = Tarquin Then Sound.WavPlay sons.eTarquin, LarryVolume
    Sleep 500
    Sound.WavPlay sons.third, LarryVolume
    
    Case InvalidVersion
     Winsock1.Close
    DataStream.Close
    StringStream.Close
    ConnectStatus = "versao invalida, favor atualizar sua versao"
    tmrConnectError.Interval = 4000
    GameStatus = 5
    GoTo NextStep
    
    Case ServerFull
     Winsock1.Close
    DataStream.Close
    StringStream.Close
    ConnectStatus = "servidor cheio"
    GameStatus = 5
    tmrConnectError.Interval = 4000
    
    Case VoceTaNoJogo
    If GameStatus <> 6 Then GoTo NextStep
    tmrControlCamera.Interval = CamSpeed
    
    'ConnectStatus = "conectado!"
    'Conectar = False
    'ConectadoAoServidor = True
    'AllChocksCreated = False
    'GameStarted = False
    'GameStatus = 6
    'PodedesenharoCarro = True
    'If PistaAtual <> command.parametro3 Then AllChocksCreated = False
    If GameStatus = 6 Then PistaAtual = command.parametro1
     
     If GameStatus = 6 Then tmrStartGame.Interval = 2000
        'If PistaAtual <> command.parametro3 Then
            'PistaAtual = command.parametro3
         '   AllChocksCreated = False
          '  If PistaAtual = 0 Then DesenhePista chem_vi, 0, 0, True
           ' If PistaAtual = 1 Then DesenhePista2 chem_vi, 0, 0, True
            'If PistaAtual = 2 Then DesenhePista3 chem_vi, 0, 0, True
        'End If
    Case ShowTab
    'tabShowStr = tabShowStr & Left(command.parametroStr, command.StringLength)
    
    AddTabString command
    
    
    Case Ping
    
        send Winsock1.SocketHandle, command, GetSizeToSend(command.command), 0
        'ReDim b(0 To GetSizeToSend(command.command) - 1)
        'CopyMemory b(0), command, GetSizeToSend(command.command)
        'Winsock1.SendData b
    Case GameFinished
    
    'Sound.WavStop sons.apresentacao
    
    'For x = 1 To UBound(OtherPlayers)
    'OtherPlayers(x).correndo = False
    'Next x
    'ZereTudo
    'Case Go
    'GameStarted = True
    'GameStatus = 6
    'PodedesenharoCarro = True
    'ConectadoAoServidor = True
    'Randomize Timer
    's = Int(2 - 1 * Rnd) + 1
    'If s = 1 Then Sound.WavPlay sons.stageset Else Sound.WavPlay sons.carneage
    
    Case sPosition
        'Sound.WavStop sons.apresentacao
        cl_gaitestimation = cl_gaitestimation Or ex_gaitestimation
        
        
        'If tmrcorridacompletada.Interval <> 0 Then
        cont = 0
        'Do
         '   DoEvents
          '  cont = cont + 1
           ' If cont > 5000 Then Exit Do
        'Loop Until tmrcorridacompletada.Interval = 0
        'End If
        
        tmrcorridacompletada.Interval = 0
        ProntoPraCorrer = True
        ZereTudo
        tmrControlCamera.Interval = 25
   
        GameStarted = False
        PodedesenharoCarro = True
        GameStatus = 6
        ConectadoAoServidor = True

        Player(0).position.x = command.parametro1
        Player(0).position.y = command.parametro2
        AllChocksCreated = False
        PistaAtual = command.parametro3
        'If PistaAtual = 0 Then DesenhePista chem_vi, 0, 0, True
        'If PistaAtual = 1 Then DesenhePista2 chem_vi, 0, 0, True
        'If PistaAtual = 2 Then DesenhePista3 chem_vi, 0, 0, True
        
        Player(0).Car_Image_Index = 15
        Player(0).velocidade = command.parametro4
        'MoverCarro Player(0), command.parametro4, Player(0).position.angulo
        'SaidText(2) = Player(0).position.x & "  " & Player(0).position.y
        'Camera.x = (Player(0).position2D.x - (80 / ZoomX)) - 250
        'Camera.y = (Player(0).position2D.y - ((250 / ZoomY) / ZoomY)) - 250
   
            
            Sound.WavStop MusicaPrincipal
            'MusicaPrincipal = MusicaPrincipal + 1
            'If MusicaPrincipal > 4 Then MusicaPrincipal = 0
            MusicaPrincipal = command.parametroDouble
            
            Sound.WavPlay MusicaPrincipal, MusicVolume
        

        If command.parametro5 = 0 Then tmrStartGame.Interval = 4000 Else tmrStartGame.Interval = 1
        
    Case ChoosePlayer
        'SelectingPlayer = True
        
    Case SendingText
    've o que tiver livre
    
    
    Case OK_Player
            
    Case PlayerJoined
    
    
    
    Case UnRegisterPlayer
    'ShowString "disregistrando player"
    For x = 1 To 100
    If OtherPlayers(x).Data.id = command.parametro1 Then
    OtherPlayers(x).Data.id = 0
    OtherPlayers(x).Active = False
    PaintCars(x).id = 0
    
    'MsgBox "desresgistrado"
    server.PlayersIn = server.PlayersIn - 1
    If server.PlayersIn < 0 Then server.PlayersIn = 0
    End If
    Next x
    
    Case GetID
    
    ConnectStatus = "pegando identificacao " & str(command.parametro1)
    
    Player(0).id = command.parametro1
    'PistaAtual = command.parametro2
    Player(0).receivedID = True

    command.command = dados1Received
    res = 0
    Do
        res = res + 1
        If res > 500 Then Exit Do
    Loop Until send(Winsock1.SocketHandle, command, GetSizeToSend(command.command), 0) <> SOCKET_ERROR
    'envia seu nome
    ''command.command = SendMyName
    
     ''command.parametroStr = MyName
    ''command.StringLength = Len(MyName)
    'versao
    ''command.parametro1 = 3
    ''command.parametro2 = Player(0).piloto
    
    ''send Winsock1.SocketHandle, ByVal VarPtr(command), GetSizeToSend(command.command), 0
    'DoEvents
    
    
    
    Case PlayersIn
    
    ConnectStatus = " passando numero de players (" & str(command.command) & ")"
    'ShowString "players In = " & str(command.parametro1)
    server.PlayersIn = command.parametro1
    ConectadoAoServidor = True
    command.command = dados2Received
    'ReDim Preserve OtherPlayers(0 To Server.PlayersIn - 1)
 res = 0
    Do
        res = res + 1
        If res > 500 Then Exit Do
    res2 = send(Winsock1.SocketHandle, command, GetSizeToSend(command.command), 0)
    Loop Until res2 <> SOCKET_ERROR
    
    
    Case Register1Player
    
    
    
    If command.parametro1 <> Player(0).id Then
    'verifica se ja nao esta esse ID
    
    For x = 1 To 100
        If OtherPlayers(x).Data.id = command.parametro1 Then GoTo NextStep
    Next x
    'aloca o ID
    
    For x = 1 To 100
    If OtherPlayers(x).Data.id = 0 Then
        OtherPlayers(x).Data.id = command.parametro1
        OtherPlayers(x).ImageIndex = x
        OtherPlayers(x).name = GetStringInArray(command.parametroStr)


        PaintCars(x).id = command.parametro1
        Exit For
    End If
    
    'DoEvents
    Next x
    
          
    End If
    
    
    Case RegisterPlayers
   
    
    
    If command.parametro1 <> Player(0).id Then
    'verifica se ja nao esta esse ID
    For x = 1 To 100
    If OtherPlayers(x).Data.id = command.parametro1 Then GoTo next1
    Next x
    'aloca o ID
    
    For x = 1 To 100
    If OtherPlayers(x).Data.id = 0 Then
        OtherPlayers(x).Data.id = command.parametro1
        If UseBots = True Then Form2.List1.AddItem "registrado " & command.parametro1 & " em " & x
        If UseBots = True Then Form2.List1.AddItem "registrado nome " & GetStringInArray(command.parametroStr) & " em " & x
        'OtherPlayers(x).name = Left(command.parametroStr, command.StringLength)
        'If IsUnicode(command.parametroStr) Then
            OtherPlayers(x).name = GetStringInArray(command.parametroStr)
        'Else
         '   OtherPlayers(x).name = Left(command.parametroStr, command.StringLength)
        'End If
        'OtherPlayers(x).name = StrConv(command.parametroStr, vbUnicode)
            'OtherPlayers(x).name = CorrigirString(OtherPlayers(x).name)
        PaintCars(x).id = command.parametro1
        Exit For
    End If
    
    'DoEvents
    Next x
    
    End If
next1:
     ConnectStatus = "passando outros players info"
    'ShowString "registrando players "
    
    'If command.parametro1 = 0 Then
    'envia seu nome
    command.command = SendMyName
    
    PutStringInArray command.parametroStr, MyName
    'command.StringLength = Len(MyName)
    'versao
    command.parametro1 = 5
    command.parametro2 = Player(0).piloto
    
    res = 0
    Do
        res = res + 1
        If res > 500 Then Exit Do
    Loop Until send(Winsock1.SocketHandle, command, GetSizeToSend(command.command), 0) <> SOCKET_ERROR
    
    DeletePasswordIntoName MyName
    
    
    End Select
    
NextStep:
SeekBuffer = SeekBuffer + GetSizeToSend(mComand)
    Loop
    
    'quando se desconectar
    'apagar ID
    'receivedID
End Sub

Public Sub ShowString(str As String, Optional ByVal x As Long, Optional ByVal y As Long, Optional ByVal tempo As Long)
ShString = str
ShX = x
ShY = y
If tempo <> 0 Then FrmDirectX.tmrShowString.Interval = tempo * 1000
End Sub

Public Sub ShowJoinedPlayer(str As String)
If str = Empty Then Exit Sub
JString = str

FrmDirectX.tmrjoined.Interval = 5000

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim x As Long
On Error Resume Next
If Number = 10050 Or Number = 10051 Or Number = 10052 Or Number = 10053 Or Number = 10054 Or Number = 10057 Or Number = 10058 Or Number = 10060 Or Number = 10061 Or Number = 10064 Or Number = 10065 Then
    ConnectStatus = Description
    
    ConectadoAoServidor = False
    Player(0).receivedID = False
    Player(0).id = 0
    'desregistra todos
    For x = 1 To 100
        OtherPlayers(x).Data.id = 0
    Next x
    GameStatus = 5
    ConectadoAoServidor = False
    Winsock1.Close
    DataStream.Close
    StringStream.Close
    tmrConnectError.Interval = 4000
  
End If

'Winsock1.Connect server.Address, 20777
End Sub

Public Sub SelectPlayer()
Dim x As Long
Dim extFont As Long
'Dim FontInfo As New StdFont

'BackBuffer.SetFontTransparency True
'BackBuffer.SetForeColor vbWhite
'FontInfo.Bold = True
'FontInfo.Size = 12
'FontInfo.name = "Verdana"

'BackBuffer.SetFont FontInfo

BackBuffer.DrawText 70, 30, "pressione [ ENTER ] para escolher", False
DisplaySprite SelectScreen, 100, 250, , True
DisplaySprite pilotos(PlayerNumber), 300, 50, , True

End Sub

Public Sub ZereTudo()
Dim x As Long
CarLastPositionX = 0
CarLastPositionY = 0
LastPolBefore = 0

CameraSeguirOutroPlayer = 0
    Player(0).VoltasDadas = 0
    Player(0).armas.Traseiras.quantidade = Player(0).armas.Traseiras.QuantasTenho
    Player(0).armas.Frontais.quantidade = Player(0).armas.Frontais.QuantasTenho
    GameStarted = False
    Player(0).velocidade = 0
    jump = 0
    Player(0).blow = 0
      
                    naoelevarAgora = False
                    jump = 0
                    RampaSaltada = False
                    processealtura_caindo = False
                    processeAltura = False
                    OrigemdoPulo = 0
                    Player(0).CarroSaltouPelaRampa = False
                    Player(0).CarroNaRampadaFrente = False
                    Player(0).CarroNaRampadeTras = False
                    Player(0).AlturaReal = Player(0).position.Z
                    alturaAnterior = Player(0).AlturaReal
                    BotaoSaltoPressionado = False
                    AplicandoGravidade = False
                    'verifica se quando tocar o chao o carro explode
                    ExtendRampaCont = 0
                    ExtendRampa = False
                    LastRampa = 0
                    RetornadoAPista = False
                    ProcesseSuspensao = True
                    SubaDescaCarroAposSalto = False
    For x = 0 To 2000
        ObjectsFromNet(x).Active = False
    Next x
    For x = 0 To UBound(ArmasTraseirasNaPista)
    ArmasTraseirasNaPista(x).Active = False
    Next x
    
    For x = 0 To UBound(ArmasFrontaisNaPista)
    ArmasFrontaisNaPista(x).Active = False
    Next x
    
    For x = 0 To UBound(SprayPixadas)
        SprayPixadas(x).Status = 0
    Next x

End Sub


Public Sub ExecuteCommands(ByVal strData As String)

Dim c1 As String
Dim p1 As String
Dim m As String
Dim command As command
Dim res As Long
m = strData
  comandosStr2 = Empty
  comandosStr1 = Empty
   ''VERIFICA se é comando
   
   mSplit strData
   c1 = LCase(comandosStr1)
   If Len(comandosStr2) <> 0 Then p1 = comandosStr2
   

   Select Case c1
   Case "timeleft"
   command.command = timeleft
   res = 0
    Do
        res = res + 1
        If res > 50 Then Exit Do
    Loop Until send(FrmDirectX.Winsock1.SocketHandle, command, GetSizeToSend(command.command), 0) <> SOCKET_ERROR
    '    ReDim b(0 To GetSizeToSend(command.command) - 1)
     '   CopyMemory b(0), command, GetSizeToSend(command.command)
      '  Winsock1.SendData b
   Exit Sub
   
   Case "kick"
   command.command = kick
   command.parametro1 = p1
    res = 0
 
    Do
        res = res + 1
        If res > 50 Then Exit Do
    Loop Until send(FrmDirectX.Winsock1.SocketHandle, command, GetSizeToSend(command.command), 0) <> SOCKET_ERROR
        'ReDim b(0 To GetSizeToSend(command.command) - 1)
        'CopyMemory b(0), command, GetSizeToSend(command.command)
        'Winsock1.SendData b
   Exit Sub
   
   
   Case "sv_hideconsole"
   If Val(p1) <> 0 And Val(p1) <> 1 Then SaidText(1) = "sv_hideconsole deve ser 0 ou 1 ": Exit Sub
   command.command = sv_hideconsole
   command.parametro1 = p1
 res = 0
    Do
        res = res + 1
        If res > 50 Then Exit Do
    Loop Until send(FrmDirectX.Winsock1.SocketHandle, command, GetSizeToSend(command.command), 0) <> SOCKET_ERROR
    '    ReDim b(0 To GetSizeToSend(command.command) - 1)
     '   CopyMemory b(0), command, GetSizeToSend(command.command)
      '  Winsock1.SendData b
        
   Exit Sub
   
   Case "gl_vsync"
   If Len(p1) = 0 Then SaidText(1) = "gl_vsync = " & VSync: Exit Sub
   If Val(p1) = 1 Then VSync = True: SaidText(1) = "gl_vsync ajustada para 1": Exit Sub
    If Val(p1) = 0 Then VSync = False: SaidText(1) = "gl_vsync ajustada para 0": Exit Sub
    SaidText(1) = "gl_vsync deve ser 0 ou 1 "
   Exit Sub
   
   Case "showfps"
   If Len(p1) = 0 Then SaidText(1) = "showfps = " & ShowFps: Exit Sub
   If Val(p1) = 1 Then ShowFps = True: SaidText(1) = "showfps ajustada para 1": Exit Sub
    If Val(p1) = 0 Then ShowFps = False: SaidText(1) = "showfps ajustada para 0": Exit Sub
    SaidText(1) = "showfps deve ser 0 ou 1 "
   Exit Sub
   
   
   Case "cl_quit"
   End
   Exit Sub
   
   Case "cl_gaitestimation"
   If Len(p1) = 0 Then SaidText(1) = "cl_gaitestimation = " & cl_gaitestimation: Exit Sub
   If Val(p1) = 1 Then cl_gaitestimation = True: SaidText(1) = "cl_gaitestimation ajustada para 1": Exit Sub
    If Val(p1) = 0 Then cl_gaitestimation = False: SaidText(1) = "cl_gaitestimation ajustada para 0": Exit Sub
    SaidText(1) = "cl_gaitestimation deve ser 0 ou 1 "
   Exit Sub
   
   Case "cl_message"
   If Len(p1) = 0 Then SaidText(1) = "cl_message = " & cl_message: Exit Sub
   If Val(p1) = 1 Then cl_message = True: SaidText(1) = "cl_message ajustada para 1": Exit Sub
    If Val(p1) = 0 Then cl_message = False: SaidText(1) = "cl_message ajustada para 0": Exit Sub
    SaidText(1) = "cl_message deve ser 0 ou 1 "
    Exit Sub
   
   Case "music_volume"
   If Len(p1) = 0 Then SaidText(1) = "music_volume = " & MusicVolume: Exit Sub
   If Abs(Val(p1)) < 0 Or Abs(Val(p1)) > 10000 Then SaidText(1) = "music_volume deve ser entre 0 a 10000": Exit Sub
   MusicVolume = Val(p1)
   SaidText(1) = "music_Volume ajustado para " & MusicVolume
   Exit Sub
   
   Case "effects_volume"
   If Len(p1) = 0 Then SaidText(1) = "effects_volume = " & EffectsVolume: Exit Sub
   If Abs(Val(p1)) < 0 Or Abs(Val(p1)) > 10000 Then SaidText(1) = "effects_volume deve ser entre 0 a 10000": Exit Sub
   EffectsVolume = Val(p1)
   SaidText(1) = "effects_Volume ajustado para " & EffectsVolume
   Exit Sub
   
   Case "larry_volume"
   If Len(p1) = 0 Then SaidText(1) = "larry_volume = " & LarryVolume: Exit Sub
   If Abs(Val(p1)) < 0 Or Abs(Val(p1)) > 10000 Then SaidText(1) = "Larry_volume deve ser entre 0 a 10000": Exit Sub
   LarryVolume = Val(p1)
   SaidText(1) = "Larry_Volume ajustado para " & LarryVolume
   Exit Sub
   
   Case "volume"
   If Len(p1) = 0 Then SaidText(1) = "volume = " & AudioVolume: Exit Sub
   If Abs(Val(p1)) < 0 Or Abs(Val(p1)) > 10000 Then SaidText(1) = "volume deve ser entre 0 a 10000": Exit Sub
   AudioVolume = Val(p1)
   SaidText(1) = "Volume ajustado para " & AudioVolume
   Exit Sub
   
   Case "gl_flag"
   If Len(p1) = 0 Then SaidText(1) = "gl_flag = " & gl_flag: Exit Sub
    gl_flag = p1
    SaidText(1) = "gl_flag ajustada " & p1
    
    Exit Sub
   
   Case "gl_noclipping"
   If Len(p1) = 0 Then SaidText(1) = "gl_noclipping = " & gl_noclipping: Exit Sub
    If Val(p1) = 1 Then gl_noclipping = True: SaidText(1) = "gl_noclipping ajustada para 1": Exit Sub
    If Val(p1) = 0 Then gl_noclipping = False: SaidText(1) = "gl_noclipping ajustada para 0": Exit Sub
    SaidText(1) = "gl_noclipping deve ser 0 ou 1 "
    Exit Sub
    
   Case "+attack"
    Attack = True
    Exit Sub
   
   Case "+attack2"
    Attack2 = True
    Exit Sub
   
   Case "name"
    If Len(p1) = 0 Then
    ''SaidText(1) = "name = " & MyName: Exit Function
    command.command = GetName
    res = 0
 
    Do
        res = res + 1
        If res > 50 Then Exit Do
    Loop Until send(FrmDirectX.Winsock1.SocketHandle, command, GetSizeToSend(command.command), 0) <> SOCKET_ERROR
    '    ReDim b(0 To GetSizeToSend(command.command) - 1)
     '   CopyMemory b(0), command, GetSizeToSend(command.command)
      '  Winsock1.SendData b
        
Exit Sub
    End If
    MyName = p1
    '' SaidText(1) = "camera ajustada para esquerda " & p1
    command.command = rename
        PutStringInArray command.parametroStr, MyName
    
    'command.StringLength = Len(MyName)
    
 res = 0
    Do
        res = res + 1
        If res > 50 Then Exit Do
    Loop Until send(FrmDirectX.Winsock1.SocketHandle, command, GetSizeToSend(command.command), 0) <> SOCKET_ERROR
        'ReDim b(0 To GetSizeToSend(command.command) - 1)
        'CopyMemory b(0), command, GetSizeToSend(command.command)
        'Winsock1.SendData b
        
    DeletePasswordIntoName MyName
    SetarValor "software\\RockNRollRacing", "name", MyName

    Exit Sub
   
   Case "camleft"
    If Len(p1) = 0 Then SaidText(1) = "CameraX = " & CamLeft: Exit Sub
    CamLeft = p1
    SaidText(1) = "camera ajustada para esquerda " & p1
    Exit Sub
    
   Case "camright"
    If Len(p1) = 0 Then SaidText(1) = "CameraX = " & CamLeft: Exit Sub
    CamLeft = -Val(Abs(p1))
    SaidText(1) = "camera ajustada para direita " & Abs(p1)
    Exit Sub
   
   Case "camup"
    If Len(p1) = 0 Then SaidText(1) = "CameraY = " & CamUp: Exit Sub
    CamUp = p1
    SaidText(1) = "camera ajustada para cima " & p1
    Exit Sub
    
   Case "camdown"
    If Len(p1) = 0 Then SaidText(1) = "CameraY = " & CamUp: Exit Sub
    CamUp = -Val(Abs(p1))
    SaidText(1) = "camera ajustada para baixo " & Abs(p1)
    Exit Sub
   
   Case "fixecam"
    If Len(p1) = 0 Then SaidText(1) = "FixeCam = " & FixeCam: Exit Sub
    If Val(p1) = 1 Then FixeCam = True: SaidText(1) = "FixeCam ajustada para 1": Exit Sub
    If Val(p1) = 0 Then FixeCam = False: SaidText(1) = "FixeCam ajustada para 0": Exit Sub
    SaidText(1) = "FixeCam deve ser 0 ou 1 "
    Exit Sub
  
   
   Case "camspeed"
    If Len(p1) = 0 Then SaidText(1) = "CamSpeed = " & CamSpeed: Exit Sub
    CamSpeed = p1
    SaidText(1) = "velocidade da camera ajustada para  " & p1 & " ms"
   Exit Sub
   
    Case Else
     SaidText(1) = "'" & c1 & "' comando nao reconhecido"
   End Select
   
 
End Sub
Public Function ProxImagem() As Long
Dim x As Long
For x = ImagemdaVez + 1 To 1000
    If ImagemAmudar(x).id <> 0 Then
        ProxImagem = x
        Exit Function
    End If
Next x

For x = 0 To ImagemdaVez - 1
    If ImagemAmudar(x).id <> 0 Then
        ProxImagem = x
        Exit Function
    End If
Next x

ProxImagem = ImagemdaVez
End Function
