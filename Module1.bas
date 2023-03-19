Attribute VB_Name = "Module1"
Option Explicit
Global ex_gaitestimation As Boolean
Global ProntoPraCorrer As Boolean
Global CoresJaMapeadas As Boolean
'Global PintandoCarro As Boolean
Type pixelCoord
    x As Long
    y As Long
End Type
Type pixelValues
    pixels0_255_0Coordenates() As pixelCoord
    pixels255_0_0Coordenates() As pixelCoord
    pixels0_0_255Coordenates() As pixelCoord
    pixels100_100_100Coordenates() As pixelCoord
    pixels150_150_150Coordenates() As pixelCoord
    pixels200_200_200Coordenates() As pixelCoord
End Type
Global pixelsCoordenates(23, 3) As pixelValues

'CodePage
Type Coresdatinta
vermelho As Byte
azul As Byte
verde As Byte
End Type

Type ImagemaMudarType
Sprite As Sprites
corSource As Long
CorDest As Long
AlturaAtual As Long
LarguraAtual As Long
id As Long
End Type

Global ImagemdaVez As Long
Global ImagemAmudar(1000) As ImagemaMudarType
Global corSource As Long
Global CorDest As Long
Global CoresTinta(0 To 254) As Coresdatinta
Global AudioOk As Boolean
Global umaVez As Boolean
Global ConsoleVisible As Boolean
Global UseBots As Boolean
Global BotConnectAt As String
Global ConsoleRoll As Long
Global ShowFps As Boolean
Global GameFps As Long
Global Attack As Boolean
Global Attack2 As Boolean
Global Attack3 As Boolean
Global gl_noclipping As Boolean
Global comandosStr1 As String
Global comandosStr2 As String
Global gl_flag As Long
Global AudioVolume As Long
Global MusicVolume As Long
Global LarryVolume As Long
Global EffectsVolume As Long
Global cl_message As Boolean
Global cl_gaitestimation As Boolean
Global ShowGeometry As Boolean
Global VSync As Boolean
Public Type COLORBYTES
   BlueByte As Byte
   GreenByte As Byte
   RedByte As Byte
   AlphaByte As Byte
End Type
Public Type COLORLONG
   longval As Long
End Type

Global PistaAtual As Long
Global MusicaPrincipal As Long
Const REG_SZ = 1 'Unicode nul terminated string
Const REG_BINARY = 3    'Free form binary
 Const REG_DWORD = 4   '32-bit number

Const HKEY_CURRENT_USER = &H80000001
'The HKEY_CURRENT_USER base key, which stores program information for the current user.
Const HKEY_LOCAL_MACHINE = &H80000002
'The HKEY_LOCAL_MACHINE base key, which stores program information for all users.
Const HKEY_USERS = &H80000003
'The HKEY_USERS base key, which has all the information for any user (not just the one provided by
'HKEY_CURRENT_USER).
Const HKEY_CURRENT_CONFIG = &H80000005
'The HKEY_CURRENT_CONFIG base key, which stores computer configuration information.
Const HKEY_DYN_DATA = &H80000006
Const KEY_ALL_ACCESS = &HF003F
'Permission for all types of access.
Const KEY_CREATE_LINK = &H20
'Permission to create symbolic links.
Const KEY_CREATE_SUB_KEY = &H4
'Permission to create subkeys.
Const KEY_ENUMERATE_SUB_KEYS = &H8
'Permission to enumerate subkeys.
Const KEY_EXECUTE = &H20019
'Same as KEY_READ.
Const KEY_NOTIFY = &H10
'Permission to give change notification.
Const KEY_QUERY_VALUE = &H1
'Permission to query subkey data.
Const KEY_READ = &H20019
'Permission for general read access.
Const KEY_SET_VALUE = &H2
'Permission to set subkey data.
Const KEY_WRITE = &H20006
 Type SECURITYATTRIBUTES
nLength As Long
lpSecurityDescriptor As Long
bInheritHandle As Long
End Type

 Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
"RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
ByVal reserved As Long, ByVal lpClass As String, ByVal dwOptions _
As Long, ByVal samDesired As Long, lpSecurityAttributes _
As SECURITYATTRIBUTES, phkResult As Long, _
lpdwDisposition As Long) As Long

 Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName _
As String, ByVal reserved As Long, ByVal dwType As Long, _
lpData As Any, ByVal cbData As Long) As Long

 Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long
    
    Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName _
As String, ByVal lpReserved As Long, lpType As Long, _
lpData As Any, lpcbData As Long) As Long


Public Type RGBColour
    r As Byte
    g As Byte
    b As Byte
End Type
Global ControlCameraX As Long
Global ControlCameraY As Long
Global CamLeft As Long
Global CamUp As Long
Global dizerOque As Long
Global FlashDrawBox As Long
Global SelectedList As Long
Global RefreshingServer As Boolean
   Global CurrRed As Integer
Global CurrGreen As Integer
Global CurrBlue As Integer
Global ShiftServerPage As Long
Global Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Global Const INTERNET_INVALID_PORT_NUMBER = 0
Global Const INTERNET_SERVICE_FTP = 1
Global Const FTP_TRANSFER_TYPE_BINARY = &H2
Global Const FTP_TRANSFER_TYPE_ASCII = &H1
Global Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000

Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As _
    Long, ByVal lpbuffer As String, ByVal dwNumberOfBytesToRead As Long, _
    lNumberOfBytesRead As Long) As Integer

Declare Function InternetOpenUrl Lib "wininet.dll" Alias _
    "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, _
    ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, _
    ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Long
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long

Global AttackString As String
Global ImageRoll As Long
Global ShiftImageY As Long
Global ShiftImageX As Long
Global OriginalRamp As DDGAMMARAMP
Global GammaControler As DirectDrawGammaControl
Global GammaRamp As DDGAMMARAMP
Global PilotChoose(0 To 5) As Sprites
Global GammaSupport As Boolean
Global FadeLevel As Long
Global AFVtemp As Long
Global GameStatus As Long
Global showingTab As Boolean
Global Atack As String
Global ContDataReceived As Long
Global PodedesenharoCarro As Boolean
Global GameStarted As Boolean
Global waitDikX As Boolean
Global OtherTime As Long
Global TempoForadaPista As Long
Global CrashDerrapagem As Boolean
Global CrashAngle As Long
Global TimerPassed As Double
Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, _
                                                    dwFlags As Long, ByVal _
                                                    lpWideCharStr As Long, ByVal _
                                                    cchWideChar As Long, ByVal _
                                                    lpMultiByteStr As Long, ByVal _
                                                    cbMultiByte As Long, ByVal _
                                                    lpDefaultChar As Long, ByVal _
                                                    lpUsedDefaultChar As Long) As Long

Global Const CP_ACP = 0 'ANSI
Global Const CP_MACCP = 2 'Mac
Global Const CP_OEMCP = 1 'OEM
Global Const CP_UTF7 = 65000
Global Const CP_UTF8 = 65001


Global GlobalString As command
Global SendingTypedTextNow As Boolean
Global Const SOCKET_ERROR = -1
Global SaidText(1 To 5) As String
Global EscrevendoTexto As Boolean
Global TextToSend As String
Global ConnectStatus  As String
Global ConectadoAoServidor As Boolean
Global ShString As String
Global ShX As Long
Global ShY As Long
Global JString As String
Global Conectar As Boolean
Global MyName As String
Global MainScreen As Sprites

Enum comandos
dataCarStream = 0
GetID = 1
PlayersIn = 2
RegisterPlayers = 3
CM_Crash = 4
UnRegisterPlayer = 5
SendMyName = 6
SendHelloServer = 7
PlayerJoined = 8
OK_Player = 9
SendingText = 10
SendFixedObjects = 11
RegisterLap = 12
WhoPlacedFirst = 13
WhoPlacedSecond = 14
WhoPlacedThird = 15
GameFinished = 16
InvalidName = 17
ServerFull = 18
ChoosePlayer = 19
ChooseCar = 20
VerifyResources = 21
Version = 22
InvalidVersion = 23
PlayAndCarReady = 24
sPosition = 25
Go = 26
Restart = 27
ShowTab = 28
Ping = 29
MorriPara = 30
VoceTaNoJogo = 31
sSpray = 32
AttackBonnus = 33
LastLapp = 34
isLeading = 35
dados1Received = 36
dados2Received = 37
dados3Received = 38
dados4Received = 39
dados5Received = 40
rename = 41
GetName = 42
Quit = 43
sv_hideconsole = 44
ban_id = 45
kick = 46
Admin_Free_slot = 47
Admin_deal = 48
sv_listPlayers = 49
sv_restart_round = 50
sv_restart_game = 51
sv_changelevel = 52
ServerHostName = 53
listId = 54
listIp = 55
map = 56
sv_MaxPlayers = 57
mp_timelimit = 59
voltas = 60
passWord = 61
removeid = 62
removeip = 63
corridacompleta = 64
spectate = 65
timeleft = 66
RePlayersIn = 67
Register1Player = 68
quemcorre = 69
Destruirobjeto = 70
End Enum



Type command
command As comandos
parametro1 As Long
parametro2 As Long
parametro3 As Long
parametro4 As Long
parametro5 As Long
parametroDouble As Double
'StringLength As Long
parametroStr(0 To 30) As Byte
'spray(0 To 74, 0 To 39) As Byte
invalidparam As Byte
End Type

Type Car_MS
command As comandos
parametro1 As Long
parametro2 As Long
parametro3 As Long
parametro4 As Long
parametro5 As Long
End Type


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(dest As Any, Source As Any, ByVal numBytes As Long)

Declare Function send Lib "wsock32" (ByVal s As Long, buffer As Any, ByVal length As Long, ByVal flags As Long) As Long

Global Crash As Boolean
Global LastCheckPoint As Long
Global LastVerticeChocked As Long
Global CarLastPositionX As Long
Global CarLastPositionY As Long
Global LastPolBefore As Long
Global RetornadoAPista As Boolean
Global DontProcessDerrapar As Boolean
Global naoelevarAgora As Boolean
Global SubaDescaCarroAposSalto As Boolean
Global ProcesseSuspensao As Boolean
Global PoligonoMaisBaixo As Long
Global MenorNivel As Long
Global ShowGrides As Boolean
Global RampaToUse As Long
Global espereFimdaRampa As Boolean
Global FixeCam As Boolean
Global CamSpeed As Long
Global RampaSaltada As Boolean
Global BotaoSaltoPressionado As Boolean
Global AplicandoGravidade As Boolean
Global ExtendRampa As Boolean
Global ExtendRampaCont As Long
Global diff As Double
Global tantoaAbaixar As Double
Global DontCheckColision As Boolean
Global alturaAnterior As Double
Global diminuirQueda As Boolean
Global lastWidth As Long
Global lastHeight As Long
Global LastPolignCreated As Long
Global poligonoAchado As Boolean
Global carrofora As Boolean
Global xAxis As Double
Global yAxis As Double
Global zAxis As Double
Global CarroNaPista As Boolean
Global CarroNaRampa As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const Retangular As Long = 0
 Private Const OFFSET_4 = 4294967296#
      Private Const MAXINT_4 = 2147483647
      Private Const OFFSET_2 = 65536
      Private Const MAXINT_2 = 32767
      Private Const OFFSET_1 = 256
      Private Const MAXINT_1 = 127
      
Enum armasdoJogo
sOil = 1
sMine = 2
sbomb = 3
slaser = 4
End Enum

Enum carros
Maraudercar = 0
End Enum

Type ArmaTras
tipo As armasdoJogo
quantidade As Integer
QuantasTenho As Integer
End Type

Type ArmaFrente
tipo As armasdoJogo
quantidade As Integer
QuantasTenho As Integer
End Type

Type armas
Traseiras As ArmaTras
Frontais As ArmaFrente
End Type
Global nAng As Double
Global elevacao As Long
Global LastPolign As Long
Global LastZ As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Enum posCarro
reto = 0
subindo = 1
descendo = 2
End Enum
Private Type LineType
    x As Integer
    y As Integer
    Z As Integer
    X1 As Integer
    Y1 As Integer
    Z1 As Integer
End Type

Type polign
indice As Long
pos(0 To 3) As POINTAPI
VideoPos(0 To 3) As POINTAPI
Type As Long
piso As PolignType
checkPertencente As Long
tangente As Double 'somente para rampas
NivelInicial As Double
PoligonoAnterior As Long
NivelFinal As Double
IsCheckPoint As Boolean
IsLagada As Boolean
altura As Long
largura As Long
ArmaUsada As armasdoJogo
IsRampaStep As Boolean
End Type

Global CarroEstaNoPoligono As Long

Type pos
altura As Double
largura As Double
x As Double
y As Double
Z As Double
angulo As Double
UltimoX_Posicionado As Double
UltimoY_Posicionado As Double
UltimoX_Poligono As Double
UltimoY_Poligono As Double
UltimoPoligonoDesenhado As PolignType
End Type

Type checks
pos As pos
Sprite As Sprites
color As Long
checked As Boolean
poligono As Long
ForceToindex As Long
index As Long
End Type
'cores 1 ate 20 são os checkpoints
Global CheckPoints() As checks
'cor 21= linha de chegada
'0 = vertical e 1 =horizontal
Global SelectingPlayer As Boolean
Global PlayerNumber As Long
Global LinhadeChegada(0 To 1) As checks
Type posCheck
x As Double
y As Double
Z As Double
angulo As Double
AtRightIsIn As Boolean
End Type
Enum pilot
IvanZypher = 0
JakeBlanders = 1
KatarinaLyons = 2
SnakeSanders = 3
Tarquin = 4
CyberHawks = 5
End Enum

Type Player_Stats
declive As posCarro
position As pos
position2D As pos
Car As carros
Car_Image_Index As Long
ChockPoints(3) As posCheck
CarroTodoNaPista As Boolean
CarroSeChocouQuinaExterna As Boolean
CarroSeChocouQuinaInterna As Boolean
CarroTodoFora As Boolean
CarroNaRampadaFrente As Boolean
CarroNaRampadeTras As Boolean
CarroPegouAlgo As Long
CarroSaltouPelaRampa As Boolean
velocidade As Double
CarroExplodiu As Boolean
CarroVaiExplodir As Boolean
VoltasDadas As Long
AlturaReal As Double
id As Long
receivedID As Boolean
armas As armas
piloto As pilot
TopSpeed As Double
TopAcceleration As Double
TopCorner  As Double
TopJumping As Double
color As Byte
name As String
blow As Long

End Type


Type Serv
PlayersIn As Long
Address As String
End Type



Type pl
id As Long
Serial As String * 10
socket As Long
DataSocket As Long
name As String * 20
voltas As Long
pontos As Long
piloto As pilot
PilotSelected As Boolean
Running As Boolean
Ping As Long
kills As Long
deads As Long
invalidparam As Byte
IsAdmin As Boolean
IsAdminMaster As Boolean
End Type

Global tabShowStr(20) As String
Type Register
commando As comandos
players As pl
End Type

Global server As Serv
Global AnguloDerrapagem As Long
Global derrapar As Boolean
Global Player() As Player_Stats
Global OtherPlayers(1 To 100) As Other_Player_Stats
Global PistaChockArea(0 To 29999) As posCheck
Global PonteiroChock As Long
Global AllChocksCreated As Boolean
Global Camera As pos
'Global velocidade As Double
Global CameraSeguirOutroPlayer As Long
Global amarelo As Long
Global azul As Long
Global Preto As Integer
Global branco As Long
Global cinza As Long
Global laranja As Long
Global verde As Long
Global vermelho As Long
Global violeta As Long
Global UltimoY As Long
Global SimularCarroY As Double
Global SimularCarroX As Double
Global processeAltura As Boolean
Global processealtura_caindo As Boolean
Global jump As Double
Global dik_left_time_press As Long
Global dik_right_time_press As Long
Global ProcesseExplosao As Boolean
Global contexplosao As Double
Global ContpixelAposRetornar As Long
Global GetAngle As Double
Global GetLastValidImageIndex As Long
Global RecPos() As pos
Global RecCounter As Long
Global lastpixel(0 To 3) As Long
Global CountPistasVerticais As Long
Global CountPistasHorizontais As Long
Type Objetos
Active As Boolean
VideoPos As POINTAPI
positionX As Integer
positionY As Integer
positionZ As Integer
extra As Byte
tipo As Byte
PolignToUse As Integer
id As Long
handle As Long
invalidparam As Byte
End Type


Global ArmasTraseirasNaPista() As Objetos
Global ArmasFrontaisNaPista() As Objetos
Type FumacaPos
position As pos
Status As Double
End Type
Global laser(0 To 23) As Sprites
Global FumacaPos(0 To 15000) As FumacaPos
Global Fumaca(0 To 3) As Sprites
Global Oleo As Sprites
Global ProcesseFumaca As Boolean
Global ProcesseFumacaOthers As Boolean
Global CountFumaca As Double
Global CountFrames As Long
Global FPS As Long
Global Const FPSBase As Double = 245
Global SaltoReal As Double
Global VelocidadeVirtual As Double 'sempre 7 no máximo para carro sem equipamento
Global VelocidadeMaxima As Double

Global vertices2D(0 To 199) As POINTAPI
Global poligono() As polign

Private Type Point_3D

    x As Double
    y As Double
    Z As Double

End Type

Private Const PI As Double = 3.14159265358979
Private Const RADIAN As Double = PI / 180

Private Const FOV As Double = 90

Private Vertex_List() As Point_3D
Private temp() As Point_3D
Private Local_Vertex() As Point_3D
Private Camera3D() As Point_3D
Private Perspective() As Point_3D
Private Screen3D() As Point_3D

Private Camera3D_Distance As Double

Private Number_Of_Vertices As Long

Private Viewport_Width As Double
Private Viewport_Height As Double

Private Viewplane_Width As Double
Private Viewplane_Height As Double

Private Distance As Double

Private Camera3D_Pos As Point_3D

Private ASPECT_RATIO As Double

Private angle As Double

Type otherPlayersData
commando As comandos
NewObject As Objetos
id As Long
positionX As Integer
positionY As Integer
positionZ As Integer
Car As Byte
Car_Image_Index As Byte
velocidade As Integer
CarroExplosao As Byte
FumacaDerrapagem As Byte
ShowSombra As Byte
elevacao As Byte
color As Long
blow As Byte
laserHit As Byte
oleoHit As Byte
JumpHit As Byte


End Type

Global Tinta(0 To 5) As Sprites
Global ConsoleScr As Sprites
Global ConsoleTxtscr As Sprites
Global ConsoleTxt As String
Type Other_Player_Stats
name As String * 255
Data As otherPlayersData
position2D As pos
LastX As Double
LastY As Double
LastZ As Double
Active As Boolean
contexplosao As Long
AcaboudeReceber As Boolean
correndo As Boolean
ImageIndex As Long
End Type

Global ObjectsFromNet(0 To 2000) As Objetos
Type SendObjects
command As comandos
objeto As Objetos
invalidparam As Byte
End Type

Global pilotos(0 To 5) As Sprites
Global SelectScreen As Sprites

Type TabParams
name As String
pontos As Long
kills As Long
deads As Long
Ping As Long
End Type
Global TabScreen(0 To 20) As TabParams
Global Logo(0 To 15) As Sprites
Global planetas(0 To 5) As Sprites
Global estrelas(0 To 5) As Sprites
Global planetaNaTela As Sprites
Global Tela1 As Sprites
Global botao() As Sprites
Global ServerAddress() As String

Type ServerInf
noIp As String
PlayersInInfo As String
Ping As String
name As String
Selected As Boolean
End Type
Global serverInfo() As ServerInf
Global ServerInfoIndexToRetrieve As Long
Global ServerStartedTime(0 To 20000) As Double
Global Roll As Sprites
Type ptCars
color As Long
id As Long
pronto As Boolean
End Type
Global PaintCars(0 To 100) As ptCars
Global NomeScreen As Sprites
Global buycarScreen As Sprites
Global interrogacao As Sprites
Global MoveDrawBoxLeft As Long
Global MoveDrawBoxTop As Long
Global sairTimer As Double
Global Sair As Boolean
Global spray As Sprites
Global SprayPixadas() As FumacaPos

Global Sound As clsDxSound


' These 3 variables are overhead for this app
Global nWaves() As Long
Global sWavName() As String
Enum sons
    paranoid = 0
    badtobone = 1
    pettergun = 2
    borntobewild = 3
    highwayStar = 4
    
    derrapagem = 5
    cornercrash = 6
    getobjects = 7
    salto = 8
    eLaser = 9
    eOleo = 10
    
    aceito = 11
    moving = 12
    naoaceito = 13
    
    eExplosao = 14
    
    eLaser2 = 15
    eLaser3 = 16
    eLaser4 = 17
    eLaser5 = 18
    eLaser6 = 19
    eLaser7 = 20
    
    eOleo2 = 21
    eOleo3 = 22
    eOleo4 = 23
    eOleo5 = 24
    eOleo6 = 25
    eOleo7 = 26
    
    eCarCrash = 27
    
    'in running
    stageset = 28
    abouttoblow = 29
    
    dollar = 31
    fadethelast = 32
    jaminthefirst = 33
    LastLap = 34
    
    'explodiu
    always = 30
    hurrysup = 35
    ouch = 36
    UaiPaud = 37
    wow = 38
    
    'finished
    First = 39
    Second = 40
    third = 41
    notcomplete = 42
    'pilotos
    ecyber = 43
    eIvan = 44
    eJake = 45
    eKatarina = 46
    eRip = 47
    eSnake = 48
    eTarquin = 49
    
    carneage = 50
    apresentacao = 51
    
End Enum

Global nPlayWav(0 To 200) As Long











Public Function DesenhePrimeiroPedacodaPista(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = vertical, Optional ByVal tZoom As Long = 1, Optional ByVal CriarPoligono As Boolean = False) As Boolean
If tamanho < 1 Then DesenhePrimeiroPedacodaPista = False

Dim p As Long
Dim a As Double
'Dim diffA As Long
Select Case mundo
Case chem_vi
    Select Case tipo
    Case vertical
        For p = 0 To (tamanho - 1)
            CountPistasVerticais = CountPistasVerticais + 1
            'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(0), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        Next p
        
             If CriarPoligono Then
                 a = 57.5 'altura do quadrado
                a = a * (tamanho)
                CreatePoligono 200, a + 70, Retangular, vertical, , x + 10, y
                
                ''Call CreatePoligono(x, -47 - a + y - 40, 205, 64.5, Retangular, vertical)
            End If
            ''f(y)= -1,356097 * x
        Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
        Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
        Camera.UltimoPoligonoDesenhado = vertical
      'relacao de cada pedaco
      '15x3
    
    Case horizontal
    For p = 0 To (tamanho - 1)
        CountPistasHorizontais = CountPistasHorizontais + 1
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(10), x + ((945 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((-225 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
        
    Next p
    Camera.UltimoX_Posicionado = x + ((945 / ZoomX) / Screen.TwipsPerPixelX * p)
    Camera.UltimoY_Posicionado = y - ((-225 / ZoomY) / Screen.TwipsPerPixelY * p)

    End Select
End Select
End Function

Public Function PegarTextoArquivo(ByVal arquivo As String)
On Error Resume Next
Dim f As Long
Dim k As Long
Dim b() As Byte
f = FreeFile
k = FileLen(arquivo)
ReDim b(0 To k - 1)
Open arquivo For Binary As #f
Get #f, 1, b
Close #f
PegarTextoArquivo = StrConv(b, vbUnicode)
End Function
Public Function DesenhepedacosdaPista(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = vertical, Optional ByVal tZoom As Long = 1, Optional ByVal CriarPoligono As Boolean = False, Optional ExtraHeight As Long = 0, Optional ExtraWidth As Long = 0, Optional deslocamento As Long = 0, Optional ByVal xInit As Long = 0, Optional ByVal yInit As Long = 0, Optional ByVal nShiftY As Long = 0, Optional ByVal nivelacao As Double = -1) As Boolean

If tamanho < 1 Then DesenhepedacosdaPista = False

Dim InitX As Long
Dim InitY As Long
Dim FimX As Long
Dim FimY As Long
Dim p As Long
Dim a As Double
'Dim diffA As Long
Select Case mundo
Case chem_vi
    Select Case tipo
    Case vertical
        For p = 0 To (tamanho - 1)
            CountPistasVerticais = CountPistasVerticais + 1
            'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(0), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        Next p
             If CriarPoligono Then
                 a = 64.5 'altura do quadrado
                a = a * (tamanho)
                
                CreatePoligono 200 + ExtraWidth, a + ExtraHeight, Retangular, vertical, , deslocamento, xInit, yInit, , nivelacao, , , nShiftY
            End If
   
        Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
        Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
        Camera.UltimoPoligonoDesenhado = vertical
      'relacao de cada pedaco
      '15x3
    
    Case horizontal
        For p = 0 To (tamanho - 1)
            CountPistasHorizontais = CountPistasHorizontais + 1
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
                DDraw.DisplaySprite pistas(10), x + ((945 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((-225 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        Next p
        If CriarPoligono Then
                 a = 42 'largura do quadrado
                a = a * (tamanho)
                CreatePoligono a + ExtraWidth, 329 + ExtraHeight, Retangular, horizontal, 0, -1 + deslocamento, xInit, yInit, , nivelacao, , , nShiftY
        End If
   
        Camera.UltimoX_Posicionado = x + ((945 / ZoomX) / Screen.TwipsPerPixelX * p)
        Camera.UltimoY_Posicionado = y - ((-225 / ZoomY) / Screen.TwipsPerPixelY * p)

    End Select
End Select
End Function
Public Function DesenheRampadaPista(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = vertical, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False, Optional ExtraHeight As Long = 0, Optional ExtraWidth As Long = 0, Optional ByVal nShiftX As Long = 0, Optional ByVal xInicial As Long = 0, Optional ByVal yInicial As Long = 0, Optional ByVal nShiftY As Long = 0) As Boolean
If tamanho < 1 Then DesenheRampadaPista = False
Dim p As Long
Dim a As Double
Dim InitX As Long
Dim InitY As Long
Dim FimX As Long
Dim FimY As Long
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(2), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
    Next p
     If CriarPoligono Then
        a = 241.5 'altura do quadrado
        a = a * (tamanho)
        CreatePoligono 210, a, Retangular, Rampa, 0.325
    End If
    Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
    Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function
Public Function DesenheLadeira(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = horizontal, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False, Optional ExtraHeight As Long = 0, Optional ExtraWidth As Long = 0, Optional ByVal nShiftX As Long = 0, Optional ByVal xInicial As Long = 0, Optional ByVal yInicial As Long = 0, Optional ByVal nShiftY As Long = 0) As Boolean
If tamanho < 1 Then DesenheLadeira = False
Dim p As Long
Dim a As Double
Dim InitX As Long
Dim InitY As Long
Dim FimX As Long
Dim FimY As Long
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(4), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100

    Next p
    
     If CriarPoligono Then
        a = 74.5 'altura do quadrado
        a = a * tamanho
        CreatePoligono 200, a, Retangular, Rampa, -0.595, nShiftX, xInicial, yInicial, , , , , nShiftY 'descendo
    End If
        
        If CriarPoligono Then
                a = 54.5 'altura do quadrado
                a = a * tamanho
                CreatePoligono 200, a, Retangular, vertical, , nShiftX, xInicial, yInicial, , , , , nShiftY
            End If
   
Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)

End Select
End Function

Public Function DesenhePista(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, Optional ByVal CriarPoligono As Boolean = False)
Dim k As Long
If CriarPoligono = True Then
'For k = 0 To UBound(poligono)
Camera.UltimoX_Posicionado = 0
Camera.UltimoY_Posicionado = 0
Camera.UltimoX_Poligono = 0
Camera.UltimoY_Poligono = 0
lastWidth = 0
lastHeight = 0

ReDim poligono(0)
poligono(0).Type = 0
poligono(0).ArmaUsada = 0
poligono(0).indice = -1
poligono(0).piso = 0
poligono(0).Type = 0
LastPolignCreated = 0
'Next k
End If
CountPistasVerticais = 0
CountPistasHorizontais = 0
DesenhePrimeiroPedacodaPista chem_vi, x, y, 12, , , CriarPoligono
DesenheRampadaPista chem_vi, (Camera.UltimoX_Posicionado - (5 / ZoomX)), (Camera.UltimoY_Posicionado - (88 / ZoomY)), 1, , , CriarPoligono
DesenheLadeira chem_vi, (Camera.UltimoX_Posicionado + (192 / ZoomX)), (Camera.UltimoY_Posicionado - (54 / ZoomY)), 1, , , CriarPoligono
desenhepista3rampa chem_vi, (Camera.UltimoX_Posicionado + (34 / ZoomX)), (Camera.UltimoY_Posicionado + (25 / ZoomY)), 1, , , CriarPoligono
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + (230 / ZoomX)), (Camera.UltimoY_Posicionado + (13 / ZoomY)), 8, , , CriarPoligono
DesenhePistaCurvaALTAESQold chem_vi, (Camera.UltimoX_Posicionado - (60 / ZoomX)), (Camera.UltimoY_Posicionado - (64 / ZoomY)), 1, , , CriarPoligono
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + (516 / ZoomX)), (Camera.UltimoY_Posicionado + (108 / ZoomY)), 6, horizontal, , CriarPoligono, , , -150
DesenheRampaHorizontalold chem_vi, (Camera.UltimoX_Posicionado - 0), (Camera.UltimoY_Posicionado - (9 / ZoomY)), 1, horizontal, , CriarPoligono
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + (162 / ZoomX)), (Camera.UltimoY_Posicionado + (35 / ZoomY)), 4, horizontal, , CriarPoligono, , 40
DesenhePista2Horizontal chem_vi, (Camera.UltimoX_Posicionado + 0), (Camera.UltimoY_Posicionado + (3 / ZoomY)), 1, , , CriarPoligono
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + (191 / ZoomX)), (Camera.UltimoY_Posicionado + (93 / ZoomY)), 15, horizontal, , CriarPoligono, -20, 95
DesenhePistaCurvaALTADIR chem_vi, (Camera.UltimoX_Posicionado + 0), (Camera.UltimoY_Posicionado - 0), 1, , , CriarPoligono

'DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado - 385)  , (Camera.UltimoY_Posicionado + 237  ), 5, vertical,

DesenhePistaCurvaBAIXAESQold chem_vi, (-219) / ZoomX, (6) / ZoomY, 1, , , CriarPoligono

'desenha já emendando
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + (162 / ZoomX)), (Camera.UltimoY_Posicionado + (163 / ZoomY)), 24, horizontal, , CriarPoligono, -5, 80, -35, 90, 5
DesenheRampinha_Horizontal chem_vi, (Camera.UltimoX_Posicionado + (20 / ZoomX)), (Camera.UltimoY_Posicionado - (58 / ZoomY)), 1, vertical, , CriarPoligono

DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + (240 / ZoomX)), (Camera.UltimoY_Posicionado + (124 / ZoomY)), 4, horizontal, , CriarPoligono, -30
'emenda horizontal
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado - (80 / ZoomX)), (Camera.UltimoY_Posicionado - (18 / ZoomY)), 1, horizontal, , CriarPoligono, -30
DesenhePistaCurvaBAIXADIR chem_vi, (Camera.UltimoX_Posicionado - (15 / ZoomX)), (Camera.UltimoY_Posicionado - (2 / ZoomY)), 1, , , CriarPoligono, , , , 274
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + (353 / ZoomX)), (Camera.UltimoY_Posicionado + (19 / ZoomY)), 12, vertical, , CriarPoligono, -165, 55
DesenheRampinha_Vertical chem_vi, (Camera.UltimoX_Posicionado + (33 / ZoomX)), (Camera.UltimoY_Posicionado - (41 / ZoomY)), 1, vertical, , CriarPoligono
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + (127 / ZoomX)), (Camera.UltimoY_Posicionado + (8 / ZoomY)), 7, vertical, , CriarPoligono, 75, 30, 0
DesenhePistaRampaVerticalDIR chem_vi, (Camera.UltimoX_Posicionado + (17 / ZoomX)), (Camera.UltimoY_Posicionado - (23 / ZoomY)), 1, vertical, , CriarPoligono
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + (147 / ZoomX)), (Camera.UltimoY_Posicionado + (18 / ZoomY)), 6, vertical, , CriarPoligono, , , 0

'criar os checkpoints e a largada (sempre por ultimo)
'  21
'15  3
'  9
'
If CriarPoligono = True Then
    CriarLinhadeLargada 42
    CreateCheckPoint 0, 21, 1
    CreateCheckPoint 6, 21, 2
    CreateCheckPoint 18, 3, 3
    CreateCheckPoint 61, 9, 4
    CreateCheckPoint 45, 15, 5
    AllChocksCreated = True
End If

End Function
Public Function MoverCarro(jogador As Player_Stats, ByVal velocidade As Double, ByVal angulo As Double, Optional ByVal pulo_altura As Long = 0, Optional ByVal CriarPoligono As Boolean = False)
  
    'pista tem 25º
    'calculo de x=sin do angulo
    'calculo de y=cos do angulo
    'angulo = 380
    '20=0
    '21=15
velocidade = velocidade * 7 / FPS

        jogador.position.x = jogador.position.x - (velocidade * Cos(anguloPI(angulo)))
        jogador.position.y = jogador.position.y - (velocidade * Sin(anguloPI(angulo)))
        If PistaAtual = 0 Then
            If AllChocksCreated = False Then
                DesenhePista chem_vi, 0, 0, True
            Else
                DesenhePista chem_vi, 0, 0
            End If
        End If
        If PistaAtual = 1 Then
            If AllChocksCreated = False Then
                DesenhePista2 chem_vi, 0, 0, True
            Else
                DesenhePista2 chem_vi, 0, 0
            End If
        End If
        If PistaAtual = 2 Then
            If AllChocksCreated = False Then
                DesenhePista3 chem_vi, 0, 0, True
            Else
                DesenhePista3 chem_vi, 0, 0
            End If 'DesenhePista chem_vi, 0, 0
        End If
        
        'DesenhePista2 chem_vi, 0, 0
        'DesenhePista3 chem_vi, 0, 0
        Exit Function

    
End Function

Public Function anguloPI(ByVal angulo As Double) As Double
Dim PI As Double
PI = 3.141592653
anguloPI = angulo * (PI / CDbl(180))
End Function
Public Function GetAngleFromCarImage(ByVal indice As Long, Optional Rampa As Long = 0) As Double
    Select Case indice
    Case 10
        GetAngleFromCarImage = 350
    Case 11
        GetAngleFromCarImage = 356
    Case 12
        GetAngleFromCarImage = 0
    Case 13
        GetAngleFromCarImage = 2
    Case 14
        GetAngleFromCarImage = 5
    Case 15
        GetAngleFromCarImage = 21.7
    Case 16
        GetAngleFromCarImage = 29
    Case 17
        GetAngleFromCarImage = 52
    Case 18
        GetAngleFromCarImage = 90
    Case 19
        GetAngleFromCarImage = 127
    Case 20
        GetAngleFromCarImage = 148
    Case 21
        
        If poligono(CarroEstaNoPoligono).piso = vertical Or poligono(CarroEstaNoPoligono).piso = horizontal Then
            GetAngleFromCarImage = 165.5
        Else
            GetAngleFromCarImage = 166
        End If
    Case 22
        GetAngleFromCarImage = 173
    Case 23
        GetAngleFromCarImage = 177
    Case 0
        GetAngleFromCarImage = 180
    Case 1
        GetAngleFromCarImage = 184
    Case 2
        GetAngleFromCarImage = 188
    Case 3
        GetAngleFromCarImage = 201.5
    Case 4
        GetAngleFromCarImage = 204
    Case 5
        GetAngleFromCarImage = 214
    Case 6
        GetAngleFromCarImage = 270
    Case 7
        GetAngleFromCarImage = 285
    Case 8
        GetAngleFromCarImage = 330
    Case 9
        GetAngleFromCarImage = 345.5
    End Select
'corrige angulo para ser mostrada nos poligonos
GetAngleFromCarImage = GetAngleFromCarImage - 75.5
End Function

'Public Function CriarChock(ByVal linear As Double, ByVal x As Long, ByVal top As Long, ByVal bottom As Long, Optional Deslocamento As Long = 0, Optional ByVal AEDireitaEPista As Boolean = False, Optional ByVal PuloAltura As Long = 0)
    'pista tem 25º
    'calculo de x=sin do angulo
    'calculo de y=cos do angulo
    'angulo = 380
    '20=0
    '21=15
'angulo = Jogador.position.angulo
'Exit Function
'Dim a As Long

 '       For a = top To bottom
        
  '      PistaChockArea(PonteiroChock).x = (x + ((a - top) * linear)) + Deslocamento
        'x = x + ((a - top) * -1)
   '     PistaChockArea(PonteiroChock).y = a
         '= x
    '    PistaChockArea(PonteiroChock).z = PuloAltura
     '   PistaChockArea(PonteiroChock).AtRightIsIn = AEDireitaEPista
        

        'DisplaySprite pontinho, CLng(PistaChockArea(PonteiroChock).x), CLng(PistaChockArea(PonteiroChock).y)
      '  PonteiroChock = PonteiroChock + 1

        
      '  Next a
        'End
        
    ' Exit Function
    
'End Function


Public Function MouseX() As Long
    Dim point As POINTAPI
   GetCursorPos point
   MouseX = point.x
End Function

Public Function MouseY() As Long
    Dim point As POINTAPI
   GetCursorPos point
   MouseY = point.y
End Function

Public Sub SetStatus()
'CarroEstaNoPoligono = LastPolign
 Dim lockrect As RECT
Dim SrcDesc         As DDSURFACEDESC2
Dim x As Long
Dim cadaponto As Long
Dim nPol As Long
Dim qualquer As Long
Dim u As Long
Dim r As Long
Dim command As command


Player(x).CarroNaRampadaFrente = False
    Player(x).CarroNaRampadeTras = False
    Player(x).CarroPegouAlgo = False
    Player(x).CarroTodoFora = False
    Player(x).CarroTodoNaPista = False
    Player(x).CarroSeChocouQuinaExterna = False
    Player(x).CarroSeChocouQuinaInterna = False
'If processeAltura = True Then Exit Sub
nPol = UBound(poligono) - 1
 If RetornandoAPista = False Then
    
    If derrapar = False Then
'        MoverCarro Player(0), Player(0).velocidade, Player(0).position.angulo, , True
    Else
        'MoverCarro Player(0), Player(0).velocidade, Player(0).position.angulo - AnguloDerrapagem, , True
    End If
Else
    'MoverCarro Player(0), -5, GetAngle, , True
End If
'MoverCarro Player(0), Player(0).velocidade, Player(0).position.angulo, , True
CarroNaPista = CarroTodoNaPista
CarroNaRampa = CarroTodoNaRampa
carrofora = CarroTodoFora
For x = 0 To 0
    ''SetChockPoints Player(X)
    
    'checka todos os poligonos
    
    poligonoAchado = False
    
    If ExtendRampa = False Then
    
    For u = 0 To nPol
        Select Case poligono(u).piso
        
        Case vertical
            If PoligIn(poligono(u).indice, qualquer) = 0 Then
            If Player(0).AlturaReal >= Player(0).position.Z Then
            elevacao = 0
            Player(0).declive = reto
            End If
            CarroEstaNoPoligono = u: poligonoAchado = True: Exit For
            End If
        Case horizontal
            If PoligIn(poligono(u).indice, qualquer) = 0 Then
            If Player(0).AlturaReal >= Player(0).position.Z Then
            elevacao = 0
            Player(0).declive = reto
            End If
            
            CarroEstaNoPoligono = u: poligonoAchado = True: Exit For
            End If
        Case Rampa
            If PoligIn(poligono(u).indice, qualquer) = 0 Then
                ''Player(0).position.Z = -GetNivel(u)
                If poligono(u).tangente > 0 Then
                    If Player(0).ChockPoints(0).y <= Player(0).ChockPoints(3).y Then
                    If Player(0).AlturaReal >= Player(0).position.Z Then
                        elevacao = 1
                        Player(0).declive = subindo
                    End If
            
                    Else
                    If Player(0).AlturaReal >= Player(0).position.Z Then
                        elevacao = 3
                        Player(0).declive = descendo
                    End If
            
                    End If
                
                     CarroEstaNoPoligono = u
                     poligonoAchado = True
                    
                    Exit For
                Else
                    If Player(0).ChockPoints(0).y >= Player(0).ChockPoints(3).y Then
                    If Player(0).AlturaReal >= Player(0).position.Z Then
                        elevacao = 1
                        Player(0).declive = subindo
                    End If
            
                    Else
                        If Player(0).AlturaReal >= Player(0).position.Z Then
                            elevacao = 3
                            Player(0).declive = descendo
                        End If
                          
                                      
                    End If
                    CarroEstaNoPoligono = u
                     poligonoAchado = True
                    
                     Exit For
                End If
            End If
        Case RampaH
            If PoligIn(poligono(u).indice, qualquer) = 0 Then
                ''Player(0).position.Z = -GetNivel(u)
                If poligono(u).tangente > 0 Then
                    If Player(0).ChockPoints(0).y <= Player(0).ChockPoints(3).y Then
                    If Player(0).AlturaReal >= Player(0).position.Z Then
                        elevacao = 1
                        Player(0).declive = subindo
                    End If
            
                    Else
                    If Player(0).AlturaReal >= Player(0).position.Z Then
                        elevacao = 3
                        Player(0).declive = descendo
                    End If
            
                    End If
                
                     CarroEstaNoPoligono = u
                     poligonoAchado = True
                    
                    Exit For
                Else
                    If Player(0).ChockPoints(0).y >= Player(0).ChockPoints(3).y Then
                    If Player(0).AlturaReal >= Player(0).position.Z Then
                        elevacao = 1
                        Player(0).declive = subindo
                    End If
            
                    Else
                    If Player(0).AlturaReal >= Player(0).position.Z Then
                        elevacao = 3
                        Player(0).declive = descendo
                    End If
                           
                                   
                    End If
                    CarroEstaNoPoligono = u
                     poligonoAchado = True
                     
                     Exit For
                End If
            End If
        
        Case sObjeto
        
        Case Else
        
        End Select
    Next u
    Else
    
    End If

    'verifica se o poligono encontrado é largada ou check point
    If poligono(CarroEstaNoPoligono).IsCheckPoint = True Then
        'procura o check anterior
        If CheckPoints(CarroEstaNoPoligono).index <> 1 Then
            For u = 0 To 999
                If CheckPoints(u).index = CheckPoints(CarroEstaNoPoligono).index - 1 Then
                    If CheckPoints(u).checked = True Then CheckPoints(CarroEstaNoPoligono).checked = True
                End If
            Next u
        Else
            CheckPoints(CarroEstaNoPoligono).checked = True
        End If
        
        LastCheckPoint = CarroEstaNoPoligono
    End If
    
    If poligono(CarroEstaNoPoligono).IsLagada = True Then
        
        For r = 0 To UBound(poligono) - 1
            'vem que é check para ser chekado
            If poligono(r).IsCheckPoint = True Then
                If CheckPoints(r).checked = False Then GoTo np1
            End If
        Next r
        'soma 1 volta
        Player(0).VoltasDadas = Player(0).VoltasDadas + 1
        'renova as armas
        'registra volta
        command.command = RegisterLap
        r = 0
        Do
            DoEvents
            If FrmDirectX.Winsock1.Tag <> "conectado" Then Exit Do
        Loop Until send(FrmDirectX.Winsock1.SocketHandle, ByVal VarPtr(command), GetSizeToSend(command.command), 0) <> SOCKET_ERROR
        'ReDim b(0 To GetSizeToSend(command.command) - 1)
        'CopyMemory b(0), command, GetSizeToSend(command.command)
        'FrmDirectX.Winsock1.SendData b
        
        Player(0).armas.Traseiras.quantidade = Player(0).armas.Traseiras.QuantasTenho
        Player(0).armas.Frontais.quantidade = Player(0).armas.Frontais.QuantasTenho
        For r = 0 To 999
            CheckPoints(r).checked = False
        Next r
    End If

np1:

    If RetornandoAPista = True Then
        DDraw.ClearBuffer
        
        If ContpixelAposRetornar >= 6 Then
            FrmDirectX.tmrexplosao.Interval = 0
            FrmDirectX.tmrexplosaowait.Interval = 0
            ContpixelAposRetornar = 0
            RetornandoAPista = False
            Player(0).CarroExplodiu = False
            Player(0).CarroVaiExplodir = False
            Player(0).velocidade = 0
            
            Exit Sub
        End If
            'verifica se no retorno o carro esta todo na pista
            If CarroNaPista = True Then ContpixelAposRetornar = ContpixelAposRetornar + 1: Exit Sub
    End If

    
    If CarroNaPista = True Then GetLastValidImageIndex = Player(0).Car_Image_Index
    

    '//carro explodiu
    ''If Player(0).CarroExplodiu = False And Player(0).CarroVaiExplodir = False And RetornandoAPista = False And jump <= 0 And Player(0).position.altura <= 0 Then
      
     ''End If
    'If Player(0).CarroExplodiu = False And Player(0).CarroVaiExplodir = False And RetornandoAPista = False And jump <= 0 And Player(0).position.altura <= 0 Then
        'If carrofora = True Then FrmDirectX.tmrexplosaowait.Interval = 1500
    'End If
    
    If Player(0).CarroExplodiu = False And Player(0).CarroVaiExplodir = False And FrmDirectX.tmrexplosaowait.Interval = 0 And jump <= 0 And Player(0).position.altura <= 0 Then
        'If carrofora = True Then Player(0).CarroVaiExplodir = True: GetAngle = GetAngleFromCarImage(GetLastValidImageIndex): FrmDirectX.tmrexplosaowait.Interval = 1500
    End If

If Player(0).CarroExplodiu = True Or Player(0).CarroVaiExplodir = True Or RetornandoAPista = True Or (processeAltura = True And Player(0).CarroSaltouPelaRampa = True) Then DDraw.ClearBuffer: Exit Sub
    
       
    
    
    If CarroNaRampa = 1 Then Player(0).CarroNaRampadeTras = False: GoTo 13
    If CarroNaRampa = 2 Then GoTo 13
    
13:     If CarroNaPista Then
        Player(x).CarroTodoNaPista = True
        UltimoY = Player(x).position.y
        SimularCarroX = 0
        SimularCarroY = 0
        ''Player(x).CarroNaRampadaFrente = False
        ''Player(x).CarroNaRampadeTras = False
    End If
If Player(0).CarroSaltouPelaRampa = True Then Exit Sub
    
NextStep1:
'10 seria a altura da calçada

If processeAltura = True Then
    Player(x).CarroSeChocouQuinaExterna = False
    Player(x).CarroSeChocouQuinaInterna = False
    DDraw.ClearBuffer
    Exit Sub
End If
If carrofora = True Then Exit Sub

   'If pixel(0) = 0 And pixel(1) = 0 And pixel(2) = 0 And pixel(3) = 0 Then Player(x).CarroTodoFora = True
        'If (pixel(0) = 0 And pixel(1) = 65503) Or (pixel(2) = 65503 And pixel(3) = 0) Then Player(x).CarroSeChocouQuinaExterna = True: Player(x).CarroNaRampadaFrente = True
        'If (pixel(0) = 65503 And pixel(1) = 0) Or (pixel(2) = 0 And pixel(3) = 65503) Then Player(x).CarroSeChocouQuinaInterna = True: Player(x).CarroNaRampadaFrente = True
        Dim Asomar As Long
        'If Player(0).Car_Image_Index = 6 Or Player(0).Car_Image_Index = 5 Or Player(0).Car_Image_Index = 4 Then
         '   Asomar = -1
        'Else
         '   Asomar = 1
        'End If
        
    If Player(x).CarroSeChocouQuinaExterna = True Or Player(x).CarroSeChocouQuinaInterna = True Then
        'corner = 2 padrao
        If Abs(Player(0).velocidade) > 1.428 Then Player(0).velocidade = Player(0).velocidade / Player(0).TopCorner
        If Player(0).velocidade > 0 And Player(0).velocidade < 1.428 Then Player(0).velocidade = 1.428
        If Player(0).velocidade < 0 Then Player(0).velocidade = 0.1571
            'Player(0).velocidade = 0.2
'        If Player(0).declive <> reto Then Player(0).velocidade = 0
        Select Case LastVerticeChocked
            Case 1 'vertice 0
            Player(x).Car_Image_Index = Player(x).Car_Image_Index + 1
            If Player(x).Car_Image_Index >= 24 Then Player(x).Car_Image_Index = 0
            If derrapar = False Then
                MoverCarro Player(0), -Player(0).velocidade * 4, Player(0).position.angulo - 45
            Else
                MoverCarro Player(0), -Player(0).velocidade * 4, AnguloDerrapagem - 45
            End If
            
            Case 2 'vertice 1
            Player(x).Car_Image_Index = Player(x).Car_Image_Index + 1
            If Player(x).Car_Image_Index >= 24 Then Player(x).Car_Image_Index = 0
            If derrapar = False Then
                MoverCarro Player(0), -Player(0).velocidade * 4, Player(0).position.angulo - 45
            Else
                    MoverCarro Player(0), -Player(0).velocidade * 4, AnguloDerrapagem - 45
            End If
            
            Case 3 'vertice 2
            Player(x).Car_Image_Index = Player(x).Car_Image_Index - 1
            If Player(x).Car_Image_Index < 0 Then Player(x).Car_Image_Index = 23
            If derrapar = False Then
                MoverCarro Player(0), Player(0).velocidade * 4, Player(0).position.angulo - 45
            Else
                MoverCarro Player(0), Player(0).velocidade * 4, AnguloDerrapagem - 45
            End If
            
            Case 4 'vertice 3
            Player(x).Car_Image_Index = Player(x).Car_Image_Index - 1
            If Player(x).Car_Image_Index < 0 Then Player(x).Car_Image_Index = 23
            If derrapar = False Then
                MoverCarro Player(0), Player(0).velocidade * 4, Player(0).position.angulo - 45
            Else
                    MoverCarro Player(0), Player(0).velocidade * 4, AnguloDerrapagem - 45
            End If
        End Select
            
        
            'Player(x).Car_Image_Index = Player(x).Car_Image_Index - 1 * Asomar
            'If Player(x).Car_Image_Index = -1 Then Player(x).Car_Image_Index = 23
            'If Player(x).Car_Image_Index = -2 Then Player(x).Car_Image_Index = 22
        
                
            GoTo NextStep
    End If
    
    
NextStep:
    'If Player(x).CarroSeChocouQuinaExterna = False And Player(x).CarroSeChocouQuinaInterna = False Then Player(0).velocidade = 5
    'If Player(x).CarroSeChocouQuinaExterna = False And Player(x).CarroSeChocouQuinaInterna = False Then Player(0).velocidade = 1
    'DDraw.ClearBuffer
    
    
    'If Player(0).CarroSeChocouQuinaExterna = True Then Do: DoEvents: Loop
Next x
End Sub

Public Function ChangeColors(tSprite As Sprites, ByVal ColorSrc As Long, ByVal ColorDest As Long, mapcolor() As pixelCoord)
 Dim lockrect As RECT
Dim SrcDesc         As DDSURFACEDESC2
Dim pixel(0 To 3) As Long
Dim x As Long
Dim altura As Long
Dim largura As Long
    
    tSprite.imagem.GetSurfaceDesc SrcDesc
    lockrect.Right = SrcDesc.lWidth
    lockrect.Bottom = SrcDesc.lHeight
    'tSprite.imagem.Lock lockrect, SrcDesc, DDLOCK_WAIT Or DDLOCK_NOSYSLOCK, 0
    Dim u As Long
    
 '   If ColorSrc = RGB(0, 255, 0) Then
     If UBound(mapcolor) <= 0 Then
     'For largura = 0 To tSprite.Width - 1
    For altura = 0 To tSprite.Height - 1
        tSprite.imagem.Lock lockrect, SrcDesc, DDLOCK_WAIT Or DDLOCK_NOSYSLOCK, 0
        For largura = 0 To tSprite.Width - 1
            'DoEvents
            
            u = tSprite.imagem.GetLockedPixel(largura, altura)
            If u = ColorSrc Then tSprite.imagem.SetLockedPixel largura, altura, ColorDest
            
        Next largura
        tSprite.imagem.Unlock lockrect
    Next altura
    
  Else
  For altura = 0 To UBound(mapcolor)
    tSprite.imagem.Lock lockrect, SrcDesc, DDLOCK_WAIT Or DDLOCK_NOSYSLOCK, 0
    tSprite.imagem.SetLockedPixel mapcolor(altura).x, mapcolor(altura).y, ColorDest
    tSprite.imagem.Unlock lockrect
  Next altura
    End If
    
End Function

   Public Function UnsignedToInteger(Value As Long) As Integer
        If Value < 0 Or Value >= OFFSET_2 Then Error 6 ' Overflow
        If Value <= MAXINT_2 Then
          UnsignedToInteger = Value
        Else
          UnsignedToInteger = Value - OFFSET_2
        End If
      End Function
Public Function PegarYCarroNaRampa(jogador As Player_Stats, ByVal velocidade As Double, Optional ByVal pulo_altura As Long = 0) As Double
'        Dim x As Long
 '       Dim y As Long
        Dim angulo As Double
        Dim carImage As Long
        carImage = jogador.Car_Image_Index - 4
        If carImage = -1 Then carImage = 23
        If carImage = -2 Then carImage = 22
        If carImage = -3 Then carImage = 21
        If carImage = -4 Then carImage = 20
        SimularCarroX = SimularCarroX - (velocidade * Cos(anguloPI(GetAngleFromCarImage(jogador.Car_Image_Index) - 24)))
        SimularCarroY = SimularCarroY - (velocidade * Sin(anguloPI(GetAngleFromCarImage(jogador.Car_Image_Index) - 24)))
        PegarYCarroNaRampa = SimularCarroY



End Function

Public Function PegarYCarroNaLadeira(jogador As Player_Stats, ByVal velocidade As Double, Optional ByVal pulo_altura As Long = 0) As Double
'        Dim x As Long
 '       Dim y As Long
        Dim angulo As Double
        Dim carImage As Long
        carImage = jogador.Car_Image_Index - 4
        If carImage = 24 Then carImage = 0
        If carImage = 25 Then carImage = 1
        If carImage = 26 Then carImage = 2
        If carImage = 27 Then carImage = 3
        SimularCarroX = SimularCarroX - (velocidade * Cos(anguloPI(GetAngleFromCarImage(jogador.Car_Image_Index) - 24)))
        SimularCarroY = SimularCarroY + (velocidade * Sin(anguloPI(GetAngleFromCarImage(jogador.Car_Image_Index) - 24)))
        PegarYCarroNaLadeira = SimularCarroY



End Function


Public Function desenhepista3rampa(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = horizontal, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False) As Boolean
If tamanho < 1 Then desenhepista3rampa = False: Exit Function
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(6), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
    
     If CriarPoligono Then
        a = 94.5 'altura do quadrado
        a = a * tamanho
        CreatePoligono 200, a, Retangular, Rampa, 0.325 'descendo
     End If
     If CriarPoligono Then
        a = 140.5 'altura do quadrado
        a = a * tamanho
        CreatePoligono 200, a, Retangular, Rampa, -0.637  'descendo
     End If
'    DesenheRampadaPista chem_vi, x + ((950/zoomx) / Screen.TwipsPerPixelX * (p + 1))  , y - ((250/zoomy) / Screen.TwipsPerPixelY * (p + 1))  , 1
Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function

Public Function DesenhePistaCurvaALTAESQold(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = horizontal, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False, Optional ByVal nShiftX As Long = 6, Optional ByVal xInicial As Long = 16, Optional ByVal yInicial As Long = -2032, Optional ByVal nShiftY As Long = 0) As Boolean
If tamanho < 1 Then DesenhePistaCurvaALTAESQold = False: Exit Function
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(8), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
      If CriarPoligono Then
        a = 54.5 'altura do quadrado
        a = a * tamanho
        CreatePoligono 194, a, Retangular, vertical, , 6
        CreatePoligono 188, a, Retangular, vertical, , 6
        CreatePoligono 182, a * 2, Retangular, vertical, , 6
        CreatePoligono 176, a, Retangular, vertical, , 6
        CreatePoligono 170, a, Retangular, vertical, , 6
        CreatePoligono 160, a - 10, Retangular, vertical, , 10
        CreatePoligono 130, a - 20, Retangular, vertical, , 30
        CreatePoligono 90, a - 28, Retangular, vertical, , 40
        a = 330
        CreatePoligono 260, a, Retangular, vertical, , 190, 16, -2032
    End If
   
'    DesenheRampadaPista chem_vi, x + ((950/zoomx) / Screen.TwipsPerPixelX * (p + 1))  , y - ((250/zoomy) / Screen.TwipsPerPixelY * (p + 1))  , 1
Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select

End Function


Public Function DesenheRampaHorizontalold(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = vertical, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False, Optional ExtraHeight As Long = 0, Optional ExtraWidth As Long = 0, Optional deslocamento As Long = 0, Optional ByVal xInit As Long = 0, Optional ByVal yInit As Long = 0, Optional ByVal nShiftY As Long = 0, Optional ByVal nivelacao As Double = -1) As Boolean
If tamanho < 1 Then DesenheRampaHorizontalold = False: Exit Function
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(12), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
    
        If CriarPoligono Then
            a = 120.5 'largura do quadrado
            a = a * (tamanho)
            CreatePoligono a, 329, Retangular, RampaH, 0.66, -10, xInit, yInit, , nivelacao, , , nShiftY
        End If
    Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
    Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function

Public Function DesenhePista2Horizontal(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = vertical, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False, Optional ExtraHeight As Long = 0, Optional ExtraWidth As Long = 0, Optional deslocamento As Long = 0, Optional ByVal xInit As Long = 0, Optional ByVal yInit As Long = 0, Optional ByVal nShiftY As Long = 0, Optional ByVal nivelacao As Double = -1) As Boolean
If tamanho < 1 Then DesenhePista2Horizontal = False
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(14), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
    
        If CriarPoligono Then
            a = 46.5 'largura do quadrado
            a = a * (tamanho)
            CreatePoligono a, 329, Retangular, RampaH, 0.65, deslocamento, xInit, yInit, , nivelacao, , , nShiftY
        End If
        
        If CriarPoligono Then
            a = 90.5 'largura do quadrado
            a = a * (tamanho)
            CreatePoligono a, 329, Retangular, RampaH, -0.6, deslocamento, xInit, yInit, , nivelacao, , , nShiftY
        End If
    Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
    Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function

Public Function DesenhePistaCurvaALTADIR(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = horizontal, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False, Optional ByVal nShiftX As Long = 0, Optional ByVal xInicial As Long = 0, Optional ByVal yInicial As Long = 0, Optional ByVal nShiftY As Long = 0) As Boolean
If tamanho < 1 Then DesenhePistaCurvaALTADIR = False: Exit Function
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(16), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
    
      If CriarPoligono Then
        a = 54.5 'altura do quadrado
        a = a * tamanho
        CreatePoligono 90, a - 28, Retangular, horizontal, , nShiftX, xInicial, yInicial, , , , , nShiftY
        CreatePoligono 90, a - 20, Retangular, vertical, , , , , True
        CreatePoligono 120, a - 10, Retangular, vertical, , , , , True
        CreatePoligono 150, a, Retangular, vertical, , , , , True
        CreatePoligono 180, a, Retangular, vertical, , , , , True
        CreatePoligono 180, a, Retangular, vertical, , , , , True
        CreatePoligono 200, a, Retangular, vertical, , , , , True
        CreatePoligono 200, a * 2, Retangular, vertical, , , , , True
        CreatePoligono 200, a, Retangular, vertical, , , , , True
        CreatePoligono 200, a, Retangular, vertical, , , , , True
       
End If
'    DesenheRampadaPista chem_vi, x + ((950/zoomx) / Screen.TwipsPerPixelX * (p + 1))  , y - ((250/zoomy) / Screen.TwipsPerPixelY * (p + 1))  , 1
Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function


Public Function DesenhePistaRampaVerticalDIR(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = horizontal, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False) As Boolean
If tamanho < 1 Then DesenhePistaRampaVerticalDIR = False: Exit Function
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(18), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
    
     If CriarPoligono Then
        CreatePoligono 200, 90, Retangular, Rampa, -0.4925, 30
    End If
'    DesenheRampadaPista chem_vi, x + ((950/zoomx) / Screen.TwipsPerPixelX * (p + 1))  , y - ((250/zoomy) / Screen.TwipsPerPixelY * (p + 1))  , 1
Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function


Public Function DesenhePistaCurvaBAIXAESQold(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = horizontal, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False) As Boolean
If tamanho < 1 Then DesenhePistaCurvaBAIXAESQold = False: Exit Function
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(20), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
    If CriarPoligono Then
        a = 54.5 'altura do quadrado
        a = a * tamanho
        CreatePoligono 200, a, Retangular, vertical, , 0, 10, 39, , 0
        CreatePoligono 194, a, Retangular, vertical, , 6, , , True
        CreatePoligono 188, a * 2, Retangular, vertical, , 6, , , True
        CreatePoligono 226, a, Retangular, vertical, , 10, , , True
        CreatePoligono 200, a, Retangular, vertical, , 17, , , True
        CreatePoligono 190, a - 10, Retangular, vertical, , 17, , , True
        CreatePoligono 160, a - 20, Retangular, vertical, , 24, , , True
        CreatePoligono 140, a - 28, Retangular, vertical, , 24, , , True
        CreatePoligono 140, a - 28, Retangular, vertical, , , , , True
        CreatePoligono 140, a - 28, Retangular, vertical, , , , , True
        ''a = 330
        ''CreatePoligono 260, a, Retangular, vertical, , 190, 16, -2032
    End If
'    DesenheRampadaPista chem_vi, x + ((950/zoomx) / Screen.TwipsPerPixelX * (p + 1))  , y - ((250/zoomy) / Screen.TwipsPerPixelY * (p + 1))  , 1
Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function

Public Function DesenheRampinha_Horizontal(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = horizontal, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False, Optional ByVal DrawOnlyFirstPolig As Boolean = False) As Boolean
If tamanho < 1 Then DesenheRampinha_Horizontal = False: Exit Function
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(22), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
    
        If CriarPoligono Then
            a = 136.5 'largura do quadrado
            a = a * (tamanho)
            CreatePoligono a, 329, Retangular, RampaH, 0.75, -7
        End If
        
        If CriarPoligono And DrawOnlyFirstPolig = False Then
            a = 60.5 'largura do quadrado
            a = a * (tamanho)
            CreatePoligono a, 329, Retangular, RampaH, -0.6
        End If
'    DesenheRampadaPista chem_vi, x + ((950/zoomx) / Screen.TwipsPerPixelX * (p + 1))  , y - ((250/zoomy) / Screen.TwipsPerPixelY * (p + 1))  , 1
Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function


Public Function DesenhePistaCurvaBAIXADIR(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = horizontal, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False, Optional ByVal nShiftX As Long = 0, Optional ByVal xInicial As Long = 0, Optional ByVal yInicial As Long = 0, Optional ByVal nShiftY As Long = 0) As Boolean
If tamanho < 1 Then DesenhePistaCurvaBAIXADIR = False: Exit Function
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(24), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
    
      If CriarPoligono Then
        a = 54.5 'altura do quadrado
        a = a * tamanho
        'CreatePoligono 90, a - 28, Retangular, horizontal, , 936, 700, -544
        CreatePoligono 90, a - 28, Retangular, horizontal, , nShiftX, xInicial, yInicial, , , , , nShiftY
        CreatePoligono 90, a - 20, Retangular, vertical
        CreatePoligono 120, a - 10, Retangular, vertical
        CreatePoligono 150, a, Retangular, vertical
        CreatePoligono 190, a, Retangular, vertical
        CreatePoligono 220, a, Retangular, vertical
        CreatePoligono 250, a, Retangular, vertical
        CreatePoligono 250, a * 2, Retangular, vertical
        CreatePoligono 250, a, Retangular, vertical
        CreatePoligono 250, a, Retangular, vertical
       
End If
'    DesenheRampadaPista chem_vi, x + ((950/zoomx) / Screen.TwipsPerPixelX * (p + 1))  , y - ((250/zoomy) / Screen.TwipsPerPixelY * (p + 1))  , 1
Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function

Public Function DesenheRampinha_Vertical(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = horizontal, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False) As Boolean
If tamanho < 1 Then DesenheRampinha_Vertical = False: Exit Function
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(26), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
    
     If CriarPoligono Then
        a = 70.5 'altura do quadrado
        a = a * (tamanho)
        CreatePoligono 240, a, Retangular, Rampa, 0.625, 30
        CreatePoligono 240, 70, Retangular, vertical, 30
        CreatePoligono 240, 50, Retangular, Rampa, -0.7425, 0
    
    End If
    
'    DesenheRampadaPista chem_vi, x + ((950/zoomx) / Screen.TwipsPerPixelX * (p + 1))  , y - ((250/zoomy) / Screen.TwipsPerPixelY * (p + 1))  , 1
Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function

Public Sub SetarOrigem()
'origem do pulo
            '1 = reta branca
            '2 = reta amarela
            '3 = pulando rampa
            '4 = descendo rampa
            If Player(0).CarroTodoNaPista = True And CarroNaPista = True Then OrigemdoPulo = 1
            ''If Player(0).CarroTodoNaPista = True And pixel(0) = amarelo Then OrigemdoPulo = 2
             
             If Player(0).CarroNaRampadaFrente = True Then
                If (Player(0).Car_Image_Index > 18 And Player(0).Car_Image_Index <= 23) Or Player(0).Car_Image_Index = 0 Or Player(0).Car_Image_Index = 1 Or Player(0).Car_Image_Index = 2 Or Player(0).Car_Image_Index = 3 Or Player(0).Car_Image_Index = 4 Then
                    OrigemdoPulo = 3
                Else
                    OrigemdoPulo = 4
                End If
            End If
        
           If Player(0).CarroNaRampadeTras = True Then
            
                If (Player(0).Car_Image_Index > 18 And Player(0).Car_Image_Index <= 23) Or Player(0).Car_Image_Index = 0 Or Player(0).Car_Image_Index = 1 Or Player(0).Car_Image_Index = 2 Or Player(0).Car_Image_Index = 3 Or Player(0).Car_Image_Index = 4 Then
                    OrigemdoPulo = 4
                Else
                    OrigemdoPulo = 3
                End If
            End If
        
        
End Sub
'Public Function ChangeColor(tSprite As Sprites, ByVal ColorSrc As Long, ByVal ColorDest As Long)
'Dim lockrect As RECT
'Dim SrcDesc As DDSURFACEDESC2
'Dim pixel As Long
'Dim x As Long
'Dim y As Long
'tSprite.imagem.GetSurfaceDesc SrcDesc
'lockrect.Right = SrcDesc.lWidth
'lockrect.Bottom = SrcDesc.lHeight
'tSprite.imagem.Lock lockrect, SrcDesc, DDLOCK_WAIT Or DDLOCK_NOSYSLOCK, 0

'For x = 1 To tSprite.width
 '   For y = 1 To tSprite.height
  '      pixel = tSprite.imagem.GetLockedPixel(x, y)
   '     If pixel = ColorSrc Then tSprite.imagem.SetLockedPixel x, y, ColorDest
'    Next y
'Next x

'tSprite.imagem.Unlock lockrect

'End Function

Public Function PuloReal(ByVal ImageIndex As Long, ByVal AlturaAtual As Double) As Double

Select Case ImageIndex
Case 21
    PuloReal = AlturaAtual * 1
Case 22
    PuloReal = AlturaAtual * 0.832
Case 23
    PuloReal = AlturaAtual * 0.637
Case 0
    PuloReal = AlturaAtual * 0.5
Case 1
    PuloReal = AlturaAtual * 0.637
Case 2
    PuloReal = AlturaAtual * 0.832
Case 3
    PuloReal = AlturaAtual * 1
Case 4
    PuloReal = AlturaAtual * 0.832
Case 5
    PuloReal = AlturaAtual * 0.637
Case 6
    PuloReal = AlturaAtual * 0.5
Case 7
    PuloReal = AlturaAtual * 0.637
Case 8
    PuloReal = AlturaAtual * 0.832
Case 9
    PuloReal = AlturaAtual * 1
Case 10
    PuloReal = AlturaAtual * 0.832
Case 11
    PuloReal = AlturaAtual * 0.637
Case 12
    PuloReal = AlturaAtual * 0.5
Case 13
    PuloReal = AlturaAtual * 0.637
Case 14
    PuloReal = AlturaAtual * 0.832
Case 15
    PuloReal = AlturaAtual * 1
Case 16
    PuloReal = AlturaAtual * 0.832
Case 17
    PuloReal = AlturaAtual * 0.637
Case 18
    PuloReal = AlturaAtual * 0.5
Case 19
    PuloReal = AlturaAtual * 0.637
Case 20
    PuloReal = AlturaAtual * 0.832
    

End Select
End Function
Public Function CreatePoligono(ByVal Width As Double, ByVal Height As Double, ByVal tipo As Long, ByVal p As PolignType, Optional ByVal tangente As Double = 0, Optional ByVal nShiftX As Long = 0, Optional ByVal xInicial As Long = 0, Optional ByVal yInicial As Long = 0, Optional ByVal Reverse As Boolean = False, Optional ByVal nivelacao As Double = -1, Optional ByVal IsLargada As Boolean = False, Optional ByVal IsCheckPoint As Boolean, Optional ByVal yShift As Long = 0, Optional ByVal ERampaStep As Boolean = False, Optional ByVal dontChangeLastPositions As Boolean = False) As Long

''Xinicial e y= vertice 3
'tipo 0 = retangular - 4 vertices
''nivelinicial = nivel em z da linha vertice 3 ate 2
Dim ToUse As Long
Dim x As Long
'If xInicial = 0 Then xInicial = Camera.UltimoX_Poligono
'If yInicial = 0 Then yInicial = Camera.UltimoY_Poligono
'xInicial = xInicial + Camera.UltimoX_Poligono
'yInicial = yInicial + Camera.UltimoY_Poligono
'01
'32

Select Case tipo
Case 0
    ToUse = UBound(poligono)
    ReDim Preserve poligono(0 To UBound(poligono) + 1)
    
    With poligono(ToUse)
        .IsRampaStep = ERampaStep
        .indice = ToUse
        ''For x = 0 To 3
          ''  .pos(x).x = vertices2D(x).x
           '' .pos(x).y = vertices2D(x).y
        ''Next x
        If p = horizontal Or p = RampaH Then
            nShiftX = nShiftX + lastWidth
            If xInicial = 0 And yInicial = 0 Then
                .pos(0).x = Camera.UltimoX_Poligono + xInicial + nShiftX
                .pos(0).y = (Camera.UltimoY_Poligono + yInicial - (nShiftX * 1.356097)) + yShift
            Else
                .pos(0).x = xInicial + nShiftX
                .pos(0).y = yInicial - (nShiftX * 1.356097) + yShift
            End If
            .pos(3).x = .pos(0).x
            .pos(3).y = .pos(0).y + Height
        ''largura
            .pos(2).x = .pos(3).x + Width
            .pos(2).y = .pos(3).y - Abs((.pos(2).x - .pos(3).x)) * 1.356097
        
            .pos(1).x = .pos(2).x
            .pos(1).y = .pos(2).y - Height
            
            
            .piso = p
        
            If poligono(ToUse).piso = Rampa Or poligono(ToUse).piso = RampaH Then
                .tangente = tangente
            Else
                .tangente = 0
            End If
        Else
        
             'vertical
            If Reverse = False Then
                If xInicial = 0 And yInicial = 0 Then
                    .pos(3).x = Camera.UltimoX_Poligono + xInicial + nShiftX
                    .pos(3).y = Camera.UltimoY_Poligono + yInicial - (nShiftX * 1.356097) + yShift
                Else
                    .pos(3).x = xInicial + nShiftX
                    .pos(3).y = yInicial - (nShiftX * 1.356097) + yShift
                End If
            End If
            If Reverse = True Then
                If xInicial = 0 And yInicial = 0 Then
                    .pos(3).x = poligono(LastPolignCreated).pos(0).x + xInicial + nShiftX
                    .pos(3).y = poligono(LastPolignCreated).pos(0).y + lastHeight + Height + yInicial - (nShiftX * 1.356097) - 2
                Else
                    .pos(3).x = xInicial + nShiftX
                    .pos(3).y = yInicial - (nShiftX * 1.356097)
                End If
            End If

        ''altura
            .pos(0).y = .pos(3).y - Height
            .pos(0).x = .pos(3).x
          ''f(y)= -1,356097 * x
        ''largura
            .pos(2).x = .pos(3).x + Width
            .pos(2).y = .pos(3).y - Abs((.pos(2).x - .pos(3).x)) * 1.356097
        
            .pos(1).x = .pos(2).x
            .pos(1).y = .pos(2).y - Height

            .piso = p
        
            If poligono(ToUse).piso = Rampa Or poligono(ToUse).piso = RampaH Then
                .tangente = tangente
            Else
                .tangente = 0
            End If
        End If
        
        
        ''.NivelInicial = NivelInicial
        'localiza o poligono valido anterior
        .PoligonoAnterior = -1
        If ToUse <> 0 Then
        For x = ToUse - 1 To 0 Step -1
            If poligono(x).piso <> largada And poligono(x).piso <> checkpoint Then
                .PoligonoAnterior = x
                Exit For
            End If
        Next x
        
        If .PoligonoAnterior <> -1 Then
            .NivelInicial = poligono(.PoligonoAnterior).NivelFinal
            If nivelacao <> -1 Then
                .NivelInicial = nivelacao
                '.NivelFinal = nivelacao
            End If
    
            If poligono(ToUse).piso = RampaH Then
                .NivelFinal = .NivelInicial - Abs(poligono(ToUse).pos(0).x - poligono(ToUse).pos(1).x) * poligono(ToUse).tangente
            Else
                .NivelFinal = .NivelInicial - Abs(poligono(ToUse).pos(0).y - poligono(ToUse).pos(3).y) * poligono(ToUse).tangente
            End If
        End If
        End If
    If .NivelFinal > MenorNivel Then MenorNivel = .NivelFinal: PoligonoMaisBaixo = ToUse
    If .NivelInicial > MenorNivel Then MenorNivel = .NivelInicial: PoligonoMaisBaixo = ToUse
    If dontChangeLastPositions = False Then
    Camera.UltimoX_Poligono = .pos(0).x
    Camera.UltimoY_Poligono = .pos(0).y + 1
    lastWidth = Width
    lastHeight = Abs(Height)
    End If
    .IsLagada = IsLargada
    .IsCheckPoint = IsCheckPoint
    .altura = Abs(Height)
    .largura = Abs(Width)
    End With
    
    CreatePoligono = ToUse
If dontChangeLastPositions = False Then LastPolignCreated = ToUse
    End Select



End Function

Public Function DrawPoligono(ByVal indice As Long, ByVal x As Long, ByVal y As Long) As Boolean

If poligono(indice).Type = 0 Then
    poligono(indice).VideoPos(0).x = poligono(indice).pos(0).x + x
    poligono(indice).VideoPos(0).y = poligono(indice).pos(0).y + y
    poligono(indice).VideoPos(1).x = poligono(indice).pos(1).x + x
    poligono(indice).VideoPos(1).y = poligono(indice).pos(1).y + y
    poligono(indice).VideoPos(2).x = poligono(indice).pos(2).x + x
    poligono(indice).VideoPos(2).y = poligono(indice).pos(2).y + y
    poligono(indice).VideoPos(3).x = poligono(indice).pos(3).x + x
    poligono(indice).VideoPos(3).y = poligono(indice).pos(3).y + y
    
    If ShowGrides = True Then
        BackBuffer.DrawLine poligono(indice).VideoPos(0).x, poligono(indice).VideoPos(0).y, poligono(indice).VideoPos(1).x, poligono(indice).VideoPos(1).y
        BackBuffer.DrawLine poligono(indice).VideoPos(1).x, poligono(indice).VideoPos(1).y, poligono(indice).VideoPos(2).x, poligono(indice).VideoPos(2).y
        BackBuffer.DrawLine poligono(indice).VideoPos(2).x, poligono(indice).VideoPos(2).y, poligono(indice).VideoPos(3).x, poligono(indice).VideoPos(3).y
        BackBuffer.DrawLine poligono(indice).VideoPos(3).x, poligono(indice).VideoPos(3).y, poligono(indice).VideoPos(0).x, poligono(indice).VideoPos(0).y
    End If
    
End If
End Function

Public Function DrawAllPoligonos(Optional ByVal x As Long = 0, Optional ByVal y As Long = 0) As Boolean
Dim k As Long
Dim angulo As Double
Dim centerX As Long
Dim CenterY As Long
Dim raio As Double
For k = 0 To UBound(poligono) - 1
DrawPoligono k, -x - 50, -y + 500
Next k
'desenha o carro(4 vertices)
Player(0).ChockPoints(0).x = 29
Player(0).ChockPoints(0).y = 250
Player(0).ChockPoints(1).x = 54
Player(0).ChockPoints(1).y = 216
Player(0).ChockPoints(2).x = 54
Player(0).ChockPoints(2).y = 266
Player(0).ChockPoints(3).x = 29
Player(0).ChockPoints(3).y = 300
angulo = 0
nAng = GetAngleFromCarImage(Player(0).Car_Image_Index) - 90

angulo = 0 + nAng
angulo = angulo Mod 360
raio = 22
centerX = Player(0).ChockPoints(0).x + (Abs(Player(0).ChockPoints(0).x - Player(0).ChockPoints(1).x) / 2)
CenterY = Player(0).ChockPoints(0).y + (Abs(Player(0).ChockPoints(0).y - Player(0).ChockPoints(2).y) / 2)

Player(0).ChockPoints(0).x = centerX + (14.84 * Cos(anguloPI(angulo + 180)))
Player(0).ChockPoints(0).y = CenterY + (14.84 * Sin(anguloPI(angulo + 180)))
''BackBuffer.DrawLine centerX, CenterY, Player(0).ChockPoints(0).X, Player(0).ChockPoints(0).Y

Player(0).ChockPoints(1).x = centerX + (44 * Cos(anguloPI(angulo - 70))) / 2.3
Player(0).ChockPoints(1).y = CenterY + (44 * Sin(anguloPI(angulo - 70))) / 2.3
''BackBuffer.DrawLine centerX, CenterY, Player(0).ChockPoints(1).X, Player(0).ChockPoints(1).Y

Player(0).ChockPoints(2).x = centerX + (14.84 * Cos(anguloPI(angulo + 0)))
Player(0).ChockPoints(2).y = CenterY + (14.84 * Sin(anguloPI(angulo + 0)))
''BackBuffer.DrawLine centerX, CenterY, Player(0).ChockPoints(2).X, Player(0).ChockPoints(2).Y

Player(0).ChockPoints(3).x = centerX + (44 * Cos(anguloPI(angulo + 110))) / 2.3
Player(0).ChockPoints(3).y = CenterY + (44 * Sin(anguloPI(angulo + 110))) / 2.3
''BackBuffer.DrawLine centerX, CenterY, Player(0).ChockPoints(3).X, Player(0).ChockPoints(3).Y

''BackBuffer.DrawCircle Player(0).ChockPoints(0).X, Player(0).ChockPoints(0).Y, 2
''BackBuffer.DrawCircle Player(0).ChockPoints(1).X, Player(0).ChockPoints(1).Y, 2
''BackBuffer.DrawCircle Player(0).ChockPoints(2).X, Player(0).ChockPoints(2).Y, 2
''BackBuffer.DrawCircle Player(0).ChockPoints(3).X, Player(0).ChockPoints(3).Y, 2

If ShowGrides = True Then
    BackBuffer.DrawText Player(0).ChockPoints(0).x, Player(0).ChockPoints(0).y, "0", False
    BackBuffer.DrawText Player(0).ChockPoints(1).x, Player(0).ChockPoints(1).y, "1", False
    BackBuffer.DrawText Player(0).ChockPoints(2).x, Player(0).ChockPoints(2).y, "2", False
    BackBuffer.DrawText Player(0).ChockPoints(3).x, Player(0).ChockPoints(3).y, "3", False
End If
Dim linha As lados


EscrevaNaTela

End Function

Public Function Atualize2D()
'atualiza e posiciona somente o carro
    Dim x As Long
    On Error Resume Next
    Viewport_Width = 640 / ZoomX
    Viewport_Height = 480 / ZoomY
    
    ASPECT_RATIO = Viewport_Width / Viewport_Height
    
    Viewplane_Width = 2
    Viewplane_Height = 2 / ASPECT_RATIO
    
    Distance = 0.5 * Viewplane_Width * Tan(PI * (FOV / 2) / 180)
    
    Camera3D_Pos.Z = 150
    
     Dim Current_Vertex As Long
    
    Dim alpha As Double
    Dim Beta As Double
    Dim Saltou As Boolean
    Number_Of_Vertices = 1
    
    ReDim Vertex_List(Number_Of_Vertices) As Point_3D
    ReDim temp(Number_Of_Vertices) As Point_3D
    ReDim Local_Vertex(Number_Of_Vertices) As Point_3D
    ReDim Camera3D(Number_Of_Vertices) As Point_3D
    ReDim Perspective(Number_Of_Vertices) As Point_3D
    ReDim Screen3D(Number_Of_Vertices) As Point_3D
    Dim Estadescendo As Boolean
    'vertices
    If BotaoSaltoPressionado = False Then
        If Player(0).declive = reto Or Player(0).declive = subindo Or Player(0).declive = descendo Then
            If ExtendRampa = False Then
                If AplicandoGravidade = False Then
                    If poligono(CarroEstaNoPoligono).tangente >= 0 Then
                        Player(0).AlturaReal = Player(0).position.Z + jump
                    Else
                        Player(0).AlturaReal = Player(0).position.Z
                    End If
                    Saltou = SaltouRampa
                    'BackBuffer.DrawText 0, 0, Saltou & "   " & Player(0).AlturaReal, False
                    
                    If Saltou = False And alturaAnterior - Player(0).AlturaReal < -diff And poligono(CarroEstaNoPoligono).tangente <> 0 Then Estadescendo = True: tantoaAbaixar = 70 Else tantoaAbaixar = 115
                         

                    If (Saltou = True Or Estadescendo = True) And Abs(Player(0).velocidade > CDbl(4)) Then
                        If Saltou Then SubaDescaCarroAposSalto = True
                        Player(0).AlturaReal = alturaAnterior + (poligono(LastRampa).tangente * 10 / SaltoReal)
                        processeAltura = True
                        AplicandoGravidade = True
                        If ExtendRampaCont = 0 And BotaoSaltoPressionado = False And Estadescendo = False Then
                            If Player(0).declive = subindo Then ExtendRampa = True
                            'If Saltou = True Then RampaSaltada = True
                        End If
                    Else
                        'Player(0).AlturaReal = Player(0).position.Z + jump
                    End If
                Else
                    Player(0).AlturaReal = alturaAnterior + (tantoaAbaixar / SaltoReal)
                End If
            Else
        ''extendrampa
                ExtendRampaCont = ExtendRampaCont + 1
                If ExtendRampaCont > CDbl(2) * Player(0).velocidade Then ExtendRampa = False
                    Player(0).AlturaReal = Player(0).AlturaReal - 140 / SaltoReal
                
            End If
        Else
    ''carro descendo
                
            
            Player(0).AlturaReal = Player(0).position.Z + jump
            AplicandoGravidade = False
        ''AplicandoGravidade = True
                    'verifica se quando tocar o chao o carro explode
            ExtendRampaCont = 0
            ExtendRampa = False
        End If
    alturaAnterior = Player(0).AlturaReal
    Else
    Player(0).AlturaReal = alturaAnterior + jump
    End If
       
    Vertex_List(0).x = Player(0).position.x: Vertex_List(0).y = Player(0).position.y
    Vertex_List(0).Z = Player(0).AlturaReal
    angle = 284.5
    'Angle = 0
    Rotate xAxis, yAxis, zAxis
    
    angle = angle Mod 360
    
    
    For Current_Vertex = 0 To Number_Of_Vertices - 1
        
        Player(0).position2D.x = Local_Vertex(Current_Vertex).x + Camera3D_Pos.x + 250
        Player(0).position2D.y = Local_Vertex(Current_Vertex).y + Camera3D_Pos.y - 25
        
    'ddraw.f
    Next Current_Vertex
    AtualizeOtherPlayers2D
    
    'agora os objetos armas frontais
    For x = 0 To UBound(ArmasFrontaisNaPista)
        If ArmasFrontaisNaPista(x).Active = True Then
            MoverObjetos ArmasFrontaisNaPista(x), 25, GetAngleFromCarImage(ArmasFrontaisNaPista(x).extra)
            Vertex_List(0).x = ArmasFrontaisNaPista(x).positionX
            Vertex_List(0).y = ArmasFrontaisNaPista(x).positionY
            Vertex_List(0).Z = ArmasFrontaisNaPista(x).positionZ
            angle = 284.5
            Rotate xAxis, yAxis, zAxis
            angle = angle Mod 360
            For Current_Vertex = 0 To Number_Of_Vertices - 1
                ArmasFrontaisNaPista(x).VideoPos.x = Local_Vertex(Current_Vertex).x + Camera3D_Pos.x + 250
                ArmasFrontaisNaPista(x).VideoPos.y = Local_Vertex(Current_Vertex).y + Camera3D_Pos.y - 25
            Next Current_Vertex
        End If
    Next x
    
    'agora os objetos vindos da net
    For x = 0 To 2000
        If ObjectsFromNet(x).Active = True Then
            If ObjectsFromNet(x).tipo = slaser Then
                MoverObjetos ObjectsFromNet(x), 25, GetAngleFromCarImage(ObjectsFromNet(x).extra)
                Vertex_List(0).x = ObjectsFromNet(x).positionX
                Vertex_List(0).y = ObjectsFromNet(x).positionY
                Vertex_List(0).Z = ObjectsFromNet(x).positionZ
                angle = 284.5
                Rotate xAxis, yAxis, zAxis
                angle = angle Mod 360
                For Current_Vertex = 0 To Number_Of_Vertices - 1
                    ObjectsFromNet(x).VideoPos.x = Local_Vertex(Current_Vertex).x + Camera3D_Pos.x + 250
                    ObjectsFromNet(x).VideoPos.y = Local_Vertex(Current_Vertex).y + Camera3D_Pos.y - 25
                Next Current_Vertex
            End If
        End If
    Next x
End Function

Public Sub swap(var1 As Double, var2 As Double)
Dim tempVar As Double
tempVar = var1
var1 = var2
var2 = tempVar
End Sub

Private Sub Rotate(Angle_X As Double, Angle_Y As Double, Angle_Z As Double)

    Dim Current_Vertex As Long
    
    For Current_Vertex = 0 To Number_Of_Vertices - 1
    
        temp(Current_Vertex).x = Vertex_List(Current_Vertex).x * (Cos(Angle_Y * RADIAN) * Cos(Angle_Z * RADIAN)) + Vertex_List(Current_Vertex).y * ((Sin(Angle_X * RADIAN) * Sin(Angle_Y * RADIAN) * Cos(Angle_Z * RADIAN)) + (Cos(Angle_X * RADIAN) * -Sin(Angle_Z * RADIAN))) + Vertex_List(Current_Vertex).Z * (Cos(Angle_X * RADIAN) * Sin(Angle_Y * RADIAN) * Cos(Angle_Z * RADIAN)) + (-Sin(Angle_X * RADIAN) * -Sin(Angle_Z * RADIAN))
        temp(Current_Vertex).y = Vertex_List(Current_Vertex).x * (Cos(Angle_Y * RADIAN) * Sin(Angle_Z * RADIAN)) + Vertex_List(Current_Vertex).y * ((Sin(Angle_X * RADIAN) * Sin(Angle_Y * RADIAN) * Sin(Angle_Z * RADIAN)) + (Cos(Angle_X * RADIAN) * Cos(Angle_Z * RADIAN))) + Vertex_List(Current_Vertex).Z * (Cos(Angle_X * RADIAN) * Sin(Angle_Y * RADIAN) * Sin(Angle_Z * RADIAN)) + (-Sin(Angle_X * RADIAN) * Cos(Angle_Z * RADIAN))
        'Teste = Temp(Current_Vertex).y
        temp(Current_Vertex).Z = Vertex_List(Current_Vertex).x * (-Sin(Angle_Y * RADIAN)) + Vertex_List(Current_Vertex).y * (Sin(Angle_X * RADIAN) * Cos(Angle_Y * RADIAN)) + Vertex_List(Current_Vertex).Z * (Cos(Angle_X * RADIAN) * Cos(Angle_Y * RADIAN))
        
        Local_Vertex(Current_Vertex) = temp(Current_Vertex)
    
    Next Current_Vertex

End Sub


Public Function PegueAngulo(ByVal Ax As Double, ByVal Ay As Double, ByVal Bx As Double, ByVal by As Double, ByVal Cx As Double, ByVal Cy As Double) As Double
Dim dot_product As Double
Dim cross_product As Double

    ' Get the dot product and cross product.
    dot_product = DotProduct(Ax, Ay, Bx, by, Cx, Cy)
    cross_product = CrossProductLength(Ax, Ay, Bx, by, Cx, Cy)

    ' Calculate the angle.
    PegueAngulo = ATan2(cross_product, dot_product)
End Function

Private Function DotProduct( _
    ByVal Ax As Single, ByVal Ay As Single, _
    ByVal Bx As Single, ByVal by As Single, _
    ByVal Cx As Single, ByVal Cy As Single _
  ) As Single
Dim BAx As Single
Dim BAy As Single
Dim BCx As Single
Dim BCy As Single

    ' Get the vectors' coordinates.
    BAx = Ax - Bx
    BAy = Ay - by
    BCx = Cx - Bx
    BCy = Cy - by

    ' Calculate the dot product.
    DotProduct = BAx * BCx + BAy * BCy
End Function


' Return the cross product AB x BC.
' The cross product is a vector perpendicular to AB
' and BC having length |AB| * |BC| * Sin(theta) and
' with direction given by the right-hand rule.
' For two vectors in the X-Y plane, the result is a
' vector with X and Y components 0 so the Z component
' gives the vector's length and direction.
Public Function CrossProductLength( _
    ByVal Ax As Single, ByVal Ay As Single, _
    ByVal Bx As Single, ByVal by As Single, _
    ByVal Cx As Single, ByVal Cy As Single _
  ) As Single
Dim BAx As Single
Dim BAy As Single
Dim BCx As Single
Dim BCy As Single

    ' Get the vectors' coordinates.
    BAx = Ax - Bx
    BAy = Ay - by
    BCx = Cx - Bx
    BCy = Cy - by

    ' Calculate the Z coordinate of the cross product.
    CrossProductLength = BAx * BCy - BAy * BCx
End Function


' Return the angle with tangent opp/hyp. The returned
' value is between PI and -PI.
Public Function ATan2(ByVal opp As Single, ByVal adj As Single) As Single
Dim angle As Single

    ' Get the basic angle.
    If Abs(adj) < 0.0001 Then
        angle = PI / 2
    Else
        angle = Abs(Atn(opp / adj))
    End If

    ' See if we are in quadrant 2 or 3.
    If adj < 0 Then
        ' angle > PI/2 or angle < -PI/2.
        angle = PI - angle
    End If

    ' See if we are in quadrant 3 or 4.
    If opp < 0 Then
        angle = -angle
    End If

    ' Return the result.
    ATan2 = angle
End Function

Public Sub CreateCheckPoint(ByVal PoligonoDono As Long, Optional ByVal ForceToindex As Long = -1, Optional ByVal index As Long)
  poligono(PoligonoDono).IsCheckPoint = True
  CheckPoints(PoligonoDono).ForceToindex = ForceToindex
  Dim x As Long
  
  CheckPoints(PoligonoDono).pos.x = (poligono(PoligonoDono).pos(0).x + (poligono(PoligonoDono).pos(1).x - poligono(PoligonoDono).pos(0).x) / 2) - 50
CheckPoints(PoligonoDono).pos.y = poligono(PoligonoDono).pos(0).y + (poligono(PoligonoDono).pos(3).y - poligono(PoligonoDono).pos(0).y) / 2
CheckPoints(PoligonoDono).pos.Z = poligono(PoligonoDono).NivelInicial
CheckPoints(PoligonoDono).index = index

End Sub

Public Function CriarLinhadeLargada(ByVal PoligonoDono As Long) As Boolean
         poligono(PoligonoDono).IsLagada = True
End Function

Public Function CarroTodoNaPista() As Boolean
Dim nPol As Long
Dim r As Long
Dim res As Long
Dim qualquer As lados
Dim s As Long

'nPol = UBound(Poligono) - 1
'For r = 0 To nPol
    'se o carro estiver todo dentro de um dos poligonos, entáo ele está na pista
    'If Poligono(r).piso <> checkpoint And Poligono(r).piso <> largada Then
        
        res = CheckColision(CarroEstaNoPoligono, qualquer)
        If res = 0 Or poligono(CarroEstaNoPoligono).IsRampaStep = True Then
            Player(0).CarroSeChocouQuinaExterna = False
            Player(0).CarroSeChocouQuinaInterna = False
 
            CarroTodoNaPista = True
            TempoForadaPista = 0
            
            If poligono(CarroEstaNoPoligono).IsRampaStep = False And processeAltura = False And RetornandoAPista = False And carrofora = False And ProcesseSuspensao = False And SubaDescaCarroAposSalto = False Then
                CarLastPositionX = Player(0).position.x
                CarLastPositionY = Player(0).position.y
                LastPolBefore = Player(0).position.Z - 40
            End If
    
            Exit Function
        Else
            'res<>0 se um dos vertices nao estiver em nenhum poligono
           
            If CheckColision2(res - 1) = False Then
            'If res <> 0 Then
                If qualquer = direita Or qualquer = baixo Then Player(0).CarroSeChocouQuinaInterna = True: LastVerticeChocked = res
                If qualquer = esquerda Or qualquer = cima Then Player(0).CarroSeChocouQuinaExterna = True: LastVerticeChocked = res
                If carrofora = False Then If carrofora = False Then Sound.WavPlay sons.cornercrash, EffectsVolume
                
                TempoForadaPista = TempoForadaPista + 1
                If TempoForadaPista > 50 Then
                    Player(0).CarroExplodiu = True
                    TempoForadaPista = 0
                    Sound.WavPlay sons.eExplosao, EffectsVolume
                    
                End If
           Else
            Player(0).CarroSeChocouQuinaExterna = False
            Player(0).CarroSeChocouQuinaInterna = False
 
            CarroTodoNaPista = True
            
           End If
           
           ' Exit Function
        End If
           ' res = CheckColision2(Poligono(r).indice, qualquer)
            'If res <> 0 Then
             '   If qualquer = direita Then Player(0).CarroSeChocouQuinaInterna = True
              '  If qualquer = esquerda Then Player(0).CarroSeChocouQuinaExterna = True
           'End If
        
        
            'If qualquer = direita Then Player(0).CarroSeChocouQuinaInterna = True
            'If qualquer = esquerda Then Player(0).CarroSeChocouQuinaExterna = True

       'End If
        ''Or qualquer = esquerda Then
            ''If res = 1 Or res = 4 Then Player(0).CarroSeChocouQuinaExterna = True
            ''If res = 2 Or res = 3 Then Player(0).CarroSeChocouQuinaInterna = True
        ''End If
    'End If
'Next r

        End Function

Public Function GetNivel(ByVal polig As Long, Optional ByVal CarVertice As Long = 0) As Double
Dim VerticeCarro As Long

Dim x As Double
Dim y As Double
Dim Ydiff As Double
Dim Xdiff As Double
Dim Xmais As Double
Dim Ymais As Double
Dim Proporcao As Double
Dim v0to1Up As Boolean
Dim v1to2Up As Boolean
Dim v2to3Up As Boolean
Dim v3to4Up As Boolean

Select Case poligono(polig).piso
Case Rampa
 VerticeCarro = CarVertice
    x = Player(0).ChockPoints(VerticeCarro).x
   '' y = 0
                        y = 0
                        Ydiff = poligono(polig).VideoPos(3).y - poligono(polig).VideoPos(2).y
                        Xdiff = (poligono(polig).VideoPos(3).x - poligono(polig).VideoPos(2).x)
                        Proporcao = Ydiff / Xdiff
                        Xmais = -poligono(polig).VideoPos(3).x
                        Ymais = poligono(polig).VideoPos(3).y
                        x = x + Xmais
                        y = (Proporcao * x)
                        y = y + Ymais - Player(0).ChockPoints(VerticeCarro).y
                       

'minimo deve ser aplicado
''If y < 0 Then y = 0
GetNivel = -y * poligono(polig).tangente + poligono(polig).NivelInicial
Case RampaH
 VerticeCarro = CarVertice
    x = Player(0).ChockPoints(VerticeCarro).x
    ''y = 0
                        y = 0
                        ''Ydiff = Poligono(polig).VideoPos(3).y - Poligono(polig).VideoPos(2).y
                        Xdiff = (poligono(polig).VideoPos(3).x - poligono(polig).VideoPos(2).x)
                        ''Proporcao = Ydiff / Xdiff
                        Xmais = -poligono(polig).VideoPos(3).x
                        ''Ymais = Poligono(polig).VideoPos(3).y
                        x = x + Xmais
                        ''y = (Proporcao * x)
                        ''y = y + Ymais - Player(0).ChockPoints(VerticeCarro).y
                       

'minimo deve ser aplicado
''If y < 0 Then y = 0
GetNivel = -x * poligono(polig).tangente + poligono(polig).NivelInicial

Case Else
GetNivel = poligono(polig).NivelInicial
End Select
End Function

Public Function CarroTodoNaRampa() As Long
'0 nao esta - 1 = rampa da frente vermelha , 2 = rampa de tras azul
Dim nPol As Long
Dim r As Long
Dim qualquer As lados
nPol = UBound(poligono) - 1
For r = 0 To nPol
    'se o carro estiver todo dentro de um dos poligonos, entáo ele está na pista
    If poligono(r).piso = Rampa Or poligono(r).piso = RampaH Then
        If CheckColision(poligono(r).indice, qualquer) = 0 Then
            'If Poligono(r).tangente < 0 Then CarroTodoNaRampa = 1: Exit Function
            'If Poligono(r).tangente > 0 Then CarroTodoNaRampa = 2: Exit Function
        CarroTodoNaRampa = 1
        End If
    End If
Next r
End Function

''Public Function CarroTodoFora() As Boolean
''Dim nPol As Long
''Dim r As Long
''Dim res As Long
''Dim qualquer As lados
''nPol = UBound(Poligono)
''For r = 0 To nPol
    'se o carro estiver todo dentro de um dos poligonos, entáo ele está na pista
    ''If Poligono(r).piso <> checkpoint And Poligono(r).piso <> largada Then
  ''      res = CheckColision(Poligono(r).indice, qualquer)
      ''  If res = 0 Then
    ''         Exit Function
       '' End If
        
   '' End If
''Next r
''CarroTodoFora = True
''End Function

Public Function CarroTodoFora() As Boolean
''checkcolision = 0 = carro nao se choca com nada
'linechock 1 = esquerda,2 = direita,3=de cima,4 = de baixo
If Player(0).AlturaReal = poligono(CarroEstaNoPoligono).NivelInicial Then Exit Function
Dim VerticeCarro As Long
Dim maxVert As Long
Dim x As Double
Dim y As Double
Dim Ydiff As Double
Dim Xdiff As Double
Dim Xmais As Double
Dim Ymais As Double
Dim Proporcao As Double
Dim v0to1Up As Boolean
Dim v1to2Up As Boolean
Dim v2to3Up As Boolean
Dim v3to4Up As Boolean
Dim nivel As Long
Dim polig As Long
''If Poligono(polig).piso = Rampa Then maxVert = 0 Else maxVert = 3
Dim vTOff(0 To 3) As Boolean
For VerticeCarro = 0 To 3
    For polig = 0 To UBound(poligono) - 1
    x = Player(0).ChockPoints(VerticeCarro).x
    y = 0
        'primeiro verifica se esta dentro pelas laterais
        If poligono(polig).VideoPos(0).x < Player(0).ChockPoints(VerticeCarro).x Then
            If poligono(polig).VideoPos(1).x > Player(0).ChockPoints(VerticeCarro).x Then
            'vertice está dentro
                ''If Poligono(Polig).VideoPos(0).X <= Poligono(Polig).VideoPos(1).X Then
                     
                    Ydiff = poligono(polig).VideoPos(0).y - poligono(polig).VideoPos(1).y
                    Xdiff = (poligono(polig).VideoPos(0).x - poligono(polig).VideoPos(1).x)
                    Proporcao = Ydiff / Xdiff
                    Xmais = -poligono(polig).VideoPos(0).x
                    Ymais = poligono(polig).VideoPos(0).y
                    x = x + Xmais
                    y = (Proporcao * x)
                    y = y + Ymais
                  
                    If y < (Player(0).ChockPoints(VerticeCarro).y) Then
                        x = Player(0).ChockPoints(VerticeCarro).x
                        y = 0
                        Ydiff = poligono(polig).VideoPos(3).y - poligono(polig).VideoPos(2).y
                        Xdiff = (poligono(polig).VideoPos(3).x - poligono(polig).VideoPos(2).x)
                        Proporcao = Ydiff / Xdiff
                        Xmais = -poligono(polig).VideoPos(3).x
                        Ymais = poligono(polig).VideoPos(3).y
                        x = x + Xmais
                        y = (Proporcao * x)
                        y = y + Ymais
                        If y > (Player(0).ChockPoints(VerticeCarro).y) Then
                            FrmDirectX.tmrexplosaowait2.Interval = 0
                            Exit Function
                        Else
              ''           vTOff(VerticeCarro) = True
                        End If
                    Else
            ''           vTOff(VerticeCarro) = True
                    End If
                   
            Else
          ''      vTOff(VerticeCarro) = True
            End If
        Else
           ''lineChock = esquerda
        ''Exit Function
        ''vTOff(VerticeCarro) = True
        End If
        
    Next
proxVert:
Next VerticeCarro

CarroTodoFora = True
TempoForadaPista = 0
FrmDirectX.tmrexplosaowait2.Interval = 5000
End Function




Public Function SaltouRampa() As Boolean
Dim altura As Double
Dim EndTang As Double
'If ExtendRampa = True Or BotaoSaltoPressionado = True Or AplicandoGravidade = True Then Exit Function
                    
If poligono(CarroEstaNoPoligono).tangente = 0 Or Player(0).declive = descendo Then espereFimdaRampa = False: LastRampa = 0: Exit Function
If espereFimdaRampa = False Then
    If poligono(CarroEstaNoPoligono).tangente <> 0 Then RampaToUse = CarroEstaNoPoligono: espereFimdaRampa = True: LastRampa = CarroEstaNoPoligono
    Exit Function
End If
If espereFimdaRampa = True Then
    'If Poligono(RampaToUse).tangente > 0 Then
        If poligono(LastRampa).NivelFinal < poligono(LastRampa).NivelInicial Then altura = poligono(LastRampa).NivelFinal Else altura = poligono(LastRampa).NivelInicial
        
        If poligono(LastRampa).IsRampaStep = True Then EndTang = 0.5 Else EndTang = 0.95
        
        If Player(0).position.Z < (altura * EndTang) Then
            'If LastRampa <> CarroEstaNoPoligono Then
        
                espereFimdaRampa = False
                SaltouRampa = True
            'Else
         
            'End If
        Else
            
            'espereFimdaRampa = False
            'SaltouRampa = False
            
        End If
    'End If
End If

End Function

Public Function AtualizeOtherPlayers2D()
'atualiza e posiciona somente o carro
    On Error Resume Next
    Viewport_Width = 640 / ZoomX
    Viewport_Height = 480 / ZoomY
    
    ASPECT_RATIO = Viewport_Width / Viewport_Height
    
    Viewplane_Width = 2
    Viewplane_Height = 2 / ASPECT_RATIO
    
    Distance = 0.5 * Viewplane_Width * Tan(PI * (FOV / 2) / 180)
    
    Camera3D_Pos.Z = 150
    
     Dim Current_Vertex As Long
    
    Dim alpha As Double
    Dim Beta As Double
    Dim Saltou As Boolean
    Number_Of_Vertices = 1
    
    ReDim Vertex_List(Number_Of_Vertices) As Point_3D
    ReDim temp(Number_Of_Vertices) As Point_3D
    ReDim Local_Vertex(Number_Of_Vertices) As Point_3D
    ReDim Camera3D(Number_Of_Vertices) As Point_3D
    ReDim Perspective(Number_Of_Vertices) As Point_3D
    ReDim Screen3D(Number_Of_Vertices) As Point_3D
    Dim Estadescendo As Boolean
    Dim sv As Long
    'vertices
For sv = 1 To 100
    If OtherPlayers(sv).Data.id <> 0 Then
        If OtherPlayers(sv).Active = True Then
            If OtherPlayers(sv).AcaboudeReceber = True Or cl_gaitestimation = True Then MoveOtherCars OtherPlayers(sv)
            OtherPlayers(sv).AcaboudeReceber = False
            Vertex_List(0).x = OtherPlayers(sv).Data.positionX: Vertex_List(0).y = OtherPlayers(sv).Data.positionY
            Vertex_List(0).Z = OtherPlayers(sv).Data.positionZ
            angle = 284.5
            Rotate xAxis, yAxis, zAxis
            angle = angle Mod 360
        
            For Current_Vertex = 0 To Number_Of_Vertices - 1
                OtherPlayers(sv).position2D.x = Local_Vertex(Current_Vertex).x + Camera3D_Pos.x + 250
                OtherPlayers(sv).position2D.y = Local_Vertex(Current_Vertex).y + Camera3D_Pos.y - 25
            Next Current_Vertex
        End If
    Else
        OtherPlayers(sv).Active = False
    End If
Next sv
End Function

Public Sub DrawOtherPlayers()
Dim pneu As Long
Dim sv As Long
Dim x As Long
Dim h As Double
Dim SomAltura As Long
On Error Resume Next
For sv = 1 To 100
    If OtherPlayers(sv).Data.id <> 0 Then
        'vai para lista de carros a ser pintados
        
        pneu = (OtherPlayers(sv).position2D.x Mod 15) Mod 3
        pneu = Abs(pneu)
        If OtherPlayers(sv).Data.CarroExplosao = 0 Then
            If OtherPlayers(sv).Data.ShowSombra Then
                DDraw.DisplaySprite sombra, CLng(OtherPlayers(sv).position2D.x) + 30, CLng(OtherPlayers(sv).position2D.y) + 60
            End If
                'DDraw.DisplaySprite cars(OtherPlayers(x).data.color, 0, OtherPlayers(sv).data.Car_Image_Index, OtherPlayers(sv).data.elevacao, pneu), CLng(OtherPlayers(sv).position2D.x), CLng(OtherPlayers(sv).position2D.y)      'Display the appropriate frame of the sprite
            DDraw.DisplaySprite cars(sv, 0, OtherPlayers(sv).Data.Car_Image_Index, OtherPlayers(sv).Data.elevacao, pneu), CLng(OtherPlayers(sv).position2D.x), CLng(OtherPlayers(sv).position2D.y)        'Display the appropriate frame of the sprite
            'SaidText(1) = OtherPlayers(sv).data.color
        Else
            'OtherPlayers(sv).contexplosao = OtherPlayers(sv).contexplosao + 1
            DDraw.DisplaySprite explosao(Int(OtherPlayers(sv).Data.CarroExplosao)), CLng(OtherPlayers(sv).position2D.x) - 250, CLng(OtherPlayers(sv).position2D.y) - 175
        End If
        
        If OtherPlayers(sv).Data.blow > 0 Then
            
            ProcesseFumacaOthers = True
            CriarFumacaDerrapagemOthers OtherPlayers(sv), sv
            
        End If

        If OtherPlayers(sv).Data.FumacaDerrapagem And pneu = 1 Then CriarFumacaDerrapagem 3, OtherPlayers(sv).position2D.x, OtherPlayers(sv).position2D.y: ProcesseFumaca = True
        '
    If OtherPlayers(sv).Data.laserHit Then
    'Sound.WavPlay sons.eLaser7, SomAltura
    Sound.WavPlay sons.eLaser7, EffectsVolume
    End If
    
    If OtherPlayers(sv).Data.oleoHit Then
    Sound.WavPlay sons.eOleo7, EffectsVolume
    End If
    
    If OtherPlayers(sv).Data.JumpHit Then
    Sound.WavPlay sons.salto, EffectsVolume
    End If
aindahafumaca:
End If


Next sv
End Sub

Public Function CarCrash() As Boolean
If CrashDerrapagem = True Then Exit Function
Dim x As Long
Dim limite As Long
Dim PlayerMessage As command
Dim a As Long
Dim b As Long
Dim c As Long
limite = 40
Dim impacto As Double


For x = 1 To 100
    If OtherPlayers(x).Data.id <> 0 Then
        
        If Player(0).position.x > OtherPlayers(x).Data.positionX - limite And Player(0).position.x < OtherPlayers(x).Data.positionX + limite Then
            If Player(0).position.y > OtherPlayers(x).Data.positionY - limite And Player(0).position.y < OtherPlayers(x).Data.positionY + limite Then
                If Player(0).AlturaReal > OtherPlayers(x).Data.positionZ - 25 And Player(0).AlturaReal < OtherPlayers(x).Data.positionZ + 25 Then
                    If Abs(Player(0).velocidade) <= 2 And Abs(OtherPlayers(x).Data.velocidade) <= 2 Then
                        'If Player(0).velocidade < OtherPlayers(x).data.velocidade Then
                        '    Player(0).velocidade = Player(0).velocidade - 0.3
                        'Else
                         '   Player(0).velocidade = Player(0).velocidade + 1
                        'End If
                        'If Player(0).velocidade < -0 Then Player(0).velocidade = -1
                        'If Player(0).velocidade > 7 Then Player(0).velocidade = 7
                        'CarCrash = True
                        'CrashDerrapagem = True
                    Else
                        
                        If Abs(OtherPlayers(x).Data.Car_Image_Index - Player(0).Car_Image_Index) >= 12 Then
                        a = OtherPlayers(x).Data.Car_Image_Index - 12
                        a = validateIndex(a)
                        b = Player(0).Car_Image_Index - a
                        
                        c = OtherPlayers(x).Data.Car_Image_Index + b
                        c = validateIndex(c)
                        Else
                        c = OtherPlayers(x).Data.Car_Image_Index
                        End If
                        Player(0).velocidade = OtherPlayers(x).Data.velocidade
                        derrapar = True
                        'OtherTime = Abs(GetAngleFromCarImage(Player(0).Car_Image_Index) - GetAngleFromCarImage(OtherPlayers(x).data.Car_Image_Index) * (Player(0).velocidade + OtherPlayers(x).data.velocidade) * 3)
                        CrashDerrapagem = True
                        AnguloDerrapagem = GetAngleFromCarImage(c)
                         CarCrash = True
                        
                     'If GetAngleFromCarImage(OtherPlayers(x).data.Car_Image_Index) >= GetAngleFromCarImage(Player(0).Car_Image_Index) - 90 And GetAngleFromCarImage(OtherPlayers(x).data.Car_Image_Index) <= GetAngleFromCarImage(Player(0).Car_Image_Index) + 90 Then
                    'transferencia simples
                        
                     '   If CrashDerrapagem = False Then CrashAngle = GetAngleFromCarImage(OtherPlayers(x).data.Car_Image_Index)
                      '  CarCrash = True
                       ' Player(0).velocidade = OtherPlayers(x).data.velocidade
                        'derrapar = True
                        'CrashDerrapagem = True
                        'AnguloDerrapagem = CrashAngle
                        'FrmDirectX.PararDerrapagem.Interval = 200 * Abs(Player(0).velocidade)
                    'Else
                        
                     '   If CrashDerrapagem = False Then CrashAngle = GetAngleFromCarImage(OtherPlayers(x).data.Car_Image_Index)
                      '  CarCrash = True
                       ' Player(0).velocidade = OtherPlayers(x).data.velocidade * 2
'                        derrapar = True
 '                       CrashDerrapagem = True
  '                      AnguloDerrapagem = CrashAngle
                        
   '                 End If
                End If
                    
            
            End If
        End If
    End If
    End If
Next x

End Function


Public Sub EscrevaNaTela()
On Error Resume Next
Dim FontInfo As New StdFont
FontInfo.name = "Arial"
FontInfo.Size = 10
BackBuffer.SetFont FontInfo
Dim extColor As Long
Dim x As Long
If ConsoleVisible = True Then
    DisplaySprite ConsoleScr, 0, ConsoleRoll, , True
    DisplaySprite ConsoleTxtscr, 40, ConsoleRoll + 180, , True
    BackBuffer.SetForeColor RGB(0, 0, 0)
    BackBuffer.DrawText 45, ConsoleRoll + 190, ConsoleTxt, False
    ConsoleRoll = ConsoleRoll + 20
    If ConsoleRoll >= 0 Then ConsoleRoll = 0
    'Exit Sub
Else
    DisplaySprite ConsoleScr, 0, ConsoleRoll, , True
    DisplaySprite ConsoleTxtscr, 40, ConsoleRoll + 180, , True
    BackBuffer.SetForeColor RGB(0, 0, 0)
    BackBuffer.DrawText 45, ConsoleRoll + 190, ConsoleTxt, False
    ConsoleRoll = ConsoleRoll - 20
    If ConsoleRoll <= -250 Then ConsoleRoll = -250
End If
If GameStatus <> 6 Then Exit Sub
BackBuffer.SetForeColor RGB(255, 255, 255)
If JString <> Empty Then

 BackBuffer.DrawText 10, 180, JString, False
End If

If Attack <> Empty Then

 BackBuffer.DrawText 20, 20, Attack, False
End If


If EscrevendoTexto = True Then
BackBuffer.DrawText 0, 0, "diga_ " & TextToSend, False
End If

'atack
FontInfo.Size = 14
BackBuffer.SetFont FontInfo
BackBuffer.SetForeColor vbBlack
BackBuffer.DrawText 200, 200, AttackString, False
If AttackString = "Attack Bonus" Then BackBuffer.SetForeColor RGB(204, 115, 17) Else BackBuffer.SetForeColor vbWhite

BackBuffer.DrawText 201, 201, AttackString, False
'said texts
FontInfo.Size = 10
BackBuffer.SetFont FontInfo
extColor = BackBuffer.GetForeColor
BackBuffer.SetForeColor vbBlue
For x = 1 To 5
BackBuffer.DrawText 59, 400 + x * 8, SaidText(x), False
Next x
BackBuffer.SetForeColor vbWhite
For x = 1 To 5
BackBuffer.DrawText 60, 400 + x * 8, SaidText(x), False
Next x
BackBuffer.SetForeColor extColor

If showingTab = True Then MostrarTab

If ShowFps = True Then BackBuffer.DrawText 500, 430, "Fps : " & GameFps, False
If CameraSeguirOutroPlayer <> 0 Then BackBuffer.DrawText 200, 400, OtherPlayers(CameraSeguirOutroPlayer).name, False
End Sub

Public Sub CriarFumacaDerrapagem(Optional ByVal max As Long = 4, Optional DX As Double = -1, Optional Dy As Double = -1)
If Player(0).blow = 0 And (Player(0).velocidade <= 4.285 Or processeAltura = True) And (DX = -1 And Dy = -1) Then Exit Sub
''SetChockPoints Player(0)

Dim x As Long
Dim pStep As Long
Dim deslocX As Long
Dim deslocY As Long
'If Player(0).blow > 0 Then pStep = 2 Else pStep = 1
pStep = 1
Randomize Timer
For x = 0 To 49 Step pStep
    'If x > 20 And Player(0).blow = 0 Then Exit Sub
    
    If FumacaPos(x).Status <= 0 Then ' Player(0).blow - 1 Then
        
        If Player(0).blow = 0 Then
           FumacaPos(x).Status = max
            If DX = -1 Then FumacaPos(x).position.x = Player(0).position2D.x + 50 Else FumacaPos(x).position.x = DX + 50
            If Dy = -1 Then FumacaPos(x).position.y = Player(0).position2D.y + 45 Else FumacaPos(x).position.y = Dy + 45
        End If
        If Player(0).blow > 0 Then
            FumacaPos(x).Status = max
            FumacaPos(x + 1).Status = max
            Randomize Timer
            deslocX = Int((20 - 1) * Rnd) + 1
            deslocY = Int((20 - 1) * Rnd) + 1
            deslocX = deslocX - 10
            deslocY = deslocY - 10
            If DX = -1 Then FumacaPos(x).position.x = Player(0).position2D.x + 50 + deslocX + (Int((10 - 0) * Rnd) + 0) Else FumacaPos(x).position.x = DX + 50 + deslocX + (Int((10 - 0) * Rnd) + 0)
            If Dy = -1 Then FumacaPos(x).position.y = Player(0).position2D.y + 30 + deslocY + (Int((10 - 0) * Rnd) + 0) Else FumacaPos(x).position.y = Dy + 30 + deslocY + (Int((10 - 0) * Rnd) + 0)
            deslocX = Int((20 - 1) * Rnd) + 1
            deslocY = Int((20 - 1) * Rnd) + 1
            deslocX = deslocX - 10
            deslocY = deslocY - 10
            
            If DX = -1 Then FumacaPos(x + 1).position.x = Player(0).position2D.x + 55 + deslocX + (Int((10 - 0) * Rnd) + 0) Else FumacaPos(x).position.x = DX + 55 + deslocX + (Int((10 - 0) * Rnd) + 0)
            If Dy = -1 Then FumacaPos(x + 1).position.y = Player(0).position2D.y + 35 + deslocY + (Int((10 - 0) * Rnd) + 0) Else FumacaPos(x).position.y = Dy + 35 + deslocY + (Int((10 - 0) * Rnd) + 0)
        End If
        
        Exit Sub
    End If
Next x

End Sub

Public Sub DesenheTodosObjetos()
Dim command As command
Dim x As Long
Dim limite As Long
Dim s As Long
Dim res As Long
Dim res2 As Long
limite = 30
 
 
 ' RotationSprite Oleo.imagem, pistas(0).imagem, CLng(Int(Timer) Mod 360), 0, 0

'Exit Sub
For x = 0 To UBound(ArmasTraseirasNaPista)
   If ArmasTraseirasNaPista(x).Active = True Then
    Select Case ArmasTraseirasNaPista(x).tipo
        Case sOil
            DisplaySprite Oleo, ArmasTraseirasNaPista(x).VideoPos.x, ArmasTraseirasNaPista(x).VideoPos.y
   
        'verifica se o player em cima
           
            
            If GameStarted = True And CheckObjectColision(ArmasTraseirasNaPista(x)) And CrashDerrapagem = False And waitDikX = False And jump = 0 Then
                CrashDerrapagem = True: AnguloDerrapagem = GetAngleFromCarImage(Player(0).Car_Image_Index)
                If ArmasTraseirasNaPista(x).extra > 5 Then
                    Player(0).Car_Image_Index = Player(0).Car_Image_Index + 1
                Else
                    Player(0).Car_Image_Index = Player(0).Car_Image_Index - 1
                End If
                
                If Player(0).Car_Image_Index > 23 Then Player(0).Car_Image_Index = 0
                If Player(0).Car_Image_Index < 0 Then Player(0).Car_Image_Index = 23
                OtherTime = 1000
            End If
        End Select
    End If
Next x

        
For x = 0 To UBound(ArmasFrontaisNaPista)
   If ArmasFrontaisNaPista(x).Active = True Then
    Select Case ArmasFrontaisNaPista(x).tipo
        Case slaser
        
        DisplaySprite laser(ArmasFrontaisNaPista(x).extra), ArmasFrontaisNaPista(x).VideoPos.x, ArmasFrontaisNaPista(x).VideoPos.y
        'obviamente o seu proprio tiro nao pode te acertar
        CheckObjectColisionOthers ArmasFrontaisNaPista(x), 70
     
        
    End Select
   End If
Next x

'desenha os objetos vindos da net
For x = 0 To 2000
   If ObjectsFromNet(x).Active = True Then
        Select Case ObjectsFromNet(x).tipo
        Case sOil
            DisplaySprite Oleo, ObjectsFromNet(x).VideoPos.x, ObjectsFromNet(x).VideoPos.y
   
        'verifica se o player em cima
        
            If GameStarted = True And CheckObjectColision(ObjectsFromNet(x)) And CrashDerrapagem = False And waitDikX = False And jump = 0 Then
                CrashDerrapagem = True: AnguloDerrapagem = GetAngleFromCarImage(Player(0).Car_Image_Index)
                Player(0).Car_Image_Index = Player(0).Car_Image_Index + ArmasTraseirasNaPista(x).extra
                If Player(0).Car_Image_Index > 23 Then Player(0).Car_Image_Index = 0
                If Player(0).Car_Image_Index < 0 Then Player(0).Car_Image_Index = 23
                OtherTime = 1000
            End If
        
        Case slaser
        
        DisplaySprite laser(ObjectsFromNet(x).extra), ObjectsFromNet(x).VideoPos.x, ObjectsFromNet(x).VideoPos.y
        
        If GameStarted = True And CheckObjectColision(ObjectsFromNet(x), 70) And contexplosao = 0 And Player(0).CarroExplodiu = False Then
            ProcesseSuspensao = True
             ObjectsFromNet(x).Active = False
            ' ObjectsFromNet(x).positionX = -51678
             'ObjectsFromNet(x).positionX = -15678
            Player(0).blow = Player(0).blow + 1
             
             If Player(0).blow = 3 Then
                If Player(0).piloto = CyberHawks Then Sound.WavPlay sons.ecyber, LarryVolume
                If Player(0).piloto = IvanZypher Then Sound.WavPlay sons.eIvan, LarryVolume
                If Player(0).piloto = JakeBlanders Then Sound.WavPlay sons.eJake, LarryVolume
                If Player(0).piloto = KatarinaLyons Then Sound.WavPlay sons.eKatarina, LarryVolume
                If Player(0).piloto = SnakeSanders Then Sound.WavPlay sons.eSnake, LarryVolume
                If Player(0).piloto = Tarquin Then Sound.WavPlay sons.eTarquin, LarryVolume
                dizerOque = sons.abouttoblow
                FrmDirectX.tmrSayIt.Interval = 500
             End If
             
             If Player(0).blow >= 4 Then
                Sound.WavPlay sons.eExplosao, EffectsVolume
                Player(0).CarroExplodiu = True
                command.command = MorriPara
                command.parametro1 = ObjectsFromNet(x).id
                command.parametro2 = ObjectsFromNet(x).handle
                res = 0
                Randomize Timer
                s = Int((5 - 1) * Rnd) + 1
                    If s = 1 Then Sound.WavPlay sons.always, LarryVolume
                    If s = 2 Then Sound.WavPlay sons.hurrysup, LarryVolume
                    If s = 3 Then Sound.WavPlay sons.ouch, LarryVolume
                    If s = 4 Then Sound.WavPlay sons.UaiPaud, LarryVolume
                    If s = 5 Then Sound.WavPlay sons.wow, LarryVolume
                Player(0).blow = 0
                Do
                    res = res + 1
                    If res > 50 Then Exit Do
                    res2 = send(FrmDirectX.Winsock1.SocketHandle, ByVal VarPtr(command), GetSizeToSend(command.command), 0)
                
            
                Loop Until res2 <> SOCKET_ERROR
                'ReDim b(0 To GetSizeToSend(command.command) - 1)
                'CopyMemory b(0), command, GetSizeToSend(command.command)
                'FrmDirectX.Winsock1.SendData b
        
            Else
                command.command = Destruirobjeto
                command.parametro1 = ObjectsFromNet(x).id
                res = 0
                Do
                    res = res + 1
                    If res > 50 Then Exit Do
                    res2 = send(FrmDirectX.Winsock1.SocketHandle, ByVal VarPtr(command), GetSizeToSend(command.command), 0)
                Loop Until res2 <> SOCKET_ERROR
                
            End If
 
    
            Exit For
            End If
        End Select
    End If
Next x
End Sub
Public Function RotateSprite(ByRef surfOrigine As Sprites, surfDestination As Sprites, lngAngle As Long, XDest As Long, Ydest As Long)
On Error Resume Next

Dim ddsdOrigine As DDSURFACEDESC2, ddsdDestination As DDSURFACEDESC2
Dim iX As Long, iY As Long
Dim iXDest As Long, iYDest As Long
Dim rEmpty As RECT, rEmpty2 As RECT
Dim sngA As Single, SinA As Single, CosA As Single
Dim dblRMax As Long
Dim lngXO As Long, lngYO As Long
Dim lngColor As Long
Dim lWidth As Long, lHeight As Long

sngA = lngAngle * PI / 180
SinA = Sin(sngA)
CosA = Cos(sngA)
'lock du source
surfOrigine.imagem.GetSurfaceDesc ddsdOrigine
lWidth = ddsdOrigine.lWidth
lHeight = ddsdOrigine.lHeight
dblRMax = Sqr(lWidth ^ 2 + lHeight ^ 2)
'lock ou on fait la rotation
surfOrigine.imagem.Lock rEmpty, ddsdOrigine, DDLOCK_WAIT, 0
surfDestination.imagem.GetSurfaceDesc ddsdDestination
surfDestination.imagem.Lock rEmpty2, ddsdDestination, DDLOCK_WAIT, 0

XDest = XDest + lWidth / 2
Ydest = Ydest + lHeight / 2
For iX = -dblRMax To dblRMax
    For iY = -dblRMax To dblRMax
        'DoEvents
        lngXO = lWidth / 2 - (iX * CosA + iY * SinA)
        lngYO = lHeight / 2 - (iX * SinA - iY * CosA)
        If lngXO >= 0 Then
            If lngYO >= 0 Then
                If lngXO < lWidth Then
                    If lngYO < lHeight Then
                        lngColor = surfOrigine.imagem.GetLockedPixel(lngXO, lngYO)
                        If lngColor <> 0 Then
                           surfDestination.imagem.SetLockedPixel XDest + iX, Ydest + iY, lngColor
                        End If
                    End If
                End If
            End If
        End If
    Next iY
Next iX
surfOrigine.imagem.Unlock rEmpty
surfDestination.imagem.Unlock rEmpty2

End Function


Public Function MoverObjetos(objeto As Objetos, ByVal velocidade As Double, ByVal angulo As Double)

velocidade = velocidade * 7 / FPS
    objeto.positionX = objeto.positionX - (velocidade * Cos(anguloPI(angulo)))
    objeto.positionY = objeto.positionY - (velocidade * Sin(anguloPI(angulo)))
    MoverPoligono objeto.PolignToUse, objeto.positionX, objeto.positionY
End Function


Public Function GetAngleToMoveObject(ByVal CarImageIndex As Long) As Double
Select Case CarImageIndex
Case 20
        GetAngleToMoveObject = 160
Case 19
       GetAngleToMoveObject = 140
       
Case 18
        GetAngleToMoveObject = 90
Case 17
        GetAngleToMoveObject = 70
Case 16
       GetAngleToMoveObject = 30
Case 15
       GetAngleToMoveObject = 10
Case 14
       GetAngleToMoveObject = -10
Case 13
       GetAngleToMoveObject = -15
Case 12
       GetAngleToMoveObject = -15
Case 11
        GetAngleToMoveObject = -16
Case 10
        GetAngleToMoveObject = -18
Case 9
        GetAngleToMoveObject = -20
Case 8
        GetAngleToMoveObject = -22
Case 7
      GetAngleToMoveObject = -44
Case 6
        GetAngleToMoveObject = 257
Case 5
        GetAngleToMoveObject = 237
Case 4
      GetAngleToMoveObject = 215
Case 3
      GetAngleToMoveObject = 185
Case 2
    GetAngleToMoveObject = 178
Case 1
    GetAngleToMoveObject = 170
Case 0
    GetAngleToMoveObject = 165
Case 23
    GetAngleToMoveObject = 160
Case 22
    GetAngleToMoveObject = 162
Case 21
    GetAngleToMoveObject = 166

End Select
'GetAngleToMoveObject = GetAngleToMoveObject - 75.5
End Function

Public Function CheckObjectColision(objeto As Objetos, Optional ByVal limiteZ As Long = 25) As Boolean
'colisao3d
'deve retornar true se ao menos 1 dos vetices tiver dentro
On Error Resume Next
If Player(0).CarroExplodiu = True Then Exit Function
Dim res As Long
Dim vertice As Long
Dim polig As Long
Dim VerticeCarro As Long
Dim maxVert As Long
Dim x As Double
Dim y As Double
Dim Ydiff As Double
Dim Xdiff As Double
Dim Xmais As Double
Dim Ymais As Double
Dim Proporcao As Double
Dim v0to1Up As Boolean
Dim v1to2Up As Boolean
Dim v2to3Up As Boolean
Dim v3to4Up As Boolean
Dim nivel As Long
Dim Ptest As Long
Dim limite As Long
Dim lk As Long
limite = 40
If Player(0).AlturaReal > objeto.positionZ - limiteZ And Player(0).AlturaReal < objeto.positionZ + limiteZ Then

polig = objeto.PolignToUse
       
''checkcolision = 0 = carro nao se choca com nada
'linechock 1 = esquerda,2 = direita,3=de cima,4 = de baixo
''If Poligono(polig).piso = Rampa Then maxVert = 0 Else maxVert = 3
Dim vTOff(0 To 3) As Boolean
For VerticeCarro = 0 To 3
    x = Player(0).ChockPoints(VerticeCarro).x
    y = 0
        'primeiro verifica se esta dentro pelas laterais
        If poligono(polig).VideoPos(0).x < Player(0).ChockPoints(VerticeCarro).x Then
            If poligono(polig).VideoPos(1).x > Player(0).ChockPoints(VerticeCarro).x Then
                    Ydiff = poligono(polig).VideoPos(0).y - poligono(polig).VideoPos(1).y
                    Xdiff = (poligono(polig).VideoPos(0).x - poligono(polig).VideoPos(1).x)
                    Proporcao = Ydiff / Xdiff
                    Xmais = -poligono(polig).VideoPos(0).x
                    Ymais = poligono(polig).VideoPos(0).y
                    x = x + Xmais
                    y = (Proporcao * x)
                    y = y + Ymais
                    If y < (Player(0).ChockPoints(VerticeCarro).y) Then
                        x = Player(0).ChockPoints(VerticeCarro).x
                        y = 0
                        Ydiff = poligono(polig).VideoPos(3).y - poligono(polig).VideoPos(2).y
                        Xdiff = (poligono(polig).VideoPos(3).x - poligono(polig).VideoPos(2).x)
                        Proporcao = Ydiff / Xdiff
                        Xmais = -poligono(polig).VideoPos(3).x
                        Ymais = poligono(polig).VideoPos(3).y
                        x = x + Xmais
                        y = (Proporcao * x)
                        y = y + Ymais
                        If y > (Player(0).ChockPoints(VerticeCarro).y) Then
                            CheckObjectColision = True
                            Exit Function
                        End If
                    End If
                End If
        End If
        
        
    
proxVert:
Next VerticeCarro

End If

End Function

Public Function CreateObjectPolign(ByVal Width As Double, ByVal Height As Double, ByVal x As Double, ByVal y As Double, ByVal ArmaAUsar As armasdoJogo) As Long

If ArmaAUsar = sOil Then
x = x + 80
y = y - 200
End If
If ArmaAUsar = slaser Then
x = x - 70
y = y + 10
End If

Dim ToUse As Long
'Dim x As Long

Dim p As PolignType
Select Case ToUse
Case 0
    ToUse = UBound(poligono)
    ReDim Preserve poligono(0 To UBound(poligono) + 1)
    
    With poligono(ToUse)
        .indice = ToUse
        .pos(3).x = x
        .pos(3).y = y
           
            
        ''altura
            .pos(0).y = .pos(3).y - Height
            .pos(0).x = .pos(3).x
          ''f(y)= -1,356097 * x
        ''largura
            .pos(2).x = .pos(3).x + Width
            .pos(2).y = .pos(3).y - Abs((.pos(2).x - .pos(3).x)) * 1.356097
        
            .pos(1).x = .pos(2).x
            .pos(1).y = .pos(2).y - Height

            .piso = sObjeto
            .largura = Width
            .altura = Height
            .ArmaUsada = ArmaAUsar
    End With
CreateObjectPolign = ToUse
End Select
End Function


Public Sub MoverPoligono(polig As Integer, ByVal x As Long, ByVal y As Long)
Dim Height As Long
Dim Width As Long

Width = poligono(polig).largura
Height = poligono(polig).altura

If poligono(polig).ArmaUsada = slaser Then
x = x + 80
y = y - 200
End If
If poligono(polig).ArmaUsada = slaser Then
x = x - 70
y = y + 10
End If

poligono(polig).pos(3).x = x
poligono(polig).pos(3).y = y

With poligono(polig)
    .pos(0).y = .pos(3).y - Height
    .pos(0).x = .pos(3).x
            
    .pos(2).x = .pos(3).x + Width
    .pos(2).y = .pos(3).y - Abs((.pos(2).x - .pos(3).x)) * 1.356097
        
    .pos(1).x = .pos(2).x
    .pos(1).y = .pos(2).y - Height

End With

End Sub
Public Sub MostrarTab()
Dim x As Long
BackBuffer.DrawText 0, 0, "Player", False
BackBuffer.DrawText 160, 0, "Pontos", False
BackBuffer.DrawText 220, 0, "Kills", False
BackBuffer.DrawText 280, 0, "Deads", False
BackBuffer.DrawText 340, 0, "Ping", False

For x = 1 To 20
    If TabScreen(x).name <> Empty Then
        BackBuffer.DrawText 0, (x * 12) + 20, TabScreen(x).name, False
        BackBuffer.DrawText 165, (x * 12) + 20, TabScreen(x).pontos, False
        BackBuffer.DrawText 225, (x * 12) + 20, TabScreen(x).kills, False
        BackBuffer.DrawText 285, (x * 12) + 20, TabScreen(x).deads, False
        BackBuffer.DrawText 345, (x * 12) + 20, Int(TabScreen(x).Ping / 2), False
    End If
Next x
End Sub

Public Sub AddTabString(command As command)
Dim x As Long
Dim tStr As String
    For x = 1 To 20
        If command.parametro1 >= TabScreen(x).pontos Then
            ShiftToDown x
            'tStr = Left(command.parametroStr, command.StringLength)
            'If IsUnicode(command.parametroStr) Then
                tStr = GetStringInArray(command.parametroStr)
            'Else
                'tStr = Left(command.parametroStr, command.StringLength)
            'End If
            'tStr = StrConv(command.parametroStr, vbUnicode)
            'tStr = CorrigirString(tStr)

            TabScreen(x).name = tStr
            TabScreen(x).pontos = command.parametro1
            TabScreen(x).kills = command.parametro2
            TabScreen(x).deads = command.parametro3
            If command.parametro4 <= 0 Then TabScreen(x).Ping = 1 Else TabScreen(x).Ping = command.parametro4
            Exit For
        End If
        
    Next x
    'organiza
    
End Sub

Public Sub ShiftToDown(ByVal From As Long)
'desloca tudo pra baixo até achar um vazio ou ate chegar no 20
Dim x As Long
For x = 20 To From + 1 Step -1
        TabScreen(x).name = TabScreen(x - 1).name
        TabScreen(x).pontos = TabScreen(x - 1).pontos
        TabScreen(x).kills = TabScreen(x - 1).kills
        TabScreen(x).deads = TabScreen(x - 1).deads
        TabScreen(x).Ping = TabScreen(x - 1).Ping
Next x
End Sub
Public Sub ShowAtack(str As String)

Atack = str
FrmDirectX.tmrShowAtack.Interval = 5000

End Sub

Public Function GetSizeToSend(opCode As comandos) As Long
On Error Resume Next
Dim p As command
'p.command=quemcorre
Select Case opCode

Case timeleft
GetSizeToSend = LenB(opCode)

Case quemcorre
GetSizeToSend = LenB(opCode) + LenB(p.parametro1)

Case RePlayersIn
GetSizeToSend = LenB(opCode) + LenB(p.parametro1)

Case spectate
GetSizeToSend = LenB(opCode) + LenB(p.parametro1) + LenB(p.parametro2)

Case corridacompleta
GetSizeToSend = LenB(opCode)

Case Quit
GetSizeToSend = LenB(opCode)

Case GetName
GetSizeToSend = LenB(opCode)

Case dados1Received
GetSizeToSend = LenB(opCode)

Case dados2Received
GetSizeToSend = LenB(opCode)

Case dados3Received
GetSizeToSend = LenB(opCode)

Case dados4Received
GetSizeToSend = LenB(opCode)

Case dados5Received
GetSizeToSend = LenB(opCode)

Case isLeading
GetSizeToSend = LenB(opCode) + LenB(p.parametro1) + LenB(p.parametro2) + LenB(p.parametro3) + LenB(p.parametro4) + LenB(p.parametro5) + LenB(p.parametroDouble) + UBound(p.parametroStr) + 1


Case Is = LastLapp
GetSizeToSend = LenB(opCode)

Case AttackBonnus
GetSizeToSend = LenB(opCode)

Case ChooseCar

Case ChoosePlayer
GetSizeToSend = LenB(opCode)

Case GetID
GetSizeToSend = LenB(opCode) + LenB(p.parametro1) + LenB(p.parametro2)

Case Go
GetSizeToSend = LenB(opCode)

Case InvalidVersion
GetSizeToSend = LenB(opCode)

Case MorriPara
GetSizeToSend = LenB(opCode) + LenB(p.parametro1) + LenB(p.parametro2)

Case OK_Player
GetSizeToSend = LenB(opCode)

Case InvalidName
GetSizeToSend = LenB(opCode)

Case Ping
GetSizeToSend = LenB(opCode) + LenB(p.parametro1)

Case PlayAndCarReady
GetSizeToSend = LenB(opCode)

Case PlayersIn
GetSizeToSend = LenB(opCode) + LenB(p.parametro1)
Case RegisterLap
GetSizeToSend = LenB(opCode)

Case Restart
GetSizeToSend = LenB(opCode)

Case SendFixedObjects
GetSizeToSend = LenB(opCode)

Case SendHelloServer
GetSizeToSend = LenB(opCode)

Case SendingText
GetSizeToSend = LenB(opCode)

Case ServerFull
GetSizeToSend = LenB(opCode)

Case sPosition
GetSizeToSend = LenB(opCode) + LenB(p.parametro1) + LenB(p.parametro2) + LenB(p.parametro3) + LenB(p.parametro4) + LenB(p.parametro5) + LenB(p.parametroDouble)

Case UnRegisterPlayer
GetSizeToSend = LenB(opCode) + LenB(p.parametro1)

Case VoceTaNoJogo
GetSizeToSend = LenB(opCode) + LenB(p.parametro1)

Case Destruirobjeto
GetSizeToSend = LenB(opCode) + LenB(p.parametro1)

Case Else
GetSizeToSend = LenB(p)
End Select
End Function

Public Function FadeSprite(tSprite As Sprites, ByVal level As Long)
 'vermelho = 63488
 'verde = 2016
 'azul = 31
 
 'trocar por 12710,2145 e 43776
 On Error Resume Next
 Dim lockrect As RECT
Dim SrcDesc         As DDSURFACEDESC2
Dim pixel(0 To 3) As Long
Dim x As Long
Dim altura As Long
Dim largura As Long
Dim ColorDest As Long
Dim rEmpty As RECT
Dim u As Integer
Dim h As Byte
Dim L As Byte
    tSprite.imagem.GetSurfaceDesc SrcDesc
    'lockrect.Right = tSprite.Height
    'lockrect.Bottom = tSprite.Width
    
    tSprite.imagem.Lock rEmpty, SrcDesc, DDLOCK_WAIT, 0
    '439, 152
    For altura = 0 To tSprite.Height - 1
        For largura = 0 To tSprite.Width - 1
            u = tSprite.imagem.GetLockedPixel(largura, altura)
            'If u <> 0 Then AFVtemp = u
            
            
            h = level And &HFF
            L = level And &HFF00
            h = h + level
            L = L + level
            If L > 255 Then L = 255
            If L < 0 Then L = 0
            If h > 255 Then h = 255
            If h < 0 Then h = 0
            
            CopyMemory u, L, 1
            CopyMemory ByVal VarPtr(u) + 1, h, 1
            
            If u < 0 Then u = 0
            'If u > &HFFFF Then u = &HFFFF
            tSprite.imagem.SetLockedPixel largura, altura, u
        Next largura
    Next altura
    tSprite.imagem.Unlock rEmpty
End Function
Sub UpdateGamma(intRed As Integer, intGreen As Integer, intBlue As Integer)
'I'm not sure who wrote this procedure; but I (Jack Hoxley) didn't.
'Full credit to whoever did...
On Error GoTo GamOut:
Dim i As Integer

If GammaSupport = True Then
'Alter the gamma ramp to the percent given by comparing to original state
'A value of zero ("0") for intRed, intGreen, or intBlue will result in the
'gamma level being set back to the original levels. Anything ABOVE zero will
'fade towards FULL colour, anything below zero will fade towards NO colour
For i = 0 To 255
    If intRed < 0 Then GammaRamp.Red(i) = ConvToSignedValue(ConvToUnSignedValue(OriginalRamp.Red(i)) * (100 - Abs(intRed)) / 100)
    If intRed = 0 Then GammaRamp.Red(i) = OriginalRamp.Red(i)
    If intRed > 0 Then GammaRamp.Red(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(OriginalRamp.Red(i))) * (100 - intRed) / 100))
    If intGreen < 0 Then GammaRamp.Green(i) = ConvToSignedValue(ConvToUnSignedValue(OriginalRamp.Green(i)) * (100 - Abs(intGreen)) / 100)
    If intGreen = 0 Then GammaRamp.Green(i) = OriginalRamp.Green(i)
    If intGreen > 0 Then GammaRamp.Green(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(OriginalRamp.Green(i))) * (100 - intGreen) / 100))
    If intBlue < 0 Then GammaRamp.Blue(i) = ConvToSignedValue(ConvToUnSignedValue(OriginalRamp.Blue(i)) * (100 - Abs(intBlue)) / 100)
    If intBlue = 0 Then GammaRamp.Blue(i) = OriginalRamp.Blue(i)
    If intBlue > 0 Then GammaRamp.Blue(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(OriginalRamp.Blue(i))) * (100 - intBlue) / 100))
Next
GammaControler.SetGammaRamp DDSGR_DEFAULT, GammaRamp

End If

Exit Sub
GamOut:
End Sub

Sub CheckForGammaSupport()
Dim Hard As DDCAPS, Soft As DDCAPS
Dim lVal As Long
dd.GetCaps Hard, Soft

If (Hard.lCaps2 And DDCAPS2_PRIMARYGAMMA) = 0 Then
GammaSupport = False
Else
GammaSupport = True
 CreateGamma
End If
End Sub

Function ConvToUnSignedValue(intValue As Integer) As Long
'This was written by the same person who did the "updateGamma" code
    If intValue >= 0 Then
        ConvToUnSignedValue = intValue
        Exit Function
    End If
    ConvToUnSignedValue = intValue + 65535
End Function

Function ConvToSignedValue(lngValue As Long) As Integer
'This was written by the same person who did the "updateGamma" code
    If lngValue <= 32767 Then
        ConvToSignedValue = CInt(lngValue)
        Exit Function
    End If
    ConvToSignedValue = CInt(lngValue - 65535)
End Function

Sub CreateGamma()
If GammaSupport = False Then Exit Sub
If GammaSupport = True Then
Set GammaControler = Primary.GetDirectDrawGammaControl
GammaControler.GetGammaRamp DDSGR_DEFAULT, OriginalRamp
End If
End Sub
Sub FadeScreenToBlack()
   Dim x As Long

   For x = 1 To 255
                CurrRed = CurrRed - 1
                CurrGreen = CurrGreen - 1
                CurrBlue = CurrBlue - 1
                UpdateGamma CurrRed, CurrGreen, CurrBlue
            Sleep 2
            Next x
End Sub

Sub FadeScreenToWhite()
   Dim x As Long

   For x = 1 To 255
                CurrRed = CurrRed + 1
                CurrGreen = CurrGreen + 1
                CurrBlue = CurrBlue + 1
                UpdateGamma CurrRed, CurrGreen, CurrBlue
            Sleep 2
            Next x
End Sub

Public Function GetServerList(ByVal URL As String) As String
    On Error Resume Next
    GetServerList = PegarTextoArquivo(App.Path & "/servidor.txt")
    
    Exit Function
    Dim hInternetSession As Long
    Dim hUrl As Long
    Dim FileNum As Integer
    Dim oK As Boolean
    Dim NumberOfBytesRead As Long
    Dim buffer As String
    Dim fileIsOpen As Boolean
    Dim DataReceived As String

    On Error GoTo ErrorHandler

    ' check obvious syntax errors
    If Len(URL) = 0 Then Err.Raise 5

    ' open an Internet session, and retrieve its handle
    hInternetSession = InternetOpen(App.EXEName, INTERNET_OPEN_TYPE_PRECONFIG, _
        vbNullString, vbNullString, 0)
    If hInternetSession = 0 Then Err.Raise vbObjectError + 1000, , _
        "An error occurred calling InternetOpen function"

    ' open the file and retrieve its handle
    hUrl = InternetOpenUrl(hInternetSession, URL, vbNullString, 0, _
        INTERNET_FLAG_EXISTING_CONNECT, 0)
    If hUrl = 0 Then Err.Raise vbObjectError + 1000, , _
        "An error occurred calling InternetOpenUrl function"

    ' ensure that there is no local file
    On Error Resume Next
    'Kill FileName

    On Error GoTo ErrorHandler
    
    ' open the local file
    FileNum = FreeFile
    'Open FileName For Binary As FileNum
    fileIsOpen = True

    ' prepare the receiving buffer
    buffer = Space(4096)
    
    Do
        ' read a chunk of the file - returns True if no error
        oK = InternetReadFile(hUrl, buffer, Len(buffer), NumberOfBytesRead)

        ' exit if error or no more data
        If NumberOfBytesRead = 0 Or Not oK Then: Exit Do
        If NumberOfBytesRead >= 1 Then
            DataReceived = DataReceived & buffer
            GetServerList = DataReceived
            
        End If
        'If NumberOfBytesRead > 1 Then CopyURLToFile = 0: Exit Do
        ' save the data to the local file
       ' Put #FileNum, , Left$(Buffer, NumberOfBytesRead)
    DoEvents
    Loop
    
    ' flow into the error handler


ErrorHandler:
    ' close the local file, if necessary
    'If fileIsOpen Then Close #FileNum
    ' close internet handles, if necessary
    If hUrl Then InternetCloseHandle hUrl
    If hInternetSession Then InternetCloseHandle hInternetSession
'processa a lista

    ' report the error to the client, if there is one
    'If err Then err.Raise err.Number, , err.Description
End Function


Public Sub ProcessGetList()
DDraw.ClearBuffer






BackBuffer.DrawText 50, 99, "Pegando a lista de servidores, por favor espere", False
DDraw.Flip

On Error Resume Next
Dim sList As String
Dim sAddress As String
Dim x As Long
sList = GetServerList("http" & ":/" & "/www.f" & "ile" & "den" & ".com/files" & "/2012/7/26/33" & "30437/rrr" & "serv" & "ers.txt")

If Len(sList) >= 0 Then
ServerAddress = Split(sList, "*")

Else
DDraw.ClearBuffer
BackBuffer.DrawText 56, 78, "falha ao pegar a lista de servidores", False
DDraw.Flip
Sleep 4000
GameStatus = 3
Exit Sub
End If

'preenche os dados
If UBound(ServerAddress) > -1 Then
    For x = 0 To UBound(ServerAddress)
        ReDim Preserve serverInfo(0 To x)
        serverInfo(x).noIp = Trim(ServerAddress(x))
        
        
    Next x
End If
RefreshServers
ShowServerList

End Sub


Public Sub ShowServerList()
On Error Resume Next
Dim x As Long
If UBound(ServerAddress) > -1 Then
BackBuffer.SetForeColor vbWhite
'limits 50, 80, 480, 300
For x = 0 To UBound(ServerAddress)
    If ServerAddress(x) <> Empty Then
            If serverInfo(x).Selected = True Then BackBuffer.SetForeColor vbWhite: BackBuffer.SetFontTransparency False: BackBuffer.SetFontBackColor RGB(236, 147, 26) Else BackBuffer.SetForeColor vbWhite: BackBuffer.SetFontTransparency True
            If ShiftServerPage + 90 + x * 12 > 82 And ShiftServerPage + 90 + x * 12 < 290 Then
            If serverInfo(x).name = Empty Then BackBuffer.DrawText 58, ShiftServerPage + 90 + x * 12, serverInfo(x).noIp, False Else BackBuffer.DrawText 58, ShiftServerPage + 90 + x * 12, serverInfo(x).name, False
            BackBuffer.DrawText 335, ShiftServerPage + 90 + x * 12, serverInfo(x).PlayersInInfo, False
            BackBuffer.DrawText 423, ShiftServerPage + 90 + x * 12, serverInfo(x).Ping, False
        End If
    End If
Next x
End If
BackBuffer.SetFontTransparency True
End Sub

Public Sub GetInfoServer(server As ServerInf, ByVal socket As Long)
On Error Resume Next
server.PlayersInInfo = "???"
Dim InitTime As Double
ServerStartedTime(socket) = Timer
If server.noIp = Empty Then Exit Sub
If FrmDirectX.GetServerInfo(socket).State = 9 Then FrmDirectX.GetServerInfo(socket).Close
'Sleep 25

FrmDirectX.GetServerInfo(socket).Connect Trim(server.noIp), 20781

'Dim k As Long
'InitTime = Timer
'Do
'BackBuffer.DrawText 0, 0, Timer - InitTime, False
'If Timer - InitTime >= 1 Then serverInfo(ServerInfoIndexToRetrieve).PlayersInInfo = "OFF": Exit Sub
'esperarespostas
'k = k + 1
'If k > 5000000 Then serverInfo(ServerInfoIndexToRetrieve).PlayersInInfo = "ERRO": Exit Sub
'If serverInfo(ServerInfoIndexToRetrieve).PlayersInInfo <> "???" Then Exit Sub
'DoEvents
'Loop

End Sub

Public Sub RefreshServers()

RefreshingServer = True
Dim x As Long
If UBound(ServerAddress) > 0 Then
    For x = 0 To UBound(ServerAddress)
        
        ServerInfoIndexToRetrieve = x
        GetInfoServer serverInfo(x), x
        DoEvents
    Next x
Else
ProcessGetList
End If
ServerInfoIndexToRetrieve = 0
RefreshingServer = False
End Sub

Public Sub UnRGB(ByRef color As OLE_COLOR, ByRef r As Byte, ByRef g As Byte, ByRef b As Byte)
    b = color And &HFF&
    g = (color And &HFF00&) \ &H100&
    r = (color And &HFF0000) \ &H10000
End Sub
Public Sub changeCarColorTo(tSprite As Sprites, color As Long, mapcolor As pixelValues)
'(0,99,173)
'cor mais escura(principal)

Dim rb As Byte
Dim bb As Byte
Dim gb As Byte

Dim r As Integer
Dim b As Integer
Dim g As Integer
Dim nCor As Long
Dim luz As Double
'color = RGB(100, 0, 0)
luz = 1
'nCor = ColorSetAlpha(RGB(R, G, B), 255)
UnRGB color, rb, gb, bb
'ChangeColors tSprite, RGB(0, 99, 173), RGB(r, g, b)
r = rb
g = gb
b = bb
'ChangeColors tSprite, RGB(0, 153, 200), RGB(R, G, B)
'ChangeColors tSprite, RGB(2, 84, 146), RGB(R, G, B)

'If r * luz > 255 Then r = 255 Else r = r * luz
'If g * luz > 255 Then g = 255 Else g = g * luz
'If b * luz > 255 Then b = 255 Else b = b * luz
'ChangeColors tSprite, 45888, RGB16(r, g, b)

'teto RGB(0,239,239)
'End

ChangeColors tSprite, RGB(100, 100, 100), Luminosidade(RGB(r, g, b), 1), mapcolor.pixels100_100_100Coordenates

'teto RGB(0,239,239)
'luz = 3
r = rb
g = gb
b = bb
'If r = 0 Then r = 1
'If g = 0 Then g = 1
'If b = 0 Then b = 1
'If r * luz > 255 Then r = 255 Else r = r * luz
'If g * luz > 255 Then g = 255 Else g = g * luz
'If b * luz > 255 Then b = 255 Else b = b * luz
ChangeColors tSprite, RGB(150, 150, 150), Luminosidade(RGB(r, g, b), 3), mapcolor.pixels150_150_150Coordenates

'ChangeColors tSprite, 52512, RGB16(r, g, b)
'luz = 7
r = rb
g = gb
b = bb
'If r = 0 Then r = 1
'If g = 0 Then g = 1
'If b = 0 Then b = 1
'If r * luz > 255 Then r = 255 Else r = r * luz
'If g * luz > 255 Then g = 255 Else g = g * luz
'If b * luz > 255 Then b = 255 Else b = b * luz
ChangeColors tSprite, RGB(200, 200, 200), Luminosidade(RGB(r, g, b), 5), mapcolor.pixels200_200_200Coordenates
'ChangeColors tSprite, 59104, RGB16(r, g, b)

'luz = 11
r = rb
g = gb
b = bb
'If r = 0 Then r = 1
'If g = 0 Then g = 1
'If b = 0 Then b = 1


'If r * luz > 255 Then r = 255 Else r = r * luz
'If g * luz > 255 Then g = 255 Else g = g * luz
'If b * luz > 255 Then b = 255 Else b = b * luz
'ChangeColors tSprite, RGB(250, 250, 250), Luminosidade(RGB(r, g, b), 7),ma


'teto RGB(0,239,239)
'End


'pneus
'r = rb
'g = gb
'b = bb
'ChangeColors tSprite, RGB(0, 255, 0), RGB(90, 90, 90)

End Sub

Public Function PegarValor(caminho As String, opcao As String) As String
Dim hregkey As Long ' receives handle to the newly created or opened registry key
Dim secattr As SECURITYATTRIBUTES  ' security settings of the key
Dim subkey As String ' name of the subkey to create
Dim neworused As Long ' receives 1 if new key was created or 2 if an existing key was opened
Dim stringbuffer As String ' the string to put into the registry
Dim retval As Long ' return value
Dim slength As Long
Dim buffer As String
' Set the name of the new key and the default security settings
'subkey = "Software\\SoftLend\\NetScuta\\Config"
secattr.nLength = Len(secattr) ' size of the structure
secattr.lpSecurityDescriptor = 0 ' default security level
secattr.bInheritHandle = True ' the default value for this setting

' Create or open the registry key
retval = RegCreateKeyEx(HKEY_LOCAL_MACHINE, caminho, 0, "", 0, KEY_READ, secattr, hregkey, neworused)
If retval <> 0 Then ' error during open

End ' terminate the program
End If

stringbuffer = Space(255) ' make room in the buffer to receive the information
slength = 255 ' this must be set if passing a string to the function
retval = RegQueryValueEx(hregkey, opcao, 0, REG_SZ, ByVal stringbuffer, slength) ' read data
stringbuffer = Left(stringbuffer, slength) ' extract the returned data from the buffer

'retira vbnullchar
Dim x As Long
x = InStr(1, stringbuffer, vbNullChar)
If x = 0 Then buffer = Empty
If x > 0 Then buffer = Mid(stringbuffer, 1, (x - 1))
PegarValor = buffer
End Function

Public Sub SetarValor(caminho As String, opcao As String, ByVal valor As String)
Dim hregkey As Long ' receives handle to the newly created or opened registry key
Dim secattr As SECURITYATTRIBUTES  ' security settings of the key
Dim subkey As String ' name of the subkey to create
Dim neworused As Long ' receives 1 if new key was created or 2 if an existing key was opened
Dim stringbuffer As String ' the string to put into the registry
Dim retval As Long ' return value

' Set the name of the new key and the default security settings
'subkey = "Software\\SoftLend\\NetScuta\\Config"
secattr.nLength = Len(secattr) ' size of the structure
secattr.lpSecurityDescriptor = 0 ' default security level
secattr.bInheritHandle = True ' the default value for this setting

' Create or open the registry key
retval = RegCreateKeyEx(HKEY_LOCAL_MACHINE, caminho, 0, "", 0, KEY_WRITE, secattr, hregkey, neworused)
If retval <> 0 Then ' error during open

End ' terminate the program
End If

' Write the string to the registry. Note that because Visual Basic is being used, the string passed to the
' function must explicitly be passed ByVal.
valor = valor & vbNullChar ' note how a null character must be appended to the string

retval = RegSetValueEx(hregkey, opcao, 0, REG_SZ, ByVal valor, Len(valor) + 1) ' write the string

' Close the registry key
retval = RegCloseKey(hregkey)


End Sub



Public Sub CriarFumacaDerrapagemOthers(otherplayer As Other_Player_Stats, ByVal index As Long)
Dim x As Long
Dim pStep As Long
Dim deslocX As Long
Dim deslocY As Long

For x = 50 * (index + 1) To (50 * (index + 2)) - 1 'Step 2
'For x = 0 To 49  'Step 2
    
    If FumacaPos(x).Status <= 0 Then
        
        
        If otherplayer.Data.blow > 0 Then
            FumacaPos(x).Status = otherplayer.Data.blow
            FumacaPos(x + 1).Status = otherplayer.Data.blow
            Randomize Timer
            deslocX = Int((20 - 1) * Rnd) + 1
            deslocY = Int((20 - 1) * Rnd) + 1
            deslocX = deslocX - 10
            deslocY = deslocY - 10
            FumacaPos(x).position.x = otherplayer.position2D.x + 50 + deslocX
            FumacaPos(x).position.y = otherplayer.position2D.y + 30 + deslocY
            deslocX = Int((20 - 1) * Rnd) + 1
            deslocY = Int((20 - 1) * Rnd) + 1
            deslocX = deslocX - 10
            deslocY = deslocY - 10
            
            FumacaPos(x + 1).position.x = otherplayer.position2D.x + 55 + deslocX
            FumacaPos(x + 1).position.y = otherplayer.position2D.y + 35 + deslocY
        End If
        
        Exit Sub
    End If
Next x

End Sub

Public Function MoveOtherCars(objeto As Other_Player_Stats)
Dim velocidade As Double
Dim angulo As Double
velocidade = objeto.Data.velocidade
angulo = GetAngleFromCarImage(objeto.Data.Car_Image_Index)
'velocidade = velocidade * 7 / FPS
    objeto.Data.positionX = objeto.Data.positionX - (velocidade * Cos(anguloPI(angulo)))
    objeto.Data.positionY = objeto.Data.positionY - (velocidade * Sin(anguloPI(angulo)))
    'MoverPoligono objeto.PolignToUse, objeto.positionX, objeto.positionY
End Function

Public Sub CheckObjectColisionOthers(objeto As Objetos, Optional ByVal limiteZ As Long = 25)
'colisao3d
'deve retornar true se ao menos 1 dos vetices tiver dentro
On Error Resume Next

Dim res As Long
Dim vertice As Long
Dim polig As Long
Dim VerticeCarro As Long
Dim maxVert As Long
Dim x As Double
Dim y As Double
Dim Ydiff As Double
Dim Xdiff As Double
Dim Xmais As Double
Dim Ymais As Double
Dim Proporcao As Double
Dim v0to1Up As Boolean
Dim v1to2Up As Boolean
Dim v2to3Up As Boolean
Dim v3to4Up As Boolean
Dim nivel As Long
Dim Ptest As Long
Dim limite As Long
Dim lk As Long
limite = 70

For x = 1 To UBound(OtherPlayers)
'If OtherPlayers(lk).data.positionZ > objeto.positionZ - limiteZ And Player(0).AlturaReal < objeto.positionZ + limiteZ Then
    If OtherPlayers(x).Data.id <> 0 Then
    
    If objeto.VideoPos.x > OtherPlayers(x).position2D.x - limite And objeto.VideoPos.x < OtherPlayers(x).position2D.x + limite Then
            If objeto.VideoPos.y > OtherPlayers(x).position2D.y - limite And objeto.VideoPos.y < OtherPlayers(x).position2D.y + limite Then
                'If objeto.positionZ > OtherPlayers(x).data.positionZ - 25 And objeto.positionZ < OtherPlayers(x).data.positionZ + 25 Then
                    Sound.WavPlay sons.always, LarryVolume
                    objeto.Active = False
                    Exit Sub
                'End If
            End If
        End If
    End If

Next x


End Sub

Public Function DesenhePista2(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, Optional ByVal CriarPoligono As Boolean = False)
If CriarPoligono = True Then
'For k = 0 To UBound(poligono)
Camera.UltimoX_Posicionado = 0
Camera.UltimoY_Posicionado = 0
Camera.UltimoX_Poligono = 0
Camera.UltimoY_Poligono = 0

lastWidth = 0
lastHeight = 0


ReDim poligono(0)
poligono(0).Type = 0
poligono(0).ArmaUsada = 0
poligono(0).indice = -1
poligono(0).piso = 0
poligono(0).Type = 0
LastPolignCreated = 0
'Next k
End If
CountPistasVerticais = 0
CountPistasHorizontais = 0
DesenhePrimeiroPedacodaPista chem_vi, x, y, 18, , , CriarPoligono
DesenhePistaCurvaALTAESQold chem_vi, (Camera.UltimoX_Posicionado - 60) / ZoomX, (Camera.UltimoY_Posicionado - 64) / ZoomY, 1, , , CriarPoligono
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + 515) / ZoomX, (Camera.UltimoY_Posicionado + 108) / ZoomY, 9, horizontal, , CriarPoligono, , 150, -85, 36, -1585
DesenhePistaCurvaBAIXADIR chem_vi, (Camera.UltimoX_Posicionado - (15 / ZoomX)), (Camera.UltimoY_Posicionado - (2 / ZoomY)), 1, , , CriarPoligono, , , , 300
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + 354) / ZoomX, (Camera.UltimoY_Posicionado + 20) / ZoomY, 13, vertical, , CriarPoligono
DesenhePistaCurvaALTAESQold chem_vi, (Camera.UltimoX_Posicionado - 60) / ZoomX, (Camera.UltimoY_Posicionado - 64) / ZoomY, 1, , , CriarPoligono, , , , 180
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + 517) / ZoomX, (Camera.UltimoY_Posicionado + 106) / ZoomY, 18, horizontal, , CriarPoligono, , 150, 460, , , -490
DesenhePistaCurvaALTADIR chem_vi, (Camera.UltimoX_Posicionado - 40), (Camera.UltimoY_Posicionado - 10), 1, , , CriarPoligono

DesenhePistaCurvaBAIXAESQold chem_vi, (-219) / ZoomX, (6) / ZoomY, 1, , , CriarPoligono
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + (162 / ZoomX)), (Camera.UltimoY_Posicionado + (163 / ZoomY)), 7, horizontal, , CriarPoligono, -5, 80, -35, 90, 5
DesenhePista2ladeirasHorizontais chem_vi, (Camera.UltimoX_Posicionado - 8) / ZoomX, (Camera.UltimoY_Posicionado - 2) / ZoomY, 1, horizontal, , CriarPoligono
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + (137 / ZoomX)), (Camera.UltimoY_Posicionado + (116 / ZoomY)), 6, horizontal, , CriarPoligono, , 20
DesenheRampaHorizontalold chem_vi, (Camera.UltimoX_Posicionado - 0), (Camera.UltimoY_Posicionado - (9 / ZoomY)), 1, horizontal, , CriarPoligono
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + (163 / ZoomX)), (Camera.UltimoY_Posicionado + (36 / ZoomY)), 6, horizontal, , CriarPoligono, , 40
DesenhePistarampinhahorizontal2 chem_vi, (Camera.UltimoX_Posicionado - (10 / ZoomX)), (Camera.UltimoY_Posicionado - (3 / ZoomY)), 1, horizontal, , CriarPoligono
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + (77 / ZoomX)), (Camera.UltimoY_Posicionado + (36 / ZoomY)), 7, horizontal, , CriarPoligono, , 100
DesenhePistaCurvaBAIXADIR chem_vi, (Camera.UltimoX_Posicionado - (15 / ZoomX)), (Camera.UltimoY_Posicionado - (2 / ZoomY)), 1, , , CriarPoligono, , , , 300
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + 354) / ZoomX, (Camera.UltimoY_Posicionado + 20) / ZoomY, 3, vertical, , CriarPoligono, -117
DesenhePistarampa4 chem_vi, (Camera.UltimoX_Posicionado + 33) / ZoomX, (Camera.UltimoY_Posicionado - 0) / ZoomY, 1, vertical, , CriarPoligono, , , 60
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + (130 / ZoomX)), (Camera.UltimoY_Posicionado + (13 / ZoomY)), 4, , , CriarPoligono, 40
DesenheRampadaPista chem_vi, (Camera.UltimoX_Posicionado - (5 / ZoomX)), (Camera.UltimoY_Posicionado - (88 / ZoomY)), 1, , , CriarPoligono
DesenheLadeira chem_vi, (Camera.UltimoX_Posicionado + (192 / ZoomX)), (Camera.UltimoY_Posicionado - (54 / ZoomY)), 1, , , CriarPoligono
DesenheRampada2 chem_vi, (Camera.UltimoX_Posicionado + (47 / ZoomX)), (Camera.UltimoY_Posicionado + (25 / ZoomY)), 1, , , CriarPoligono, 50
'DesenhePistarampa4 chem_vi, (Camera.UltimoX_Posicionado + 25) / ZoomX, (Camera.UltimoY_Posicionado + 2) / ZoomY, 1, , , CriarPoligono, -20, , , , , , -1.037
DesenhePistarampa4 chem_vi, (Camera.UltimoX_Posicionado + 25) / ZoomX, (Camera.UltimoY_Posicionado + 2) / ZoomY, 1, , , CriarPoligono, -20, 30, , , , -0.487
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + (130 / ZoomX)), (Camera.UltimoY_Posicionado + (13 / ZoomY)), 19, , , CriarPoligono, , , 30

'criar os checkpoints e a largada (sempre por ultimo)
'  21
'15  3
'  9
'
'BackBuffer.DrawText 0, 0, CarroEstaNoPoligono, False
If CriarPoligono = True Then
    CriarLinhadeLargada 1
    CreateCheckPoint 10, 3, 1
    CreateCheckPoint 20, 21, 2
    CreateCheckPoint 31, 3, 3
    CreateCheckPoint 78, 9, 4
    CreateCheckPoint 59, 18, 5
    AllChocksCreated = True
End If

End Function

Public Function DesenhePista2ladeirasHorizontais(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = horizontal, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False) As Boolean
If tamanho < 1 Then Exit Function
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(27), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
    
     If CriarPoligono Then
        a = 54.5 'altura do quadrado
        a = a * tamanho
        CreatePoligono a, 329, Retangular, RampaH, -1.06, 0
     End If
     If CriarPoligono Then
        a = 40.5 'altura do quadrado
        a = a * tamanho
        CreatePoligono a, 329, Retangular, RampaH, -0.66, 0
     End If
'    DesenheRampadaPista chem_vi, x + ((950/zoomx) / Screen.TwipsPerPixelX * (p + 1))  , y - ((250/zoomy) / Screen.TwipsPerPixelY * (p + 1))  , 1
Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function


Public Function DesenhePistarampinhahorizontal2(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = horizontal, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False) As Boolean
If tamanho < 1 Then Exit Function
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(28), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
    
     If CriarPoligono Then
        a = 54.5 'altura do quadrado
        a = a * tamanho
        CreatePoligono a, 329, Retangular, RampaH, 0.66, 0
     End If
     
Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function



Public Function DesenhePistarampa4(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = horizontal, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False, Optional ExtraHeight As Long = 0, Optional ExtraWidth As Long = 0, Optional ByVal nShiftX As Long = 0, Optional ByVal xInicial As Long = 0, Optional ByVal yInicial As Long = 0, Optional ByVal nShiftY As Long = 0, Optional ByVal tangente As Double = -0.637) As Boolean
If tamanho < 1 Then Exit Function
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(29), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
    
     If CriarPoligono Then
        a = 132.5 'altura do quadrado
        a = a * tamanho
        
        CreatePoligono 200 + ExtraWidth, a + ExtraHeight, Retangular, Rampa, tangente, nShiftY, xInicial, yInicial, , , , , nShiftY 'descendo
     End If
'
     
Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function



Public Function DesenheRampada2(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = vertical, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False, Optional ExtraHeight As Long = 0, Optional ExtraWidth As Long = 0, Optional ByVal nShiftX As Long = 0, Optional ByVal xInicial As Long = 0, Optional ByVal yInicial As Long = 0, Optional ByVal nShiftY As Long = 0) As Boolean
If tamanho < 1 Then DesenheRampada2 = False
Dim p As Long
Dim a As Double
Dim InitX As Long
Dim InitY As Long
Dim FimX As Long
Dim FimY As Long
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(30), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
    Next p
     If CriarPoligono Then
        a = 71.5 'altura do quadrado
        a = a * (tamanho)
        CreatePoligono 210 + ExtraWidth, a + ExtraHeight, Retangular, Rampa, 0.66, nShiftX, xInicial, yInicial, , , , , nShiftY
    End If
    Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
    Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function

Public Sub DesenhePista3(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, Optional ByVal CriarPoligono As Boolean = False)
If CriarPoligono = True Then
Camera.UltimoX_Posicionado = 0
Camera.UltimoY_Posicionado = 0
Camera.UltimoX_Poligono = 0
Camera.UltimoY_Poligono = 0

lastWidth = 0
lastHeight = 0

'For k = 0 To UBound(poligono)
ReDim poligono(0)
poligono(0).Type = 0
poligono(0).ArmaUsada = 0
poligono(0).indice = -1
poligono(0).piso = 0
poligono(0).Type = 0
LastPolignCreated = 0
'Next k
End If
CountPistasVerticais = 0
CountPistasHorizontais = 0

DesenhePrimeiroPedacodaPista chem_vi, x, y, 8, , , CriarPoligono
DesenhePistarampa4 chem_vi, (Camera.UltimoX_Posicionado + 25) / ZoomX, (Camera.UltimoY_Posicionado + 2) / ZoomY, 1, , , CriarPoligono, , , , , , -0.487
DesenheCruzEsquerda chem_vi, (Camera.UltimoX_Posicionado + 158) / ZoomX, (Camera.UltimoY_Posicionado - 97) / ZoomY, 1, , , CriarPoligono, , -10
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + 355) / ZoomX, (Camera.UltimoY_Posicionado + 13) / ZoomY, 11, , , CriarPoligono
DesenhePistaCurvaALTADIR chem_vi, (Camera.UltimoX_Posicionado + 7), (Camera.UltimoY_Posicionado - 137), 1, , , CriarPoligono, -200, , , -460
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado - 758) / ZoomX, (Camera.UltimoY_Posicionado - 145) / ZoomY, 11, horizontal, , CriarPoligono, , 60, -700, , , -460
DesenhePistaCurvaALTAESQ chem_vi, (Camera.UltimoX_Posicionado - 1080) / ZoomX, (Camera.UltimoY_Posicionado - 210) / ZoomY, 1, , , CriarPoligono, -180, 0, 0, 426
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado - 762) / ZoomX, (Camera.UltimoY_Posicionado + 282) / ZoomY, 11, , , CriarPoligono, 170, , -115, , , 1070
DesenhePistaCurvaBAIXAESQ chem_vi, (Camera.UltimoX_Posicionado - 789) / ZoomX, (Camera.UltimoY_Posicionado + 156) / ZoomY, 1, , , CriarPoligono, , , , , , 870
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + 97) / ZoomX, (Camera.UltimoY_Posicionado + 147) / ZoomY, 9, horizontal, , CriarPoligono, , 27, -50, , , -275
DesenheRampaStepHorizontal chem_vi, (Camera.UltimoX_Posicionado - 5) / ZoomX, (Camera.UltimoY_Posicionado - 3) / ZoomY, 1, horizontal, , CriarPoligono
DesenheRampaHorizontal chem_vi, (Camera.UltimoX_Posicionado + 420) / ZoomX, (Camera.UltimoY_Posicionado + 160) / ZoomY, 1, horizontal, , CriarPoligono, , 8, 250, , , 50, 84.721
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado + 162) / ZoomX, (Camera.UltimoY_Posicionado + 37) / ZoomY, 7, horizontal, , CriarPoligono, , 150
DesenhePistaCurvaALTADIR chem_vi, (Camera.UltimoX_Posicionado + 3), (Camera.UltimoY_Posicionado - 2), 1, , , CriarPoligono
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado - 704) / ZoomX, (Camera.UltimoY_Posicionado + 321) / ZoomY, 10, vertical, , CriarPoligono, 280, , , , , 700
DesenhePistaCurvaBAIXADIR chem_vi, (Camera.UltimoX_Posicionado - (1081 / ZoomX)), (Camera.UltimoY_Posicionado + 173) / ZoomY, 1, , , CriarPoligono, -210, , , 1145
DesenhepedacosdaPista chem_vi, (Camera.UltimoX_Posicionado - 815) / ZoomX, (Camera.UltimoY_Posicionado - 163) / ZoomY, 12, horizontal, , CriarPoligono, , 210, -860, , , 255
DesenhePistaCurvaBAIXAESQ chem_vi, (Camera.UltimoX_Posicionado - 937) / ZoomX, (Camera.UltimoY_Posicionado - 315) / ZoomY, 1, , , CriarPoligono, , , -180

'  21
'15  3
'  9
If CriarPoligono = True Then
    CriarLinhadeLargada 0
    CreateCheckPoint 4, 21, 1
    CreateCheckPoint 15, 15, 2
    CreateCheckPoint 24, 9, 3
    CreateCheckPoint 35, 3, 4
    CreateCheckPoint 38, 3, 5
    CreateCheckPoint 49, 9, 6
    CreateCheckPoint 60, 15, 7
    AllChocksCreated = True
End If

End Sub

Public Function DesenheCruzEsquerda(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = vertical, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False, Optional ExtraHeight As Long = 0, Optional ExtraWidth As Long = 0, Optional ByVal nShiftX As Long = 0, Optional ByVal xInicial As Long = 0, Optional ByVal yInicial As Long = 0, Optional ByVal nShiftY As Long = 0) As Boolean
If tamanho < 1 Then tamanho = 1
Dim p As Long
Dim a As Double
Dim InitX As Long
Dim InitY As Long
Dim FimX As Long
Dim FimY As Long
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(31), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
    Next p
     If CriarPoligono Then
        a = 71.5 'altura do quadrado
        a = a * (7)
        CreatePoligono 210 + ExtraWidth, a + ExtraHeight, Retangular, vertical, 0, nShiftX, xInicial, yInicial, , , , , nShiftY
    a = 330
        CreatePoligono 70, a, Retangular, horizontal, , -10, , , , , , , 60, , True
    End If
   
    Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
    Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function

Public Function DesenhePistaCurvaALTAESQ(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = horizontal, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False, Optional ByVal nShiftX As Long = 0, Optional ByVal xInicial As Long = 0, Optional ByVal yInicial As Long = 0, Optional ByVal nShiftY As Long = 0) As Boolean
If tamanho < 1 Then tamanho = 1
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(8), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
      If CriarPoligono Then
        
        a = 54.5 'altura do quadrado
        a = a * tamanho
        CreatePoligono 194, a, Retangular, vertical, , nShiftX, xInicial, yInicial, , , , , nShiftY
        CreatePoligono 188, a, Retangular, vertical, , 6
        CreatePoligono 182, a * 2, Retangular, vertical, , 6
        CreatePoligono 176, a, Retangular, vertical, , 6
        CreatePoligono 170, a, Retangular, vertical, , 6
        CreatePoligono 160, a - 10, Retangular, vertical, , 10
        CreatePoligono 130, a - 20, Retangular, vertical, , 30
        CreatePoligono 90, a - 28, Retangular, vertical, , 40
        a = 330
        'CreatePoligono 230, a, Retangular, vertical, , -120, xInicial, yInicial, , , , , 630
    End If
   
'    DesenheRampadaPista chem_vi, x + ((950/zoomx) / Screen.TwipsPerPixelX * (p + 1))  , y - ((250/zoomy) / Screen.TwipsPerPixelY * (p + 1))  , 1
Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function

Public Function DesenhePistaCurvaBAIXAESQ(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = horizontal, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False, Optional ExtraHeight As Long = 0, Optional ExtraWidth As Long = 0, Optional ByVal nShiftX As Long = 0, Optional ByVal xInicial As Long = 0, Optional ByVal yInicial As Long = 0, Optional ByVal nShiftY As Long = 0) As Boolean
If tamanho < 1 Then DesenhePistaCurvaBAIXAESQ = False: Exit Function
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(20), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
    If CriarPoligono Then
        a = 54.5 'altura do quadrado
        a = a * tamanho
        CreatePoligono 200, a, Retangular, vertical, , nShiftX, xInicial, yInicial, , , , , nShiftY
        CreatePoligono 194, a, Retangular, vertical, , 6, , , True
        CreatePoligono 188, a * 2, Retangular, vertical, , 6, , , True
        CreatePoligono 226, a, Retangular, vertical, , 10, , , True
        CreatePoligono 200, a, Retangular, vertical, , 17, , , True
        CreatePoligono 190, a - 10, Retangular, vertical, , 17, , , True
        CreatePoligono 160, a - 20, Retangular, vertical, , 24, , , True
        CreatePoligono 140, a - 28, Retangular, vertical, , 24, , , True
        CreatePoligono 140, a - 28, Retangular, vertical, , , , , True
        CreatePoligono 140, a - 28, Retangular, vertical, , , , , True
        ''a = 330
        ''CreatePoligono 260, a, Retangular, vertical, , 190, 16, -2032
    End If
'    DesenheRampadaPista chem_vi, x + ((950/zoomx) / Screen.TwipsPerPixelX * (p + 1))  , y - ((250/zoomy) / Screen.TwipsPerPixelY * (p + 1))  , 1
Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function

Public Function DesenheRampaStepHorizontal(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = horizontal, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False, Optional ExtraHeight As Long = 0, Optional ExtraWidth As Long = 0, Optional ByVal nShiftX As Long = 0, Optional ByVal xInicial As Long = 0, Optional ByVal yInicial As Long = 0, Optional ByVal nShiftY As Long = 0) As Boolean
tamanho = 1
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(32), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
       If CriarPoligono Then
            a = 76.5 'largura do quadrado
            a = a * (tamanho)
            CreatePoligono a, 329, Retangular, RampaH, 2, , , , , , , , , True
        End If
'    DesenheRampadaPista chem_vi, x + ((950/zoomx) / Screen.TwipsPerPixelX * (p + 1))  , y - ((250/zoomy) / Screen.TwipsPerPixelY * (p + 1))  , 1
Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function


Public Function DesenheRampaHorizontal(ByVal mundo As Mundos, ByVal x As Long, ByVal y As Long, ByVal tamanho As Long, Optional ByVal tipo As PolignType = vertical, Optional ByVal tZoom As Long, Optional ByVal CriarPoligono As Boolean = False, Optional ExtraHeight As Long = 0, Optional ExtraWidth As Long = 0, Optional deslocamento As Long = 0, Optional ByVal xInit As Long = 0, Optional ByVal yInit As Long = 0, Optional ByVal nShiftY As Long = 0, Optional ByVal nivelacao As Double = -1) As Boolean
If tamanho < 1 Then tamanho = 1
Dim p As Long
Dim a As Double
Select Case mundo
Case chem_vi
    For p = 0 To (tamanho - 1)
        'sprite pistas.Picture1, Pistas_masks.Picture1, x + (950 * p), y - (250 * p) '(x - 950) ,(y+250)
            DDraw.DisplaySprite pistas(12), x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p), y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p), 100
        
    Next p
    
        If CriarPoligono Then
            a = 120.5 'largura do quadrado
            a = a * (tamanho)
            CreatePoligono a + ExtraWidth, 329 + ExtraHeight, Retangular, RampaH, 0.66, deslocamento, xInit, yInit, , nivelacao, , , nShiftY
        End If
    Camera.UltimoX_Posicionado = x + ((950 / ZoomX) / Screen.TwipsPerPixelX * p)
    Camera.UltimoY_Posicionado = y - ((250 / ZoomY) / Screen.TwipsPerPixelY * p)
End Select
End Function

Public Function SendString(socket As Winsock, ByVal strData As String) As Boolean
Dim c1 As String
Dim p1 As String
Dim m As String
Dim command As command
Dim res As Long
m = strData
On Error Resume Next
12
   Dim WSAResult As Long, i As Long, L As Long
  '
    L = Len(strData)
    ReDim buff(L + 1) As Byte
    
    For i = 1 To L
       buff(i - 1) = Asc(Mid(strData, i, 1))
    Next
    buff(L) = 0

    WSAResult = send(socket.SocketHandle, buff(0), L, 0)
    If WSAResult = SOCKET_ERROR Then
        
        SendString = False
    Else
        SendString = True
    End If
End Function

Public Function ColorSetAlpha(ByVal lColor As Long, ByVal alpha As Byte) As Long
   Dim bytestruct As COLORBYTES
   Dim result As COLORLONG
   
   result.longval = lColor
   LSet bytestruct = result

   bytestruct.AlphaByte = alpha

   LSet result = bytestruct
   ColorSetAlpha = result.longval
End Function

Public Function ColorARGB(ByVal alpha As Byte, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte) As Long
   Dim bytestruct As COLORBYTES
   Dim result As COLORLONG
   
   With bytestruct
      .AlphaByte = alpha
      .RedByte = Red
      .GreenByte = Green
      .BlueByte = Blue
   End With
   
   LSet result = bytestruct
   ColorARGB = result.longval
End Function
Public Function GetRGB_GDIP2VB(ByVal lColor As Long) As Long
   Dim argb As COLORBYTES
   CopyMemory argb, lColor, 4
   GetRGB_GDIP2VB = RGB(argb.RedByte, argb.GreenByte, argb.BlueByte)
End Function

Public Function validateIndex(ByVal a As Long)
If a >= 0 And a <= 23 Then validateIndex = a: Exit Function
Do
 If a < 0 Then a = 24 + a
Loop Until a >= 0
 
Do
 If a > 23 Then a = a - 24
Loop Until a <= 23
 validateIndex = a
End Function

Public Function GetColorAtMousePos() As Long
 'vermelho = 63488
 'verde = 2016
 'azul = 31
 On Error Resume Next
 Dim r As Byte
 Dim g As Byte
 Dim b As Byte
 'trocar por 12710,2145 e 43776
 Dim lockrect As RECT
Dim SrcDesc         As DDSURFACEDESC2
Dim pixel(0 To 3) As Long
Dim x As Long
Dim altura As Long
Dim largura As Long
    
    BackBuffer.GetSurfaceDesc SrcDesc
    lockrect.Right = SrcDesc.lWidth
    lockrect.Bottom = SrcDesc.lHeight
    BackBuffer.Lock lockrect, SrcDesc, DDLOCK_WAIT Or DDLOCK_NOSYSLOCK, 0
    Dim u As Long
            GetColorAtMousePos = BackBuffer.GetLockedPixel(MouseX, MouseY)
            'If u <> 0 Then
            'UnRGB u, R, G, b
            'Print #122, u & "  RGB(" & R & " , " & G & " , " & b & " )"
            'End If
            
    BackBuffer.Unlock lockrect

End Function

Public Function RGB16(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
    On Error Resume Next
    'swap CDbl(r), CDbl(b)
    Dim res As Long
    Dim Cor5bits As Long
    
    'res = r * 31 / 255
    res = b / 8
    
    If res > 31 Then res = 31
    Cor5bits = res
    res = g / 8
    If res > 31 Then res = 31
    res = res * 63
    
    Cor5bits = res Or Cor5bits
    res = r / 8
    If res > 31 Then res = 31
    res = res * 2047
    
    Cor5bits = res Or Cor5bits
    
    RGB16 = Cor5bits
End Function

Public Function mSplit(ByVal mstring As String) As Boolean
''primeiro comando
If Len(mstring) = 0 Then Exit Function
Dim x As Long
Dim y As Long
Dim p1 As Long
Dim palavra As String
For x = 1 To Len(mstring)
If Mid(mstring, x, 1) = Empty Or Mid(mstring, x, 1) = " " Then Exit For
palavra = palavra & Mid(mstring, x, 1)
Next x
comandosStr1 = palavra
mSplit = True
If x = Len(mstring) Then Exit Function

For y = x To Len(mstring)
If Mid(mstring, y, 1) <> " " Then Exit For
Next y
If y = Len(mstring) And Mid(mstring, y, 1) = " " Then Exit Function

palavra = Empty
For p1 = y To Len(mstring)
If Mid(mstring, p1, 1) = " " Then Exit For
palavra = palavra & Mid(mstring, p1, 1)
Next p1

comandosStr2 = palavra
End Function
Public Sub DeletePasswordIntoName(str As String)
Dim x As Long
For x = 1 To Len(str)
If Mid(str, x, 1) = "@" Then
    str = Left(str, x - 1)
End If
Next x
End Sub

Public Function CorrigirString(ByVal str As String)
Dim x As Long

For x = 1 To Len(str)

If Mid(str, x, 2) = vbCrLf Then
    'CorrigirString = Right(str, Len(str) - x)
    CorrigirString = Left(str, x - 1)
    
    Exit Function
End If
Next x
End Function


Public Function GetOtherPlayersFromID(ByVal id As Long) As Long
Dim x As Long
For x = 1 To UBound(OtherPlayers)
    If OtherPlayers(x).Data.id = id Then
    GetOtherPlayersFromID = OtherPlayers(x).Data.id
    Exit Function
    End If
Next x
End Function

Public Function IsUnicode(s As String) As Boolean
If Len(s) = LenB(s) Then IsUnicode = False Else IsUnicode = True
End Function

Public Sub PutStringInArray(arr() As Byte, str As String)
On Error Resume Next

Dim x As Long
For x = LBound(arr) To UBound(arr)
    
    
    If Mid(str, x + 1, 1) = Empty Then arr(x) = Asc(vbLf): Exit Sub
    arr(x) = Asc(Mid(str, x + 1, 1))
   
    arr(x + 1) = Asc(vbLf)
Next x
End Sub

Public Function GetStringInArray(arr() As Byte) As String
On Error Resume Next

Dim x As Long

For x = LBound(arr) To UBound(arr)

   
    If Chr(arr(x)) = vbLf Then Exit Function
    
    GetStringInArray = GetStringInArray & Chr(arr(x))
Next x
End Function



Public Sub SortearTintas()
Randomize Timer
Dim ULimit As Long
Dim LLimit As Long
Dim ShiftStar As Long
Dim rotatePlanet As Long
Dim vermelho As Byte
Dim azul As Byte
Dim verde As Byte
Dim jatem As Boolean
Dim x As Long
Dim y As Long
ULimit = 0
LLimit = 255


For x = 0 To 254
    Do
        jatem = False
        vermelho = Int((ULimit - LLimit) * Rnd) + LLimit
        azul = Int((ULimit - LLimit) * Rnd) + LLimit
        verde = Int((ULimit - LLimit) * Rnd) + LLimit
    'verfifica se já tem
       
        For y = 0 To 254
            If CoresTinta(y).azul = azul And CoresTinta(y).vermelho = vermelho And CoresTinta(y).verde = verde Then jatem = True: Exit For
        Next y
        
        If jatem = False Then
            
            CoresTinta(x).verde = verde
            CoresTinta(x).azul = azul
            CoresTinta(x).vermelho = vermelho
            Exit Do
        End If
    Loop
Next x

End Sub

Public Function Luminosidade(ByVal color As Long, ByVal luz As Double) As Long
Dim rb As Byte
Dim gb As Byte
Dim bb As Byte
UnRGB color, rb, gb, bb
Dim r As Integer
Dim b As Integer
Dim g As Integer
Dim nCor As Long

r = rb
g = gb
b = bb
If r = 0 Then r = 1
If g = 0 Then g = 1
If b = 0 Then b = 1
If r * luz > 255 Then r = 255 Else r = r * luz
If g * luz > 255 Then g = 255 Else g = g * luz
If b * luz > 255 Then b = 255 Else b = b * luz

Luminosidade = RGB(r, g, b)
End Function

Public Sub AlterarCorCarro(ByVal cor As Long, Optional ByVal Todos As Boolean = True, Optional ImageIndex As Long)
On Error Resume Next
Dim cCor As Long
        Dim a As Long
        Dim rEmpty As RECT, rEmpty2 As RECT
        Dim ddsdOrigine As DDSURFACEDESC2
        Dim CorRGB As RGBColour
        Dim rC As Byte
        Dim bC As Byte
        Dim gC As Byte
        Dim corToUse As Long
        corToUse = cor

'porenquanto nao mudar cores

    'If cCor = 0 Then corToUse = RGB(CoresTinta(0).vermelho, CoresTinta(0).verde, CoresTinta(0).azul)
    'If cCor = 1 Then corToUse = RGB(CoresTinta(1).vermelho, CoresTinta(1).verde, CoresTinta(1).azul)
    'If cCor = 2 Then corToUse = RGB(CoresTinta(2).vermelho, CoresTinta(2).verde, CoresTinta(2).azul)
    'If cCor = 3 Then corToUse = RGB(CoresTinta(3).vermelho, CoresTinta(3).verde, CoresTinta(3).azul)
    'If cCor = 4 Then corToUse = RGB(CoresTinta(4).vermelho, CoresTinta(4).verde, CoresTinta(4).azul)
    'If cCor = 5 Then corToUse = RGB(CoresTinta(5).vermelho, CoresTinta(5).verde, CoresTinta(5).azul)
    'If cCor = 6 Then corToUse = RGB(CoresTinta(6).vermelho, CoresTinta(6).verde, CoresTinta(6).azul)
    'If cCor = 7 Then corToUse = RGB(CoresTinta(7).vermelho, CoresTinta(7).verde, CoresTinta(7).azul)
    
    'corToUse = RGB(0, 0, 100)
cCor = ImageIndex
        For a = 0 To 23
   
   'cars,angles,upDown,pneus
   'updown ->0= nivel normal , 1 =pra cima 1 nivel , 2=pra cima nivel2,3=pra baixo nivel, 4=pra baixo nivel 2
        LoadSprite cars(cCor, 0, a, 0, 0), App.Path & "\graficos\carros\marauder\a1\marauder" & a & ".bmp", 123, 72, CLng(Preto)
        'pneus
        MapearCor cars(cCor, 0, a, 0, 0), RGB(0, 255, 0), pixelsCoordenates(a, 0).pixels0_255_0Coordenates
        MapearCor cars(cCor, 0, a, 0, 0), RGB(255, 0, 0), pixelsCoordenates(a, 0).pixels255_0_0Coordenates
        MapearCor cars(cCor, 0, a, 0, 0), RGB(0, 0, 255), pixelsCoordenates(a, 0).pixels0_0_255Coordenates
        MapearCor cars(cCor, 0, a, 0, 0), RGB(100, 100, 100), pixelsCoordenates(a, 0).pixels100_100_100Coordenates
        MapearCor cars(cCor, 0, a, 0, 0), RGB(150, 150, 150), pixelsCoordenates(a, 0).pixels150_150_150Coordenates
        MapearCor cars(cCor, 0, a, 0, 0), RGB(200, 200, 200), pixelsCoordenates(a, 0).pixels200_200_200Coordenates
        
        
        
        ChangeColors cars(cCor, 0, a, 0, 0), RGB(0, 255, 0), RGB(90, 90, 90), pixelsCoordenates(a, 0).pixels0_255_0Coordenates
        ChangeColors cars(cCor, 0, a, 0, 0), RGB(255, 0, 0), RGB(54, 54, 54), pixelsCoordenates(a, 0).pixels255_0_0Coordenates
        ChangeColors cars(cCor, 0, a, 0, 0), RGB(0, 0, 255), RGB(15, 15, 15), pixelsCoordenates(a, 0).pixels0_0_255Coordenates
        changeCarColorTo cars(cCor, 0, a, 0, 0), corToUse, pixelsCoordenates(a, 0)
        
        If Todos = True Then
        LoadSprite cars(cCor, 0, a, 0, 1), App.Path & "\graficos\carros\marauder\a1\marauder" & a & ".bmp", 123, 72, CLng(Preto)
        
        'changeCarColorTo cars(cCor, 0, a, 0, 1), corToUse
        'pneus
        ChangeColors cars(cCor, 0, a, 0, 1), RGB(0, 0, 255), RGB(90, 90, 90), pixelsCoordenates(a, 0).pixels0_0_255Coordenates
        ChangeColors cars(cCor, 0, a, 0, 1), RGB(0, 255, 0), RGB(54, 54, 54), pixelsCoordenates(a, 0).pixels0_255_0Coordenates
        ChangeColors cars(cCor, 0, a, 0, 1), RGB(255, 0, 0), RGB(15, 15, 15), pixelsCoordenates(a, 0).pixels255_0_0Coordenates
        changeCarColorTo cars(cCor, 0, a, 0, 1), corToUse, pixelsCoordenates(a, 0)
        
        LoadSprite cars(cCor, 0, a, 0, 2), App.Path & "\graficos\carros\marauder\a1\marauder" & a & ".bmp", 123, 72, CLng(Preto)
        
        'changeCarColorTo cars(cCor, 0, a, 0, 2), corToUse
        'pneus
        ChangeColors cars(cCor, 0, a, 0, 2), RGB(255, 0, 0), RGB(90, 90, 90), pixelsCoordenates(a, 0).pixels255_0_0Coordenates
        ChangeColors cars(cCor, 0, a, 0, 2), RGB(0, 0, 255), RGB(54, 54, 54), pixelsCoordenates(a, 0).pixels0_0_255Coordenates
        ChangeColors cars(cCor, 0, a, 0, 2), RGB(0, 255, 0), RGB(15, 15, 15), pixelsCoordenates(a, 0).pixels0_255_0Coordenates
        changeCarColorTo cars(cCor, 0, a, 0, 2), corToUse, pixelsCoordenates(a, 0)
        
        LoadSprite cars(cCor, 0, a, 1, 0), App.Path & "\graficos\carros\marauder\up1\marauder" & a & ".bmp", 123, 72, CLng(Preto)
         'pneus
        MapearCor cars(cCor, 0, a, 1, 0), RGB(0, 255, 0), pixelsCoordenates(a, 1).pixels0_255_0Coordenates
        MapearCor cars(cCor, 0, a, 1, 0), RGB(255, 0, 0), pixelsCoordenates(a, 1).pixels255_0_0Coordenates
        MapearCor cars(cCor, 0, a, 1, 0), RGB(0, 0, 255), pixelsCoordenates(a, 1).pixels0_0_255Coordenates
        MapearCor cars(cCor, 0, a, 1, 0), RGB(100, 100, 100), pixelsCoordenates(a, 1).pixels100_100_100Coordenates
        MapearCor cars(cCor, 0, a, 1, 0), RGB(150, 150, 150), pixelsCoordenates(a, 1).pixels150_150_150Coordenates
        MapearCor cars(cCor, 0, a, 1, 0), RGB(200, 200, 200), pixelsCoordenates(a, 1).pixels200_200_200Coordenates
        
        ChangeColors cars(cCor, 0, a, 1, 0), RGB(0, 255, 0), RGB(90, 90, 90), pixelsCoordenates(a, 1).pixels0_255_0Coordenates
        ChangeColors cars(cCor, 0, a, 1, 0), RGB(255, 0, 0), RGB(54, 54, 54), pixelsCoordenates(a, 1).pixels255_0_0Coordenates
        ChangeColors cars(cCor, 0, a, 1, 0), RGB(0, 0, 255), RGB(15, 15, 15), pixelsCoordenates(a, 1).pixels0_0_255Coordenates
        changeCarColorTo cars(cCor, 0, a, 1, 0), corToUse, pixelsCoordenates(a, 1)
        
        LoadSprite cars(cCor, 0, a, 1, 1), App.Path & "\graficos\carros\marauder\up1\marauder" & a & ".bmp", 123, 72, CLng(Preto)
        ChangeColors cars(cCor, 0, a, 1, 1), RGB(0, 0, 255), RGB(90, 90, 90), pixelsCoordenates(a, 1).pixels0_0_255Coordenates
        ChangeColors cars(cCor, 0, a, 1, 1), RGB(0, 255, 0), RGB(54, 54, 54), pixelsCoordenates(a, 1).pixels0_255_0Coordenates
        ChangeColors cars(cCor, 0, a, 1, 1), RGB(255, 0, 0), RGB(15, 15, 15), pixelsCoordenates(a, 1).pixels255_0_0Coordenates
        changeCarColorTo cars(cCor, 0, a, 1, 1), corToUse, pixelsCoordenates(a, 1)
        
        LoadSprite cars(cCor, 0, a, 1, 2), App.Path & "\graficos\carros\marauder\up1\marauder" & a & ".bmp", 123, 72, CLng(Preto)
        'pneus
        ChangeColors cars(cCor, 0, a, 1, 2), RGB(255, 0, 0), RGB(90, 90, 90), pixelsCoordenates(a, 1).pixels255_0_0Coordenates
        ChangeColors cars(cCor, 0, a, 1, 2), RGB(0, 0, 255), RGB(54, 54, 54), pixelsCoordenates(a, 1).pixels0_0_255Coordenates
        ChangeColors cars(cCor, 0, a, 1, 2), RGB(0, 255, 0), RGB(15, 15, 15), pixelsCoordenates(a, 1).pixels0_255_0Coordenates
        changeCarColorTo cars(cCor, 0, a, 1, 2), corToUse, pixelsCoordenates(a, 1)
        
        LoadSprite cars(cCor, 0, a, 3, 0), App.Path & "\graficos\carros\marauder\d1a\marauder" & a & ".bmp", 123, 72, CLng(Preto)
        'pneus
        MapearCor cars(cCor, 0, a, 3, 0), RGB(0, 255, 0), pixelsCoordenates(a, 2).pixels0_255_0Coordenates
        MapearCor cars(cCor, 0, a, 3, 0), RGB(255, 0, 0), pixelsCoordenates(a, 2).pixels255_0_0Coordenates
        MapearCor cars(cCor, 0, a, 3, 0), RGB(0, 0, 255), pixelsCoordenates(a, 2).pixels0_0_255Coordenates
        MapearCor cars(cCor, 0, a, 3, 0), RGB(100, 100, 100), pixelsCoordenates(a, 2).pixels100_100_100Coordenates
        MapearCor cars(cCor, 0, a, 3, 0), RGB(150, 150, 150), pixelsCoordenates(a, 2).pixels150_150_150Coordenates
        MapearCor cars(cCor, 0, a, 3, 0), RGB(200, 200, 200), pixelsCoordenates(a, 2).pixels200_200_200Coordenates
        
        ChangeColors cars(cCor, 0, a, 3, 0), RGB(255, 0, 0), RGB(90, 90, 90), pixelsCoordenates(a, 2).pixels255_0_0Coordenates
        ChangeColors cars(cCor, 0, a, 3, 0), RGB(0, 0, 255), RGB(54, 54, 54), pixelsCoordenates(a, 2).pixels0_0_255Coordenates
        ChangeColors cars(cCor, 0, a, 3, 0), RGB(0, 255, 0), RGB(15, 15, 15), pixelsCoordenates(a, 2).pixels0_255_0Coordenates
        changeCarColorTo cars(cCor, 0, a, 3, 0), corToUse, pixelsCoordenates(a, 2)
        
        LoadSprite cars(cCor, 0, a, 3, 1), App.Path & "\graficos\carros\marauder\d1a\marauder" & a & ".bmp", 123, 72, CLng(Preto)
        'pneus
        ChangeColors cars(cCor, 0, a, 3, 1), RGB(0, 255, 0), RGB(90, 90, 90), pixelsCoordenates(a, 2).pixels0_255_0Coordenates
        ChangeColors cars(cCor, 0, a, 3, 1), RGB(255, 0, 0), RGB(54, 54, 54), pixelsCoordenates(a, 2).pixels255_0_0Coordenates
        ChangeColors cars(cCor, 0, a, 3, 1), RGB(0, 0, 255), RGB(15, 15, 15), pixelsCoordenates(a, 2).pixels0_0_255Coordenates
        changeCarColorTo cars(cCor, 0, a, 3, 1), corToUse, pixelsCoordenates(a, 2)
        
        LoadSprite cars(cCor, 0, a, 3, 2), App.Path & "\graficos\carros\marauder\d1a\marauder" & a & ".bmp", 123, 72, CLng(Preto)
        ChangeColors cars(cCor, 0, a, 3, 2), RGB(0, 0, 255), RGB(90, 90, 90), pixelsCoordenates(a, 2).pixels0_0_255Coordenates
        ChangeColors cars(cCor, 0, a, 3, 2), RGB(0, 255, 0), RGB(54, 54, 54), pixelsCoordenates(a, 2).pixels0_255_0Coordenates
        ChangeColors cars(cCor, 0, a, 3, 2), RGB(255, 0, 0), RGB(15, 15, 15), pixelsCoordenates(a, 2).pixels255_0_0Coordenates
        changeCarColorTo cars(cCor, 0, a, 3, 2), corToUse, pixelsCoordenates(a, 2)
        
        
            LoadSprite cars(cCor, 0, a, 4, 0), App.Path & "\graficos\carros\marauder\d1h\marauder" & a & ".bmp", 123, 72, CLng(Preto)
            MapearCor cars(cCor, 0, a, 4, 0), RGB(0, 255, 0), pixelsCoordenates(a, 3).pixels0_255_0Coordenates
            MapearCor cars(cCor, 0, a, 4, 0), RGB(255, 0, 0), pixelsCoordenates(a, 3).pixels255_0_0Coordenates
            MapearCor cars(cCor, 0, a, 4, 0), RGB(0, 0, 255), pixelsCoordenates(a, 3).pixels0_0_255Coordenates
            MapearCor cars(cCor, 0, a, 4, 0), RGB(100, 100, 100), pixelsCoordenates(a, 3).pixels100_100_100Coordenates
            MapearCor cars(cCor, 0, a, 4, 0), RGB(150, 150, 150), pixelsCoordenates(a, 3).pixels150_150_150Coordenates
            MapearCor cars(cCor, 0, a, 4, 0), RGB(200, 200, 200), pixelsCoordenates(a, 3).pixels200_200_200Coordenates
        
        'pneus
            ChangeColors cars(cCor, 0, a, 4, 0), RGB(255, 0, 0), RGB(90, 90, 90), pixelsCoordenates(a, 3).pixels255_0_0Coordenates
            ChangeColors cars(cCor, 0, a, 4, 0), RGB(0, 0, 255), RGB(54, 54, 54), pixelsCoordenates(a, 3).pixels0_0_255Coordenates
            ChangeColors cars(cCor, 0, a, 4, 0), RGB(0, 255, 0), RGB(15, 15, 15), pixelsCoordenates(a, 3).pixels0_255_0Coordenates
            changeCarColorTo cars(cCor, 0, a, 4, 0), corToUse, pixelsCoordenates(a, 3)
        
            LoadSprite cars(cCor, 0, a, 4, 1), App.Path & "\graficos\carros\marauder\d1h\marauder" & a & ".bmp", 123, 72, CLng(Preto)
         'pneus
            ChangeColors cars(cCor, 0, a, 4, 1), RGB(0, 255, 0), RGB(90, 90, 90), pixelsCoordenates(a, 3).pixels0_255_0Coordenates
            ChangeColors cars(cCor, 0, a, 4, 1), RGB(255, 0, 0), RGB(54, 54, 54), pixelsCoordenates(a, 3).pixels255_0_0Coordenates
            ChangeColors cars(cCor, 0, a, 4, 1), RGB(0, 0, 255), RGB(15, 15, 15), pixelsCoordenates(a, 3).pixels0_0_255Coordenates
            changeCarColorTo cars(cCor, 0, a, 4, 1), corToUse, pixelsCoordenates(a, 3)
        
            LoadSprite cars(cCor, 0, a, 4, 2), App.Path & "\graficos\carros\marauder\d1h\marauder" & a & ".bmp", 123, 72, CLng(Preto)
            ChangeColors cars(cCor, 0, a, 4, 2), RGB(0, 0, 255), RGB(90, 90, 90), pixelsCoordenates(a, 3).pixels0_0_255Coordenates
            ChangeColors cars(cCor, 0, a, 4, 2), RGB(0, 255, 0), RGB(54, 54, 54), pixelsCoordenates(a, 3).pixels0_255_0Coordenates
            ChangeColors cars(cCor, 0, a, 4, 2), RGB(255, 0, 0), RGB(15, 15, 15), pixelsCoordenates(a, 3).pixels255_0_0Coordenates
            changeCarColorTo cars(cCor, 0, a, 4, 2), corToUse, pixelsCoordenates(a, 3)
        End If
        Next a

CoresJaMapeadas = True
End Sub


Public Function MapearCor(tSprite As Sprites, ByVal cor As Long, buffer() As pixelCoord)
 If CoresJaMapeadas = True Then Exit Function
 'vermelho = 63488
 'verde = 2016
 'azul = 31
 'FrmDirectX.TmrChangeColor.Interval = 1
 'acha disponivel
 'Dim p As Long
 'For p = 0 To 1000
 'If ImagemAmudar(p).id = 0 Then
  '  ImagemAmudar(p).id = p + 1
   '  ImagemAmudar(p).sprite = tSprite
    ' ImagemAmudar(p).corSource = ColorSrc
     'ImagemAmudar(p).CorDest = ColorDest
     'Exit For
 'End If
 'Next p
 

 'Exit Function
 'trocar por 12710,2145 e 43776
 Dim lockrect As RECT
Dim SrcDesc         As DDSURFACEDESC2
Dim pixel(0 To 3) As Long
Dim x As Long
Dim altura As Long
Dim largura As Long
    
    tSprite.imagem.GetSurfaceDesc SrcDesc
    lockrect.Right = SrcDesc.lWidth
    lockrect.Bottom = SrcDesc.lHeight
    'tSprite.imagem.Lock lockrect, SrcDesc, DDLOCK_WAIT Or DDLOCK_NOSYSLOCK, 0
    Dim u As Long
    For altura = 0 To tSprite.Height - 1
        tSprite.imagem.Lock lockrect, SrcDesc, DDLOCK_WAIT Or DDLOCK_NOSYSLOCK, 0
        For largura = 0 To tSprite.Width - 1
            'DoEvents
            
            u = tSprite.imagem.GetLockedPixel(largura, altura)
            
            If u = cor Then
                ReDim Preserve buffer(0 To x)
                buffer(x).x = largura
                buffer(x).y = altura
                x = x + 1
                
            End If
            
        Next largura
        tSprite.imagem.Unlock lockrect
    Next altura
    'tSprite.imagem.Unlock lockrect
End Function
