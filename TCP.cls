VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCPIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private mytag As String
Private mwsaData As WSA_DATA
Private mlngSocket As Long

Private mlngRemotePort As Long

Private mstrLocalHostIP As String
Private msaLocalAddr As SOCK_ADDR

Private mstrRemoteHostIP As String
Private msaRemoteAddr As SOCK_ADDR

Private mlngLastErrorCode As Long
Private mstrLastErrorDesc As String
Private DataReceivedTemp() As Byte
Private wStatus As Integer
Private wState As Integer


Public Event connect()
Public Event Closed()
Public Event ConnectionRequest(ByVal connectionID As Long)
Public Event DataArrival(ByVal bytestotais As Long)
Public Event Error(ByVal Number As Integer, Description As String)
Public Event SendComplete()
Public Event SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)


Public Function connect2() As Boolean
    Dim WSAResult As Long
    
    connect2 = False
     
    'Initialize Winsock API
    WSAResult = WSAStartup(&H101, mwsaData)
    If WSAResult <> WSANOERROR Then
       SetLastErrorCode "Error en connect::WSAStartup"
       Exit Function
    End If
    
     'Create new socket
     mlngSocket = socket(PF_INET, SOCK_STREAM, 0)
     If (mlngSocket = INVALID_SOCKET) Then
        SetLastErrorCode "Error in connect::socket"
        Exit Function
     End If
    
     'Bind socket to LocalHostIP
     msaLocalAddr.sin_family = PF_INET
     msaLocalAddr.sin_port = 0
     msaLocalAddr.sin_addr.S_addr = inet_addr(mstrLocalHostIP)
     If (msaLocalAddr.sin_addr.S_addr = INADDR_NONE) Then
        SetLastErrorCode "Error in connect::inet_addr"
        Exit Function
     End If
     WSAResult = bind(mlngSocket, msaLocalAddr, Len(msaLocalAddr))
     If (WSAResult = SOCKET_ERROR) Then
        SetLastErrorCode "Error in connect::bind"
        Exit Function
     End If
    'resolve address
'    sHostName = Trim$(sHostName)
   Dim lphost As Long
   lphost = gethostbyname(mstrRemoteHostIP)
  
   If lphost = 0 Then
      'GetIPAddress = ""
      'SetLastErrorCode "Windows Sockets are not responding. " & _
              "Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
   'to extract the returned IP address, we have to copy
  'the HOST structure and its members
    Dim HOST      As HOSTENT
    Dim dwIPAddr  As Long
    Dim tmpIPAddr() As Byte
    Dim i As Long
    Dim sIPAddr  As String

   CopyMemory HOST, lphost, Len(HOST)
   CopyMemory dwIPAddr, HOST.hAddrList, 4

   'create an array to hold the result
   ReDim tmpIPAddr(1 To HOST.hLen)
   CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
   
    'and with the array, build the actual address,
  'appending a period between members
   For i = 1 To HOST.hLen
      sIPAddr = sIPAddr & tmpIPAddr(i) & "."
   Next
   
  mstrRemoteHostIP = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)



  ' mstrRemoteHostIP = lphost
    
    
     'Connect with remote host
     msaRemoteAddr.sin_family = PF_INET
     msaRemoteAddr.sin_port = htons(mlngRemotePort)
     msaRemoteAddr.sin_addr.S_addr = inet_addr(mstrRemoteHostIP)
     If (msaLocalAddr.sin_addr.S_addr = INADDR_NONE) Then
        SetLastErrorCode "Error in connect::inet_addr=INADDR_NONE"
        Exit Function
     End If
     msaRemoteAddr.sin_zero(0) = 0
     
     WSAResult = connect(mlngSocket, msaRemoteAddr, Len(msaRemoteAddr))
     If (WSAResult = SOCKET_ERROR) Then
        SetLastErrorCode "Error in connect::connect"
     Else
        connect2 = True
     End If
End Function

Public Function GetData(ByVal BufferToReceiveAddress) As Long
CopyMemory ByVal BufferToReceiveAddress, DataReceivedTemp(0), UBound(DataReceivedTemp) + 1

End Function
Public Function ReceiveData() As Long
    Const MAX_BUFF_SIZE = 10000
    Dim buff(0 To MAX_BUFF_SIZE) As Byte
    Dim WSAResult As Long
    
    WSAResult = recv(mlngSocket, buff(0), MAX_BUFF_SIZE, 0)
'Form9.Visible = True
'Form9.Caption = WSAResult
    If WSAResult = 0 Then
        ReceiveData = 0
        
        RaiseEvent Error(30017, WSAGetLastError)
'        strData = ""
        'lngbytesreceived = 0
        
    End If
    If WSAResult > 0 Then
    '        lngbytesreceived = WSAResult
        buff(WSAResult) = 0
        'strData = Left(StrConv(buff(), vbUnicode), lngBytesReceived)
        If WSAResult > 0 Then
            ReDim DataReceivedTemp(0 To WSAResult - 1)
            CopyMemory DataReceivedTemp(0), buff(0), WSAResult
            ReceiveData = WSAResult
            RaiseEvent DataArrival(WSAResult)
        End If
    End If
End Function

Public Function SendStringData(ByVal strData As String) As Boolean

   Dim WSAResult As Long, i As Long, L As Long
  '
    L = Len(strData)
    ReDim buff(L + 1) As Byte
    
    For i = 1 To L
       buff(i - 1) = Asc(Mid(strData, i, 1))
    Next
    buff(L) = 0

    WSAResult = send(mlngSocket, buff(0), L, 0)
    If WSAResult = SOCKET_ERROR Then
        SetLastErrorCode "Error en SendData::send"
        SendStringData = False
    Else
        SendStringData = True
    End If
End Function

Public Function IsDataAvailable() As Boolean
    
    Dim readfds As FD_SET, writefds As FD_SET, exceptfds As FD_SET
    Dim timeout As TIME_VAL
    Dim lngResult As Long, nfds As Long
    
    nfds = 0
    timeout.tv_sec = 0
    timeout.tv_usec = 1
    
    readfds.fd_count = 1
    readfds.fd_array(0) = mlngSocket
    writefds.fd_count = 0
    exceptfds.fd_count = 0
    
    lngResult = sselect(nfds, readfds, writefds, exceptfds, timeout)
    If lngResult = SOCKET_ERROR Then
        SetLastErrorCode "Error in IsDataAvailable::select"
        IsDataAvailable = False
    Else
        If lngResult > 0 Then IsDataAvailable = True Else IsDataAvailable = False
    End If
End Function

Public Sub ShutdownConnection()
     wState = 0
     Dim lngResult As Long
     lngResult = shutdown(mlngSocket, 2)
     lngResult = closesocket(mlngSocket)
     lngResult = WSACancelBlockingCall
     lngResult = WSACleanup
     mlngSocket = 0
     mytag = Empty
End Sub

Private Sub SetLastErrorCode(ByVal strAdditionalInfo As String)
    mlngLastErrorCode = WSAGetLastError
    Select Case mlngLastErrorCode
    Case INADDR_NONE
                            mstrLastErrorDesc = "INNADDR_NONE: The passed-in string does not contain a legitimate Internet address (for example, if a portion of an 'a.b.c.d' address exceeds 255), inet_addr returns the value INADDR_NONE"
    Case WSASYSNOTREADY
                            mstrLastErrorDesc = "WSASYSNOTREADY: Indicates that the underlying network subsystem is not ready for network communication"
    Case WSAVERNOTSUPPORTED
                            mstrLastErrorDesc = "WSAVERNOTSUPPORTED: The version of Windows Sockets support requested is not provided by this particular Windows Sockets implementation"
    Case WSAEINVAL
                            mstrLastErrorDesc = "WSAEINVAL: The Windows Sockets version specified by the application is not supported by this DLL"
    Case WSANOTINITIALISED
                            mstrLastErrorDesc = "WSANOTINITIALISED: A successful WSAStartup must occur before using this function"
    Case WSAENETDOWN
                            mstrLastErrorDesc = "WSAENETDOWN: The Windows Sockets implementation has detected that the network subsystem has failed"
    Case WSAEAFNOSUPPORT
                            mstrLastErrorDesc = "WSAEAFNOSUPPORT: The specified address family is not supported"
    Case WSAEINPROGRESS
                            mstrLastErrorDesc = "WSAEINPROGRESS: A blocking Windows Sockets operation is in progress"
    Case WSAEMFILE
                            mstrLastErrorDesc = "WSAEMFILE: No more file descriptors are available"
    Case WSAENOBUFS
                            mstrLastErrorDesc = "WSAENOBUFS: No buffer space is available. The socket cannot be created"
    Case WSAEPROTONOSUPPORT
                            mstrLastErrorDesc = "WSAEPROTONOSUPPORT: The specified protocol is not supported"
    Case WSAEPROTOTYPE
                            mstrLastErrorDesc = "WSAEPROTOTYPE: The specified protocol is the wrong type for this socket"
    Case WSAESOCKTNOSUPPORT
                            mstrLastErrorDesc = "WSAESOCKTNOSUPPORT: The specified socket type is not supported in this address family"
    Case WSAEADDRINUSE
                            mstrLastErrorDesc = "WSAEADDRINUSE: The specified address is already in use"
    Case WSAEINTR
                            mstrLastErrorDesc = "WSAEINTR: The (blocking) call was canceled using WSACancelBlockingCall"
    Case WSAEADDRNOTAVAIL
                            mstrLastErrorDesc = "WSAEADDRNOTAVAIL: The specified address is not available from the local computer"
    Case WSAECONNREFUSED
                            mstrLastErrorDesc = "WSAECONNREFUSED: The attempt to connect was forcefully rejected"
    Case WSAEFAULT
                            mstrLastErrorDesc = "WSAEFAULT: The namelen argument is incorrect"
    Case WSAEISCONN
                            mstrLastErrorDesc = "WSAEISCONN: The socket is already connected"
    Case WSAENETUNREACH
                            mstrLastErrorDesc = "WSAENETUNREACH: The network cant be reached from this host at this time"
    Case WSAENOTSOCK
                            mstrLastErrorDesc = "WSAENOTSOCK: The descriptor is not a socket"
    Case WSAETIMEDOUT
                            mstrLastErrorDesc = "WSAETIMEDOUT: Attempt to connect timed out without establishing a connection"
    Case WSAEWOULDBLOCK
                            mstrLastErrorDesc = "WSAEWOULDBLOCK: The socket is marked as nonblocking and the connection cannot be completed immediately. It is possible to select the socket while it is connecting by selecting it for writing"
    Case WSAEACCES
                            mstrLastErrorDesc = "The requested address is a broadcast address, but the appropriate flag was not set"
    Case WSAENETRESET
                            mstrLastErrorDesc = "The connection must be reset because the Windows Sockets implementation dropped it"
    Case WSAENOTCONN
                            mstrLastErrorDesc = "The socket is not connected"
    Case WSAEOPNOTSUPP
                            mstrLastErrorDesc = "MSG_OOB was specified, but the socket is not of type SOCK_STREAM"
    Case WSAESHUTDOWN
                            mstrLastErrorDesc = "The socket has been shutdown it is not possible to send on a socket after shutdown has been invoked with how set to 1 or 2"
    Case WSAEMSGSIZE
                            mstrLastErrorDesc = "The socket is of type SOCK_DGRAM, and the datagram is larger than the maximum supported by the Windows Sockets implementation"
    Case WSAECONNABORTED
                            mstrLastErrorDesc = "The virtual circuit was aborted due to timeout or other failure"
    Case WSAECONNRESET
                            mstrLastErrorDesc = "The virtual circuit was reset by the remote side"
    Case Else
         mstrLastErrorDesc = "Case unknown in TCPIP::SetLastErrorCode = " & mlngLastErrorCode
    End Select
    
    mstrLastErrorDesc = strAdditionalInfo & vbCrLf & mstrLastErrorDesc
    
End Sub

Public Property Get RemotePort() As Long
    RemotePort = mlngRemotePort
End Property
Public Property Get RemoteHost() As String
    RemoteHost = mstrRemoteHostIP
End Property
Public Property Let RemotePort(ByVal lngRemotePort As Long)
    mlngRemotePort = Trim(lngRemotePort)
End Property

Public Property Let RemoteHost(ByVal strRemoteIP As String)
    mstrRemoteHostIP = Trim(strRemoteIP)
End Property

Public Property Let LocalHostIP(ByVal strLocalHostIP As String)
    mstrLocalHostIP = strLocalHostIP
End Property

Public Property Get State() As Long
    State = wState
End Property

Public Property Get ErrorDescription() As String
    ErrorDescription = mstrLastErrorDesc
End Property

Public Property Get errorcode() As Long
    errorcode = mlngLastErrorCode
End Property

Private Sub Class_Initialize()
mstrLocalHostIP = "0.0.0.0"
End Sub

Private Sub Class_Terminate()
    If mlngSocket <> 0 Then ShutdownConnection
End Sub

Public Property Get Tag() As String
    Tag = mytag
End Property

Public Property Get Status() As Long
    Status = wStatus
End Property

Public Property Let Tag(valor As String)
    mytag = valor
End Property

Public Property Get SocketHandle() As Long
    SocketHandle = mlngSocket
End Property
Public Property Get DataReceived() As Long
    DataReceived = VarPtr(DataReceivedTemp(0))
End Property


Public Function SendData(ByVal lpbuffer As Long, ByVal quantidade As Long) As Boolean
Dim totaltransmitido As Long
Dim res As Long
Dim socket As Long
socket = mlngSocket

res = send(socket, ByVal lpbuffer, quantidade, 0)

If res = SOCKET_ERROR Then
    SetLastErrorCode "Error en SendData::send"
    If res = 10050 Or res = 10051 Or res = 10052 Or res = 10053 Or res = 10054 Or res = 10057 Or res = 10058 Or res = 10060 Or res = 10061 Or res = 10064 Or res = 10065 Then wState = 9
    SendData = -1
    DoEvents
Else
    SendData = res
    DoEvents
End If

End Function


Public Function openconnection() As Long
    wState = 0
    Dim WSAResult As Long
    openconnection = False
    If wState = 9 Then openconnection = 1: RaiseEvent Error(30003, "shtudown the conection"): Exit Function
     If mlngRemotePort = 0 Then openconnection = 2: RaiseEvent Error(30001, "port not specified"): Exit Function
     If Trim(mstrRemoteHostIP) = Empty Then openconnection = 3: RaiseEvent Error(30002, "IP address not specified"): Exit Function
    'Initialize Winsock API
    WSAResult = WSAStartup(&H101, mwsaData)
    If WSAResult <> WSANOERROR Then
       SetLastErrorCode "Error en connect::WSAStartup"
       openconnection = 4
       Exit Function
    End If
    
     'Create new socket
     mlngSocket = socket(AF_INET, SOCK_STREAM, 0)
     If (mlngSocket = INVALID_SOCKET) Then
        SetLastErrorCode "Error in connect::socket"
        openconnection = 5
        RaiseEvent Error(30004, "Error in connect::socket")

        Exit Function
     End If
    wState = 1
     'Bind socket to LocalHostIP
     msaLocalAddr.sin_family = AF_INET
     msaLocalAddr.sin_port = 0
     msaLocalAddr.sin_addr.S_addr = inet_addr(mstrLocalHostIP)

     If (msaLocalAddr.sin_addr.S_addr = INADDR_NONE) Then
        SetLastErrorCode "Error in connect::inet_addr"
        openconnection = 6
        wState = 9
        RaiseEvent Error(30005, "Error in connect::inet_addr")
        Exit Function
     End If
     WSAResult = bind(mlngSocket, msaLocalAddr, Len(msaLocalAddr))
     If (WSAResult = SOCKET_ERROR) Then
        SetLastErrorCode "Error in connect::bind"
        openconnection = 7
        wState = 9
        RaiseEvent Error(30006, "Error in connect::bind")
        Exit Function
     End If
    'resolve address
'    sHostName = Trim$(sHostName)

wState = 4
Dim errocode As Long
mstrRemoteHostIP = GetIPFromHostName(mstrRemoteHostIP, errocode)

If errocode <> 0 Then wState = 9: openconnection = 10: RaiseEvent Error(30007, "can't resolve name - " & "error : " & errocode & " " & GetWsErrorCodeDescription(errocode)): Exit Function
wState = 5

  ' mstrRemoteHostIP = lphost
    
    
     'Connect with remote host
     msaRemoteAddr.sin_family = AF_INET
     msaRemoteAddr.sin_port = htons(mlngRemotePort)
     msaRemoteAddr.sin_addr.S_addr = inet_addr(mstrRemoteHostIP)
     If (msaLocalAddr.sin_addr.S_addr = INADDR_NONE) Then
        SetLastErrorCode "Error in connect::inet_addr=INADDR_NONE"
        openconnection = 11
        wState = 11
        Exit Function
     End If
     msaRemoteAddr.sin_zero(0) = 0
'        wState = 7
     WSAResult = connect(mlngSocket, msaRemoteAddr, Len(msaRemoteAddr))
     If (WSAResult = SOCKET_ERROR) Then
        
        SetLastErrorCode "Error in connect::connect"
        openconnection = 12
        wState = 12
     Else
        openconnection = 0
        wState = 7
        RaiseEvent connect
        'Do
        'If IsDataAvailable Then
        '    ReceiveData
        'End If
        'DoEvents
        'Loop
     End If
End Function


Private Function GetWsErrorCodeDescription(code As Long) As String
Select Case code
    Case 6
    GetWsErrorCodeDescription = "Invalid Handle"
    Exit Function
    
    Case 8
    GetWsErrorCodeDescription = "Not enough memory"
    Exit Function
    
    Case 87
    GetWsErrorCodeDescription = "Invalid Parameter"
    Exit Function
    
    Case 995
    GetWsErrorCodeDescription = "Overlapped operation aborted"
    Exit Function
    
    Case 996
    GetWsErrorCodeDescription = "Overlapped I/O event object not in signaled state"
    Exit Function
    
    Case 997
    GetWsErrorCodeDescription = "Overlapped operations will complete later"
    Exit Function
    
    Case 10004
    GetWsErrorCodeDescription = "Interrupted function call"
    Exit Function
    
    Case 10013
    GetWsErrorCodeDescription = "Permission denied"
    Exit Function
    
    Case 10014
    GetWsErrorCodeDescription = "Bad address"
    Exit Function
    
    Case 10022
    GetWsErrorCodeDescription = "Invalid argument"
    Exit Function
    
    Case 10024
    GetWsErrorCodeDescription = "Too many open files"
    Exit Function
    
    Case 10035
    GetWsErrorCodeDescription = "Resource temporarily unvaliable"
    Exit Function
    
    Case 10036
    GetWsErrorCodeDescription = "Operation now in progress"
    Exit Function
    
    Case 10037
    GetWsErrorCodeDescription = "Operation already in progress"
    Exit Function
    
    Case 10038
    GetWsErrorCodeDescription = "Socket operation in nonsocket"
    Exit Function
    
    Case 10039
    GetWsErrorCodeDescription = "Destination adress required"
    Exit Function
    
    Case 10040
    GetWsErrorCodeDescription = "Message too long"
    Exit Function
    
    Case 10041
    GetWsErrorCodeDescription = "Protocol wrong type for socket"
    Exit Function
    
    Case 10042
    GetWsErrorCodeDescription = "Bad protocol option"
    Exit Function
    
    Case 10043
    GetWsErrorCodeDescription = "protocol not supported"
    Exit Function
    
    Case 10044
    GetWsErrorCodeDescription = "socket type not supported"
    Exit Function
    
    Case 10045
    GetWsErrorCodeDescription = "operation not supported"
    Exit Function
    
    Case 10046
    GetWsErrorCodeDescription = "Protocol family not supported"
    Exit Function
    
    Case 10047
    GetWsErrorCodeDescription = "Address family not supported by protocol family"
    Exit Function
    
    Case 10048
    GetWsErrorCodeDescription = "Address already in use"
    Exit Function
    
    Case 10049
    GetWsErrorCodeDescription = "Cannot assign requested address"
    Exit Function
    
    Case 10050
    GetWsErrorCodeDescription = "Network is down"
    Exit Function
    
    Case 10051
    GetWsErrorCodeDescription = "Network is unreachable"
    Exit Function
    
    Case 10052
    GetWsErrorCodeDescription = "Network dropped connection on reset"
    Exit Function
    
    Case 10053
    GetWsErrorCodeDescription = "Software caused connection abort"
    Exit Function
    
    Case 10054
    GetWsErrorCodeDescription = "Connection reset by peer"
    Exit Function
    
    Case 10055
    GetWsErrorCodeDescription = "No buffer space avaliable"
    Exit Function
    
    Case 10056
    GetWsErrorCodeDescription = "socket is already connected"
    Exit Function
    
    Case 10057
    GetWsErrorCodeDescription = "socket is not connected"
    Exit Function
    
    Case 10058
    GetWsErrorCodeDescription = "cannot send after socket shutdown"
    Exit Function
    
    Case 10059
    GetWsErrorCodeDescription = "too many references"
    Exit Function
    
    Case 10060
    GetWsErrorCodeDescription = "connection timeout"
    Exit Function
    
    Case 10061
    GetWsErrorCodeDescription = "connection refused"
    Exit Function
    
    Case 10062
    GetWsErrorCodeDescription = "cannot translate name"
    Exit Function
    
    Case 10063
    GetWsErrorCodeDescription = "name too long"
    Exit Function
    
    Case 10064
    GetWsErrorCodeDescription = "host is down"
    Exit Function
    
    Case 10065
    GetWsErrorCodeDescription = "no route to host"
    Exit Function
    
    Case 10066
    GetWsErrorCodeDescription = "directory not empty"
    Exit Function
    
    Case 10067
    GetWsErrorCodeDescription = "too many processes"
    Exit Function
    
    Case 10068
    GetWsErrorCodeDescription = "user quota exceeded"
    Exit Function
    
    Case 10069
    GetWsErrorCodeDescription = "disk quota exceeded"
    Exit Function
    
    Case 10070
    GetWsErrorCodeDescription = "stale file handle reference"
    Exit Function
End Select
End Function







