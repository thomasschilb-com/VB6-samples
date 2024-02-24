VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimpleServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' ---------------------------------------------------------------------------------
' File...........: SimpleServer.cls
' Author.........: J.A. Coutts
' Created........: 11/01/2016
' Updated........: 07/07/2017 - Added uInBuffer, uOutBuffer
' Bug fix........: 07/19/2017 - in DeleteByte routine in Modsocket
' Version........: 1.1
' Website........: http://www.yellowhead.com
' Contact........: allecnarf@hotmail.com
'
'Copyright (c) 2016 by JAC Computing
'Vernon, BC, Canada
'
'Subclassing based on CSocketMaster by Emiliano Scavuzzo
'
'SimpleServer basically performs the same functions as NewSocket. Like NewSocket,
'it supports IPv6 as well as IPv4. As long as IPv4 is used, it will operate on
'pre Windows Vista systems that do not support dual stack. This was accomplished
'by using inet_ntoa for IPv4 functions instead of inet_ntop. Unlike NewSocket,
'it cannot be used as a Control Array because of the way it handles listening
'sockets.
'
'All error messages are returned to the calling program through the Error event
'(using the CallBack function). NewSocket and it's predecessors simply reported
'errors and caried on attempting communication. This sometimes made debugging
'difficult as mutiple errors would get reported. SimpleServer attempts to exit the
'problem function on an error that would cause communication to fail.
'

'================================
'Member Socket Variables
'================================
Private m_Callback As SimpleServer
Private m_Index    As Long

'================================
'General Socket Constants
'================================
Private Const AF_UNSPEC As Long = 0
Private Const AF_INET As Long = 2
Private Const AF_INET6 As Long = 23
Private Const SOCK_STREAM As Long = 1
Private Const SOCK_DGRAM As Long = 2
Private Const INADDR_ANY As Long = 0
Private Const IPPROTO_TCP As Long = 6
Private Const IPPROTO_UDP As Long = 17
Private Const SOL_SOCKET As Long = 65535
Private Const SO_SNDBUF         As Long = &H1001&
Private Const SO_RCVBUF         As Long = &H1002&
Private Const SO_MAX_MSG_SIZE   As Long = &H2003
Private Const SO_BROADCAST      As Long = &H20
Private Const SOMAXCONN         As Long = 5
Private Const LOCAL_HOST_BUFF   As Integer = 256
Private Const FIONREAD          As Long = &H4004667F
Private Const SO_REUSEADDR As Long = &H4            'Allow local address reuse
Private Const SO_LINGER As Long = &H80              'Set TIME_WAIT
Private Const SCK_TCP As Long = 0
Private Const SCK_UDP As Long = 1
Private Const MSG_PEEK As Long = &H2

' Length of string fields for IPv4 and IPv6
Private Const INET_ADDRSTRLEN   As Long = 16
Private Const INET6_ADDRSTRLEN  As Long = 46
Private Const AI_PASSIVE        As Long = 1

'WINSOCK CONTROL ERROR CODES
Private Const sckBadState = 40006
Private Const sckInvalidArg = 40014
Private Const sckUnsupported = 40018
Private Const sckInvalidOp = 40020

'================================
'Socket states
'================================
Private Enum SockState
    sckClosed = 0
    sckOpen
    scklistening
    sckConnectionPending
    sckResolvingHost
    sckHostResolved
    sckconnecting
    sckConnected
    sckClosing
    sckError
End Enum

'================================
'MEMBER VARIABLES
'================================
Private m_hSocket           As Long    'socket handle
Private m_hListen           As Long    'listening socket handle
Private m_lMemoryPointer    As Long    'memory pointer used as buffer when resolving host
Private m_lMemoryHandle     As Long    'buffer memory handle
Private m_sRemoteHost       As String  'remote host
Private m_lRemotePort       As Long    'remote port
Private m_sRemotePort       As String  'remote port
Private m_sRemoteHostIP     As String  'remote host ip
Private m_lLocalPort        As Long    'local port
Private m_lPortListen       As Long    'listening port
Private m_sLocalIP          As String  'local IP
Private m_State             As SockState    'socket state
Private m_Protocol          As Long    'defaults to 0 = TCP (UDP = 1)
Private m_IPv               As Long    'IP Version
Private m_bAcceptClass      As Boolean 'if True then this is a Accept socket class
Private m_bEncr             As Boolean 'True for secure sockets
Private m_lSendBufferLen    As Long    'winsock buffer size for sends
Private m_lRecvBufferLen    As Long    'winsock buffer size for receives

Private Type sockaddr
    sa_family           As Integer  '2 bytes
    sa_data(25)         As Byte     '26 bytes
End Type                            'Total 28 bytes

'Basic IPv4 addressing structures.
Private Type in_addr
   s_addr   As Long
End Type
Private Type sockaddr_in
    sin_family          As Integer  '2 bytes
    sin_port            As Integer  '2 bytes
    sin_addr            As in_addr  '4 bytes
    sin_zero(0 To 7)    As Byte     '8 bytes
End Type                            'Total 16 bytes

'Basic IPv6 addressing structures.
Private Type in6_addr
    s6_addr(0 To 15)      As Byte
End Type
Private Type sockaddr_in6
    sin6_family         As Integer  '2 bytes
    sin6_port           As Integer  '2 bytes
    sin6_flowinfo       As Long     '4 bytes
    sin6_addr           As in6_addr '16 bytes
    sin6_scope_id       As Long     '4 bytes
End Type                            'Total 28 bytes

Private Type addrinfo
    ai_flags As Long
    ai_family As Long
    ai_socktype As Long
    ai_protocol As Long
    ai_addrlen As Long
    ai_canonname As Long 'strptr
    ai_addr As Long 'p sockaddr
    ai_next As Long 'p addrinfo
End Type

Private m_bSendBuffer() As Byte     'local outbound byte buffer
Private m_bRecvBuffer() As Byte     'local inbound byte buffer
Private m_bData()       As Byte     'Temporary store for recovered data
Private sa_UDPdest      As sockaddr 'UDP destination socket address

'================================
'API Functions
'================================
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function API_Socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal stype As Long, ByVal Protocol As Long) As Long
Private Declare Function API_CloseSocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function API_Connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef name As sockaddr, ByVal namelen As Long) As Long
Private Declare Function API_GetAddrInfo Lib "ws2_32.dll" Alias "getaddrinfo" (ByVal NodeName As String, ByVal ServName As String, ByVal lpHints As Long, lpResult As Long) As Long
Private Declare Function API_FreeAddrInfo Lib "ws2_32.dll" Alias "freeaddrinfo" (ByVal res As Long) As Long
Private Declare Function API_Send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long) As Long
Private Declare Function API_SendTo Lib "ws2_32.dll" Alias "sendto" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long, ByRef toaddr As sockaddr, ByVal tolen As Long) As Long
Private Declare Function API_Bind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef name As sockaddr, ByRef namelen As Long) As Long
Private Declare Function API_Listen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function API_Accept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef addr As sockaddr, ByRef addrlen As Long) As Long
Private Declare Function API_GetSockOpt Lib "ws2_32.dll" Alias "getsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function API_SetSockOpt Lib "ws2_32.dll" Alias "setsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function API_GetPeerName Lib "ws2_32.dll" Alias "getpeername" (ByVal s As Long, ByRef name As sockaddr, ByRef namelen As Long) As Long
Private Declare Function API_GetSockName Lib "ws2_32.dll" Alias "getsockname" (ByVal s As Long, ByRef name As sockaddr, ByRef namelen As Long) As Long
Private Declare Function API_GetHostName Lib "ws2_32.dll" Alias "gethostname" (ByVal host_name As String, ByVal namelen As Long) As Long
Private Declare Function API_IoctlSocket Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
Private Declare Function API_Recv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long) As Long
Private Declare Function API_RecvFrom Lib "ws2_32.dll" Alias "recvfrom" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long, ByRef fromaddr As sockaddr, ByRef fromlen As Long) As Long

Private Declare Function inet_ntop Lib "ws2_32.dll" (ByVal af As Long, ByRef ppAddr As Any, ByRef pStringBuf As Any, ByVal StringBufSize As Long) As Long
Private Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal iaddr As Long) As Long
Private Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Long) As Integer

'================================
'Generalized functions
'================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Friend Property Get BufferSize() As Long
    BufferSize = m_lSendBufferLen
End Property

Public Sub WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    '  ****  WARNING WARNING WARNING WARNING ******
    'This sub MUST be the first on the class. DO NOT attempt
    'to change it's location or the code will CRASH.
    'This sub receives system messages from our WndProc.
    Dim LoWord As Long
    Dim HiWord As Long
    Select Case uMsg
    Case SOCKET_MESSAGE
        LoWord = lParam And &HFFFF&
        If (lParam And &H80000000) = &H80000000 Then
            HiWord = ((lParam And &H7FFF0000) \ &H10000) Or &H8000&
        Else
            HiWord = (lParam And &HFFFF0000) \ &H10000
        End If
        PostSocket LoWord, HiWord
    End Select
End Sub


Friend Function Accept(requestID As Long, lRemotePort As Long, sRemoteHostIP As String) As Boolean
    Const Routine As String = "SimpleServer.Accept"
    'Socket should be either Closed or in Error
    If m_State <> sckClosed Then
        Call m_Callback.Error(0, sckInvalidOp, "Socket is not Available", Routine)
    End If
    'If the calling program indicates that the connection wasn't accepted by
    'setting lRemotePort to zero, we must close the socket that was created.
    If lRemotePort = 0 Then
        API_CloseSocket requestID
        Call PrintDebug("OK Closed accepted socket: " & CStr(requestID))
        Exit Function
    End If
    m_hSocket = requestID 'Use the handle created by FD_ACCEPT message
    m_sRemoteHostIP = sRemoteHostIP
    m_lRemotePort = lRemotePort
    m_Protocol = SCK_TCP
    Call ProcessOptions
    If IsSocketRegistered(requestID) Then
        Call m_Callback.Error(0, sckBadState, "Wrong protocol or connection state for the requested transaction or request", Routine)
        Exit Function
    Else
        m_State = sckConnected
        Call PrintDebug("STATE: sckConnected")
        Call modSocket.RegisterSocket(m_hSocket, ObjPtr(Me), False)
    End If
    Accept = True
End Function


Friend Property Get bInBuffer() As Byte()
    bInBuffer = m_bData
End Property

Friend Property Let bOutBuffer(bNewValue() As Byte)
    Call AddByte(m_bSendBuffer, bNewValue)
End Property

Friend Function BuildArray(ByVal lSize As Long, ByVal blnPeek As Boolean, ByRef lErrorCode As Long) As Byte()
    Dim lRet        As Long
    Dim tmpSa       As sockaddr
    Dim lFlags      As Long
    Dim bTmp() As Byte
    If m_Protocol = SCK_TCP Then 'TCP transfers data from m_bRecvBuffer
        ReDim bTmp(lSize - 1)
        CopyMemory bTmp(0), m_bRecvBuffer(0), lSize
        BuildArray = bTmp
        If Not blnPeek Then
            Call DeleteByte(m_bRecvBuffer, lSize)
        End If
    Else 'UDP recovers data from Winsock buffer
        If blnPeek Then lFlags = MSG_PEEK
        ReDim bTmp(lSize - 1)
        lRet = API_RecvFrom(m_hSocket, bTmp(0), lSize, lFlags, tmpSa, LenB(tmpSa))
        If lRet = SOCKET_ERROR Then
            lErrorCode = Err.LastDllError
        End If
        BuildArray = bTmp
        GetRemoteInfoFromSI tmpSa, m_lRemotePort, m_sRemoteHostIP, m_sRemoteHost
    End If
End Function

Friend Property Get BytesReceived() As Long
    If m_Protocol = SCK_TCP Then
        BytesReceived = GetbSize(m_bRecvBuffer)
    Else
        BytesReceived = GetBufferLenUDP
    End If
End Property

Friend Function CloseSocket() As Boolean
    Const Routine As String = "SimpleServer.CloseSocket"
    Dim lErrorCode As Long
    Dim lRet As Long
    Dim bTmp() As Byte
    If Not m_hSocket = SOCKET_ERROR Then
        lRet = API_CloseSocket(m_hSocket)
        If lRet = -1 Then
            m_State = sckError
            lErrorCode = Err.LastDllError
            Call m_Callback.Error(m_Index, lErrorCode, "Could not Close Socket!", Routine)
        Else
            Call modSocket.UnregisterSocket(m_hSocket)
            m_hSocket = SOCKET_ERROR
            m_State = sckClosed
            Call PrintDebug("STATE: sckClosed")
            m_sLocalIP = vbNullString
            m_sRemoteHostIP = ""
            m_bRecvBuffer = bTmp
            m_bSendBuffer = bTmp
            m_lSendBufferLen = 0
            m_lRecvBufferLen = 0
            CloseSocket = True
        End If
    End If
End Function



Private Sub EmptyBuffer()
    'Empty winsock incoming buffer from a UDP socket.
    Dim B As Byte
    API_Recv m_hSocket, B, Len(B), 0&
End Sub

Friend Property Get Callback(Optional ByVal Index As Long) As SimpleServer
'The Friend scope modifier prevents the following properties and methods from
'being Implemented yet still allows them to be accessible within the project
    Set Callback = m_Callback
End Property

Friend Property Set Callback(ByVal Index As Long, ByRef RHS As SimpleServer)
    m_Index = Index
    Set m_Callback = RHS
End Property

Public Sub CloseSck(ByVal Index As Long)

End Sub

Public Sub Connect(ByVal Index As Long)

End Sub

Public Sub ConnectionRequest(ByVal Index As Long, ByVal requestID As Long, ByVal lRemotePort As Long, ByVal sRemoteHostIP As String)

End Sub

Public Sub DataArrival(ByVal Index As Long, ByVal bytesTotal As Long)

End Sub

Public Sub EncrDataArrival(ByVal Index As Long, ByVal bytesTotal As Long)

End Sub

Friend Property Let EncrFlg(ByVal NewValue As Boolean)
    m_bEncr = NewValue
End Property


Public Sub Error(ByVal Index As Long, ByVal Number As Long, Description As String, ByVal Source As String)

End Sub

Public Sub SendComplete(ByVal Index As Long)

End Sub

Public Sub SendProgress(ByVal Index As Long, ByVal bytesSent As Long, ByVal bytesRemaining As Long)

End Sub

Private Function GetBufferLenUDP() As Long
    'Returns winsock incoming buffer length from a UDP socket.
    Dim lRet    As Long
    Dim lBuff   As Long
    lRet = API_IoctlSocket(m_hSocket, FIONREAD, lBuff)
    If lRet = SOCKET_ERROR Then
        GetBufferLenUDP = 0
    Else
        GetBufferLenUDP = lBuff
    End If
End Function

Private Function GetLocalHostName() As String
    Const Routine As String = "SimpleServer.GetLocalHostName"
    Dim HostNameBuf As String * LOCAL_HOST_BUFF
    Dim lRet As Long
    Dim lErrorCode As Long
    lRet = API_GetHostName(HostNameBuf, LOCAL_HOST_BUFF)
    If lRet = SOCKET_ERROR Then
        GetLocalHostName = vbNullString
        lErrorCode = Err.LastDllError
        Call m_Callback.Error(m_Index, lErrorCode, GetErrorDescription(lErrorCode), Routine)
    Else
        GetLocalHostName = Left(HostNameBuf, InStr(1, HostNameBuf, Chr(0)) - 1)
    End If
End Function

Private Function GetLocalPort(ByVal hSocket As Long) As Long
    'Returns local port from a connected or bound socket.
    'Returns SOCKET_ERROR if fails.
    Dim newSa As sockaddr
    Dim Sin4 As sockaddr_in
    Dim Sin6 As sockaddr_in6
    Dim lRet As Long
    lRet = API_GetSockName(hSocket, newSa, LenB(newSa))
    If lRet = SOCKET_ERROR Then
        GetLocalPort = SOCKET_ERROR
    Else
        If newSa.sa_family = AF_INET6 Then
            CopyMemory Sin6, newSa, LenB(Sin6)  'Save to sockaddr_in6
            GetLocalPort = modSocket.IntegerToUnsigned(ntohs(Sin6.sin6_port))
        Else
            CopyMemory Sin4, newSa, LenB(Sin4)  'Save to sockaddr_in
            GetLocalPort = modSocket.IntegerToUnsigned(ntohs(Sin4.sin_port))
        End If
    End If
End Function

Private Function GetRemoteInfo(ByVal hSocket As Long, ByRef lRemotePort As Long, ByRef strRemoteHostIP As String, ByRef strRemoteHost As String) As Boolean
    'Retrieves remote info from a connected socket.
    'If succeeds returns TRUE and loads the arguments.
    'If fails returns FALSE and arguments are not loaded.
    Dim lRet As Long
    Dim tmpSa As sockaddr
    GetRemoteInfo = False
    lRet = API_GetPeerName(hSocket, tmpSa, LenB(tmpSa))
    If lRet = 0 Then
        GetRemoteInfo = True
        GetRemoteInfoFromSI tmpSa, m_lRemotePort, m_sRemoteHostIP, m_sRemoteHost
    Else
       lRemotePort = 0
       strRemoteHostIP = ""
       strRemoteHost = ""
    End If
End Function

Private Sub GetRemoteInfoFromSI(ByRef newSa As sockaddr, ByRef lRemotePort As Long, ByRef sRemoteHostIP As String, ByRef sRemoteHost As String)
    'Gets remote info from a sockaddr_in structure.
    Dim lRet As Long
    Dim Sin4 As sockaddr_in
    Dim Sin6 As sockaddr_in6
    Dim aLen As Long
    Dim bBuffer() As Byte
    If newSa.sa_family = AF_INET6 Then
        aLen = INET6_ADDRSTRLEN
        CopyMemory Sin6, newSa, LenB(Sin6)  'Save to sockaddr_in6
        lRemotePort = IntegerToUnsigned(ntohs(Sin6.sin6_port))
        ReDim bBuffer(0 To aLen - 1)    'Resize string buffer
        'Get IPv6 address as string
        lRet = inet_ntop(AF_INET6, Sin6.sin6_addr, bBuffer(0), aLen)
    Else 'Must be IPv4
        aLen = INET_ADDRSTRLEN
        CopyMemory Sin4, newSa, LenB(Sin4)  'Save to sockaddr_in
        lRemotePort = IntegerToUnsigned(ntohs(Sin4.sin_port))
        ReDim bBuffer(0 To aLen - 1)  'Resize string buffer
        'Get IPv4 address as string
        lRet = inet_ntoa(Sin4.sin_addr.s_addr)
    End If
    If lRet Then sRemoteHostIP = StringFromPointer(lRet)
    m_sRemoteHost = ""
End Sub

Private Function Initialize() As Boolean
    'socket's handle default value
    m_hSocket = SOCKET_ERROR
    'Set IPv Flag default value
    IPvFlg = 4 'Defaults to IPv4
    Initialize = modSocket.Initialize
End Function

Friend Property Get IPvFlg() As Long
    IPvFlg = m_IPv
End Property

Friend Property Let IPvFlg(ByVal IPvFlg As Long)
    m_IPv = IPvFlg
End Property

Friend Function Listen(LocalPort As Long) As Boolean
    Const Routine As String = "SimpleServer.Listen"
    Dim sa_local As sockaddr
    Dim lRet As Long
    Dim lErrorCode As Long
    Dim IPFamily As Long
    If m_IPv = 4 Then
        IPFamily = AF_INET
    Else
        IPFamily = AF_INET6
    End If
    'Create TCP socket
    m_Protocol = SCK_TCP
    m_State = sckconnecting
    m_hSocket = API_Socket(IPFamily, SOCK_STREAM, IPPROTO_TCP)
    If m_hSocket = SOCKET_ERROR Then
        m_State = sckError
        Call m_Callback.Error(0, m_hSocket, GetErrorDescription(m_hSocket), Routine)
        Exit Function
    Else
        ProcessOptions 'Get Send/Recv buffer lengths & register socket
        If Not modSocket.RegisterSocket(m_hSocket, ObjPtr(Me), True) Then Exit Function
    End If
    'Populate the local sockaddr
    sa_local.sa_family = IPFamily
    sa_local.sa_data(0) = PeekB(VarPtr(LocalPort) + 1)
    sa_local.sa_data(1) = PeekB(VarPtr(LocalPort))
    'Bind to socket
    lRet = API_Bind(m_hSocket, sa_local, LenB(sa_local))
    If lRet = SOCKET_ERROR Then
        m_State = sckError
        Call m_Callback.Error(0, lRet, GetErrorDescription(lRet), Routine)
        Exit Function
    Else 'Set into listening mode
        m_State = sckOpen
        lRet = API_Listen(m_hSocket, SOMAXCONN)
        If lRet = SOCKET_ERROR Then
            lErrorCode = Err.LastDllError
            Call m_Callback.Error(0, lErrorCode, GetErrorDescription(lErrorCode), Routine)
        Else
            m_State = scklistening
            m_lPortListen = LocalPort
            Listen = True
            Call PrintDebug("STATE: sckListening")
        End If
    End If
End Function

Friend Property Get LocalHostName() As String
    LocalHostName = GetLocalHostName
End Property

Friend Property Get LocalPort() As Long
    LocalPort = GetLocalPort(m_hSocket)
End Property

Friend Sub PeekData(ByRef data() As Byte, Optional maxLen As Long)
    Const Routine As String = "SimpleServer.PeekData"
    Dim lBytesRecv  As Long
    If m_Protocol = SCK_TCP Then
        If m_State <> sckConnected Then
            Call m_Callback.Error(m_Index, sckBadState, "Wrong protocol or connection state for the requested transaction or request", Routine)
            Exit Sub
        End If
    Else
        If m_State <> sckOpen Then
            Call m_Callback.Error(m_Index, sckBadState, "Wrong protocol or connection state for the requested transaction or request", Routine)
            Exit Sub
        End If
        If GetBufferLenUDP = 0 Then Exit Sub
    End If
    If maxLen < 1 Then
        If m_Protocol = SCK_TCP Then
            maxLen = GetbSize(m_bRecvBuffer)
        Else
            maxLen = GetBufferLenUDP
        End If
    End If
    lBytesRecv = RecvData(data, True, maxLen)
    Call PrintDebug("OK Bytes obtained from buffer: " & CStr(lBytesRecv))
End Sub

Private Sub PostSocket(ByVal lEventID As Long, ByVal lErrorCode As Long)
    'This procedure is called by the WindowProc callback function
    'from the modSocket module. The lEventID argument is an ID of the
    'network event that occurred for the socket. The lErrorCode
    'argument contains an error code only if an error was occurred
    'during an asynchronous execution.
    Const Routine As String = "SimpleServer.PostSocket"
    Dim newSa As sockaddr
    Dim Sin4 As sockaddr_in
    Dim Sin6 As sockaddr_in6
    Dim lRet As Long
    Dim lBytesRecvd As Long
    If lErrorCode <> 0 Then
        m_State = sckError
        Call PrintDebug("STATE: sckError")
        Call m_Callback.Error(m_Index, lErrorCode, GetErrorDescription(lErrorCode), "Windows Message")
        CloseSocket
        Exit Sub
    End If
    Select Case lEventID
    '======================================================================
    Case FD_CONNECT
        'Arrival of this message means that the connection initiated by the call
        'of the connect Winsock API function was successfully established.
        Call PrintDebug("FD_CONNECT " & CStr(m_hSocket))
        If m_State <> sckconnecting Then
            Call PrintDebug("WARNING: Omitting FD_CONNECT")
            Exit Sub
        End If
        'Get the connection local end-point parameters
        lRet = API_GetPeerName(m_hSocket, newSa, LenB(newSa))
        If lRet = 0 Then
            If newSa.sa_family = AF_INET6 Then
                ReDim bBuffer(0 To INET6_ADDRSTRLEN - 1)    'Resize string buffer
                CopyMemory Sin6, newSa, LenB(Sin6)  'Save to sockaddr_in6
                m_lRemotePort = modSocket.IntegerToUnsigned(ntohs(Sin6.sin6_port))
                lRet = inet_ntop(AF_INET6, Sin6.sin6_addr, bBuffer(0), INET6_ADDRSTRLEN)
            Else
                ReDim bBuffer(0 To INET_ADDRSTRLEN - 1)    'Resize string buffer
                CopyMemory Sin4, newSa, LenB(Sin4)  'Save to sockaddr_in
                m_lRemotePort = modSocket.IntegerToUnsigned(ntohs(Sin4.sin_port))
                lRet = inet_ntoa(Sin4.sin_addr.s_addr)
            End If
            If lRet Then m_sRemoteHostIP = StringFromPointer(lRet)
        End If
        m_State = sckConnected
        Call PrintDebug("STATE: sckConnected")
        Call m_Callback.Connect(0)
    '======================================================================
    Case FD_WRITE
        'This message means that the socket in a write-able
        'state, that is, buffer for outgoing data of the transport
        'service is empty and ready to receive data to send through
        'the network.
        Call PrintDebug("FD_WRITE " & CStr(m_hSocket))
        If m_State <> sckConnected Then
            Call PrintDebug("WARNING: Omitting FD_WRITE")
            Exit Sub
        End If
        If GetbSize(m_bSendBuffer) > 0 Then
            If m_Protocol = SCK_TCP Then
                TCPSend
            Else
                UDPSend
            End If
        End If
    '======================================================================
    Case FD_READ
        'Some data has arrived for this socket.
        Call PrintDebug("FD_READ " & CStr(m_hSocket))
        If m_Protocol = SCK_TCP Then
            If m_State <> sckConnected Then
                Call PrintDebug("WARNING: Omitting FD_READ")
                Exit Sub
            End If
            'Call the RecvDataToBuffer function that move arrived data
            'from the Winsock buffer to the local one and returns number
            'of bytes received.
            lBytesRecvd = RecvDataToBuffer
            If lBytesRecvd > 0 Then
                If m_bEncr Then
                    Call m_Callback.EncrDataArrival(m_Index, GetbSize(m_bRecvBuffer))
                Else
                    Call m_Callback.DataArrival(m_Index, GetbSize(m_bRecvBuffer))
                End If
            End If
        Else 'UDP protocol
            If m_State <> sckOpen Then
                Call PrintDebug("WARNING: Omitting FD_READ")
                Exit Sub
            End If
            'If we use UDP we don't remove data from winsock buffer.
            'We just let the user know the amount received so
            'he/she can decide what to do.
            lBytesRecvd = GetBufferLenUDP
            If lBytesRecvd > 0 Then
                If m_bEncr Then
                    Call m_Callback.EncrDataArrival(m_Index, lBytesRecvd)
                Else
                    Call m_Callback.DataArrival(m_Index, lBytesRecvd)
                End If
            End If
            'Now the buffer is emptied no matter what the user
            'dicided to do with the received data
            EmptyBuffer
        End If
    '======================================================================
    Case FD_ACCEPT
        'When the socket is in a listening state, arrival of this message
        'means that a connection request was received.
        Call PrintDebug("FD_ACCEPT " & CStr(m_hSocket))
        If m_State <> scklistening Then
            Call PrintDebug("WARNING: Omitting FD_ACCEPT")
            Exit Sub
        End If
        'Create a new socket for the requested connection.
        lRet = API_Accept(m_hSocket, newSa, LenB(newSa))
        If lRet = SOCKET_ERROR Then
            lErrorCode = Err.LastDllError
            Call m_Callback.Error(m_Index, lErrorCode, GetErrorDescription(lErrorCode), Routine)
        Else
            Call GetRemoteInfo(lRet, m_lRemotePort, m_sRemoteHostIP, m_sRemoteHost)
            Call PrintDebug("OK Accept socket? " & CStr(lRet))
            Call m_Callback.ConnectionRequest(0, lRet, m_lRemotePort, m_sRemoteHostIP)
        End If
    '======================================================================
    Case FD_CLOSE
        'This message means that the remote host is closing the conection
        Call PrintDebug("FD_CLOSE " & CStr(m_hSocket))
        If m_State <> sckConnected Then
            Call PrintDebug("WARNING: Omitting FD_CLOSE")
            Exit Sub
        End If
        m_State = sckClosing
        Call PrintDebug("REMOTE: sckClosing")
        Call m_Callback.CloseSck(m_Index)
    End Select
End Sub

Private Sub ProcessOptions()
    'Retrieves some socket options.
    'If it is an UDP socket also sets SO_BROADCAST option.
    Const Routine As String = "SimpleServer.ProcessOptions"
    Dim lRet As Long
    Dim lBuffer As Long
    Dim lErrorCode As Long
    If m_Protocol = SCK_TCP Then
        lRet = API_GetSockOpt(m_hSocket, SOL_SOCKET, SO_RCVBUF, lBuffer, LenB(lBuffer))
        If lRet = SOCKET_ERROR Then
            lErrorCode = Err.LastDllError
            Call m_Callback.Error(m_Index, lErrorCode, GetErrorDescription(lErrorCode), Routine)
        Else
            m_lRecvBufferLen = lBuffer
        End If
        lRet = API_GetSockOpt(m_hSocket, SOL_SOCKET, SO_SNDBUF, lBuffer, LenB(lBuffer))
        If lRet = SOCKET_ERROR Then
            lErrorCode = Err.LastDllError
            Call m_Callback.Error(m_Index, lErrorCode, GetErrorDescription(lErrorCode), Routine)
        Else
            m_lSendBufferLen = lBuffer
        End If
    Else
        lBuffer = 1
        lRet = API_SetSockOpt(m_hSocket, SOL_SOCKET, SO_BROADCAST, lBuffer, LenB(lBuffer))
        lRet = API_GetSockOpt(m_hSocket, SOL_SOCKET, SO_MAX_MSG_SIZE, lBuffer, LenB(lBuffer))
        If lRet = SOCKET_ERROR Then
            lErrorCode = Err.LastDllError
            Call m_Callback.Error(m_Index, lErrorCode, GetErrorDescription(lErrorCode), Routine)
        Else
            m_lRecvBufferLen = lBuffer
            m_lSendBufferLen = lBuffer
        End If
    End If
    Call PrintDebug("Winsock buffer size for sends: " & CStr(m_lRecvBufferLen))
    Call PrintDebug("Winsock buffer size for receives: " & CStr(m_lSendBufferLen))
End Sub

Friend Property Get Protocol() As Long
    Protocol = m_Protocol
End Property

Friend Sub RecoverData(Optional maxLen As Long)
    Const Routine As String = "SimpleServer.RecoverData"
    Dim lBytesRecv  As Long
    If m_Protocol = SCK_TCP Then
        If m_State <> sckConnected And Not m_bAcceptClass Then
            Call m_Callback.Error(m_Index, sckBadState, "Wrong protocol or connection state for the requested transaction or request", Routine)
            Exit Sub
        End If
    Else
        If m_State <> sckOpen Then
            Call m_Callback.Error(m_Index, sckBadState, "Wrong protocol or connection state for the requested transaction or request", Routine)
            Exit Sub
        End If
        If GetBufferLenUDP = 0 Then Exit Sub
    End If
    If maxLen < 1 Then 'MaxLen not provided
        If m_Protocol = SCK_TCP Then
            maxLen = GetbSize(m_bRecvBuffer)
        Else
            maxLen = GetBufferLenUDP
        End If
    End If
    lBytesRecv = RecvData(m_bData, False, maxLen)
    Call PrintDebug("OK Bytes obtained from buffer: " & CStr(lBytesRecv))
 End Sub

Private Function RecvData(ByRef bData() As Byte, ByVal blnPeek As Boolean, Optional maxLen As Long) As Long
    Const Routine As String = "SimpleServer.RecvData"
    'Using TCP, the data is retrieved from a local buffer (m_bRecvBuffer).
    'Using UDP, the data is retrieved from the winsock buffer.
    'If blnPeek is TRUE, the function returns number of bytes in the
    'buffer, and copies the buffer data into the data argument.
    'If blnPeek is FALSE, then this function returns number of bytes in the
    'buffer and moves the buffer data into the data argument.
    Dim lBuffLen        As Long
    Dim lErrorCode      As Long
    Dim bTmp()          As Byte
    If m_Protocol = SCK_TCP Then
        lBuffLen = GetbSize(m_bRecvBuffer)
    Else
        lBuffLen = GetBufferLenUDP
    End If
    If maxLen = 0 Then 'if maxLen argument is missing
        RecvData = lBuffLen
        If RecvData = 0 Then
            bData = bTmp
            Exit Function
        Else
            bData = BuildArray(lBuffLen, blnPeek, lErrorCode)
        End If
    Else 'if maxLen argument is not missing
        If lBuffLen = 0 Then
            RecvData = 0
            bData = bTmp
            If m_Protocol = SCK_UDP Then
                EmptyBuffer
                Call m_Callback.Error(m_Index, WSAEMSGSIZE, GetErrorDescription(WSAEMSGSIZE), Routine)
            End If
            Exit Function
        ElseIf maxLen > lBuffLen Then
            RecvData = lBuffLen
            bData = BuildArray(lBuffLen, blnPeek, lErrorCode)
        Else
            RecvData = maxLen
            bData = BuildArray(maxLen, blnPeek, lErrorCode)
        End If
    End If
    'if BuildArray returns an error is handled here
    If lErrorCode <> 0 Then
        Call m_Callback.Error(m_Index, lErrorCode, GetErrorDescription(lErrorCode), Routine)
    End If
End Function

Private Function RecvDataToBuffer() As Long
    'This function retrieves data from the Winsock buffer
    'into the class local buffer. The function returns number
    'of bytes retrieved (received).
    Const Routine As String = "SimpleServer.RecvDataToBuffer"
    Dim bInArray()     As Byte
    Dim lBytesRecv      As Long
    Dim lErrorCode      As Long
    ReDim bInArray(m_lRecvBufferLen - 1)
    lBytesRecv = API_Recv(m_hSocket, bInArray(0), m_lRecvBufferLen, 0&)
    If lBytesRecv = SOCKET_ERROR Then
        m_State = sckError
        Call PrintDebug("STATE: sckError")
        lErrorCode = Err.LastDllError
        Call m_Callback.Error(m_Index, lErrorCode, GetErrorDescription(lErrorCode), Routine)
    ElseIf lBytesRecv > 0 Then
        Call AddByte(m_bRecvBuffer, bInArray, lBytesRecv)
        RecvDataToBuffer = lBytesRecv
    End If
End Function

Friend Property Get RemoteHostIP() As String
    RemoteHostIP = m_sRemoteHostIP
End Property

Friend Sub TCPSend()
    Const Routine As String = "SimpleServer.TCPSend"
    Dim bData()         As Byte
    Dim lBufLen         As Long
    Dim lSendLen        As Long
    Dim lRet            As Long
    Dim lTotalSent      As Long
    Dim lErrorCode      As Long
    Dim lTemp           As Long
    Dim bTmp() As Byte
    If m_State <> sckConnected Then
        Call m_Callback.Error(m_Index, sckBadState, "Wrong connection state for TCP Send!", Routine)
        Exit Sub
    End If
    lBufLen = GetbSize(m_bSendBuffer)
    Do Until lRet = SOCKET_ERROR Or lBufLen = 0
        If GetbSize(m_bSendBuffer) = 0 Then
            Exit Sub
        ElseIf lBufLen > m_lSendBufferLen Then
            lSendLen = m_lSendBufferLen
            ReDim bData(lSendLen - 1)
            CopyMemory bData(0), m_bSendBuffer(0), lSendLen
        Else
            lSendLen = lBufLen
            bData = m_bSendBuffer
        End If
        lRet = API_Send(m_hSocket, bData(0), lSendLen, 0&)
        If lRet = SOCKET_ERROR Then
            lErrorCode = Err.LastDllError
            If lErrorCode = WSAEWOULDBLOCK Then
                Call PrintDebug("WARNING: Send buffer full, waiting...")
                If lTotalSent > 0 Then Call m_Callback.SendProgress(m_Index, lTotalSent, lBufLen)
            Else
                m_State = sckError
                Call PrintDebug("STATE: sckError")
                Call m_Callback.Error(m_Index, lErrorCode, GetErrorDescription(lErrorCode), Routine)
            End If
        Else
            Call PrintDebug("OK Bytes sent: " & CStr(lRet))
            lTotalSent = lTotalSent + lRet
            If lBufLen > lRet Then
                Call DeleteByte(m_bSendBuffer, lRet)
            Else
                Call PrintDebug("OK Finished SENDING")
                m_bSendBuffer = bTmp
                lTemp = lTotalSent
                lTotalSent = 0
                Call m_Callback.SendProgress(m_Index, lTemp, 0)
                Call m_Callback.SendComplete(m_Index)
            End If
        End If
        lBufLen = GetbSize(m_bSendBuffer)
    Loop
End Sub

Friend Sub UDPSend()
    Const Routine As String = "SimpleServer.UDPSend"
    Dim bData()     As Byte
    Dim lBufLen     As Long
    Dim lRet        As Long
    Dim lErrorCode  As Long
    Dim lTotalSent  As Long
    Dim lTemp       As Long
    Dim bTmp()      As Byte
    If m_State <> sckOpen Then
        Call m_Callback.Error(m_Index, sckBadState, "Wrong socket state for UDP Send!", Routine)
        Exit Sub
    End If
    lBufLen = GetbSize(m_bSendBuffer)
    bData = m_bSendBuffer
    m_bSendBuffer = bTmp
    lRet = API_SendTo(m_hSocket, bData(0), lBufLen, 0&, sa_UDPdest, LenB(sa_UDPdest))
    If lRet = SOCKET_ERROR Then
        lErrorCode = Err.LastDllError
        m_State = sckError
        Call PrintDebug("STATE: sckError")
        Call m_Callback.Error(m_Index, lErrorCode, GetErrorDescription(lErrorCode), Routine)
    Else
        Call PrintDebug("OK Bytes sent: " & CStr(lRet))
        lTotalSent = lTotalSent + lRet
        If lBufLen > lRet Then
            Call DeleteByte(m_bSendBuffer, lRet)
        Else
            Call PrintDebug("OK Finished SENDING")
            m_bSendBuffer = bTmp
            lTemp = lTotalSent
            lTotalSent = 0
            Call m_Callback.SendProgress(m_Index, lTemp, 0)
            Call m_Callback.SendComplete(m_Index)
        End If
    End If
End Sub

Friend Property Get uInBuffer() As String
    uInBuffer = ByteToUni(m_bData)
End Property

Friend Property Let uOutBuffer(sNewValue As String)
    Dim bTmp() As Byte
    bTmp = sNewValue
    Call AddByte(m_bSendBuffer, bTmp)
End Property


Friend Property Get sInBuffer() As String
    sInBuffer = ByteToStr(m_bData)
End Property

Friend Property Let sOutBuffer(sNewValue As String)
    Dim bTmp() As Byte
    bTmp = StrToByte(sNewValue)
    Call AddByte(m_bSendBuffer, bTmp)
End Property

Friend Property Get State() As Long
    State = m_State
End Property

Friend Function TCPConnect(RemoteHost As String, RemotePort As Long, Optional LocalPort As Long) As Boolean
    Const Routine As String = "SimpleServer.TCPConnect"
    Dim lErrorCode As Long
    Dim sa_local As sockaddr
    Dim sa_dest As sockaddr
    Dim lRet As Long
    Dim ptrResult As Long
    Dim Hints As addrinfo
    Dim IPFamily As Long
    If m_IPv = 4 Then
        IPFamily = AF_INET
    Else
        IPFamily = AF_INET6
    End If
    'Create TCP socket
    m_Protocol = SCK_TCP
    m_State = sckconnecting
    m_hSocket = API_Socket(IPFamily, SOCK_STREAM, IPPROTO_TCP)
    If m_hSocket = SOCKET_ERROR Then
        m_State = sckError
        Call m_Callback.Error(m_Index, m_hSocket, GetErrorDescription(m_hSocket), Routine)
        Exit Function
    Else
        'Get Send/Recv buffer lengths & register socket
        ProcessOptions
        If Not modSocket.RegisterSocket(m_hSocket, ObjPtr(Me), True) Then Exit Function
    End If
    'Recover info about the destination.
    Hints.ai_family = IPFamily
    lRet = API_GetAddrInfo(RemoteHost, RemotePort, VarPtr(Hints), ptrResult)
    If lRet <> 0 Then
        m_State = sckError
        Call m_Callback.Error(m_Index, lRet, GetErrorDescription(lRet), Routine)
        Exit Function
    End If
    Hints.ai_next = ptrResult   'Point to first structure in linked list
    CopyMemory Hints, ByVal Hints.ai_next, LenB(Hints) 'Copy next address info to Hints
    CopyMemory sa_dest, ByVal Hints.ai_addr, LenB(sa_dest) 'Save sockaddr portion
    API_FreeAddrInfo (ptrResult) ' free the linked list
    If LocalPort > 0 Then
        'Populate the local sockaddr
        sa_local.sa_family = IPFamily
        sa_local.sa_data(0) = PeekB(VarPtr(LocalPort) + 1)
        sa_local.sa_data(1) = PeekB(VarPtr(LocalPort))
        lRet = API_SetSockOpt(m_hSocket, SOL_SOCKET, SO_LINGER, 1&, 4&)
        If lRet = SOCKET_ERROR Then
            m_State = sckError
            Call m_Callback.Error(m_Index, lRet, GetErrorDescription(lRet), Routine)
            Exit Function
        End If
        lRet = API_SetSockOpt(m_hSocket, SOL_SOCKET, SO_REUSEADDR, 1&, 4&)
        If lRet = SOCKET_ERROR Then
            m_State = sckError
            Call m_Callback.Error(m_Index, lRet, GetErrorDescription(lRet), Routine)
            Exit Function
        End If
        lRet = API_Bind(m_hSocket, sa_local, LenB(sa_local))
        If lRet = SOCKET_ERROR Then
            m_State = sckError
            Call m_Callback.Error(m_Index, lRet, GetErrorDescription(lRet), Routine)
            Exit Function
        End If
    End If
    'Connect to sockaddr
    m_State = sckconnecting
    lRet = API_Connect(m_hSocket, sa_dest, LenB(sa_dest))
    'Check and handle errors
    If lRet = SOCKET_ERROR Then
        lErrorCode = Err.LastDllError
        If lErrorCode <> WSAEWOULDBLOCK Then
            m_State = sckError
            Call m_Callback.Error(m_Index, lErrorCode, GetErrorDescription(lErrorCode), Routine)
            Exit Function
        End If
    End If
    TCPConnect = True
End Function

Friend Function UDPInit(RemoteHost As String, RemotePort As Long, LocalPort As Long) As Boolean
    Const Routine As String = "SimpleServer.UDPInit"
    'Using UDP, we just bind the socket to a known port
    Dim lErrorCode As Long
    Dim sa_local As sockaddr
    Dim lRet As Long
    Dim ptrResult As Long
    Dim Hints As addrinfo
    Dim IPFamily As Long
    Dim SocketExists As Boolean
    If m_hSocket > 1 Then SocketExists = True
    If m_IPv = 4 Then
        IPFamily = AF_INET
    Else
        IPFamily = AF_INET6
    End If
    m_Protocol = SCK_UDP
    'If UDP socket does not exist, create and register it
    Debug.Print "UDP Socket " & CStr(m_hSocket)
    If Not SocketExists Then
        m_hSocket = API_Socket(IPFamily, SOCK_DGRAM, IPPROTO_UDP)
        If m_hSocket = SOCKET_ERROR Then
            m_State = sckError
            Call m_Callback.Error(m_Index, m_hSocket, GetErrorDescription(m_hSocket), Routine)
            Exit Function
        Else
            'Get Send/Recv buffer lengths & register socket
            ProcessOptions
            If Not modSocket.RegisterSocket(m_hSocket, ObjPtr(Me), True) Then Exit Function
        End If
    End If
    'Recover info about the destination.
    Hints.ai_family = IPFamily
    lRet = API_GetAddrInfo(RemoteHost, RemotePort, VarPtr(Hints), ptrResult)
    If lRet <> 0 Then
        m_State = sckError
        Call m_Callback.Error(m_Index, lRet, GetErrorDescription(lRet), Routine)
        Exit Function
    End If
    Hints.ai_next = ptrResult   'Point to first structure in linked list
    CopyMemory Hints, ByVal Hints.ai_next, LenB(Hints) 'Copy next address info to Hints
    CopyMemory sa_UDPdest, ByVal Hints.ai_addr, LenB(sa_UDPdest)    'Save sockaddr portion
    API_FreeAddrInfo (ptrResult) ' free the linked list
    If SocketExists Then
        UDPInit = True
        Exit Function
    End If
    'Populate the local sockaddr
    sa_local.sa_family = IPFamily
    sa_local.sa_data(0) = PeekB(VarPtr(LocalPort) + 1)
    sa_local.sa_data(1) = PeekB(VarPtr(LocalPort))
    'Bind to socket
    lRet = API_Bind(m_hSocket, sa_local, LenB(sa_local))
    If lRet = SOCKET_ERROR Then
        m_State = sckError
        Call m_Callback.Error(m_Index, lRet, GetErrorDescription(lRet), Routine)
    Else
        m_State = sckOpen
        UDPInit = True
    End If
End Function

Private Sub Class_Initialize()
    Initialize
End Sub


Private Sub Class_Terminate()
    'clean processes and finish winsock service
    Call modSocket.Finalize
End Sub


