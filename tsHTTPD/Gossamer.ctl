VERSION 5.00
Begin VB.UserControl Gossamer 
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   PaletteMode     =   4  'None
   ScaleHeight     =   450
   ScaleWidth      =   450
   ToolboxBitmap   =   "Gossamer.ctx":0000
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   400
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   370
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillStyle       =   7  'Diagonal Cross
      Height          =   400
      Left            =   0
      Top             =   0
      Width           =   400
   End
End
Attribute VB_Name = "Gossamer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'Gossamer
'========
'
'A tiny HTTP server control.
'
'This control uses SimplServer as a listener on Index = 0 for incoming HTTP
'client connections which are handed off to higher index values.
'
'Changes
'-------
'
'Version: 1.1
'
' o Added Port property.
' o Changed StartListening so Port parameter is optional, using
'   current value of the Port property setting if not supplied.
' o Added procedure attributes, set LogEvent as default event.
' o Marked several Public members "hidden" since they're only meant
'   for use by GossClients.
' o Methods: LogEvent is now RaiseLogEvent, DynamicRequest is now
'   RaiseDynamicRequest.
' o New handling of VDir, property can now only be set while not
'   listening.
'
'Version: 1.2
'
' o Split out function UTCDateTime from UTCString.
' o Added function UTCParseString to convert HTTP timestamps to
'   Date values.
' o VDirPath R/O property is no longer hidden, since it can be useful
'   in handling dynamic requests.
'
'Version: 1.3
'
' no change
'
'Version: 1.4
'
' o Added new (hidden) property ServerHeader for use by GossClient.
' o Added EntityEncode method.
'
'Version: 1.5
'
' o Corrected ResolvePath() so it no longer allocates a buffer 1 char too
'   long nor needs to then truncate prior to return.
'
'Version 1.6
'
' o Treat sckWouldBlock as a getWSSoftError: log it as such and then
'   ignore it.
' o Stop setting .Timestamp since it gets set when GE instances are
'   created.
'
'Version: 1.7
'
' o Added GOSS_VERSION_MAJOR and GOSS_VERSION_MINOR Consts for use in the
'   default ServerHeader value.
' o Removed sending an extraneous space after header ":" separators.  Not
'   important, and quite commonly done but a silly waste of bandwidth.
'
Implements SimpleServer
Private mServer() As New SimpleServer
Private IPVersion As String
Private RemotePort As Long
Private RemoteHostIP As String
Private lIndex As Long
Private LastConnect As String
Private mBlnConnClose As Boolean      'Received a "Connection: close" header.
Private mBlnInUse As Boolean          'In-use status of this GossClient.
Private mBytBuffer() As Byte          'Buffer for reading static resource requested.
Private mBytResponse() As Byte        'Response data from dynamic request.
Private mLngResponseLen As Long       'Valid bytes in mBytResponse.
Private mReqState As RequestStates    'Where we are receiving a request.
Private mStrBuffer As String          'Buffered incoming request text.
Private mStrContent As String         'Request Content body.
Private mStrHTTPVersion As String     'Version of request.
Private mStrReqLine As String         'HTTP request line.
Private mColReqHeaders As Collection  'Request headers Collection.
Private mLngContentLen                'Request content length from header.
Private mIntFile As Integer           'Native I/O file number of static resource requested.
                                      'If non-0 a file is open.
Private mIntFullBlocks As Integer     'Remaining count of full blocks to send in static resource.
Private mIntIndex As Integer          'Index in control array of this Client.
Private mLngLastBlockSize As Long     'Size in bytes of final block in static resource.
Private mStrRespBuffer As String      'We want to buffer responses so we can handle
                                      '  ReqCloseConn = True properly.
Private mLngRespUsed As Long

Private Const GOSS_VERSION_MAJOR As String = "1"
Private Const GOSS_VERSION_MINOR As String = "7"
Private Const DOUBLE_CRLF As String = vbCrLf & vbCrLf
'Private Const STATIC_BUFFER_SZ As Long = 8192
Private mBufferSize As Long

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private Declare Function GetFullPathName Lib "kernel32" _
    Alias "GetFullPathNameW" ( _
    ByVal lpFileName As Long, _
    ByVal nBufferLength As Long, _
    ByVal lpBuffer As Long, _
    ByVal lpFilePart As Long) As Long

Private Declare Function GetTimeZoneInformation Lib "kernel32" ( _
    lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2

Private mMaxConnections As Integer
Private mVDir As String
Private mVDirPath As String
Private mServerHeader As String

Public DefaultPage As String 'Default page for directory requests.
Attribute DefaultPage.VB_VarDescription = "Simple file name of Gossamer's default page to return on directory requests"
Public Port As Long          'Server listen port.
Attribute Port.VB_VarDescription = "HTTP port to listen on"

Public Event DynamicRequest(ByVal Method As String, _
                            ByVal URI As String, _
                            ByVal Params As String, _
                            ByVal ReqHeaders As Collection, _
                            ByRef RespStatus As Single, _
                            ByRef RespStatusText As String, _
                            ByRef RespMIME As String, _
                            ByRef RespExtraHeaders As String, _
                            ByRef RespBody() As Byte, _
                            ByVal ClientIndex As Integer)
                            
Public Event LogEvent(ByVal GossEvent As GossEvent, _
                      ByVal ClientIndex As Integer) '-1 for Gossamer events.
Attribute LogEvent.VB_Description = "Raised when a loggable event has occurred"
Attribute LogEvent.VB_MemberFlags = "200"

Private Enum RequestStates
    rqsIdle = 0
    rqsHeadersComplete
    rqsReqLineComplete
    rqsRequestComplete
End Enum

Private Sub AppendResp(ByVal Text As String)
    Dim Length As Long
    
    Length = Len(Text)
    If mLngRespUsed + Length > Len(mStrRespBuffer) Then
        If Len(mStrRespBuffer) < 1 Then
            mStrRespBuffer = Space$(Length + 200)
        Else
            mStrRespBuffer = mStrRespBuffer & Space$(Length + 100)
        End If
    End If
    Mid$(mStrRespBuffer, mLngRespUsed + 1, Length) = Text
    mLngRespUsed = mLngRespUsed + Length
End Sub

Public Property Let BufferSize(ByVal NewValue As Long)
    mBufferSize = NewValue
End Property

Public Property Get MaxConnections() As Long
Attribute MaxConnections.VB_Description = "Maximum number of active client connections to accept"
    MaxConnections = mMaxConnections
End Property

Public Property Let MaxConnections(ByVal Max As Long)
    If 0 < Max Or Max <= 1000 Then
        mMaxConnections = Max
    Else
        Err.Raise &H8004A702, "Gossamer", "MaxConnections must be 1 to 1000"
    End If
End Property

Private Sub ProcessRequest()
    Dim blnReturnFile As Boolean
    Dim dtSince As Date
    Dim GE As GossEvent
    Dim lngFileBytes As Long
    Dim lngInStr As Long
    Dim sngStatus As Single
    Dim strFile As String
    Dim strKeepAlive As String
    Dim strMIME As String
    Dim strParts() As String
    Dim strRespExtraHeaders As String
    Dim strStatusText As String

    strParts = Split(mStrReqLine, " ", 3)
    strParts(0) = UCase$(strParts(0))
    Set GE = New GossEvent
    With GE
        .EventType = getHTTP
        .IP = RemoteHostIP
        .Port = RemotePort
    End With
    If UBound(strParts) <> 2 Then
        With GE
            .EventSubtype = gesHTTPError
            .Method = "ERROR"
            .Text = "Bad Request Line: " & mStrReqLine
        End With
        RaiseLogEvent GE, mIntIndex
        
        mReqState = rqsIdle
        SimpleServer_CloseSck (mIntIndex)
        mStrBuffer = ""
        mBlnInUse = False
    Else
        'Process "Connection: close" headers.
        On Error Resume Next
        strKeepAlive = mColReqHeaders("CONNECTION")(1)
        If Err.Number = 0 Then
            On Error GoTo 0
            If UCase$(strKeepAlive) = "CLOSE" Then mBlnConnClose = True
        End If
        
        strParts(2) = UCase$(strParts(2))
        mStrHTTPVersion = strParts(2)
        With GE
            .Method = strParts(0)
            .Text = strParts(1)
            .HTTPVersion = strParts(2)
        End With
        strParts(1) = Replace$(strParts(1), "/", "\")
        Select Case strParts(0)
            Case "GET", "HEAD"
                lngInStr = InStr(strParts(1), "?")
                If lngInStr > 0 Then
                    'Request parameters present, assume dynamic content request.
                    If strParts(0) = "HEAD" Then
                        SendCanned 501
                    Else
                        'GET dynamic content.
                        GE.EventSubtype = gesGETDynamic
                        RaiseDynamicRequest strParts(0), _
                                                   Left$(strParts(1), lngInStr - 1), _
                                                   Mid$(strParts(1), lngInStr + 1), _
                                                   mColReqHeaders, _
                                                   sngStatus, _
                                                   strStatusText, _
                                                   strMIME, _
                                                   strRespExtraHeaders, _
                                                   mBytResponse, _
                                                   mIntIndex
                        If sngStatus = 0 Then
                            SendCanned 501
                        Else
                            SendResponse sngStatus, _
                                         strStatusText, _
                                         strMIME, _
                                         strRespExtraHeaders
                        End If
                    End If
                Else
                    'Request for static resource.
                    GE.EventSubtype = gesGETStatic
                    strFile = strParts(1)
                    If Right$(strFile, 1) = "\" Then strFile = strFile & DefaultPage
                    strFile = ResolvePath(VDirPath & strFile)
                    If Left$(strFile, Len(VDirPath)) <> VDirPath Then
                        'Bad request, trying to snoop outside VDir?
                        SendCanned 403
                    Else
                        'Locate file.
                        On Error Resume Next
                        GetAttr strFile
                        If Err.Number Then
                            'No such file.
                            On Error GoTo 0
                            SendCanned 404
                        Else
                            'Found file.
                            On Error GoTo 0
                            If strParts(0) = "HEAD" Then
                                'Return only HEADers.
                                SendStaticHeader FileLen(strFile), strFile
                            Else
                                'GET of static content.
                                On Error Resume Next
                                dtSince = _
                                    UTCParseString(mColReqHeaders("IF-MODIFIED-SINCE")(1))
                                If Err.Number Then
                                    On Error GoTo 0
                                    blnReturnFile = True
                                Else
                                    On Error GoTo 0
                                    blnReturnFile = _
                                        dtSince < UTCDateTime(FileDateTime(strFile))
                                End If
                                
                                If blnReturnFile Then
                                    On Error Resume Next
                                    mIntFile = GetFreeFile()
                                    If Err.Number Then
                                        On Error GoTo 0
                                        RaiseLogEvent GE, mIntIndex
                                        GE.EventSubtype = gesServerError
                                        GE.Text = "Ran out of file numbers"
                                        SendCanned 500.13
                                    Else
                                        'Open file, format and send headers and prime for
                                        'transmission of content.
                                        On Error GoTo 0
                                        Open strFile For Binary Access Read As #mIntFile
'                                        ReDim mBytBuffer(STATIC_BUFFER_SZ - 1)
                                        ReDim mBytBuffer(mBufferSize - 1)
                                       lngFileBytes = LOF(mIntFile)
'                                        mIntFullBlocks = lngFileBytes \ STATIC_BUFFER_SZ
                                        mIntFullBlocks = lngFileBytes \ mBufferSize
'                                        mIntLastBlockSize = lngFileBytes Mod STATIC_BUFFER_SZ
                                        mLngLastBlockSize = lngFileBytes Mod mBufferSize
                                        SendStaticHeader lngFileBytes, strFile
                                        'File content will be sent via SendComplete handler.
                                    End If
                                Else
                                    SendCanned 304
                                End If
                            End If
                        End If
                    End If
                End If
            
            Case "POST"
                GE.EventSubtype = gesPOST
                RaiseDynamicRequest strParts(0), _
                                           strParts(1), _
                                           mStrContent, _
                                           mColReqHeaders, _
                                           sngStatus, _
                                           strStatusText, _
                                           strMIME, _
                                           strRespExtraHeaders, _
                                           mBytResponse, _
                                           mIntIndex
                If sngStatus = 0 Then
                    SendCanned 501
                Else
                    SendResponse sngStatus, _
                                 strStatusText, _
                                 strMIME, _
                                 strRespExtraHeaders
                End If
            
            Case Else
                GE.EventSubtype = gesUnknown
                SimpleServer_CloseSck (mIntIndex)
        End Select

        mStrContent = ""
        RaiseLogEvent GE, mIntIndex
        mReqState = rqsIdle
        
        Set mColReqHeaders = Nothing
    End If
End Sub

Public Property Get ServerHeader() As String
Attribute ServerHeader.VB_Description = "Returns versioned server header string for responses"
Attribute ServerHeader.VB_MemberFlags = "40"
    ServerHeader = mServerHeader
End Property

Public Property Get State() As Integer
Attribute State.VB_Description = "Returns State of the listener Winsock control"
    State = State
End Property

Public Property Get VDir() As String
Attribute VDir.VB_Description = "Directory containing static resources of Gossamer site"
    VDir = mVDir
End Property

Public Property Let VDir(ByVal Directory As String)
'    If wskRequest.State <> sckListening Then
        mVDir = Directory
        If Len(mVDir) = 0 Then
            mVDirPath = CurDir$()
        Else
            mVDirPath = ResolvePath(mVDir)
        End If
'    Else
'        Err.Raise &H8004A704, "Gossamer", "Can't change VDir while listening"
'    End If
End Property

Public Property Get VDirPath() As String
Attribute VDirPath.VB_Description = "Returns fully qualified path for current VDir setting"
    VDirPath = mVDirPath
End Property

Public Function EntityEncode(ByVal Text As String) As String
Attribute EntityEncode.VB_Description = "Encode a text string to be inserted into HTML as text with HTML entity encoding"
    EntityEncode = Join$(Split(Text, "&"), "&amp;")
    EntityEncode = Join$(Split(EntityEncode, """"), "&quot;")
    EntityEncode = Join$(Split(EntityEncode, "<"), "&lt;")
    EntityEncode = Join$(Split(EntityEncode, ">"), "&gt;")
End Function

Public Function ExtensionToMIME(ByVal Extension As String) As String
Attribute ExtensionToMIME.VB_Description = "Return MIME type corresponding to the supplied file extension value (without .)"
    Extension = UCase$(Extension)
    Select Case Extension
        Case "CSS"
            ExtensionToMIME = "text/css"
        Case "GIF"
            ExtensionToMIME = "image/gif"
        Case "HTM", "HTML"
            ExtensionToMIME = "text/html"
        Case "ICO"
            ExtensionToMIME = "image/vnd.microsoft.icon"
        Case "JPG", "JPEG"
            ExtensionToMIME = "image/jpeg"
        Case "JS", "JSE"
            ExtensionToMIME = "application/javascript"
        Case "PNG"
            ExtensionToMIME = "image/png"
        Case "RTF"
            ExtensionToMIME = "application/rtf"
        Case "TIF", "TIFF"
            ExtensionToMIME = "image/tiff"
        Case "TXT"
            ExtensionToMIME = "text/plain"
        Case "VBS", "VBE"
            ExtensionToMIME = "application/vbscript"
        Case "XML", "XSD"
            ExtensionToMIME = "text/xml"
        Case "ZIP"
            ExtensionToMIME = "application/zip"
        Case Else
            ExtensionToMIME = "application/octet-stream"
    End Select
End Function

Private Function ExtractResp() As String
    ExtractResp = Left$(mStrRespBuffer, mLngRespUsed)
    mLngRespUsed = 0
End Function

Public Function GetFreeFile() As Integer
Attribute GetFreeFile.VB_Description = "Calls FreeFile(0) to get file number, if exhausted tries FreeFile(1)"
    On Error Resume Next
    GetFreeFile = FreeFile(0)
    If Err.Number Then
        On Error GoTo 0
        GetFreeFile = FreeFile(1)
    End If
End Function

Private Function GetNextSocket() As Long
    Dim lNum As Long
    For lNum = 0 To MaxConnections
        If mServer(lNum).State = 0 Then Exit For 'Closed socket found
    Next lNum
    GetNextSocket = lNum
End Function

Public Sub RaiseDynamicRequest(ByVal Method As String, _
                               ByVal URI As String, _
                               ByVal Params As String, _
                               ByVal ReqHeaders As Collection, _
                               ByRef RespStatus As Single, _
                               ByRef RespStatusText As String, _
                               ByRef RespMIME As String, _
                               ByRef RespExtraHeaders As String, _
                               ByRef RespBody() As Byte, _
                               ByVal Index As Integer)
    RaiseEvent DynamicRequest(Method, _
                              URI, _
                              Params, _
                              ReqHeaders, _
                              RespStatus, _
                              RespStatusText, _
                              RespMIME, _
                              RespExtraHeaders, _
                              RespBody, _
                              Index)
End Sub

Public Sub RaiseLogEvent(ByVal GossEvent As GossEvent, ByVal Index As Integer)
Attribute RaiseLogEvent.VB_Description = "Only for use by GossClient"
Attribute RaiseLogEvent.VB_MemberFlags = "40"
    RaiseEvent LogEvent(GossEvent, Index)
End Sub

Public Function ResolvePath(ByVal RelativePath As String) As String
Attribute ResolvePath.VB_Description = "Only for use by GossClient"
Attribute ResolvePath.VB_MemberFlags = "40"
    'Returns full path to RelativePath, "" if any error.
    Dim strFullPath As String
    Dim lngLen As Long
    Dim lngFilePart As Long
    
    lngLen = GetFullPathName(StrPtr(RelativePath), 0, StrPtr(strFullPath), lngFilePart)
    If lngLen Then
        'If the lpBuffer buffer is too small to contain the path, the return value
        'is the size, in TCHARs, of the buffer that is required to hold the path
        'and the terminating null character.
        strFullPath = String$(lngLen - 1, 0)
        lngLen = GetFullPathName(StrPtr(RelativePath), lngLen, StrPtr(strFullPath), lngFilePart)
        If lngLen Then
            'If the function succeeds, the return value is the length, in TCHARs,
            'of the string copied to lpBuffer, not including the terminating null
            'character.
            ResolvePath = strFullPath
        End If
    End If
End Function

Private Sub SendCanned(ByVal Status As Single)
    Const MSG304 As String = "304 Not Modified"
    Const MSG403 As String = "403 Forbidden"
    Const MSG404 As String = "404 Not Found"
    Const MSG500 As String = "500 Internal Server Error"
    Const MSG500_13 As String = "500.13 Server busy"
    Const MSG501 As String = "501 Not Implemented"
    
    Select Case Status
        Case 304
            SendCannedFormatted MSG304
        
        Case 403
            SendCannedFormatted MSG403
        
        Case 404
            SendCannedFormatted MSG404
        
        Case 500.13
            SendCannedFormatted MSG500_13
        
        Case 501
            SendCannedFormatted MSG501
        
        Case Else
            SendCannedFormatted MSG500
    End Select
End Sub

Private Sub SendCannedFormatted(ByVal StatusText As String)
    AppendResp mStrHTTPVersion & " " & StatusText & vbCrLf
    AppendResp "Date:" & UTCString(Now()) & vbCrLf
    AppendResp "Content-Type:text/html" & vbCrLf
    AppendResp "Content-Length:" & CStr(Len(StatusText) + 35) & vbCrLf
    AppendResp "Accept-Ranges:none" & vbCrLf
    AppendResp ServerHeader & DOUBLE_CRLF
    AppendResp "<html><body><h1>" & StatusText & "</h1></body></html>"
'    wskClient.SendData ExtractResp()
    mServer(mIntIndex).sOutBuffer = ExtractResp
    mServer(mIntIndex).TCPSend
End Sub

Private Sub SendResponse(ByVal Status As Single, ByVal StatusText As String, ByVal MIME As String, ByVal ExtraHeaders As String)
    AppendResp mStrHTTPVersion & " " & CStr(Status) & " " & StatusText & vbCrLf
    AppendResp "Date:" & UTCString(Now()) & vbCrLf
    If InStr(1, ExtraHeaders, "Last-Modified:", vbTextCompare) = 0 Then
        AppendResp "Last-Modified:" & UTCString(Now()) & vbCrLf
    End If
    If Len(MIME) > 0 Then AppendResp "Content-Type:" & MIME & vbCrLf
    mLngResponseLen = 0
    On Error Resume Next
    mLngResponseLen = UBound(mBytResponse) + 1
    On Error GoTo 0
    AppendResp "Content-Length:" & CStr(mLngResponseLen) & vbCrLf
    AppendResp "Accept-Ranges:none" & vbCrLf
    AppendResp ServerHeader & vbCrLf
    If Len(ExtraHeaders) > 0 Then
        AppendResp ExtraHeaders
        If Right$(ExtraHeaders, 2) <> vbCrLf Then AppendResp vbCrLf
    End If
    AppendResp vbCrLf 'Second CRLF, terminating headers.
    mServer(mIntIndex).sOutBuffer = ExtractResp
    mServer(mIntIndex).TCPSend
End Sub

Private Sub SendStaticHeader(ByVal Length As Long, ByVal Resource As String)
    Dim strMIME As String
    strMIME = ExtensionToMIME(Mid$(Resource, InStrRev(Resource, ".") + 1))
    AppendResp mStrHTTPVersion & " 200 Ok" & vbCrLf
    AppendResp "Date:" & UTCString(Now()) & vbCrLf
    AppendResp "Last-Modified:" & UTCString(FileDateTime(Resource)) & vbCrLf
    AppendResp "Content-Type:" & strMIME & vbCrLf
    AppendResp "Content-Length:" & CStr(Length) & vbCrLf
    AppendResp "Accept-Ranges:none" & vbCrLf
    AppendResp ServerHeader & DOUBLE_CRLF
    mServer(mIntIndex).sOutBuffer = ExtractResp
    mServer(mIntIndex).TCPSend
End Sub

Public Sub StartListening(Optional ByVal Port As Long = -1, Optional ByVal AdapterIP As String = "")
Attribute StartListening.VB_Description = "Begin accepting HTTP connections, may specify listen port and adapter IP to bind to"
    Dim GE As GossEvent
    Dim lNum As Long
    
    If Port > -1 Then Me.Port = Port
    
    IPVersion = "4"
    ReDim mServer(MaxConnections)
    For lNum = 0 To MaxConnections
        Set mServer(lNum).Callback(lNum) = Me
        mServer(lNum).EncrFlg = False
        mServer(lNum).IPvFlg = CLng(IPVersion)
    Next
    ReDim RecLen(MaxConnections)
    ReDim RecType(MaxConnections)
    mServer(0).Listen (Port)
    mBufferSize = mServer(0).BufferSize
    
    Set GE = New GossEvent
    With GE
        .EventType = getServer
        .EventSubtype = gesStarted
        .IP = AdapterIP
        .Port = Me.Port
        .Text = "Service started."
    End With
    RaiseEvent LogEvent(GE, -1)
End Sub

Public Sub StopListening()
Attribute StopListening.VB_Description = "Shuts down any active GossClients and stops listening for client connection requests"
    Dim GE As GossEvent
    Dim Index As Integer
    
    SimpleServer_CloseSck (0) 'Close listening socket
    
    Set GE = New GossEvent
    With GE
        .EventType = getServer
        .EventSubtype = gesStopped
        .Text = "Service stopped."
    End With
    RaiseEvent LogEvent(GE, -1)
End Sub

Public Function URLDecode(ByVal URLEncoded As String) As String
Attribute URLDecode.VB_Description = "Converts URLEncoded string to plaintext string"
    Dim intPart As Integer
    Dim strParts() As String
    
    URLDecode = Replace$(URLEncoded, "+", " ")
    strParts = Split(URLDecode, "%")
    For intPart = 1 To UBound(strParts)
        strParts(intPart) = _
                Chr$(CLng("&H" & Left$(strParts(intPart), 2))) _
              & Mid$(strParts(intPart), 3)
    Next
    URLDecode = Join$(strParts, "")
End Function

Public Function UTCDateTime(ByVal DateTime As Date) As Date
Attribute UTCDateTime.VB_Description = "Convert Date value from local time to UTC equivalent"
    Dim tzi As TIME_ZONE_INFORMATION
    Dim lngRet As Long
    Dim lngOffsetMinutes As Long
    
    'Return the time difference between local & GMT time in minutes.
    lngRet = GetTimeZoneInformation(tzi)
    lngOffsetMinutes = -tzi.Bias
    
    'If we are in daylight saving time, apply the bias if applicable.
    If lngRet = TIME_ZONE_ID_DAYLIGHT Then
        If tzi.DaylightDate.wMonth Then
            lngOffsetMinutes = lngOffsetMinutes - tzi.DaylightBias
        End If
    End If
    
    UTCDateTime = DateAdd("n", lngOffsetMinutes, DateTime)
End Function

Public Function UTCParseString(ByVal UTCString As String) As Date
Attribute UTCParseString.VB_Description = "Convert HTTP UTC timestamp string to Date value, if badly formatted return Now() value"
    On Error Resume Next
    UTCParseString = CDate(Mid$(UTCString, 6, 20))
    If Err.Number Then UTCParseString = UTCDateTime(Now())
End Function

Public Function UTCString(ByVal DateTime As Date) As String
Attribute UTCString.VB_Description = "Converts Date value in local time zone to HTTP timestamp in GMT form"
    UTCString = Format$(UTCDateTime(DateTime), _
                        "Ddd, dd Mmm YYYY HH:NN:SS \G\M\T")
End Function

Private Sub SimpleServer_CloseSck(ByVal Index As Long)
    On Error Resume Next
    Call mServer(Index).CloseSocket
End Sub

Private Sub SimpleServer_Connect(ByVal Index As Long)

End Sub


Private Sub SimpleServer_ConnectionRequest(ByVal Index As Long, ByVal requestID As Long, ByVal lRemotePort As Long, ByVal sRemoteHostIP As String)
    Dim lTmp As Long
    lTmp = GetNextSocket
    If lTmp > MaxConnections Then 'Request will exceed maximum, close the socket
        Call mServer(lIndex).Accept(requestID, 0, "")
        Exit Sub
    Else 'Accept the connection request
        lIndex = lTmp
        RemoteHostIP = sRemoteHostIP
        RemotePort = lRemotePort
        Call mServer(lIndex).Accept(requestID, RemotePort, RemoteHostIP)
        LastConnect = mServer(lIndex).RemoteHostIP
    End If
End Sub


Private Sub SimpleServer_DataArrival(ByVal Index As Long, ByVal bytesTotal As Long)
    Dim GE As GossEvent
    Dim strFragment As String
    Dim strHeadBlock As String
    Static strBuffer As String  'the buffer of the loading message
    Dim strChar As String
    Dim strContentLen As String
    Dim lInStr As Long
    Dim lHeader As Long
    Dim strHeaders() As String
    Dim strParts() As String
    
    mServer(Index).RecoverData 'Get inbound byte data
    strFragment = mServer(Index).sInBuffer 'Recover as string data
    mStrBuffer = mStrBuffer & strFragment
    Debug.Print strFragment
    If mReqState = rqsIdle Then
        'Erratic POST cleanup:
        'There is not supposed to be anything after the POST content but many
        'clients submit an extra CRLF.  Delete them if found here, which will
        'be a leftover from a previous request on a persistent connection.
        Do
            lInStr = InStr(mStrBuffer, vbCrLf)
            If lInStr > 0 Then
                If lInStr > 1 Then
                    'We found a complete Request Line.
                    mStrReqLine = Left$(mStrBuffer, lInStr - 1)
                    mReqState = rqsReqLineComplete
                End If
                mStrBuffer = Mid$(mStrBuffer, lInStr + 2)
            End If
        Loop Until lInStr = 0 Or lInStr > 1
    End If

    'Look for the Headers block if we have the Request Line.
    If mReqState = rqsReqLineComplete Then
        lInStr = InStr(mStrBuffer, DOUBLE_CRLF)
        If lInStr > 0 Then
            'We have the Headers.
            strHeadBlock = Left$(mStrBuffer, lInStr - 1)
            mStrBuffer = Mid$(mStrBuffer, lInStr + 4)
            
            'Parse Headers into Collection. Keys are stored UPPERCASED.
            Set mColReqHeaders = New Collection
            strHeaders = Split(strHeadBlock, vbCrLf)
            For lHeader = 0 To UBound(strHeaders)
                strParts = Split(strHeaders(lHeader), ":", 2)
                'Strip whitespace from Attribute.
                strChar = Right$(strParts(0), 1)
                Do While strChar = vbTab Or strChar = " "
                    strParts(1) = Left$(strParts(0), Len(strParts(0)) - 1)
                    strChar = Right$(strParts(0), 1)
                Loop
                If UBound(strParts) > 0 Then
                    'Strip whitespace from Value.
                    strChar = Left$(strParts(1), 1)
                    Do While strChar = vbTab Or strChar = " "
                        strParts(1) = Mid$(strParts(1), 2)
                        strChar = Left$(strParts(1), 1)
                    Loop
                End If
                'Watch for and remove duplicate headers (keep last one).
                On Error Resume Next
                mColReqHeaders.Add strParts, UCase$(strParts(0))
                If Err.Number Then
                    mColReqHeaders.Remove strParts(0)
                    mColReqHeaders.Add strParts, UCase$(strParts(0))
                End If
                On Error GoTo 0
            Next
            
            'Look for Content-Length.
            On Error Resume Next
            strContentLen = mColReqHeaders("CONTENT-LENGTH")(1)
            If Err.Number Then
                'No Content-Length header.  Bypass checking for it.
                On Error GoTo 0
                mReqState = rqsRequestComplete
            Else
                'Process Content-Length.
                On Error GoTo 0
                If IsNumeric(strContentLen) Then
                    mLngContentLen = CLng(strContentLen)
                    mReqState = rqsHeadersComplete
                Else
                    'Bad Content-Length error.
                    Set GE = New GossEvent
                    With GE
                        .EventType = getHTTP
                        .EventSubtype = gesHTTPError
                        .IP = RemoteHostIP
                        .Port = RemotePort
                        .Method = "ERROR"
                        .Text = "Bad Content-Length header value: " & strContentLen
                    End With
                    RaiseLogEvent GE, mIntIndex
                    
                    'wskClient_Close
                    SimpleServer_CloseSck (Index)
                    Exit Sub
                End If
            End If
        End If
    End If

    'Look for the end of the Request if we have processed the Headers.
    If mReqState = rqsHeadersComplete Then
        If Len(mStrBuffer) >= mLngContentLen Then
            mStrContent = Left$(mStrBuffer, mLngContentLen)
            mStrBuffer = Mid$(mStrBuffer, mLngContentLen + 1)
            mReqState = rqsRequestComplete
        End If
    End If
    
    'Process completed Request (all of Content-Length rcvd or no Content-Length header).
    mIntIndex = Index
    If mReqState = rqsRequestComplete Then ProcessRequest
End Sub


Private Sub SimpleServer_EncrDataArrival(ByVal Index As Long, ByVal bytesTotal As Long)

End Sub


Private Sub SimpleServer_Error(ByVal Index As Long, ByVal Number As Long, Description As String, ByVal Source As String)

End Sub


Private Sub SimpleServer_SendComplete(ByVal Index As Long)
    Debug.Print "SendComplete"
    If mLngResponseLen > 0 Then
        mLngResponseLen = 0
        mServer(Index).bOutBuffer = mBytResponse
        mServer(Index).TCPSend
        Erase mBytResponse
        Exit Sub 'Bypass CheckClose until next SendComplete.
    End If
    
    If mIntFile Then
        'We're sending a static (file) resource.  Continue.
        If mIntFullBlocks > 0 Then
            Get #mIntFile, , mBytBuffer
            mIntFullBlocks = mIntFullBlocks - 1
        Else
            If mLngLastBlockSize > 0 Then
                ReDim mBytBuffer(mLngLastBlockSize - 1)
                Get #mIntFile, , mBytBuffer
            End If
            Close #mIntFile
            mIntFile = 0
            If mLngLastBlockSize <= 0 Then GoTo CheckClose
        End If
        'wskClient.SendData mBytBuffer
        mServer(Index).bOutBuffer = mBytBuffer
        mServer(Index).TCPSend
        Exit Sub 'Bypass CheckClose until next SendComplete.
    End If

CheckClose:
    If mBlnConnClose Then
        'Request had a "Connection: close" header.
        'wskClient_Close
        SimpleServer_CloseSck (Index)
    End If
End Sub


Private Sub SimpleServer_SendProgress(ByVal Index As Long, ByVal bytesSent As Long, ByVal bytesRemaining As Long)

End Sub


Private Sub SimpleServer_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

End Sub

Private Sub UserControl_Initialize()
    ChDir App.Path
    mServerHeader = "Server:TS.Gossamer/" & GOSS_VERSION_MAJOR & "." & GOSS_VERSION_MINOR _
                  & vbCrLf _
                  & "X-Powered-By:Microsoft.Visual.Basic.6.0"
End Sub

Private Sub UserControl_InitProperties()
    DefaultPage = "index.html"
    MaxConnections = 320
    Port = 8080
    VDir = ""
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    DefaultPage = PropBag.ReadProperty("DefaultPage", "index.html")
    MaxConnections = PropBag.ReadProperty("MaxConnections", 320)
    Port = PropBag.ReadProperty("Port", 8080)
    VDir = PropBag.ReadProperty("VDir", "")
End Sub

Private Sub UserControl_Resize()
    Const DIMENSIONS As Single = 420
    With UserControl
        .Height = DIMENSIONS
        .Width = DIMENSIONS
    End With
    With Shape1
        .Height = DIMENSIONS
        .Width = DIMENSIONS
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "DefaultPage", DefaultPage, "index.htm"
    PropBag.WriteProperty "MaxConnections", MaxConnections, 32
    PropBag.WriteProperty "Port", Port, 8080
    PropBag.WriteProperty "VDir", VDir, ""
End Sub

