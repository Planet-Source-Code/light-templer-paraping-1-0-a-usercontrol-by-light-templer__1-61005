VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDNS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'
'   clsDNS.cls
'


Option Explicit

' Stuff for Get DNS Servers
Private Const MAX_HOSTNAME_LEN = 128
Private Const MAX_DOMAIN_NAME_LEN = 128
Private Const MAX_SCOPE_ID_LEN = 256
Private Const ERROR_SUCCESS As Long = 0

Private Type IP_ADDRESS_STRING
    IpAddr(0 To 15)     As Byte
End Type

Private Type IP_MASK_STRING
    IpMask(0 To 15)     As Byte
End Type

Private Type IP_ADDR_STRING
    dwNext              As Long
    IpAddress(0 To 15)  As Byte
    IpMask(0 To 15)     As Byte
    dwContext           As Long
End Type

Private Type FIXED_INFO
    HostName(0 To (MAX_HOSTNAME_LEN + 3))       As Byte
    DomainName(0 To (MAX_DOMAIN_NAME_LEN + 3))  As Byte
    CurrentDnsServer                            As Long
    DnsServerList                               As IP_ADDR_STRING
    NodeType                                    As Long
    ScopeId(0 To (MAX_SCOPE_ID_LEN + 3))        As Byte
    EnableRouting                               As Long
    EnableProxy                                 As Long
    EnableDns                                   As Long
End Type

Private Declare Function GetNetworkParams Lib "iphlpapi.dll" _
        (pFixedInfo As Any, _
         pOutBufLen As Long) As Long

Private Declare Function lstrcpyA Lib "kernel32" _
        (ByVal RetVal As String, _
         ByVal ptr As Long) As Long
                        
Private Declare Function lstrlenA Lib "kernel32" _
        (ByVal ptr As Any) As Long
   
Private Declare Function inet_ntoa Lib "wsock32.dll" _
        (ByVal addr As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, _
         Source As Any, _
         ByVal Length As Long)


' Stuff for DNS resolve
Const WSADescription_Len = 256
Const WSASYS_Status_Len = 128

Private Type HOSTENT
    hName       As Long
    hAliases    As Long
    hAddrType   As Integer
    hLength     As Integer
    hAddrList   As Long
End Type

Private Type WSADATA
    wversion                                As Integer
    wHighVersion                            As Integer
    szDescription(0 To WSADescription_Len)  As Byte
    szSystemStatus(0 To WSASYS_Status_Len)  As Byte
    iMaxSockets                             As Integer
    iMaxUdpDg                               As Integer
    lpszVendorInfo                          As Long
End Type

Private Declare Function WSAStartup Lib "wsock32" _
        (ByVal VersionReq As Long, _
         WSADataReturn As WSADATA) As Long
         
Private Declare Function WSACleanup Lib "wsock32" () As Long

Private Declare Function WSAGetLastError Lib "wsock32" () As Long

Private Declare Function gethostbyaddr Lib "wsock32" _
        (addr As Long, _
         addrLen As Long, _
         addrType As Long) As Long
         
Private Declare Function gethostbyname Lib "wsock32" _
        (ByVal HostName As String) As Long
        
Private Declare Sub RtlMoveMemory Lib "kernel32" _
        (hpvDest As Any, _
         ByVal hpvSource As Long, _
         ByVal cbCopy As Long)

Public Event Error(ByVal Number As Long, Description As String)
Public Event ResolveCompleted()
'
'
'

Private Sub Class_Initialize()
    
    Dim udtWSAData As WSADATA
    
    If WSAStartup(257, udtWSAData) Then
        RaiseEvent Error(Err.LastDllError, Err.Description)
    End If
    
End Sub

Private Sub Class_Terminate()

    WSACleanup
    
End Sub

Public Function AddressToName(strIP As String) As String
    
    AddressToName = NameByAddr(strIP)
    RaiseEvent ResolveCompleted
    
End Function

Public Function NameToAddress(strName As String) As String
    
    NameToAddress = AddrByName(strName)
    RaiseEvent ResolveCompleted
    
End Function

Private Function IsIP(ByVal strIP As String) As Boolean
    'checks if string is valid IP address
    
    Dim t As String
    Dim s As String
    Dim i As Integer
    
    On Error Resume Next
    
    s = strIP
    While InStr(s, ".") <> 0
        t = Left(s, InStr(s, ".") - 1)
        If IsNumeric(t) And Val(t) >= 0 And Val(t) <= 255 Then
            s = Mid(s, InStr(s, ".") + 1)
        Else
            
            Exit Function
        End If
        
        i = i + 1
    Wend
    
    t = s
    If IsNumeric(t) And _
            InStr(t, ".") = 0 And _
            Len(t) = Len(Trim(Str(Val(t)))) _
            And Val(t) >= 0 And _
            Val(t) <= 255 And _
            strIP <> "255.255.255.255" And _
            i = 3 Then
        IsIP = True
    End If
    
    If Err.Number > 0 Then
      RaiseEvent Error(Err.Number, Err.Description)
      Err.Clear
    End If
    
End Function

Private Function MakeIP(strIP As String) As Long
    'converts IP address from string to sin_addr
    
    Dim lIP As Long
    
    On Error Resume Next
    
    lIP = Left(strIP, InStr(strIP, ".") - 1)
    strIP = Mid(strIP, InStr(strIP, ".") + 1)
    lIP = lIP + Left(strIP, InStr(strIP, ".") - 1) * 256
    strIP = Mid(strIP, InStr(strIP, ".") + 1)
    lIP = lIP + Left(strIP, InStr(strIP, ".") - 1) * 256 * 256
    strIP = Mid(strIP, InStr(strIP, ".") + 1)
    If strIP < 128 Then
      lIP = lIP + strIP * 256 * 256 * 256
    Else
      lIP = lIP + (strIP - 256) * 256 * 256 * 256
    End If
    MakeIP = lIP
    
    If Err.Number > 0 Then
      RaiseEvent Error(Err.Number, Err.Description)
      Err.Clear
    End If
    
End Function

Private Function NameByAddr(strAddr As String) As String
    'resolves IP address to host name
  
    Dim nRet      As Long
    Dim lIP       As Long
    Dim strHost   As String * 255
    Dim strTemp   As String
    Dim hst       As HOSTENT
    
    On Error Resume Next
    
    If IsIP(strAddr) Then
        lIP = MakeIP(strAddr)
        nRet = gethostbyaddr(lIP, 4, 2)
        If nRet <> 0 Then
            RtlMoveMemory hst, nRet, Len(hst)
            RtlMoveMemory ByVal strHost, hst.hName, 255
            strTemp = strHost
            If InStr(strTemp, Chr(10)) <> 0 Then strTemp = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
            strTemp = Trim(strTemp)
            NameByAddr = strTemp
        Else
            RaiseEvent Error(9003, "Host name not found")
            
            Exit Function
        End If
    Else
      RaiseEvent Error(9002, "Invalid IP address")
      
      Exit Function
    End If
    
    If Err.Number > 0 Then
        RaiseEvent Error(Err.Number, Err.Description)
        Err.Clear
    End If
  
End Function

Private Function AddrByName(ByVal strHost As String) As String
    'resolves host name to IP address
  
    Dim hostent_addr        As Long
    Dim hst                 As HOSTENT
    Dim hostip_addr         As Long
    Dim temp_ip_address()   As Byte
    Dim i                   As Integer
    Dim ip_address          As String
    
    On Error Resume Next
    
    If IsIP(strHost) Then
        AddrByName = strHost
        
        Exit Function
    End If
    
    hostent_addr = gethostbyname(strHost)
    If hostent_addr = 0 Then
        RaiseEvent Error(9001, "Can't resolve hst")
        
        Exit Function
    End If
    
    RtlMoveMemory hst, hostent_addr, LenB(hst)
    RtlMoveMemory hostip_addr, hst.hAddrList, 4
    
    ReDim temp_ip_address(1 To hst.hLength)
    
    RtlMoveMemory temp_ip_address(1), hostip_addr, hst.hLength
    For i = 1 To hst.hLength
        ip_address = ip_address & temp_ip_address(i) & "."
    Next
    ip_address = Mid(ip_address, 1, Len(ip_address) - 1)
    
    AddrByName = ip_address
    If Err.Number > 0 Then
        RaiseEvent Error(Err.Number, Err.Description)
        
        Err.Clear
    End If
    
End Function


' ========================
' = Get DNS Server infos =
' ========================

' Requires W2k !

Public Function GetDNSServers(sCurrentDNSserver As String, sArrDNSserver() As String) As Long

    Dim buff()      As Byte
    Dim cbRequired  As Long
    Dim nStructSize As Long
    Dim ptr         As Long
    Dim fi          As FIXED_INFO
    Dim ipas        As IP_ADDR_STRING
    Dim cnt         As Long
    Dim lResult     As Long
    
    Const ERROR_NOT_SUPPORTED = 50&
    
    ReDim sArrDNSserver(0) As String
     
    nStructSize = LenB(ipas)
    
    ' call the api first to determine the size required for the values to be returned
    lResult = GetNetworkParams(ByVal 0&, cbRequired)
    
    If lResult = ERROR_NOT_SUPPORTED Then
        GetDNSServers = -1                      ' not supported by this OS
    
        Exit Function
    End If
    
    If cbRequired > 0 Then
    
        ReDim buff(0 To cbRequired - 1) As Byte
        
        If GetNetworkParams(buff(0), cbRequired) = ERROR_SUCCESS Then
            ptr = VarPtr(buff(0))
            CopyMemory fi, ByVal ptr, Len(fi)
    
            With fi
                ' identify the current dns server
                CopyMemory ipas, ByVal VarPtr(.CurrentDnsServer) + 4, nStructSize
                 
                sCurrentDNSserver = TrimNull(StrConv(ipas.IpAddress, vbUnicode))
                
                ' obtain a pointer to the DnsServerList array
                ptr = VarPtr(.DnsServerList)
                
                ' the IP_ADDR_STRING dwNext member indicates that more than one DNS server may be listed,
                ' so a loop is needed
                Do While (ptr <> 0)
                
                    'copy each into an IP_ADDR_STRING type
                    CopyMemory ipas, ByVal ptr, nStructSize
                    
                    With ipas
                    
                      'extract the server address and cast to the array
                       ReDim Preserve sArrDNSserver(0 To cnt) As String
                       sArrDNSserver(cnt) = TrimNull(StrConv(ipas.IpAddress, vbUnicode))
                       ptr = .dwNext
                       
                    End With
                    
                    cnt = cnt + 1
                 Loop
            End With
        End If  'If GetNetworkParams
    End If  'If cbRequired > 0

    'return number of servers found
    GetDNSServers = cnt
      
End Function

Private Function TrimNull(item As String)

    Dim pos As Integer
   
    pos = InStr(item, Chr$(0))
    If pos Then
        TrimNull = Left$(item, pos - 1)
    Else
        TrimNull = item
    End If
  
End Function

Private Function GetInetStrFromPtr(Address As Long) As String

    GetInetStrFromPtr = GetStrFromPtrA(inet_ntoa(Address))

End Function

Private Function GetStrFromPtrA(ByVal lpszA As Long) As String

    GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
    Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
   
End Function


' #*#
