VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnumComps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'
'   clsEnumComps.cls
'

'   9/3/2004
'   LightTempler

'   V. 2.00

'   Last edit:  4/6/2005


'   New in V. 2.0   * Enum machines in Active Directory, so you can get ALL windows computers
'                     in your domain. (API_NetServerEnum() enums switched on only ...)
'                     You get this by using new CompType (10) in enum 'enCompType'.
'                   * Working with current domain as default.
'                   * New property 'Domain'
'                   * New property 'LastErrMsg'
'                   * New property 'MachinesFound'
'                   * Redesign of code and interface, too! (Sry, bad style; I know...)
'                   * More comments added ;-)



Option Explicit


' *******************************
' *            EVENTS           *
' *******************************
Public Event CompFound(sCompName As String)
Public Event Error(sError As String)


' *******************************
' *            ENUMS            *
' *******************************
Public Enum enCompType
    EC_CT_AllMachines = 1
    EC_CT_AllWindowsComps = 2
    EC_CT_TerminalServer = 3
    EC_CT_Unix = 4
    EC_CT_PrintServer = 5
    EC_CT_DomainController = 6
    EC_CT_SQLServer = 7
    EC_CT_TimeServer = 8
    EC_CT_DFS = 9
    EC_CT_MachinesInActiveDirectory = 10
End Enum


' *******************************
' *           TYPES             *
' *******************************
Private Type tpSERVER_INFO_API              ' Level 101
    PlatformId  As Long
    servername  As Long
    VerMajor    As Long
    VerMinor    As Long
    Type        As Long
    comment     As Long
End Type

Private Type tpServerInfoFull
    PlatformId  As Long
    servername  As String
    Type        As Long
    VerMajor    As Long
    VerMinor    As Long
    comment     As String
    Platform    As String
    ServerType  As Long
    LanGroup    As String
    LanRoot     As String
End Type



' *************************************
' *          API CONSTANTS            *
' *************************************

' Const list from famous http://vbnet.mvps.org
Private Const SV_TYPE_WORKSTATION         As Long = &H1
Private Const SV_TYPE_SERVER              As Long = &H2
Private Const SV_TYPE_SQLSERVER           As Long = &H4
Private Const SV_TYPE_DOMAIN_CTRL         As Long = &H8
Private Const SV_TYPE_DOMAIN_BAKCTRL      As Long = &H10
Private Const SV_TYPE_TIME_SOURCE         As Long = &H20
Private Const SV_TYPE_AFP                 As Long = &H40
Private Const SV_TYPE_NOVELL              As Long = &H80
Private Const SV_TYPE_DOMAIN_MEMBER       As Long = &H100
Private Const SV_TYPE_PRINTQ_SERVER       As Long = &H200
Private Const SV_TYPE_DIALIN_SERVER       As Long = &H400
Private Const SV_TYPE_XENIX_SERVER        As Long = &H800
Private Const SV_TYPE_SERVER_UNIX         As Long = SV_TYPE_XENIX_SERVER
Private Const SV_TYPE_NT                  As Long = &H1000
Private Const SV_TYPE_WFW                 As Long = &H2000
Private Const SV_TYPE_SERVER_MFPN         As Long = &H4000
Private Const SV_TYPE_SERVER_NT           As Long = &H8000
Private Const SV_TYPE_POTENTIAL_BROWSER   As Long = &H10000
Private Const SV_TYPE_BACKUP_BROWSER      As Long = &H20000
Private Const SV_TYPE_MASTER_BROWSER      As Long = &H40000
Private Const SV_TYPE_DOMAIN_MASTER       As Long = &H80000
Private Const SV_TYPE_SERVER_OSF          As Long = &H100000
Private Const SV_TYPE_SERVER_VMS          As Long = &H200000
Private Const SV_TYPE_WINDOWS             As Long = &H400000            ' Windows95 and above
Private Const SV_TYPE_DFS                 As Long = &H800000            ' Root of a DFS tree
Private Const SV_TYPE_CLUSTER_NT          As Long = &H1000000           ' NT Cluster
Private Const SV_TYPE_TERMINALSERVER      As Long = &H2000000           ' Terminal Server
Private Const SV_TYPE_DCE                 As Long = &H10000000          ' IBM DSS
Private Const SV_TYPE_ALTERNATE_XPORT     As Long = &H20000000          ' rtn alternate transport
Private Const SV_TYPE_LOCAL_LIST_ONLY     As Long = &H40000000          ' rtn local only
Private Const SV_TYPE_DOMAIN_ENUM         As Long = &H80000000
Private Const SV_TYPE_ALL                 As Long = &HFFFFFFFF

Private Const NERR_Success          As Long = 0&
Private Const NERR_MoreData         As Long = 234&



' *************************************
' *        API DEFINITIONS            *
' *************************************
Private Declare Function API_NetServerEnum Lib "NETAPI32" Alias "NetServerEnum" _
        (ByRef lpServer As Any, _
         ByVal lLevel As Long, _
         ByRef vBuffer As Any, _
         ByRef lPreferedMaxLen As Long, _
         ByRef lEntriesRead As Long, _
         ByRef lTotalEntries As Long, _
         ByVal lServerType As Long, _
         ByVal sDomain$, _
         ByRef vResume As Any) As Long

Private Declare Sub API_CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" _
        (ByRef pTo As Any, _
         ByRef uFrom As Any, _
         ByVal lSize As Long)

Private Declare Function API_lstrlenW Lib "KERNEL32" Alias "lstrlenW" _
        (ByVal PointerToString As Long) As Long


Private Declare Function API_NetAPIBufferFree Lib "NETAPI32" Alias "NetApiBufferFree" _
        (ByVal lpBuffer As Long) As Long



' *************************************
' *       PRIVATE VARS                *
' *************************************
Private Type tpMvar
    sFilterIn               As String
    sFilterOut              As String
    lMachinesFound          As Long
    sLastErrMsg             As String
    sDomain_NETBIOS         As String
    sDomain_AD              As String
End Type
Private Mvar As tpMvar
'
'
'



' *************************************
' *           INIT/TERM               *
' *************************************
Private Sub Class_Initialize()

    With Mvar
        .sFilterIn = "*"
        .lMachinesFound = -1
    End With

End Sub

Private Sub Class_Terminate()

    ' Nothing yet

End Sub



' *************************************
' *         PUBLIC FUNCTIONS          *
' *************************************
Public Function EnumComps(ByVal CompType As enCompType) As Long

    ' For CompType <> EC_CT_MachinesInActiveDirectory property 'Domain_NETBIOS' must be set in NetBIOS form, e.g. 'MS'
    ' For CompType =  EC_CT_MachinesInActiveDirectory property 'Domain_ActiveDirectory' must be set in DNS form, e.g. 'microsoft.com'
    ' When domain isn't set explicitly your current domain is used for enumerating machines.
    

    ' Parameter check
    If CompType < enCompType.EC_CT_AllMachines Or CompType > enCompType.EC_CT_MachinesInActiveDirectory Then
        RaiseEvent Error("Wrong Parameter 'CompType' (" & CompType & ")")
        
        Exit Function
    End If
    
    ' For comparing with LIKE we need the asterik to get all machine names (means: No filtering)
    If Mvar.sFilterIn = "" Then
        Mvar.sFilterIn = "*"
    End If
    
    If CompType = EC_CT_MachinesInActiveDirectory Then
        EnumComps = EnumCompsAD()
    Else
        EnumComps = EnumCompsNETBIOS(CompType)
    End If

End Function
    
    
' *************************************
' *         PRIVATE FUNCTIONS         *
' *************************************
Private Function EnumCompsNETBIOS(CompType As Long) As Long
    
    Dim i                   As Long
    Dim lReturn             As Long
    Dim lServerInfo         As Long
    Dim lServerInfoPtr      As Long
    Dim lPreferedMaxLen     As Long
    Dim lEntriesRead        As Long
    Dim lTotalEntries       As Long
    Dim bytArrServer()      As Byte
    Dim lServerType         As Long
    Dim sCompName           As String
    Dim lFilterOutLength    As Long
    Dim vResume             As Variant
    Dim tServerInfo         As tpSERVER_INFO_API
    Dim sDomainUC           As String
    
    On Local Error GoTo error_handler
    
    
    ' Add here what you need from const list / enum above
    lServerType = Choose(CompType, SV_TYPE_ALL, _
                                    SV_TYPE_NT, _
                                    SV_TYPE_TERMINALSERVER, _
                                    SV_TYPE_SERVER_UNIX, _
                                    SV_TYPE_PRINTQ_SERVER, _
                                    SV_TYPE_DOMAIN_CTRL, _
                                    SV_TYPE_SQLSERVER, _
                                    SV_TYPE_TIME_SOURCE, _
                                    SV_TYPE_DFS)
        
    With Mvar
        bytArrServer = "" + vbNullChar
        lPreferedMaxLen = 65536
        lFilterOutLength = Len(Mvar.sFilterOut)
        
        If .sDomain_NETBIOS = "" Then
            .sDomain_NETBIOS = GetDomainName_NETBIOS()
        End If
        sDomainUC = StrConv(.sDomain_NETBIOS, vbUnicode)
        

        
        ' Until there is no more data, get list by calling API 'NetServerEnum()'
        lReturn = NERR_MoreData
        Do While (lReturn = NERR_MoreData)
            
            lReturn = API_NetServerEnum(bytArrServer(0), _
                    101, _
                    lServerInfo, _
                    lPreferedMaxLen, _
                    lEntriesRead, _
                    lTotalEntries, _
                    lServerType, _
                    sDomainUC, _
                    vResume)
            
            If (lReturn <> NERR_Success And lReturn <> NERR_MoreData) Then
                
                Exit Do
            End If
            
            ' Get comp names names from list
            i = 1
            lServerInfoPtr = lServerInfo
            Do While i <= lTotalEntries
                
                ' Get computer name from filled buffer
                API_CopyMemory tServerInfo, ByVal lServerInfoPtr, Len(tServerInfo)
                sCompName = UCase$(PointerToStringW(tServerInfo.servername))
                
                ' Filter to match given conditions
                If sCompName Like .sFilterIn Then
                    If lFilterOutLength Then
                        If Left$(sCompName, lFilterOutLength) <> .sFilterOut Then
                            RaiseEvent CompFound(sCompName)
                            .lMachinesFound = .lMachinesFound + 1
                        End If
                    Else
                        RaiseEvent CompFound(sCompName)
                        .lMachinesFound = .lMachinesFound + 1
                    End If
                End If
                i = i + 1
                lServerInfoPtr = lServerInfoPtr + Len(tServerInfo)
            Loop
            API_NetAPIBufferFree lServerInfo
        Loop
        EnumCompsNETBIOS = .lMachinesFound
    End With

    Exit Function


error_handler:

    RaiseError "[" + Err.Description + "] when enumerating with API."
    EnumCompsNETBIOS = -1 ' Error

End Function


Private Function PointerToStringW(ByVal lpStringW As Long) As String
    ' Copy unicode string to VB string
    
    Dim Buffer()    As Byte
    Dim nLen        As Long
    
    If lpStringW Then
        nLen = API_lstrlenW(lpStringW) * 2
        
        If nLen Then
            ReDim Buffer(0 To (nLen - 1)) As Byte
            API_CopyMemory Buffer(0), ByVal lpStringW, nLen
            PointerToStringW = Buffer
        End If
        
    End If
    
End Function

Private Function GetDomainName_NETBIOS()
    ' You need a reference to 'Active DS Type Library' for this!
    
    Dim oSysInfo As WinNTSystemInfo
    
    Set oSysInfo = New WinNTSystemInfo
    GetDomainName_NETBIOS = oSysInfo.DomainName
    Set oSysInfo = Nothing
    
End Function



Private Function EnumCompsAD() As Long
    
    Dim oADcomps            As IADsContainer
    Dim oADComp             As IADs
    Dim sCompName           As String
    Dim lFilterOutLength    As Long
    
    On Local Error GoTo EnumCompsAD_Error

    With Mvar
        If .sDomain_AD = "" Then
            .sDomain_AD = GetDomainName_DNS()
        End If
        
        Set oADcomps = GetObject("LDAP://CN=Computers," + .sDomain_AD)
        If oADcomps Is Nothing Then
            
            Exit Function
        End If
        
        lFilterOutLength = Len(.sFilterOut)
        For Each oADComp In oADcomps
            sCompName = Mid$(oADComp.Name, 4)     ' oADComp.Name has form "CN=MyComputer01"
            
            ' Filter to match given conditions
            If sCompName Like .sFilterIn Then
                If lFilterOutLength Then
                    If Left$(sCompName, lFilterOutLength) <> .sFilterOut Then
                        RaiseEvent CompFound(sCompName)
                        .lMachinesFound = .lMachinesFound + 1
                    End If
                Else
                    RaiseEvent CompFound(sCompName)
                    .lMachinesFound = .lMachinesFound + 1
                End If
            End If
        
        Next oADComp
        EnumCompsAD = .lMachinesFound
    End With
    
   
    Exit Function


EnumCompsAD_Error:
    
    RaiseError "[" + Err.Description + "] when enumerating with LDAP."
    EnumCompsAD = -1 ' Error
    
End Function

Private Function GetDomainName_DNS() As String

    Dim rootDSE         As IADs
    
    On Local Error Resume Next
    
    Set rootDSE = GetObject("LDAP://RootDSE")
    GetDomainName_DNS = rootDSE.Get("defaultnamingcontext")
    Set rootDSE = Nothing

End Function



Private Function RaiseError(sErrMsg As String)
    ' Centralized error reporting for easy changes

    Mvar.sLastErrMsg = sErrMsg
    RaiseEvent Error(sErrMsg)

End Function




' *************************************
' *           PROPERTIES              *
' *************************************
Public Property Let FilterIn(ByVal sNewFilterIn As String)
    ' Found machine names will compared using VB's 'LIKE' command with this pattern.
    ' On 'match' the machine is raised if 'FilterOut' doesn't prevent this.
    
    Mvar.sFilterIn = UCase$(Trim$(sNewFilterIn))

End Property

Public Property Get FilterIn() As String
    
    FilterIn = Mvar.sFilterIn

End Property


Public Property Let FilterOut(ByVal sNewFilterOut As String)
    ' If 'FilterOut' isn't empty the left part of found machine names will compared using
    ' VB's 'LEFT$()' function with this string.
    ' Only 'no match' the machine is raised.
    
    Mvar.sFilterOut = UCase$(Trim$(sNewFilterOut))

End Property

Public Property Get FilterOut() As String
    
    FilterOut = Mvar.sFilterOut

End Property


Public Property Get MachinesFound() As Long
    ' Result:   How many machines matched last search/filtering
    '           -1 means: There was an error
    
    FilterIn = Mvar.lMachinesFound

End Property


Public Property Get LastErrMsg() As String
    ' Empty, when no error
    
    LastErrMsg = Mvar.sLastErrMsg

End Property


Public Property Let Domain_NETBIOS(ByVal sNewDomain_NETBIOS As String)
    ' Leave empty to enum your current domain
    
    Mvar.sDomain_NETBIOS = Trim$(sNewDomain_NETBIOS)

End Property

Public Property Get Domain_NETBIOS() As String
    ' Filled with your current domain when calling EnumComps() with empty 'Domain_NETBIOS'
    
    Domain_NETBIOS = Mvar.sDomain_NETBIOS

End Property


Public Property Let Domain_AD(ByVal sNewDomain_AD As String)
    ' Leave empty to enum your current domain
    
    Mvar.sDomain_AD = Trim$(sNewDomain_AD)

End Property

Public Property Get Domain_AD() As String
    ' Filled with your current domain when calling EnumComps() with empty 'Domain_AD'
    
    Domain_NETBIOS = Mvar.sDomain_AD

End Property

' *#*
