VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl ucParaPing 
   CanGetFocus     =   0   'False
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3930
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ucParaPing.ctx":0000
   ScaleHeight     =   1560
   ScaleWidth      =   3930
   ToolboxBitmap   =   "ucParaPing.ctx":0974
   Begin MSWinsockLib.Winsock wsckArrPing 
      Index           =   1
      Left            =   1500
      Top             =   645
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   7
   End
   Begin VB.Timer tmrSchedule 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   915
      Top             =   630
   End
End
Attribute VB_Name = "ucParaPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'
'   ParaPING 1.0
'

'   Multiple parallel PINGs to many destination hosts.
'
'   If you wants to know within a few seconds which of
'   the 5000 machines in your local area network are really
'   running this ones for you! ;-)

'   Project started in April 2005 by Light Templer

'   Last edit:  6/9/2005


Option Explicit


'   Description:
'
'   Using the AddHost() function IP adresses are added to a ring buffer.
'   From there the timer loop (worker loop) fetches them one bye one,
'   put them into a list with infomations about running Pings and sends
'   a Ping (in fact it opens a connection to port 7) to the destination host.
'   The Pings are realized by a control array of WinSock controls, not by
'   any API (ICMP) calls. This way we have a kind of a "multi threading" in
'   pure VB without any dirty tricks, memory errors or any other problems.
'
'   That means: You can run 50 or more Pings asyncron the same time and
'   your GUI (form) stays fully responsible and interactive ...  ;-)
'   So checking hundreds or thousands of machines can be done in just a few
'   secondes.



' ******************************
' *           EVENTS           *
' ******************************
Public Event Pong(sIPadr As String, lID As Long, flgSuccess As Boolean)     ' Gives result of a Ping: 'Dead' or 'alive'
Public Event StateChanged(New_State As enState, sNewState As String)        ' Informs about any changes in ParaPings state
Public Event Error(Errorcode As enErrorCodes, sErr As String)               ' An error has occourd




' ******************************
' *      PUBLIC ENUMS          *
' ******************************
Public Enum enState
    PP_IDLE = 0
    PP_BUSY = 1
    PP_DISABLED = 2
    PP_ABORT_PENDING = 3
End Enum

Public Enum enErrorCodes
    PP_ERR_NO_ERROR = 0
    PP_ERR_GENERAL_ERROR = 1
    PP_ERR_BUFFER_OVERFLOW = 2
    PP_ERR_WRONG_PARAMETER = 3
    PP_ERR_NOT_IDLE = 4
    PP_ERR_WINSOCK = 5
End Enum



' ******************************
' *           CONSTS           *
' ******************************
Const THREADLIMIT           As Long = 100&              ' Upper limit for number of simultan Pings
Const QUEUELIMIT            As Long = 20000&            ' Max IP adresses in queue waiting for a Ping
Const MAXTIMEOUT            As Long = 10&               ' Max number of seconds to wait for an answer to a Ping
Const PING_PORT             As Long = 7&                ' TCP port for Ping (ICMP)



' ******************************
' *       DEFAULT VALUES       *
' ******************************
Const DEFAULT_MAXTHREADS    As Long = 10&               ' Default value for upper limit for number of simultan Pings
Const DEFAULT_QUEUESIZE     As Long = 500&              ' Default value for max IP adresses in queue waiting for a Ping
Const DEFAULT_TIMEOUT       As Long = 3&                ' Default value for timeout for Ping




' ******************************
' *         LOCAL UDTs         *
' ******************************

Private Type tpDestHost
    sIPadr                  As String                   ' IP adress of destination host in usual form, e.g.  "130.112.50.10"
    lID                     As Long                     ' Any Id (number) the uc user wants, e.g. an index into an array or listview
End Type

Private Type tpThread
    sIPadr                  As String                   ' IP adress of destination host in usual form, e.g.  "130.112.50.10"
    lID                     As Long                     ' Any Id (number) the uc user wants, e.g. an index into an array or listview
    lStartTime              As Long                     ' Got with VBs timer function:  Int(Timer())
    flgPong                 As Boolean                  ' Set to TRUE, when we get an answer from destination host
End Type

Private Type tpVAR
    flgEnabled              As Boolean                  ' TRUE: Adding a new host with 'AddHost' immediatly starts checking
    State                   As enState                  ' Current state of ParaPING control
    lOpenPings              As Long                     ' Running 'Pings' waiting for their 'Pongs'
    lMaxThreads             As Long                     ' How many Pings at the same time
    lTimeout                As Long                     ' Timeout for Ping result in seconds (min. 1)
    lWaitingInQueue         As Long                     ' How many host entries in queue are waiting for their check by Ping
    lQueueSize              As Long                     ' Size of queue with host entries
    lNxtFreePosInQueue      As Long                     ' Pointer into ringbuffer:  Next free position to save an entry.
    lNxtItemToTakeFromQueue As Long                     ' Pointer into ringbuffer:  Position of next entry to handle.
    sLastErr                As String                   ' The last resulted error message raised by 'RaiseError'. Empty when no error.
    LastErrorCode           As enErrorCodes             ' The last resulted error code.
End Type



' ******************************
' *         LOCAL VARs         *
' ******************************

Private VAR                 As tpVAR                    ' Holds all vars of the control (no the arrays!)
Private arrQueueDestHosts() As tpDestHost               ' Host entries waiting for their check by Ping. Organiced as a ringbuffer
Private arrThreads()        As tpThread                 ' Running 'Pings'. Organiced with empty and used slots.
'
'
'






' ******************************
' *        ALL USERCONTROL     *
' ******************************
Private Sub UserControl_Initialize()
    
    ' Nothing yet
    
End Sub


Private Sub UserControl_Terminate()
    
    Call Me.Abort
    If VAR.State = PP_IDLE Then
        Call FreeResources
    End If
    
End Sub


Private Sub UserControl_InitProperties()
    
    With VAR
        .lQueueSize = DEFAULT_QUEUESIZE
        .lMaxThreads = DEFAULT_MAXTHREADS
        .lTimeout = DEFAULT_TIMEOUT
    End With
    
    With wsckArrPing(1)
        .Protocol = sckUDPProtocol                          ' We need the connectionless UDP protocol
        .RemotePort = PING_PORT
    End With
    
End Sub

Private Sub UserControl_Resize()
    
    Const ucWIDTH = 420
    
    UserControl.Width = ucWIDTH
    UserControl.Height = ucWIDTH
    
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With VAR
        .lMaxThreads = PropBag.ReadProperty("MaxThreads", DEFAULT_MAXTHREADS)
        .lQueueSize = PropBag.ReadProperty("QueueSize", DEFAULT_QUEUESIZE)
        .lTimeout = PropBag.ReadProperty("Timeout", DEFAULT_TIMEOUT)
        
        ' Init on app start
        If Ambient.UserMode = True Then
        
            ' Setup arrays
            ReDim arrQueueDestHosts(1 To .lQueueSize)
            ReDim arrThreads(1 To .lMaxThreads)
            
            ' Set pointers into ringbuffer to start position
            .lNxtFreePosInQueue = 1
            .lNxtItemToTakeFromQueue = 1
            
            ' All to 'idle'
            SetStateTo PP_IDLE
            ClearError
            
        End If
    End With
    
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("MaxThreads", VAR.lMaxThreads, DEFAULT_MAXTHREADS)
        Call .WriteProperty("QueueSize", VAR.lQueueSize, DEFAULT_QUEUESIZE)
        Call .WriteProperty("Timeout", VAR.lTimeout, DEFAULT_TIMEOUT)
    End With
    
End Sub





' ******************************
' *       PUBLIC METHODS       *
' ******************************

Public Function About()
Attribute About.VB_Description = "Gives some information to the ParaPing control in a message box."
Attribute About.VB_UserMemId = -552
    ' Show a little 'About this control' message box

    MsgBox "ParaPing V." & App.Major & "." & App.Minor & App.Revision & _
            " - June 2005 by Light Templer", vbInformation + vbOKOnly, " About  'ParaPing'"

End Function


Public Function AddHost(sIPadress As String, Optional lID As Long = 0) As Boolean
Attribute AddHost.VB_UserMemId = 0
    ' Adds an IP adress of a destination host and an (optinal, user random defined)
    ' ID number to the ring buffer. From there its taken to create a 'thread' to
    ' check up or down by PING.
    ' If current state isn't 'Disabled' adding a new host starts checking!
    '
    ' RESULT:  TRUE,  if adding to queue was successful.
    '          FALSE, if no more empty space in queue to hold the adress or
    '                 IP adress has invalid format.
    
    On Local Error GoTo AddHost_Error
    
    ClearError
    
    With VAR
        ' Is there a free slot to hold another IP adress?
        If .lWaitingInQueue >= .lQueueSize Then
            RaiseError PP_ERR_BUFFER_OVERFLOW, "Buffer limit (" & .lQueueSize & ") reached. Cannot add more IP adresses!"
        
            Exit Function
        End If
        
        ' IP adress must be valid
        If IsValidIP(sIPadress) = False Then
            RaiseError PP_ERR_WRONG_PARAMETER, "Not a valid IP adress. ID= " & lID & " / IP= '" & sIPadress & "'"
            
            Exit Function
        End If
        
        ' Put the new entry into the ringbuffer
        With arrQueueDestHosts(.lNxtFreePosInQueue)
            .sIPadr = sIPadress
            .lID = lID
        End With
        .lWaitingInQueue = .lWaitingInQueue + 1
        
        ' Increment buffer pointer. If upper border reached restart from slot no 1
        .lNxtFreePosInQueue = .lNxtFreePosInQueue + 1
        If .lNxtFreePosInQueue > .lQueueSize Then
            .lNxtFreePosInQueue = 1
        End If
        
        ' Enable pinging to hosts if control is in enabled state
        If .flgEnabled = True And tmrSchedule.Enabled = False Then
            tmrSchedule.Enabled = True
        End If
        
    End With
    
    ' Success!
    AddHost = True

   
    Exit Function

AddHost_Error:

    RaiseError PP_ERR_GENERAL_ERROR, "[" & Err.Number & "] - '" & Err.Description & "' in AddHost() of 'ucParaPing'"
    
End Function


Public Sub Enable()
    ' Start checking hosts in queue until queue is empty and stay on 'enabled' state
    ' when all checks are done. In 'enabled' state adding a new host with 'AddHost'
    ' immediatly starts checking again!

    VAR.flgEnabled = True
    If tmrSchedule.Enabled = False Then
        tmrSchedule.Enabled = True
    End If
    ClearError
    
End Sub


Public Sub ClearQueue()
    ' Host entries in queue will be cleared. (In fact just a reset for the pointers,
    ' no need to run through the whole array and clear every item.
    ' Running checks (current 'threads') WILL BE raised with success/fail
    ' until all open Pings are done.
    
    With VAR
        .lWaitingInQueue = 0
        .lNxtFreePosInQueue = 1
        .lNxtItemToTakeFromQueue = 1
    End With
    ClearError

End Sub


Public Sub Disable()
    ' Don't check hosts from queue anymore.
    ' Host entries remains in queue waiting for a call to 'Enable' .
    ' Running checks (current threads) ARE raised with success/fail until
    ' all open threads are done.

    SetStateTo PP_DISABLED
    VAR.flgEnabled = False
    If VAR.lOpenPings = 0 Then
        ' No more checking the lists
        tmrSchedule.Enabled = False
    End If
    
End Sub


Public Sub Abort()
    ' Don't check hosts from queue anymore.
    ' Host entries remains in queue.
    ' Running checks (current threads) WILL NOT BE raised anymore.
    ' Thread list will be cleared.
    ' 'State' is going to be 'disabled'
    
    VAR.lOpenPings = 0
    SetStateTo IIf(VAR.lOpenPings > 0, PP_ABORT_PENDING, PP_DISABLED)
    VAR.flgEnabled = False

End Sub


Public Sub FreeResources()
    ' Unload no more needed elements from WinSock control array
    ' If you worked with lot of 'threads' (more than 20?) this could be a good
    ' idea to save resources.
    
    Dim i As Long

    If VAR.State <> PP_IDLE And wsckArrPing.Count > 1 Then
        RaiseError PP_ERR_NOT_IDLE, "IDLE state needed to unload no more needed elements from WinSock control array!"
    
        Exit Sub
    End If
    
    ' Top down unloading the control array elements up to the first which remains always.
    For i = wsckArrPing.Count To 2 Step -1
        Unload wsckArrPing(i)
    Next i
    
End Sub




' ******************************
' *   PRIVATE SUBS/FUNCTIONS   *
' ******************************
Private Sub RaiseError(Errorcode As enErrorCodes, sErr As String)
    ' Centralized for easy changes/additions
    
    VAR.sLastErr = sErr
    VAR.LastErrorCode = Errorcode
    
    RaiseEvent Error(Errorcode, sErr)

End Sub


Private Sub ClearError()

    VAR.sLastErr = ""
    VAR.LastErrorCode = PP_ERR_NO_ERROR
    
End Sub

Private Sub SetStateTo(NewState As enState)
    ' Centralized for easy changes/additions
        
    Dim sState As String
    
    VAR.State = NewState
    
    sState = Choose(NewState + 1, "Idle", "Busy", "Disabled", "AbortPending")
    RaiseEvent StateChanged(NewState, sState)
    
End Sub

Private Sub wsckArrPing_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    ' PONG! We just get an answer from the destination host so we save this result quickly!
    ' Rasing this result to the user is a job of the timer loop "tmrSchedule_Timer()".
    
    arrThreads(Index).flgPong = True
    
End Sub

Private Sub wsckArrPing_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' Give the error to the user
    
    RaiseError PP_ERR_WINSOCK, "WinSock control pinging to '" & arrThreads(Index).sIPadr & "' gots error [" & Description & "]"

End Sub


Private Sub tmrSchedule_Timer()
    ' Main action starts here:
    '
    ' When enabled and there are hosts to check in queue create 'threads'
    ' to check the destination hosts by using a WinSock control array.
    ' 'Ping' is realized by opening a connection to ICMP (Ping) port 7.


    Dim i               As Long
    Dim lStartSearch    As Long
    Dim lIndexFree      As Long
    Dim EmptySlot       As tpThread
    
    
    On Local Error GoTo tmrSchedule_Timer_Error


    With VAR
    
        ' === PART 1 :     Handle running 'threads' first to get free slots for new 'threads'
        If .lOpenPings > 0 Then
            For i = 1 To wsckArrPing.Count                                          ' Check allocated slots (1..n)
                With arrThreads(i)
                    If .flgPong = True Then                                         ' Did we got an answer to the Ping call?
                        RaiseEvent Pong(.sIPadr, .lID, True)                        ' Let the world know we got an answer!
                        Let arrThreads(i) = EmptySlot                               ' Free thread
                        VAR.lOpenPings = VAR.lOpenPings - 1                         ' One less
                        
                    ElseIf Len(.sIPadr) Then                                        ' Used slot?
                        If .lStartTime + VAR.lTimeout < Int(Timer()) Then           ' Is 'thread' timed out?
                            RaiseEvent Pong(.sIPadr, .lID, False)                   ' Let the world know we didn't got an answer!
                            Let arrThreads(i) = EmptySlot                           ' Free thread
                            VAR.lOpenPings = VAR.lOpenPings - 1                     ' One less
                            
                        End If
                    
                    ' Else   Do nothing and go on waiting
                    
                    End If
                End With
                
                ' Abort check if no more runnings or user called Abort()
                If .lOpenPings = 0 Then
                    
                    Exit For
                End If
                
            Next i
            
            ' User has disabled and we are not waiting for results of running Pings anymore
            If .lOpenPings = 0 And VAR.State = PP_DISABLED Then
                ' No more checking the lists
                tmrSchedule.Enabled = False
                
                Exit Sub
            End If
                
        ElseIf .lWaitingInQueue = 0 Then            ' All is done -> stop timer!
            ' No more checking the lists
            tmrSchedule.Enabled = False
                
            Exit Sub
            
        End If
        
        ' Abort request by user ?
        If VAR.State = PP_ABORT_PENDING Then
            
            ' Clear 'thread' list
            For i = 1 To wsckArrPing.Count
                Let arrThreads(i) = EmptySlot
            Next i
            
            ' Don' check lists anymore
            tmrSchedule.Enabled = False
            
            ' Set new state
            SetStateTo PP_DISABLED
            
            Exit Sub
        End If
        
        DoEvents
        
        
        
        ' === PART 2 :     Create new 'threads' if entries are waiting in queue
        lStartSearch = 1                                                            ' Start index for search in "thread"-list
        Do While .lWaitingInQueue > 0 And _
                .lOpenPings < .lMaxThreads And _
                VAR.flgEnabled = True                                               ' Entries are waiting, we have free slots and
                                                                                    ' control is enabled
            SetStateTo PP_BUSY
            
            ' Search for a free slot within current lists dimension
            lIndexFree = 0                                                          ' 0 = Not found a free slot
            For i = lStartSearch To wsckArrPing.Count
                If Len(arrThreads(i).sIPadr) = 0 Then
                    lIndexFree = i                                                  ' Found a free slot
                    lStartSearch = lIndexFree + 1
                    
                    Exit For
                End If
            Next i
            
            ' If we still need a free slot
            If lIndexFree = 0 And wsckArrPing.Count < .lMaxThreads Then             ' Not found a free slot, but space for a new one
                lIndexFree = wsckArrPing.Count + 1
                Load wsckArrPing(lIndexFree)                                        ' Load a new WinSock control in control array
                lStartSearch = lIndexFree + 1                                       ' In this way we don't search for a free slot
            End If                                                                  ' on the next cycle of this big loop.
            
            
            ' If we still don't have a free slot so we must abort for this time
            If lIndexFree = 0 Then
            
                Exit Do
            End If

            ' Now all its done for the Ping
            With arrThreads(lIndexFree)
                .sIPadr = arrQueueDestHosts(VAR.lNxtItemToTakeFromQueue).sIPadr     ' Take next item from ringbuffer and
                .lID = arrQueueDestHosts(VAR.lNxtItemToTakeFromQueue).lID           ' put it into 'thread' list
                .lStartTime = Int(Timer())
                .flgPong = False
                            
                wsckArrPing(lIndexFree).RemoteHost = .sIPadr                        ' HERE we "PING" to the dest host just by opening
                wsckArrPing(lIndexFree).SendData "<PONG this!>"                     ' a connection to port 7 and sending a string.
            End With
            
            ' Handle ringbuffer
            .lOpenPings = .lOpenPings + 1                                           ' One more open Pings we wait for
            .lWaitingInQueue = .lWaitingInQueue - 1                                 ' One less in queue to do
            If .lWaitingInQueue > 0 Then
                .lNxtItemToTakeFromQueue = .lNxtItemToTakeFromQueue + 1
                If .lNxtItemToTakeFromQueue > .lQueueSize Then
                    .lNxtItemToTakeFromQueue = 1
                End If
                
            Else
                ' Reset pointers to start when buffer empty
                .lNxtFreePosInQueue = 1
                .lNxtItemToTakeFromQueue = 1
                
            End If
        Loop
        
        ' Ready
        If .lOpenPings = 0 Then
            SetStateTo PP_IDLE
        End If
        
    End With
   
    Exit Sub

tmrSchedule_Timer_Error:

    RaiseEvent Error(PP_ERR_GENERAL_ERROR, "Error " & Err.Number & " [" & Err.Description & "] in sub tmrSchedule_Timer() / 'ucParaPing'")

End Sub


Private Function IsValidIP(sIPadress As String) As Boolean
    ' My solution to the "IsValidIP()" theme ;-)
    ' Optimized for speed and readability.
    ' Ignores the 'logical' checks which are related to subnet mask.
    '
    ' 6/6/2005 - Light Templer
    
    Dim i       As Long
    Dim lLen    As Long
    Dim lDigit  As Long
    Dim varArr  As Variant
    
    
    
    ' Min/max length
    lLen = Len(sIPadress)
    If lLen < 7 Or lLen > 15 Then Exit Function
    
    ' Valid chars only.
    For i = 1 To lLen
        If InStr(".0123456789", Mid$(sIPadress, i, 1)) = 0 Then Exit Function
    Next i
    
    ' 3 dots
    varArr = Split(sIPadress, ".")
    If UBound(varArr) <> 3 Then Exit Function
    
    ' Check all 4 entries
    For i = 0 To 3
    
        ' No empty entry
        If Len(varArr(i)) = 0 Then Exit Function
    
        ' Max valid value
        lDigit = Val(varArr(i))
        If lDigit > 255 Then Exit Function
        
        ' Special check for first entry: 0 and 255 not allowed
        If i = 0 And (lDigit = 0 Or lDigit = 255) Then Exit Function
                
    Next i
    
    IsValidIP = True

End Function



' ******************************
' *         PROPERTIES         *
' ******************************

Public Property Get PingsEnabled() As Boolean
    
    PingsEnabled = VAR.flgEnabled
    
End Property


Public Property Get State() As enState
Attribute State.VB_Description = "Idle or running Pings?"
Attribute State.VB_MemberFlags = "400"
    
    State = VAR.State
    
End Property


Public Property Get OpenPings() As Long
Attribute OpenPings.VB_Description = "How many open Pings are running?"
    
    OpenPings = VAR.lOpenPings

End Property

Public Property Get LastErrorMsg() As String
    
    LastErrorMsg = VAR.sLastErr

End Property

Public Property Get MaxThreads() As Long
Attribute MaxThreads.VB_Description = "How many parallel threads allowed?"
    
    MaxThreads = VAR.lMaxThreads

End Property


Public Property Let MaxThreads(ByVal lNew_MaxThreads As Long)
    ' Every additional 'thread' loads when used another Winsock control,
    ' so be careful with this resource ... ;-)
    
    If VAR.State <> PP_IDLE Then
        RaiseError PP_ERR_NOT_IDLE, "IDLE state needed to change max number of threads!"
    
        Exit Property
    End If
    
    If lNew_MaxThreads > 0 And lNew_MaxThreads <= THREADLIMIT Then
        VAR.lMaxThreads = lNew_MaxThreads
        ReDim arrThreads(1 To lNew_MaxThreads)
        PropertyChanged "MaxThreads"
        ClearError
    Else
        RaiseError PP_ERR_WRONG_PARAMETER, "Invalid parameter for 'MaxThreads' (" & lNew_MaxThreads & ") Valid is 1 to " & THREADLIMIT & "."
    End If
    
End Property


Public Property Get QueueSize() As Long
Attribute QueueSize.VB_Description = "Max number of stacked requests allowed to wait for a PING"
    
    QueueSize = VAR.lQueueSize

End Property

Public Property Let QueueSize(ByVal lNew_QueueSize As Long)
    ' The queue is an array build of 'tpDestHost', one long and one
    ' string per entry. Even for larger LANs with 5000 machines this shouldn't
    ' be a problem of available memory. Anyway, to recycle no more used slots
    ' this queue is organized as a ringbuffer and starts adding from beginning
    ' using resolved (free) slots when the end is reached.
    
    If VAR.State <> PP_IDLE Then
        RaiseError PP_ERR_NOT_IDLE, "IDLE state needed to change queue size!"
    
        Exit Property
    End If
    
    If lNew_QueueSize > 0 And lNew_QueueSize <= QUEUELIMIT Then
        VAR.lQueueSize = lNew_QueueSize
        ReDim arrQueueDestHosts(1 To lNew_QueueSize)
        PropertyChanged "QueueSize"
        ClearError
    Else
        RaiseError PP_ERR_WRONG_PARAMETER, "Invalid parameter for 'QueueSize' (" & lNew_QueueSize & ") Valid is 1 to " & _
                QUEUELIMIT & "."
    End If
    
End Property


Public Property Get WaitingInQueue() As Long
Attribute WaitingInQueue.VB_MemberFlags = "400"
    
    WaitingInQueue = VAR.lWaitingInQueue

End Property


Public Property Get Timeout() As Long
Attribute Timeout.VB_Description = "How many seconds to wait for an answer to a Ping call."
    
    Timeout = VAR.lTimeout

End Property

Public Property Let Timeout(ByVal lNew_Timeout As Long)
    ' If timeout is reached for a destination host the request will
    ' resolved as 'not reached'.
    
    If lNew_Timeout > 0 And lNew_Timeout <= MAXTIMEOUT Then
        VAR.lTimeout = lNew_Timeout
        PropertyChanged "QueueSize"
    Else
        RaiseError PP_ERR_WRONG_PARAMETER, "Invalid value for new timeout (" & lNew_Timeout & ")!  Must be from 1 to " & MAXTIMEOUT
    End If
        
End Property


Public Property Get WinSockArrSize() As Long
    ' How many elements of the WinSock control array are currently loaded
    
    WinSockArrSize = wsckArrPing.Count

End Property


' #*#
