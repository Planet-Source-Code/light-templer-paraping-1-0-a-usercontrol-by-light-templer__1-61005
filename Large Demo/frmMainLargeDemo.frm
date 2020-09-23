VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMainLargeDemo 
   BackColor       =   &H00F0E8DF&
   Caption         =   "  A  Demo  for  my   'ParaPing'   userControl   -   Light Templer"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMainLargeDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8280
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainLargeDemo.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainLargeDemo.frx":28FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainLargeDemo.frx":2A56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtMsg 
      BackColor       =   &H00FAD3BC&
      CausesValidation=   0   'False
      Height          =   1350
      Left            =   300
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   5355
      Width           =   8520
   End
   Begin VB.CommandButton btnParaPing 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Check them all with ParaPing !"
      Height          =   615
      Left            =   3045
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   195
      Width           =   3075
   End
   Begin MSComctlLib.ListView lvComps 
      Height          =   3480
      Left            =   285
      TabIndex        =   1
      Top             =   1575
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   6138
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   " Machine"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   " IP adress "
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   " Number "
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   " Alive "
         Object.Width           =   1764
      EndProperty
   End
   Begin DemoForParaPing.ucParaPing ucParaPing1 
      Left            =   7290
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      MaxThreads      =   50
      QueueSize       =   1000
      Timeout         =   2
   End
   Begin VB.CommandButton btnEnumComps 
      BackColor       =   &H00FAD3BC&
      Caption         =   "Enum Comps in network"
      Height          =   615
      Left            =   285
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   180
      Width           =   2340
   End
   Begin VB.Label lblChecked 
      Appearance      =   0  '2D
      BackColor       =   &H00C9FDFE&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   " 0"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   7290
      TabIndex        =   9
      Top             =   1035
      Width           =   1530
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Checked"
      Height          =   285
      Index           =   3
      Left            =   6375
      TabIndex        =   8
      Top             =   1050
      Width           =   810
   End
   Begin VB.Label lblMachines 
      BackStyle       =   0  'Transparent
      Caption         =   "Machines"
      Height          =   285
      Left            =   465
      TabIndex        =   7
      Top             =   1320
      Width           =   1875
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Messages"
      Height          =   285
      Index           =   1
      Left            =   405
      TabIndex        =   6
      Top             =   5100
      Width           =   960
   End
   Begin VB.Label lblState 
      Appearance      =   0  '2D
      BackColor       =   &H00C9FDFE&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   " Idle"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3720
      TabIndex        =   4
      Top             =   1035
      Width           =   2400
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   285
      Index           =   0
      Left            =   3150
      TabIndex        =   3
      Top             =   1050
      Width           =   570
   End
End
Attribute VB_Name = "frmMainLargeDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   Another (larger) demo for ParaPing
'


Option Explicit


Private lComps              As Long
Private lChecked            As Long
Private sngTimer            As Single
Private oDNS                As clsDNS
Private WithEvents oCOMPS   As clsEnumComps
Attribute oCOMPS.VB_VarHelpID = -1
'
'
'




' ===================================================================================
'
'                        Part 1:   Enumerate computers in Windows LAN
'
' ===================================================================================

Private Sub btnEnumComps_Click()
    ' START enumerating all windows computers in LAN
    ' (This isn't important for the ParaPing control function.)
    
    Screen.MousePointer = vbArrowHourglass
    lComps = 0
    lvComps.ListItems.Clear
    Set oDNS = New clsDNS
    Set oCOMPS = New clsEnumComps
    oCOMPS.EnumComps EC_CT_AllWindowsComps
    Set oCOMPS = Nothing
    Set oDNS = Nothing
    lblMachines.Caption = "Machines:  " & lvComps.ListItems.Count
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub oCOMPS_CompFound(sCompName As String)
    ' ADD found machines into listview
    ' (This isn't important for the ParaPing control function.)
    
    Dim oLstItem    As ListItem
    Dim sIPadress   As String
    
    lComps = lComps + 1
    sIPadress = oDNS.NameToAddress(sCompName)
    Set oLstItem = lvComps.ListItems.Add(, , " " + sCompName, , 3)
    oLstItem.SubItems(1) = sIPadress
    oLstItem.SubItems(2) = lComps
    DoEvents
    
End Sub



' ===================================================================================
'
'             Part 2:   Check 'Running' or 'Switched off' by using   -ParaPing-
'
' ===================================================================================

Private Sub btnParaPing_Click()
    ' START pinging to destination hosts

    Dim i   As Long

    With lvComps.ListItems
        
        ' Something to do?
        If .Count < 1 Then                                  ' No entries in listview (no hosts)? -> Leave!
        
            Exit Sub
        End If
        
        ' Prepare
        lblChecked.Caption = " 0"
        lChecked = 0
        sngTimer = Timer                                    ' Save start time for statistic
        
        ' Start checking with ParaPing
        ucParaPing1.QueueSize = .Count                      ' This way we don't even need the ringbuffer feature
        ucParaPing1.Enable                                  ' Start when first destination host is added
        For i = 1 To .Count
            ucParaPing1.AddHost .item(i).SubItems(1), i     ' Add all IP adresses of the desination hosts from listview. We use
        Next i                                              ' the index into the listview as ID for the requests we give to ParaPing.
        
    End With

End Sub


Private Sub ucParaPing1_Pong(sIPadr As String, lID As Long, flgSuccess As Boolean)
    ' HANDLE RESULTS of check
    
    With lvComps.ListItems.item(lID)
        .SmallIcon = IIf(flgSuccess = True, 1, 2)
        .SubItems(3) = IIf(flgSuccess = True, "Alive!", "Dead")
    End With
    
    lChecked = lChecked + 1
    lblChecked = " " & lChecked
    
End Sub


Private Sub ucParaPing1_StateChanged(New_State As enState, sNewState As String)
    ' SHOW changed state
    
    lblState.Caption = " " + sNewState
    
    If New_State = PP_IDLE Then
        txtMsg.Text = txtMsg.Text & vbCrLf & "Checking done in about " & Format(Timer - sngTimer, "#.##") & " seconds." & vbCrLf
    End If
    
End Sub


Private Sub ucParaPing1_Error(Errorcode As enErrorCodes, sErr As String)
    ' SHOW error
    
    txtMsg.Text = txtMsg.Text & "ERROR:  " & sErr & vbCrLf & String$(60, "_") & vbCrLf
    
End Sub


' #*#
