VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "   Demo for ParaPING"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows-Standard
   Begin ParaPing.ucParaPing ucParaPing1 
      Left            =   3930
      Top             =   345
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.ListBox lbResults 
      Height          =   3375
      Left            =   255
      TabIndex        =   1
      Top             =   1215
      Width           =   6750
   End
   Begin VB.CommandButton btnSendOnePing 
      Caption         =   "Only ONE Ping!"
      Height          =   555
      Left            =   345
      TabIndex        =   0
      Top             =   285
      Width           =   1485
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Minimal demo for ParaPing

Private Sub btnSendOnePing_Click()

    lbResults.Clear
    DoEvents

    With ucParaPing1
        
        ' Add some ip adresses,
        .AddHost "132.112.151.11", 1
        .AddHost "132.112.150.11", 2
        .AddHost "132.112.151.12", 3
        .AddHost "10.54.248.221", 4
        
        ' enable control, and ...
        .Enable
        
    End With

End Sub


Private Sub ucParaPing1_Pong(sIPadr As String, lID As Long, flgSuccess As Boolean)
        
    ' you 'll get response!
    lbResults.AddItem " IP adress= " & sIPadr & ",  UsedDefined ID= " & lID & ",  Pong= " & flgSuccess
    
End Sub


Private Sub ucParaPing1_Error(Errorcode As enErrorCodes, sErr As String)
    
    ' Or just an error ;-)
    MsgBox "Error: " & sErr, vbExclamation, " Error from ParaPing modul:"
    
End Sub


' #*#
