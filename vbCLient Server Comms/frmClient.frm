VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   Caption         =   "Client"
   ClientHeight    =   4680
   ClientLeft      =   1275
   ClientTop       =   4350
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   6585
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4080
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Connect"
      Height          =   255
      Left            =   5520
      TabIndex        =   3
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send..."
      Default         =   -1  'True
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   5415
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Port"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Address"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   735
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This whole form is only 28 lines of code.
'Copyright Haze Productions 2001, Dan Violet Sagmiller (hpdvs2)
'This code is free, and can be used however the user wishes.
'please leave a link to http://www.hazepro.com in any re-rendition of this code.



'******************
'Attemp to connect to the specified server
'******************

Private Sub Command2_Click()
'if currently attached to server, break connection
Winsock1(0).Close
'set remoteport
Winsock1(0).RemotePort = CInt(Text3.Text)
'set remotehost
Winsock1(0).RemoteHost = Text2.Text
'let winsock discover it's own suitable local port by giving it a null port.
Winsock1(0).LocalPort = 0
'Attempt to Connect
Winsock1(0).Connect
End Sub






'****************
'To handle winsock Errors
'****************
Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'formally announce errors
MsgBox "Err #" & Number & ": " & Description
'the errors should only be Connection failures and Unexpected server disconnects
End Sub





'***********************
'Send Messages
'***********************

Private Sub Command1_Click()
'check to see if this is connected.
If Winsock1(0).State = StateConstants.sckOpen Or Winsock1(0).State = StateConstants.sckConnected Then
    'Connected, Sending the text
    Winsock1(0).SendData Text1.Text & Chr(10) & Chr(13)
    'select the text box and give it focus again.
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.SetFocus
Else
    'if it wasn't an open connection then announce it.
    'MsgBox "Not Currently Connected to any server"
End If
End Sub





'*************************
'Display Server Communications
'*************************

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'one major errors in simple programs like this is the poor use of list boxes.
'List boxes have a limit of 32767 items.
'the program will error out if it reaches that.
'if the line count is that high then we will delete the oldest item in it before adding
'see if the list is higher than out limit
If List1.ListCount > 30000 Then
    'delete the first list item to make space for the new one
    List1.RemoveItem (0)
End If
'The server is assumed to have sent a complete line.
'create a temp string to put the arriving data
Dim tmp As String
'place the arrving data in the temp string
Winsock1(0).GetData tmp
'we'll just put any given text in it's own spot
List1.AddItem Mid(tmp, 1, Len(tmp) - 2) ' get rid of closing lines.

'now for a little repair for ease of use.
'by default when an item is added to the list it will add it to the bottom,
'while keeping the visibility at the top.  if you get more than a few
'messages, you'll have to keep scrolling down to see every message.
'the fix is to reset the viewable position to the last member of the list
List1.ListIndex = List1.ListCount - 1 '(by default the list counting starts @ 0)

End Sub


