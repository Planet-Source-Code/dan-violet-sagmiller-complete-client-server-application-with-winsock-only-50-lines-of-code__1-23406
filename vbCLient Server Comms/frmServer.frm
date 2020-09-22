VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   Caption         =   "Server"
   ClientHeight    =   1275
   ClientLeft      =   1635
   ClientTop       =   2235
   ClientWidth     =   1890
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   1890
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "4444"
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   1920
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local Port"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This whole form is only 22 lines of code.
'Copyright Haze Productions 2001, Dan Violet Sagmiller (hpdvs2)
'This code is free, and can be used however the user wishes.
'please leave a link to http://www.hazepro.com in any re-rendition of this code.


'**********************
'Start/Stop Button
'**********************

Private Sub Command1_Click()
'decide weather to start, stop or complain about the users input.
'check to see if the user click on the button when it said start.
If Command1.Caption = "&Start" Then
    'it did say start, is the port valid?
    If IsNumeric(Text1.Text) And CInt(Text1.Text) > 0 And CInt(Text1.Text) < 65536 Then
        ' it is, start the server
        StartServer
        'server is active
    Else
        'the port is not a valid port.
        MsgBox "Please choose a port number. (values generally range from 1 to 65535)"
    End If
Else
    'The command was not to start the server.
    StopServer
    'server is inactive

End If

End Sub





'******************
'Stop the server from running
'******************

Private Function StopServer()
'Close all active connections.
'close the server
Winsock1(0).Close
'create a counter
Dim i As Long
'discover how many winsock objects have been created
i = Winsock1.Count - 1 '(don't want to get rid of the server)
'Lopp them off in a loop
Do While i <> 0
    'is there someone connected?
    If Winsock1(i).State = StateConstants.sckOpen Then
        'there is, announce the Disconnection
        Winsock1(i).SendData "Sorry...  the server is now being shut down"
        'close this connection
        Winsock1(i).Close
    End If
    'remove the object from memory
    Unload Winsock1(i)
    'reduce the count by 1
    i = i - 1
Loop
'change the command buttons caption to reflect this change in state.
Command1.Caption = "&Start"
'Announce via the caption the current state of the server
Me.Caption = "Server Innactive"
End Function





'******************
'Start the server
'******************

Private Function StartServer()
'Set the winsock's local port
Winsock1(0).LocalPort = CInt(Text1.Text)
'Make the winsock Listen for connections.
Winsock1(0).Listen
'Annoucne via the caption that it is active and where.
Me.Caption = "Simple server Active on port " & Text1.Text & ", on IP " & Winsock1(0).LocalIP
'reset the name of the command button that started me.
Command1.Caption = "&Stop"
End Function






'******************
'a Request to connect
'******************

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'Create a temp Variable
Dim i As Long
'Find the last winsock and add 1 to it's position
i = Winsock1.UBound + 1
'use this new position to place the new winsock
Load Winsock1(i)
'Let the New winsock Take the connection
Winsock1(i).Accept requestID
'Formaly acknowledge the users presence
SendAll "User" & i & " has connect to the server."
End Sub





'********************
'Tend to arriving Data
'********************

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'because this is a simple server, we're only going to do standard error corection.
'we will assume that any connection (including telnet) can be made act
'as a client.  so we must be able to seperate messages properly.

'create a double variable
Dim l As String, m As String: l = Chr(10): m = Chr(13)
'create a temporary storage for the arriving string
Dim tmp As String
'create a temporary counter
Dim i As Long



'We will do this in 2 steps.
'step1: storing the data that arrived.
'Place the arriving data into the into the temporary string
Winsock1(Index).GetData tmp
'append the tag(winsock1(index).tag) with the tempdata
Winsock1(Index).Tag = Winsock1(Index).Tag & tmp



'This will help to begin again, and is called later in the program
ReprocessTagData:
'step2:figure out if the data should be sent to all yet.
'start by placing all the data from the winsock tag into the temp string
tmp = Winsock1(Index).Tag
'then loop through it to see if you find either Chr(10) or chr(13)
'(these characters ar used to signify the end of a line.
'Loop through each character of the string.
For i = 1 To Len(tmp)
    'Is the character (10) or (13)?
    If Mid(tmp, i, 1) = l Or Mid(tmp, i, 1) = m Then
        'Found a line break, send all previous text to everyone
        SendAll "User" & Index & ": " & Mid(tmp, 1, i + 1)
        ' returns text with line breaks and psuedo user names
        'reset the value of the winsock tag
        Winsock1(Index).Tag = Mid(tmp, i + 2) ' mid() will retrive all data after start point of no length is provided
        'restart this
        GoTo ReprocessTagData
    End If
    'All done
Next i
'even though there still may be data in the winsock tag, it would be with out a closing character and there fore assumed unfinished.
End Sub




'************************
'Send a message to every one
'************************

Private Function SendAll(ByVal sendString As String)
'create a counter
Dim i As Long
'discover how many winsock objects have been created
i = Winsock1.Count - 1 '(don't want to include the server)
'send them off in a loop
Do While i <> 0
    'is there someone connected?
    If Winsock1(i).State = StateConstants.sckConnected Or Winsock1(i).State = StateConstants.sckOpen Then
        'there is, send the message
        Winsock1(i).SendData sendString
    End If
    'reduce the count by 1
    i = i - 1
Loop
'message sent to every active connection
End Function










