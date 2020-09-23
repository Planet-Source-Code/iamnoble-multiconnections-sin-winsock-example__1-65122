VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock ClientSck 
      Index           =   0
      Left            =   240
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   5160
      Width           =   7335
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   255
         Left            =   4080
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdListen 
         Caption         =   "Listen"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Text            =   "5005"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Port"
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView lstConTable 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9128
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   776
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IP"
         Object.Width           =   4322
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Server Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Server Details"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "PC Details"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mnuMSG 
         Caption         =   "Send Message"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=====================================
'= Author : Evil_Inside              =
'= Website: www.DarkDevelopments.com =
'=====================================
'
'IF YOU USE THIS SOURCE GIVE CREDITS TO AUTHOR
'
'This is a Basic example on how to work with Winsock
'and Multiconnections

Dim Servers As Integer
Dim SckNumber As Integer
Private Sub ClientSck_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Servers = Servers + 1
Load ClientSck(Servers)
ClientSck(Servers).Accept requestID
ClientSck(Servers).SendData "SrvDtl"
End Sub

Private Sub ClientSck_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim SData As String
Dim Clean_Data As String
Dim Temp_Data As String
Dim Multi() As String

ClientSck(Index).GetData SData

Comamnd = Left(SData, 6)
All_Data = Right(SData, Len(SData) - 6)
Multi() = Split(All_Data, "[#]")

If Comamnd = "SrvDtl" Then
Set NewItem = lstConTable.ListItems.Add(, , Servers)
    NewItem.ListSubItems.Add , , Multi(1)
    NewItem.ListSubItems.Add , , Multi(2)
    NewItem.ListSubItems.Add , , Multi(3)
    NewItem.ListSubItems.Add , , Multi(4)
End If
End Sub

Private Sub cmdListen_Click()
ClientSck(Index).Close
ClientSck(Index).LocalPort = txtPort.Text
ClientSck(Index).Listen
cmdListen.Enabled = False
End Sub


Private Sub lstConTable_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 On Error GoTo NotThere
   
 If lstConTable.SelectedItem.Checked = True Then
 If Button = 2 Then
      PopupMenu mnuOptions
   End If
   End If
   
NotThere:
Exit Sub
End Sub

Private Sub mnuMSG_Click()
SckNumber = lstConTable.ListItems.Item(lstConTable.SelectedItem.Index)
Dim MSG As String
MSG = InputBox("Send Message to server", "Send Message", "Testing Connection")
ClientSck(SckNumber).SendData "SndMsg" & MSG
End Sub
