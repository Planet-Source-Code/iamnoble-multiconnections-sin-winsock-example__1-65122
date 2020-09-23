VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Text            =   "5005"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   480
      Top             =   0
   End
   Begin MSWinsockLib.Winsock ServerSck 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Connect to IP:"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1215
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

Private Sub ServerSck_DataArrival(ByVal bytesTotal As Long)
Dim SData As String
Dim All_Data As String
Dim Comamnd As String
Dim Multi() As String
Dim Temp_String As String

ServerSck.GetData SData

Comamnd = Left(SData, 6)
All_Data = Right(SData, Len(SData) - 6)
Multi() = Split(All_Data, "[#]")

If Comamnd = "SrvDtl" Then
Temp_String = "SrvDtl"
Temp_String = Temp_String & "[#]" & ServerSck.LocalIP
Temp_String = Temp_String & "[#]" & "SIN Example"
Temp_String = Temp_String & "[#]" & "No Details"
Temp_String = Temp_String & "[#]" & "Evil_Inside"

ServerSck.SendData Temp_String
End If

If Comamnd = "SndMsg" Then
MsgBox All_Data
End If
End Sub

Private Sub Timer1_Timer()
If Not ServerSck.State = sckConnected Then
    ServerSck.Close
    ServerSck.Connect txtIP.Text, txtPort.Text
Else
    Exit Sub
End If
End Sub
