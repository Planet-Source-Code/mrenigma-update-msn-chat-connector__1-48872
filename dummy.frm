VERSION 5.00
Object = "{3A7D68FE-BA27-4D2F-8980-E799C6D0D804}#1.3#0"; "MSNConnector.ocx"
Begin VB.Form Form1 
   BackColor       =   &H008080FF&
   Caption         =   "Form1"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   Begin MSNConnect.ChatConnectOCX MSN 
      Left            =   12210
      Top             =   330
      _ExtentX        =   1270
      _ExtentY        =   1270
      OCXClsID        =   "ECCDBA05-B58F-4509-AE26-CF47B2FFC3FE"
      AutoJoinRoom    =   -1  'True
      UseMSNChatXOCX  =   -1  'True
   End
   Begin VB.CheckBox chkHEX 
      Caption         =   "HEX"
      Height          =   285
      Left            =   4170
      TabIndex        =   10
      Top             =   540
      Width           =   1995
   End
   Begin VB.TextBox txtHEX 
      Height          =   345
      Left            =   3960
      TabIndex        =   9
      Text            =   "495243446F6D696E61746F725854"
      Top             =   90
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Access List"
      Height          =   645
      Left            =   9030
      TabIndex        =   8
      Top             =   240
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Passport"
      Height          =   375
      Left            =   9900
      TabIndex        =   7
      Top             =   150
      Width           =   2085
   End
   Begin VB.TextBox txtAccess 
      Height          =   2730
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   6120
      Width           =   14130
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Mass Kick"
      Height          =   645
      Left            =   8160
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create"
      Height          =   555
      Left            =   6660
      TabIndex        =   4
      Top             =   225
      Width           =   1410
   End
   Begin VB.TextBox txtNickName 
      Height          =   330
      Left            =   1410
      TabIndex        =   3
      Text            =   "AG-XTOCXTester"
      Top             =   390
      Width           =   2490
   End
   Begin VB.TextBox txtRoom 
      Height          =   330
      Left            =   1410
      TabIndex        =   2
      Text            =   "IRCDominatorXT"
      Top             =   90
      Width           =   2535
   End
   Begin VB.TextBox txtIN 
      Height          =   4830
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1200
      Width           =   14100
   End
   Begin VB.CommandButton command1 
      Caption         =   "Connect"
      Height          =   465
      Left            =   300
      TabIndex        =   0
      Top             =   390
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' old clisid 7a32634b-029c-4836-a023-528983982a49
' old dl ver 7,00,0206,0401
' New clsid F58E1CEF-A068-4c15-BA5E-587CAF3EE8C6
' New Ver 8,00,0210,2201


Private Sub Command1_Click()
Dim iniFile As New cIniFile
      If MSN.isConnected = True Then
         MSN.Disconnect
         Exit Sub
      End If

      txtIN = ""
      ' txtOut = ""
      MSN.RoomName = Me.txtRoom
      MSN.NickName = Me.txtNickName
      MSN.HexRoom = False
      If chkHEX Then
         MSN.RoomName = Me.txtHEX
         MSN.HexRoom = True
      End If
      If Me.Check1.Value Then
      
         iniFile.Path = "D:\VBstuff\IRCDominator\XT\IRCDominator.dat"
         iniFile.Section = "Passport"
         iniFile.Key = "PassportProfile"
         MSN.PassportProfile = iniFile.Value
         ' MSN.PassportProfile = TstGate
         iniFile.Key = "PassportTicket"
         MSN.PassportTicket = iniFile.Value
         iniFile.Key = "Cookie"
         MSN.MSNREGCookie = iniFile.Value
         MSN.UsePassPort = True
         MSN.ProfileIcon = 99
         MSN.AutoJoinRoom = True
      
      End If
      MSN.MSNReg = True
      MSN.UseRoomPassWord = True
      MSN.RoomPassword = "XT"
      
      MSN.Connect
      
      
      
End Sub
Public Function TstGate() As String
      TstGate = Chr$(Int(Rnd * 25) + 65) + Chr$(Int(Rnd * 25) + 65) + Chr$(Int(Rnd * 25) + 65) + Chr$(Int(Rnd * 25) + 65) + Chr$(Int(Rnd * 25) + 65) + Chr$(Int(Rnd * 25) + 65) + Chr$(Int(Rnd * 25) + 65) + Chr$(Int(Rnd * 25) + 65) + Chr$(Int(Rnd * 25) + 65) + Chr$(Int(Rnd * 25) + 65) + Chr$(Int(Rnd * 25) + 65) + Chr$(Int(Rnd * 25) + 65) + Chr$(Int(Rnd * 25) + 65) + Chr$(Int(Rnd * 25) + 65) + Chr$(Int(Rnd * 25) + 65) + Chr$(Int(Rnd * 25) + 65)
End Function

Private Sub Command2_Click()
Stop
      MSN.SetRoomProp SETPROPLag, 1
      
      MSN.SetRoomMode MSNMODEGuestBans_OFF
      MSN.SendRAW "access %#ircdominatorxt2 CLEAR"


End Sub

Private Sub Command3_Click()
      ' Stop

      ' Dim tFont As New MSNFontInfo
      '
      ' Dim i As Integer
      '
      '
      ' Me.txtAccess = ""
      ' For i = 1 To MSN.CurrentUsersList.Count
      ' Me.txtAccess = Me.txtAccess & vbCrLf & MSN.CurrentUsersList.Item(i).RealName & vbTab & MSN.CurrentUsersList.Item(i).DisplayName & "   " & MSN.CurrentUsersList.Item(i).UserJoinInfo & " Away is " & MSN.CurrentUsersList.Item(i).Away
      '
Dim i As Integer
Dim sTemp As String


      sTemp = ""
      For i = 1 To 15
         If i > 1 Then sTemp = sTemp & ","

         sTemp = sTemp & ">" & i
      Next


      MSN.PerformMassKick sTemp, "Test Mass Ban", True, True, 15, "15 Min Ban", False, True
  
      
End Sub



Private Sub Command4_Click()

      txtAccess = MSN.GetAccessList
  

End Sub

Private Sub MSN_IncommingMessage(MSNUser As MSNConnect.MSNUserObject, MessageData As String, OCXFontInfo As MSNConnect.MSNFontInfo)
      ' txtAccess = txtAccess & MessageData & vbCrLf
End Sub

Private Sub MSN_MSNError(ErrorCode As MSNConnect.erOCXErrors, ErrorDescription As String)
      MsgBox "Error " & ErrorCode & " - " & ErrorDescription
      ' Stop
End Sub

Private Sub MSN_RAWMessageIN(ByVal EventMessage As String)
EventMessage = ">" & EventMessage
      With txtIN
         If Right$(.Text, 2) <> vbCrLf Then
            .Text = .Text & vbCrLf
         End If
         .Text = .Text & EventMessage
         .SelStart = Len(.Text)
         If Len(.Text) > 32000 Then
            .Text = Right$(.Text, 1000)
         End If
      End With
End Sub

Private Sub MSN_RAWMessageOUT(ByVal EventMessage As String)
EventMessage = "<" & EventMessage
      With txtIN
         If Right$(.Text, 2) <> vbCrLf Then
            .Text = .Text & vbCrLf
         End If
         .Text = .Text & EventMessage
         .SelStart = Len(.Text)
         If Len(.Text) > 32000 Then
            .Text = Right$(.Text, 1000)
         End If
      End With

End Sub

Private Sub MSN_RoomModeChange(MSNUser As MSNConnect.MSNUserObject, NewMode As String)
      Me.txtAccess = Me.txtAccess & vbCrLf & "Nick = " & MSNUser.DisplayName & " (" & MSNUser.GateKeeperID & ") - Mode (" & NewMode & ")"
End Sub
Private Sub MSN_UserJoinedRoom(MSNUser As MSNConnect.MSNUserObject, UserStatus As String)
      ' Me.txtAccess = Me.txtAccess & vbCrLf & MSNUser.RealName & vbTab & MSNUser.DisplayName & vbTab & MSNUser.HostType
End Sub

Private Sub MSN_UserKicked(MSNHostUser As MSNConnect.MSNUserObject, KickedUser As MSNConnect.MSNUserObject, ReasonMessage As String)
      ' Debug.Print MSNHostUser.DisplayName & " Has Kicked " & KickedUser.DisplayName & " from the room :" & ReasonMessage

End Sub
Sub StringToHex()

End Sub

Private Sub MSN_UserModeChange(MSNUserChanging As MSNConnect.MSNUserObject, MSNUserChanged As MSNConnect.MSNUserObject, NewMode As String, ModeratedRoomChange As Boolean)
      Me.txtAccess = Me.txtAccess & vbCrLf & "Nick = " & MSNUserChanging.DisplayName & " (" & MSNUserChanged.GateKeeperID & ") - Mode (" & NewMode & ")" & " - " & ModeratedRoomChange
      Me.txtAccess.SelStart = Len(Me.txtAccess)

End Sub


Private Sub MSN_UserNameChange(MSNUser As MSNConnect.MSNUserObject, MSNOldUser As MSNConnect.MSNUserObject)
      Debug.Print MSNUser.RealName, MSNOldUser.RealName
End Sub
