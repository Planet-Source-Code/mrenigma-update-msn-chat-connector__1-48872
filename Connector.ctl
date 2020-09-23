VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl ChatConnectOCX 
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2595
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Connector.ctx":0000
   PropertyPages   =   "Connector.ctx":0ECA
   ScaleHeight     =   2760
   ScaleWidth      =   2595
   ToolboxBitmap   =   "Connector.ctx":0EE8
   Begin VB.Timer tmrJoinRoom 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   540
      Top             =   810
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   90
      Top             =   810
   End
   Begin MSWinsockLib.Winsock Svr2 
      Left            =   540
      Top             =   1290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Svr1 
      Left            =   90
      Top             =   1305
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WB 
      CausesValidation=   0   'False
      Height          =   1650
      Left            =   990
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   810
      ExtentX         =   1429
      ExtentY         =   2910
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "ChatConnectOCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Dim iPort As Long

Const m_def_RoomName = "TestRoom"
Const m_def_NickName = "NickName"
Const m_def_ServerAddress = "207.68.167.253"
Const m_def_Socket = 6667
Const m_def_HexRoom = False
Const m_def_OCXVersion = "8,00,0210,2201"
Const m_def_OCXClsID = "F58E1CEF-A068-4C15-BA5E-587CAF3EE8C6"
Const m_def_UsePassPort = False
Const m_def_MSNREGCookie = ""
Const m_def_PassportTicket = ""
Const m_def_PassportProfile = ""
Const m_def_AutoJoinRoom = False
Const m_def_RoomPassWord = ""
Const m_def_UseRoomPassWord = False
Const m_def_UseMSNChatXOCX = False
Const m_def_ServerConnectTimeOut = 5
Const m_def_RoomMode = ""
Const m_def_isJoinedRoom = False
Const m_def_UseSingleEvent = False


Dim sNamesList As String
Dim m_ActualNickName As String
Dim m_isConnected As Boolean
Dim m_RoomTopic As String
Dim m_RoomJoinMessage As String
Dim m_ServerConnectTimeOut As Integer
Dim sWhoIs As String
Dim sAccessList As String
' Dim sRoomMode As String
Dim m_RoomMode As String
Dim m_isJoinedRoom As Boolean
Dim m_UseSingleEvent As Boolean
Dim m_RoomName As String
Dim m_NickName As String
Dim m_ServerAddress As String
Dim m_Socket As Integer
Dim m_HexRoom As Boolean
Dim m_OCXVersion As String
Dim m_OCXClsID As String
Dim m_UsePassPort As Boolean
Dim m_MSNREGCookie As String
Dim m_PassportTicket As String
Dim m_PassportProfile As String
Dim bFirstConnect As Boolean
Dim m_AutoJoinRoom As Boolean
Dim m_UserCount As Integer
' Dim m_RoomProperties As Variant
Dim m_HostCount As Integer
Dim sJoinedRoom As String
Dim m_ActualRoomName As String
Dim m_RoomPassWord As String
Dim m_UseRoomPassWord As Boolean
Dim m_MSNChatHostName As String
Dim m_UseMSNChatXOCX As Boolean

' Event Declarations:
Event RAWMessageOUT(ByVal EventMessage As String)
Event RAWMessageIN(ByVal EventMessage As String)
Event UserJoinedRoom(MSNUser As MSNUserObject, UserStatus As String)
Event UserPartedRoom(MSNUser As MSNUserObject)
Event UserQuitRoom(MSNUser As MSNUserObject, QuitReason As String)
Event UserKilled(MSNUser As MSNUserObject, MSNSysop As MSNUserObject, KillReason As String)
Event UserAway(MSNUser As MSNUserObject)
Event UserUnAway(MSNUser As MSNUserObject)
Event UserKicked(MSNHostUser As MSNUserObject, KickedUser As MSNUserObject, ReasonMessage As String)
Event UserNameChange(MSNUser As MSNUserObject, MSNOldUser As MSNUserObject)
Event RoomModeChange(MSNUserChanging As MSNUserObject, NewMode As String)
Event UserModeChange(MSNUserChanging As MSNUserObject, MSNUserChanged As MSNUserObject, NewMode As String, ModeratedRoomChange As Boolean)
Event RoomKnock(NickName As String, GateKeeperID As String, KnockCode As String)
Event IncommingWhisper(MSNUser As MSNUserObject, MessageData As String, OCXFontInfo As MSNFontInfo)
Event IncommingMessage(MSNUser As MSNUserObject, MessageData As String, OCXFontInfo As MSNFontInfo)
Event IncommingNotice(MSNUser As MSNUserObject, MessageData As String)
Event IncommingAction(MSNUser As MSNUserObject, MessageData As String)
Event IncommingTimeRequest(MSNUser As MSNUserObject)
Event Connected()
Event HandShake()
Event Disconnected()
Event JoinedRoom()
Event LeftRoom()
Event MSNError(ErrorCode As erOCXErrors, ErrorDescription As String)
Event SelfBeenKicked(MSNHostUser As MSNUserObject, ReasonMessage As String)
Event PropertyChange(MSNUser As MSNUserObject, PropertyName As String, PropertyValue As String)
Event MSNEvent(EventName As String, MSNUser As MSNUserObject, IncommingSource As String)
Event OCXEvent(EventName As OCXEventsList, IncommingSource As String)
Public Enum OCXEventsList
   OCXPong = 1
   OCXJoinMessage = 2
   OCXOther = 3 ' "Other Inlisted Event"
End Enum
Public Enum erOCXErrors
   ERR_NickNameUsed = -10001
   ERR_InvalidNickName = -10002
   ERR_NoAccess = -10003
   ERR_AuthFailed = -10004
   ERR_RoomAlreadyExists = -10005
   ERR_RoomLimit = -10006
   ERR_TimeOut = -10007
   ERR_NoNickName = -10008
   ERR_InsufficientParams = -10009
   ERR_InvalidRoomName = -10010
   ERR_NotConected = -10011
   ERR_NotJoinedRoom = -10012
   ERR_AlreadyJoined = -10013
   ERR_AlreadyLeftRoom = -10014
   ERR_SockerInUse = -10015
   ERR_ChatOCXMissing = -10016
   ERR_ChatPatchedMissing = -10017
   ERR_FailedRetrieveAccess = -10018
   ERR_FailedRetrieveWhoIs = -10019
   ERR_FailedRetrievePUID = -1020
   ERR_FailedRetrieveProperty = -1021
   ERR_CannotSendToChannel = -1022
   ERR_InvalidValue = -19999
End Enum

Public Enum MSNSendMethod
   MSNPRIVMSG = 1
   MSNACTION = 2
   MSNNOTICE = 3
End Enum

Public Enum MSNRoomModes
   MSNMODESecret_ON = -1001
   MSNMODESecret_OFF = -1002
   MSNMODEInviteOnly_ON = -1003
   MSNMODEInviteOnly_OFF = -1004
   MSNMODEPrivate_ON = -1005
   MSNMODEPrivate_OFF = -1006
   MSNMODEGuestBans_ON = -1007
   MSNMODEGuestBans_OFF = -1008
   MSNMODENoGuestWhispers_on = -1009
   MSNMODENoGuestWhispers_OFF = -1010
   MSNMODEOnlyHostWhispers_ON = -1011
   MSNMODEOnlyHostWhispers_OFF = -1012
   MSNMODERoomLimit = -1013
   MSNMODEKnock_ON = -1014
   MSNMODEKnock_OFF = -1015
   MSNMODEModerated_On = -1016
   MSNMODEModerated_Off = -1017
   MSNMODEHiddenRoom_ON = -1018
   MSNMODEHiddenRoom_OFF = -1019
   MSNMODEOnlyHostChangeTopic_ON = -1020
   MSNMODEOnlyHostChangeTopic_OFF = -1021
End Enum
Public Enum TestOCXRoomModes
   testMODESecret_ON = -1001
   testMODEInviteOnly_ON = -1003
   testMODEPrivate_ON = -1005
   testMODENoGuestWhispers_ON = -1009
   testMODEOnlyHostWhispers_ON = -1011
   testMODEKnock_ON = -1014
   testMODEModerated_ON = -1016
   testMODEHiddenRoom_ON = -1018
   testMODEOnlyHostChangeTopic_ON = 1020
End Enum
Public Enum ePropCodes
   SETPROPOnJoin = 1001
   SETPROPOwnerKey = 1002
   SETPROPHostKey = 1003
   SETPROPOnPart = 1004
   SETPROPTopic = 1005
   SETPROPMemberKey = 1006
   SETPROPLag = 1007
   SETPROPLanguage = 1008
   SETPROPPics = 1009
   SETPROPClient = 1012
End Enum
Public Enum ePropCodesRead
   GETPROPOnJoin = 1001
   GETPROPOnPart = 1004
   GETPROPTopic = 1005
   GETPROPLag = 1007
   GETPROPLanguage = 1008
   GETPROPPics = 1009
   GETPROPName = 1010
   GETPROPCreation = 1011
   GETPROPClient = 1012
   GETPROPSubject = 1013
End Enum


Dim m_CurrentUsersList As MSNUserList
Dim m_MSNReg As Boolean
Dim sRoomProperty As String

Const m_def_MSNReg = False
Const m_def_ByPassSecurity = False
Const m_def_ProfileIcon = 0
Const m_def_CNick = ""
Const m_def_OverRide = False

Dim m_ByPassSecurity As Boolean
Dim iCurSecurity As Integer
Dim sSecurPath As String
Dim bCreateRoom As Boolean
Dim m_CreateModes As New MSNRoomCreationSettings
Dim m_ProfileIcon As Integer
Dim m_CNick As String
Dim m_OverRide As Boolean

Public Sub SendTimeRequest(Optional NickName As String)
Dim sTemp As String

      ' Send a Time Request
      If TestForError(True) = True Then Exit Sub
      If NickName = "" Then
         ' No NickName specified - raise error
         RaiseEvent MSNError(ERR_InvalidValue, "Invalid NickName")
      End If
      sTemp = sTemp & "PRIVMSG " & NickName & " :" & Chr(1) & "TIME" & Chr(1)
      SendServer2 sTemp
End Sub
Public Function RefreshRoomMode() As String
Dim sTemp As String
      ' Dim i As Long

      ' sRoomMode = ""
      If TestForError(False) = True Then Exit Function

      sTemp = "MODE %#" & FixSpaces(m_ActualRoomName, True)
      SendServer2 sTemp
End Function
Public Function GetAccessList() As String
Dim sTemp As String
Dim i As Long
Static lSize As Long

      If TestForError(True) = True Then Exit Function
      
      sAccessList = ""
      sTemp = "ACCESS %#" & m_ActualRoomName & " LIST"
      SendServer2 sTemp
      i = 0
      Do
         i = i + 1
         If i > 300000 Then
            RaiseEvent MSNError(ERR_FailedRetrieveAccess, "Access List was not retrieved in a timely manner")
            Exit Function
         End If
         If Locate(sAccessList, m_MSNChatHostName & " 805 ") Or Locate(sAccessList, m_MSNChatHostName & " 913 ") Then
            Exit Do
         End If
         If Len(sAccessList) > lSize Then
            ' Still loading
            i = 0
         End If
         lSize = Len(sAccessList)
         DoEvents
      Loop
      GetAccessList = sAccessList
      sAccessList = ""

End Function
Public Function GetWhoIs(sNickName As String) As String
Dim sTemp As String
Dim i As Long


      If TestForError(False) = True Then Exit Function

      sWhoIs = ""
      sTemp = "WHOIS " & sNickName
      SendServer2 sTemp
      
      i = 0
      Do
         i = i + 1
         If i > 300000 Then
            RaiseEvent MSNError(ERR_FailedRetrieveWhoIs, "Whois Information was not retrieved in a timely manner")
            Exit Function
         End If
         If Locate(sWhoIs, m_MSNChatHostName & " 318 ") Then
            Exit Do
         End If
         DoEvents
      Loop
      
      GetWhoIs = Replace(sWhoIs, ":" & m_MSNChatHostName & " ", "")
      sWhoIs = ""

End Function
Public Function GetGateKeeperID(NickName As String, bUpdateUser As Boolean) As String
Dim sTemp As String

      On Error Resume Next
      If TestForError(True) = True Then Exit Function

      sTemp = GetWhoIs(NickName)
         
      If sTemp <> "" Then
         GetGateKeeperID = Split(GetLine(sTemp, "311 "), " ")(3)
         If bUpdateUser = True Then
            m_CurrentUsersList.Item(NickName).GateKeeperID = GetGateKeeperID
         End If
         
      End If

End Function

Public Sub PerformMassKick(sNickNames As String, sMessage As String, bBan As Boolean, Optional bBanPassport As Boolean = False, Optional iTime As Integer = -1, Optional sTime As String = "", Optional bSilent As Boolean = False, Optional bBanBeforeKick As Boolean = True)
Dim sTemp As String
Dim sKickString As String
Dim sNewNick As String
Dim i As Integer
Dim iLoop As Long
Dim sBanString As String
Dim asNicks() As String
Dim sGate As String

      If TestForError(True) = True Then Exit Sub
      On Error GoTo Hell:

      sBanString = "(Access ban set for %s)"
      If sTime = "" Or iTime = -1 Then
         sTime = "15 Minutes"
         iTime = 15
      End If
      
      sNickNames = TestNick(sNickNames, True)
      
      If Right$(sNickNames, 1) = "," Then
         sNickNames = Mid$(sNickNames, 1, Len(sNickNames) - 1)
      End If
      
      asNicks = Split(sNickNames, ",")
      
      sKickString = "KICK %#" & m_ActualRoomName & " " & sNickNames & " : " & Replace(Trim(sMessage), vbCrLf, "")

      If bBan Then
      
         If bBanBeforeKick = False Then
            If Not (bSilent) Then
               sTemp = sKickString & " " & Replace(sBanString, "%s", sTime)
               SendServer2 sTemp
            End If
         End If
      
         If bBanPassport Then
            For i = 0 To UBound(asNicks)
               sNewNick = asNicks(i)
               
               sGate = m_CurrentUsersList.Item(sNewNick).GateKeeperID
               
               If sGate = "" Then
                  sTemp = GetWhoIs(sNewNick)
                  If sTemp <> "" Then
                     sNewNick = GetLine(sTemp, "311 ")
                     sGate = Split(sNewNick, " ")(3)
                  End If
               End If
               
               If bBanBeforeKick = True Then
                  sNewNick = "*!" & sGate & "*$*"
                  sTemp = "ACCESS %#" & m_ActualRoomName & " ADD DENY " & sNewNick & " " & iTime & " : " & sMessage & " "
                  sTemp = sTemp & Replace(sBanString, "%s", sTime)
                  sTemp = sTemp & " - " & asNicks(i)
                  SendServer2 sTemp
               End If
               
               iLoop = iLoop + 1
               If iLoop > 5 Then
                  For iLoop = 1 To 100000
                     DoEvents
                  Next
                  iLoop = 0
               End If
               
            Next
            If bBanBeforeKick = True Then
               If Not (bSilent) Then
                  sTemp = sKickString & " " & Replace(sBanString, "%s", sTime)
                  SendServer2 sTemp
               End If
            End If
         Else
            ' Ban Nick Name
            For i = 0 To UBound(asNicks)
               sNewNick = asNicks(i)

               sTemp = "ACCESS %#" & m_ActualRoomName & " ADD DENY " & sNewNick & " " & iTime & " : " & sMessage & " "
               sTemp = sTemp & Replace(sBanString, "%s", sTime)
               SendServer2 sTemp
                           
               iLoop = iLoop + 1
               If iLoop > 5 Then
                  For iLoop = 1 To 100000
                     DoEvents
                  Next
                  iLoop = 0
               End If
               
            Next

            If bBanBeforeKick = True Then
               If Not (bSilent) Then
                  sTemp = sKickString & " " & Replace(sBanString, "%s", sTime)
                  SendServer2 sTemp
               End If
            End If
         End If
      Else
         SendServer2 sKickString
         DoEvents
      End If
      Erase asNicks
      Exit Sub
Hell:

End Sub

Public Sub PerformKick(sNickName As String, sMessage As String, bBan As Boolean, Optional sGate As String = "", Optional bBanPassport As Boolean = False, Optional iTime As Integer = -1, Optional sTime As String = "", Optional bSilent As Boolean = False, Optional bBanBeforeKick As Boolean = True)
Dim sTemp As String
Dim sKickString As String
Dim sNewNick As String
Dim sBanString As String

      If TestForError(True) = True Then Exit Sub

      sBanString = "(Access ban set for %s)"
      If sTime = "" Or iTime = -1 Then
         sTime = "15 Minutes"
         iTime = 15
      End If
      
      sNickName = TestNick(sNickName, True)
      
      sKickString = "KICK %#" & m_ActualRoomName & " " & sNickName & " : " & Replace(Trim(sMessage), vbCrLf, "")

      If bBan Then
         If bBanPassport Then
            If sGate = "" Then
               sTemp = GetWhoIs(sNickName)
               If sTemp <> "" Then
                  sNewNick = GetLine(sTemp, "311 ")
                  sGate = Split(sNewNick, " ")(3)
               End If
            End If
            If bBanBeforeKick = False Then
               If Not (bSilent) Then
                  sTemp = sKickString & " " & Replace(sBanString, "%s", sTime)
                  SendServer2 sTemp
               End If
            Else
               sNewNick = "*!" & sGate & "*$*"
               sTemp = "ACCESS %#" & m_ActualRoomName & " ADD DENY " & sNewNick & " " & iTime & " : " & sMessage & " "
               sTemp = sTemp & Replace(sBanString, "%s", sTime)
               sTemp = sTemp & " - " & sNickName
               SendServer2 sTemp
            End If
            If bBanBeforeKick = True Then
               If Not (bSilent) Then
                  sTemp = sKickString & " " & Replace(sBanString, "%s", sTime)
                  SendServer2 sTemp
               End If
            End If
         Else
            ' Ban Nick Name
            If bBanBeforeKick = False Then
               If Not (bSilent) Then
                  sTemp = sKickString & " " & Replace(sBanString, "%s", sTime)
                  SendServer2 sTemp
               End If
            Else
               sTemp = "ACCESS %#" & m_ActualRoomName & " ADD DENY " & sNickName & " " & iTime & " : " & sMessage & " "
               sTemp = sTemp & Replace(sBanString, "%s", sTime)
               SendServer2 sTemp
            End If
            If bBanBeforeKick = True Then
               If Not (bSilent) Then
                  sTemp = sKickString & " " & Replace(sBanString, "%s", sTime)
                  SendServer2 sTemp
               End If
            End If
         End If
      Else
         SendServer2 sKickString
         DoEvents
      End If
End Sub
Public Sub SendWHISPER(sNickName As String, sMessage As String, tFontInfo As MSNFontInfo)
Dim sTemp As String
        
      If TestForError(True) = True Then Exit Sub
      sTemp = "WHISPER %#" & m_ActualRoomName
      sTemp = sTemp & " " & TestNick(sNickName, True)
      If tFontInfo.FontName <> "" Then
         sTemp = sTemp & " :" & Chr(1) & "S " & Chr(tFontInfo.MSNColour) & Chr(tFontInfo.FontStyle) & FixSpaces(tFontInfo.FontName, True) & ";0 " & sMessage & Chr(1)
         sTemp = Replace(sTemp, Chr(13), "\r")
         sTemp = Replace(sTemp, Chr(10), "\n")
         sTemp = Replace(sTemp, Chr(9), "\t")
      Else
         sTemp = sTemp & " :" & sMessage
      End If
      SendServer2 sTemp
End Sub

Public Sub SendMESSAGE(Method As MSNSendMethod, SendData As String, bAddRoomName As Boolean, tFontInfo As MSNFontInfo, Optional NickName As String = "")
Dim sTemp As String
      
      If TestForError(True) = True Then Exit Sub
      Select Case Method
         Case MSNACTION
            ' Send the data as an ACTION
            If bAddRoomName Then
               sTemp = "PRIVMSG %#" & FixSpaces(m_ActualRoomName, True)
            Else
               sTemp = "PRIVMSG"
            End If
            If NickName <> "" Then
               sTemp = sTemp & " " & NickName
            End If
            sTemp = sTemp & " :" & Chr(1) & "ACTION " & SendData & Chr(1)
         Case MSNPRIVMSG
            ' Send the data as a PRIVMSG
            If bAddRoomName Then
               sTemp = "PRIVMSG %#" & FixSpaces(m_ActualRoomName, True)
            Else
               sTemp = "PRIVMSG"
            End If
            If NickName <> "" Then
               sTemp = sTemp & " " & NickName
            End If
            sTemp = sTemp & " :"
            If tFontInfo.FontName <> "" Then
                
               ' if valid font info then use
               sTemp = sTemp & Chr(1) & "S " & Chr(tFontInfo.MSNColour) & Chr(tFontInfo.FontStyle) & FixSpaces(tFontInfo.FontName, True) & ";0 "
            End If
            sTemp = sTemp & SendData
            If tFontInfo.FontName <> "" Then
               sTemp = sTemp & Chr(1)
            End If
            sTemp = Replace(sTemp, Chr(13), "\r")
            sTemp = Replace(sTemp, Chr(10), "\n")
            sTemp = Replace(sTemp, Chr(9), "\t")
         Case MSNNOTICE
            sTemp = "NOTICE "
            If bAddRoomName Then
               sTemp = sTemp & "%#" & FixSpaces(m_ActualRoomName, True) & " "
               If NickName = "" Then
                  sTemp = sTemp & ":"
               End If
            End If
            If NickName <> "" Then
               sTemp = sTemp & NickName & " :"
            End If
            sTemp = sTemp & SendData
      End Select
      SendServer2 sTemp
End Sub
Public Sub SendRAW(SendData As String)
      If TestForError(True) = True Then Exit Sub
      SendServer2 SendData & vbCrLf
End Sub
Public Function TestRoomMode(RoomMode As TestOCXRoomModes) As Boolean
Dim sMode As String
Dim sTemp As String

      If TestForError(False) = True Then Exit Function
      TestRoomMode = False
      Select Case RoomMode
         Case testMODESecret_ON
            If Locate(m_RoomMode, "s") Then
               TestRoomMode = True
            End If
         Case testMODEInviteOnly_ON
            If Locate(m_RoomMode, "i") Then
               TestRoomMode = True
            End If
         Case testMODEPrivate_ON
            If Locate(m_RoomMode, "p") Then
               TestRoomMode = True
            End If
         Case testMODENoGuestWhispers_ON
            If Locate(m_RoomMode, "W") Then
               TestRoomMode = True
            End If
         Case testMODEOnlyHostWhispers_ON
            If Locate(m_RoomMode, "w") Then
               TestRoomMode = True
            End If
         Case testMODEKnock_ON
            If Locate(m_RoomMode, "u") Then
               TestRoomMode = True
            End If
         Case testMODEModerated_ON
            If Locate(m_RoomMode, "m") Then
               TestRoomMode = True
            End If
         Case testMODEHiddenRoom_ON
            If Locate(m_RoomMode, "h") Then
               TestRoomMode = True
            End If
         Case testMODEOnlyHostChangeTopic_ON
            If Locate(m_RoomMode, "t") Then
               TestRoomMode = True
            End If
      End Select
End Function

Private Sub tmrJoinRoom_Timer()
      tmrJoinRoom.Enabled = False
      JoinRoom
End Sub
Private Sub tmrTimeOut_Timer()
      tmrTimeOut.Enabled = False
      Disconnect
      RaiseEvent MSNError(ERR_TimeOut, "Timed out waiting to connect to Server")
End Sub
Private Sub UserControl_Paint()
      UserControl.BackColor = Ambient.BackColor
      UserControl.BackColor = UserControl.Parent.BackColor
End Sub
Private Sub UserControl_Resize()
      UserControl.Height = 720
      UserControl.Width = 720
End Sub
Private Function BindPort(iPort As Long) As Long
Dim iOldPort As Long

      ' On Error Resume Next
RepeatBind:
        
      Err.Clear
      iOldPort = iPort
        
      Svr1.Close
      Svr1.Bind iPort, "127.0.0.1"
      Svr1.LocalPort = iPort
      If Err.Number = "10048" Then
         Err.Clear
         iPort = iPort + 1
         GoTo RepeatBind
      End If
      BindPort = iPort
      If iOldPort <> iPort Then
         Debug.Print "Port " & iOldPort & " was already in use trying - now using port (" & iPort & ")"
      End If
End Function

Private Sub DoConnect()
Dim bHex As Boolean
Dim i As Integer
Dim SUrl As String
Dim sFileName As String
Dim sBaseName As String
Dim sCabPath As String

    
      Set m_CurrentUsersList = Nothing
    
      sBaseName = fso.GetBaseName(fso.GetTempName) & ".html"
      sFileName = fso.GetSpecialFolder(TemporaryFolder) & "\" & sBaseName
    
      tmrTimeOut.Enabled = False
      iPort = m_Socket
      Svr1.Close
      Svr2.Close
      Svr1.Close
      bFirstConnect = True
      On Error Resume Next
      iPort = BindPort(iPort)
      ' iPort = m_Socket
      If m_RoomName <> "" Then
         If m_UseMSNChatXOCX = True Then
            sCabPath = Encrypt1("†GG£òòwwwâwV&vVæGsâö'vòö6‡òÔ5ä4†G…â6&")
         Else
            sCabPath = ""
         End If
         If bCreateRoom = True Then
            Call BuildHTML(m_RoomName, m_NickName, m_ServerAddress, Str(iPort), m_OCXVersion, m_OCXClsID, m_UsePassPort, m_MSNREGCookie, m_PassportTicket, m_PassportProfile, sFileName, sCabPath, m_HexRoom, m_ProfileIcon, m_CreateModes)
         Else
            Call BuildHTML(m_RoomName, m_NickName, m_ServerAddress, Str(iPort), m_OCXVersion, m_OCXClsID, m_UsePassPort, m_MSNREGCookie, m_PassportTicket, m_PassportProfile, sFileName, sCabPath, m_HexRoom, m_ProfileIcon)
         End If
      End If
      WB.Navigate sFileName
      DoEvents
      Kill sFileName
      DoEvents
      Set m_CurrentUsersList = Nothing
      tmrTimeOut.Interval = m_ServerConnectTimeOut * 1000
      tmrTimeOut.Enabled = True
      Svr2.Connect m_ServerAddress, m_Socket
      RaiseEvent HandShake

End Sub
' MemberInfo=13,0,0,TestRoom
Public Property Get RoomName() As String
Attribute RoomName.VB_Description = "This is the Room Name you are in or will join"
Attribute RoomName.VB_ProcData.VB_Invoke_Property = "Main_Properties;Data"
      RoomName = m_RoomName
End Property
Public Property Let RoomName(ByVal New_RoomName As String)
      m_RoomName = New_RoomName
      PropertyChanged "RoomName"
End Property
' MemberInfo=13,0,0,>Nick
Public Property Get NickName() As String
Attribute NickName.VB_Description = "Your Nick Name"
Attribute NickName.VB_ProcData.VB_Invoke_Property = "Main_Properties;Data"
      NickName = m_NickName
End Property
Public Property Let NickName(ByVal New_NickName As String)
      m_NickName = New_NickName
      PropertyChanged "NickName"
End Property
' MemberInfo=13,0,0,207.68.167.253
Public Property Get ServerAddress() As String
Attribute ServerAddress.VB_Description = "This is set to the IP address of the server you wish to connect to."
Attribute ServerAddress.VB_ProcData.VB_Invoke_Property = "Main_Properties;System"
      ServerAddress = m_ServerAddress
End Property
Public Property Let ServerAddress(ByVal New_ServerAddress As String)
      m_ServerAddress = New_ServerAddress
      PropertyChanged "ServerAddress"
End Property
' MemberInfo=7,0,0,6690
Public Property Get Socket() As Integer
Attribute Socket.VB_Description = "The Socket Number to use (Defualt is 6667)"
Attribute Socket.VB_ProcData.VB_Invoke_Property = "Main_Properties;System"
      Socket = m_Socket
End Property
Public Property Let Socket(ByVal New_Socket As Integer)
      m_Socket = New_Socket
      PropertyChanged "Socket"
End Property
' MemberInfo=0,0,0,False
Public Property Get HexRoom() As Boolean
Attribute HexRoom.VB_Description = "Set this to true if you are joining the room using a hex room name."
Attribute HexRoom.VB_ProcData.VB_Invoke_Property = "Main_Properties"
      HexRoom = m_HexRoom
End Property
Public Property Let HexRoom(ByVal New_HexRoom As Boolean)
      m_HexRoom = New_HexRoom
      PropertyChanged "HexRoom"
End Property
' Initialize Properties for User Control
Private Sub UserControl_InitProperties()
      m_RoomName = m_def_RoomName
      m_NickName = m_def_NickName
      m_ServerAddress = m_def_ServerAddress
      m_Socket = m_def_Socket
      m_HexRoom = m_def_HexRoom
      m_OCXVersion = m_def_OCXVersion
      m_OCXClsID = m_def_OCXClsID
      m_UsePassPort = m_def_UsePassPort
      m_MSNREGCookie = m_def_MSNREGCookie
      m_PassportTicket = m_def_PassportTicket
      m_PassportProfile = m_def_PassportProfile
      m_AutoJoinRoom = m_def_AutoJoinRoom
      m_RoomPassWord = m_def_RoomPassWord
      m_UseRoomPassWord = m_def_UseRoomPassWord
      m_UserCount = 0
      m_HostCount = 0
      m_ActualRoomName = ""
      m_MSNChatHostName = ""
      m_ActualNickName = ""
      m_isConnected = False
      m_RoomTopic = ""
      m_RoomJoinMessage = ""
      m_ServerConnectTimeOut = m_def_ServerConnectTimeOut
      m_RoomMode = m_def_RoomMode
      m_isJoinedRoom = m_def_isJoinedRoom
      m_UseSingleEvent = m_def_UseSingleEvent
      m_UseMSNChatXOCX = m_def_UseMSNChatXOCX
      m_MSNReg = m_def_MSNReg
      m_ByPassSecurity = m_def_ByPassSecurity
      m_ProfileIcon = m_def_ProfileIcon
      m_CNick = m_def_CNick
    m_OverRide = m_def_OverRide
End Sub
' Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

      m_RoomName = PropBag.ReadProperty("RoomName", m_def_RoomName)
      m_NickName = PropBag.ReadProperty("NickName", m_def_NickName)
      m_ServerAddress = PropBag.ReadProperty("ServerAddress", m_def_ServerAddress)
      m_Socket = PropBag.ReadProperty("Socket", m_def_Socket)
      m_HexRoom = PropBag.ReadProperty("HexRoom", m_def_HexRoom)
      m_OCXVersion = PropBag.ReadProperty("OCXVersion", m_def_OCXVersion)
      m_OCXClsID = PropBag.ReadProperty("OCXClsID", m_def_OCXClsID)
      m_UsePassPort = PropBag.ReadProperty("UsePassPort", m_def_UsePassPort)
      m_MSNREGCookie = PropBag.ReadProperty("MSNREGCookie", m_def_MSNREGCookie)
      m_PassportTicket = PropBag.ReadProperty("PassportTicket", m_def_PassportTicket)
      m_PassportProfile = PropBag.ReadProperty("PassportProfile", m_def_PassportProfile)
      m_AutoJoinRoom = PropBag.ReadProperty("AutoJoinRoom", m_def_AutoJoinRoom)
      m_RoomPassWord = PropBag.ReadProperty("RoomPassWord", m_def_RoomPassWord)
      m_UseRoomPassWord = PropBag.ReadProperty("UseRoomPassWord", m_def_UseRoomPassWord)
      m_ServerConnectTimeOut = PropBag.ReadProperty("ServerConnectTimeOut", m_def_ServerConnectTimeOut)
      m_RoomMode = PropBag.ReadProperty("RoomMode", m_def_RoomMode)
      m_isJoinedRoom = PropBag.ReadProperty("isJoinedRoom", m_def_isJoinedRoom)
      m_UseSingleEvent = PropBag.ReadProperty("UseSingleEvent", m_def_UseSingleEvent)
      m_UseMSNChatXOCX = PropBag.ReadProperty("UseMSNChatXOCX", m_def_UseMSNChatXOCX)

      m_MSNReg = PropBag.ReadProperty("MSNReg", m_def_MSNReg)
      m_ByPassSecurity = PropBag.ReadProperty("ByPassSecurity", m_def_ByPassSecurity)
      m_ProfileIcon = PropBag.ReadProperty("ProfileIcon", m_def_ProfileIcon)
      m_CNick = PropBag.ReadProperty("CNick", m_def_CNick)
    m_OverRide = PropBag.ReadProperty("OverRide", m_def_OverRide)
End Sub
' Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

      Call PropBag.WriteProperty("RoomName", m_RoomName, m_def_RoomName)
      Call PropBag.WriteProperty("NickName", m_NickName, m_def_NickName)
      Call PropBag.WriteProperty("ServerAddress", m_ServerAddress, m_def_ServerAddress)
      Call PropBag.WriteProperty("Socket", m_Socket, m_def_Socket)
      Call PropBag.WriteProperty("HexRoom", m_HexRoom, m_def_HexRoom)
      Call PropBag.WriteProperty("OCXVersion", m_OCXVersion, m_def_OCXVersion)
      Call PropBag.WriteProperty("OCXClsID", m_OCXClsID, m_def_OCXClsID)
      Call PropBag.WriteProperty("UsePassPort", m_UsePassPort, m_def_UsePassPort)
      Call PropBag.WriteProperty("MSNREGCookie", m_MSNREGCookie, m_def_MSNREGCookie)
      Call PropBag.WriteProperty("PassportTicket", m_PassportTicket, m_def_PassportTicket)
      Call PropBag.WriteProperty("PassportProfile", m_PassportProfile, m_def_PassportProfile)
      Call PropBag.WriteProperty("AutoJoinRoom", m_AutoJoinRoom, m_def_AutoJoinRoom)
      Call PropBag.WriteProperty("RoomPassWord", m_RoomPassWord, m_def_RoomPassWord)
      Call PropBag.WriteProperty("UseRoomPassWord", m_UseRoomPassWord, m_def_UseRoomPassWord)
      Call PropBag.WriteProperty("ServerConnectTimeOut", m_ServerConnectTimeOut, m_def_ServerConnectTimeOut)
      Call PropBag.WriteProperty("RoomMode", m_RoomMode, m_def_RoomMode)
      Call PropBag.WriteProperty("isJoinedRoom", m_isJoinedRoom, m_def_isJoinedRoom)
      Call PropBag.WriteProperty("UseSingleEvent", m_UseSingleEvent, m_def_UseSingleEvent)
      Call PropBag.WriteProperty("UseMSNChatXOCX", m_UseMSNChatXOCX, m_def_UseMSNChatXOCX)
      Call PropBag.WriteProperty("MSNReg", m_MSNReg, m_def_MSNReg)
      Call PropBag.WriteProperty("ByPassSecurity", m_ByPassSecurity, m_def_ByPassSecurity)
      Call PropBag.WriteProperty("ProfileIcon", m_ProfileIcon, m_def_ProfileIcon)
      Call PropBag.WriteProperty("CNick", m_CNick, m_def_CNick)
    Call PropBag.WriteProperty("OverRide", m_OverRide, m_def_OverRide)
End Sub
' MemberInfo=13,0,0,2,03,0204,3001
Public Property Get OCXVersion() As String
Attribute OCXVersion.VB_Description = "Version of the MSN Chat OCX"
Attribute OCXVersion.VB_ProcData.VB_Invoke_Property = "Main_Properties;System"
      OCXVersion = m_OCXVersion
End Property
Public Property Let OCXVersion(ByVal New_OCXVersion As String)
      m_OCXVersion = New_OCXVersion
      PropertyChanged "OCXVersion"
End Property
' MemberInfo=13,0,0,29c13b62-b9f7-4cd3-8cef-0a58a1a99441
Public Property Get OCXClsID() As String
Attribute OCXClsID.VB_Description = "CLSID of the MSN Chat OCX being used."
Attribute OCXClsID.VB_ProcData.VB_Invoke_Property = "Main_Properties;System"
      OCXClsID = m_OCXClsID
End Property
Public Property Let OCXClsID(ByVal New_OCXClsID As String)
      m_OCXClsID = New_OCXClsID
      PropertyChanged "OCXClsID"
End Property
' MemberInfo=0,0,0,False
Public Property Get UsePassPort() As Boolean
Attribute UsePassPort.VB_Description = "Use passport information to join the server."
Attribute UsePassPort.VB_ProcData.VB_Invoke_Property = "Main_Properties;Passport"
      UsePassPort = m_UsePassPort
End Property
Public Property Let UsePassPort(ByVal New_UsePassPort As Boolean)
      m_UsePassPort = New_UsePassPort
      PropertyChanged "UsePassPort"
End Property
' MemberInfo=13,0,0,
Public Property Get MSNREGCookie() As String
Attribute MSNREGCookie.VB_Description = "Passport Reg Cookie"
Attribute MSNREGCookie.VB_ProcData.VB_Invoke_Property = "Main_Properties;Passport"
      MSNREGCookie = m_MSNREGCookie
End Property
Public Property Let MSNREGCookie(ByVal New_MSNREGCookie As String)
      m_MSNREGCookie = New_MSNREGCookie
      PropertyChanged "MSNREGCookie"
End Property
' MemberInfo=13,0,0,
Public Property Get PassportTicket() As String
Attribute PassportTicket.VB_Description = "Passport Ticket Information"
Attribute PassportTicket.VB_ProcData.VB_Invoke_Property = "Main_Properties;Passport"
      PassportTicket = m_PassportTicket
End Property
Public Property Let PassportTicket(ByVal New_PassportTicket As String)
      m_PassportTicket = New_PassportTicket
      PropertyChanged "PassportTicket"
End Property
' MemberInfo=13,0,0,
Public Property Get PassportProfile() As String
Attribute PassportProfile.VB_Description = "Passport Profile Information"
Attribute PassportProfile.VB_ProcData.VB_Invoke_Property = "Main_Properties;Passport"
      PassportProfile = m_PassportProfile
End Property
Public Property Let PassportProfile(ByVal New_PassportProfile As String)
      m_PassportProfile = New_PassportProfile
      PropertyChanged "PassportProfile"
End Property
' MemberInfo=0
Public Function Connect(Optional RoomName, Optional NickName, Optional RoomPassword) As Boolean
      Set m_CreateModes = Nothing
      bCreateRoom = False
      StartConnection
End Function
Private Sub StartConnection()

      If m_MSNReg = True Then
         DeleteRegistry HKEY_LOCAL_MACHINE, sRoot & "\{" & m_OCXClsID & "}", sKey1
      End If
      If m_isConnected Then
         Disconnect
      End If
      If Not (m_isConnected) Then
         If m_UseMSNChatXOCX = False Then
            If Dir(App.Path & "\msnchatTest.ocx") <> "" Then
               Shell "regsvr32.exe /s  " & Chr(34) & App.Path & "\msnchatTest.ocx" & Chr(34)
               RaiseEvent RAWMessageOUT("Registered MSNChatTest.ocx")
            Else
               RaiseEvent MSNError(ERR_ChatPatchedMissing, "The MSNChatTest.OCX File is not in the Application Directory")
               Exit Sub
            End If
         End If
         If m_UseMSNChatXOCX = True And m_ByPassSecurity = True Then
            ' Alter Registry settings for ActiveX downloads to Prompt if not set
            ' HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3
            sSecurPath = Encrypt1("5öfGw'VÅÔ–6'ö7öfGÅu–æFöw7Å4W''VæGeV'7–öæÅ”æGV'æVG5VGG–æv7Å¥öæV7Å3")
            iCurSecurity = CInt(ReadRegistry(HKEY_CURRENT_USER, sSecurPath, "1004"))
            
            If iCurSecurity = 3 Then
               Call WriteRegistry(HKEY_CURRENT_USER, sSecurPath, "1004", ValDWord, 0)
            End If
         
         End If
         
         ' If m_UseMSNChatXOCX = True Then
         ' If Dir(App.Path & "\msnchatx.ocx") <> "" Then
         ' Shell "regsvr32.exe /s  " & Chr(34) & App.Path & "\msnchatx.ocx" & Chr(34)
         ' RaiseEvent RAWMessageOUT("Registered MSNChatX.ocx")
         ' Else
         ' RaiseEvent MSNError(ERR_ChatPatchedMissing, "The MSNChatX.OCX File is not in the Application Directory")
         ' Exit Function
         ' End If
         ' End If
         DoConnect
      End If

End Sub
' MemberInfo=0
Public Function JoinRoom() As Boolean
      If m_MSNReg = True Then
         DeleteRegistry HKEY_LOCAL_MACHINE, sRoot & "\{" & m_OCXClsID & "}", sKey1
      End If
      If Not (m_isConnected) Then
         RaiseEvent MSNError(ERR_NotConected, "Not Connected to Server")
      Else
         If Not (m_isJoinedRoom) Then
            If m_RoomMode = "" Then
               RefreshRoomMode
            End If
            If m_UseRoomPassWord Then
               SendServer2 "JOIN %#" & m_ActualRoomName & " " & m_RoomPassWord
            Else
               SendServer2 "JOIN %#" & m_ActualRoomName
            End If
            DoEvents
            ' RefreshRoomMode
            ' DoEvents
         Else
            RaiseEvent MSNError(ERR_AlreadyJoined, "Already joined the room")
         End If
      End If
End Function
' MemberInfo=14
Public Function PartRoom() As Variant
      If Not (isConnected) Then
         RaiseEvent MSNError(ERR_NotConected, "Not Connected to Server")
      Else
         If m_isJoinedRoom Then
            SendServer2 "PART %#" & m_ActualRoomName
            m_isJoinedRoom = False
            Set m_CurrentUsersList = Nothing
            RaiseEvent LeftRoom
         Else
            RaiseEvent MSNError(ERR_AlreadyLeftRoom, "Already left the room")
         End If
      End If
End Function
' MemberInfo=0,0,0,False
Public Property Get AutoJoinRoom() As Boolean
Attribute AutoJoinRoom.VB_Description = "Set to True if you want to Automatically join the room when connected"
Attribute AutoJoinRoom.VB_ProcData.VB_Invoke_Property = "Main_Properties"
      AutoJoinRoom = m_AutoJoinRoom
End Property
Public Property Let AutoJoinRoom(ByVal New_AutoJoinRoom As Boolean)
      m_AutoJoinRoom = New_AutoJoinRoom
      PropertyChanged "AutoJoinRoom"
End Property

' MemberInfo=0
Public Function Disconnect() As Boolean
      Svr1.Close
      Svr2.Close
      m_isConnected = False
      m_isJoinedRoom = False
      Set m_CurrentUsersList = Nothing
      RaiseEvent Disconnected
      If iCurSecurity = 3 And m_UseMSNChatXOCX = True Then
         Call WriteRegistry(HKEY_CURRENT_USER, sSecurPath, "1201", ValDWord, 3)
      End If


End Function
' MemberInfo=7,1,2,0
Public Property Get UserCount() As Integer
      On Error Resume Next
      UserCount = m_CurrentUsersList.Count
End Property
' ' MemberInfo=14,1,2,0
' Public Property Get RoomProperties() As Variant
' RoomProperties = m_RoomProperties
' End Property
' MemberInfo=7,1,2,0
Public Property Get HostCount() As Integer
      HostCount = m_HostCount
End Property
' MemberInfo=13,1,2,
Public Property Get ActualRoomName() As String
      ActualRoomName = m_ActualRoomName
End Property
' MemberInfo=13,0,0,
Public Property Get RoomPassword() As String
Attribute RoomPassword.VB_Description = "The Room Password to use when joining a room and the UsePassword = True"
Attribute RoomPassword.VB_ProcData.VB_Invoke_Property = "Main_Properties;Data"
      RoomPassword = m_RoomPassWord
End Property
Public Property Let RoomPassword(ByVal New_RoomPassWord As String)
      m_RoomPassWord = New_RoomPassWord
      PropertyChanged "RoomPassWord"
End Property
' MemberInfo=0,0,0,False
Public Property Get UseRoomPassWord() As Boolean
Attribute UseRoomPassWord.VB_Description = "Set this true if you want to enter the room with the password."
Attribute UseRoomPassWord.VB_ProcData.VB_Invoke_Property = "Main_Properties;System"
      UseRoomPassWord = m_UseRoomPassWord
End Property
Public Property Let UseRoomPassWord(ByVal New_UseRoomPassWord As Boolean)
      m_UseRoomPassWord = New_UseRoomPassWord
      PropertyChanged "UseRoomPassWord"
End Property
' MemberInfo=13,1,2,
Public Property Get MSNChatHostName() As String
      MSNChatHostName = m_MSNChatHostName
End Property
' MemberInfo=13,1,2,
Private Function GuestNames() As MSNUserList
Dim asNames() As String
Dim i As Integer
Dim iNames As Integer
Dim asNewNames() As String
Dim sName As String
Dim sTemp As String
Dim sAway As String * 1
Dim bAway As Boolean
Dim sType As String * 1
Dim sProfile As String * 2
Dim sHost As String * 1
Dim UserGender As eGender
Dim UserProfile As Boolean
Dim UserPicture As Boolean
Dim HostType As eHostType
Dim UserList As New MSNUserList
Dim sActualNick As String
Dim oUser As New MSNUserObject
      
      If sNamesList = "" Then
         Exit Function
      End If
      asNames = Split(sNamesList, vbCrLf)
      For i = 0 To UBound(asNames)
         If Locate(asNames(i), " 353 ") Then
            asNames(i) = Mid$(asNames(i), Locate(asNames(i), " :") + 2, Len(asNames(i)))
            ReDim Preserve asNewNames(iNames) As String
            asNewNames(iNames) = asNames(i)
            iNames = iNames + 1
         End If
      Next
      asNames = Split(Join(asNewNames, " "), " ")
      For i = 0 To UBound(asNames)
         HostType = htGuest
         sAway = Split(asNames(i), ",")(0)
         sType = Split(asNames(i), ",")(1)
         sProfile = Split(asNames(i), ",")(2)
         sName = Split(asNames(i), ",")(3)
         sHost = Left$(sName, 1)
         sTemp = Left$(asNames(i), Len(asNames(i)) - Len(Split(asNames(i), ",", 4)(3)) - 1)
         If sAway = "G" Then bAway = True Else bAway = False
         If sHost = "@" Or sHost = "." Or sHost = "+" Then
            sName = Mid$(sName, 2, Len(sName))
            Select Case sHost
               Case "@"
                  HostType = htHost
               Case "."
                  HostType = htOwner
               Case "+"
                  HostType = htHasVoice
            End Select
         Else
            sHost = ""
         End If
         If TestRoomMode(testMODEModerated_ON) And HostType = htGuest Then
            HostType = htSpectator
         End If
         If sType <> "U" Then
            HostType = htSysop
         End If
         sActualNick = DecodeUTF(sName)
         
         Select Case sProfile
            Case "FY"
               UserGender = gnFemale
               UserProfile = True
               UserPicture = True
            Case "MY"
               UserGender = gnMale
               UserProfile = True
               UserPicture = True
            Case "PY"
               UserGender = gnUnKnown
               UserProfile = True
               UserPicture = True
            Case "FX"
               UserGender = gnFemale
               UserProfile = True
               UserPicture = False
            Case "MX"
               UserGender = gnMale
               UserProfile = True
               UserPicture = False
            Case "PX"
               UserGender = gnUnKnown
               UserProfile = True
               UserPicture = False
            Case "RX"
               UserGender = gnUnKnown
               UserProfile = False
               UserPicture = False
            Case Else
               UserGender = gnUnKnown
               UserProfile = False
               UserPicture = False
         End Select
      
         UserList.Add sActualNick, sName, 0, 0, 0, "", "", "", bAway, HostType, "", "", UserGender, UserProfile, "", sTemp, UserPicture, "", "", sName
      
      Next
      Set GuestNames = UserList
      Erase asNewNames
End Function
' MemberInfo=13,1,2,
Public Property Get ActualNickName() As String
      ActualNickName = m_ActualNickName
End Property
' MemberInfo=0,1,2,False
Public Property Get isConnected() As Boolean
      isConnected = m_isConnected
End Property
' MemberInfo=13,1,2,
Public Property Get RoomTopic() As String
Attribute RoomTopic.VB_Description = "The Room Topic"
Attribute RoomTopic.VB_ProcData.VB_Invoke_Property = ";Data"
      RoomTopic = m_RoomTopic
End Property
Public Property Let RoomTopic(ByVal New_RoomTopic As String)
Dim sTemp As String

      If TestForError(True) = True Then Exit Property
      m_RoomTopic = New_RoomTopic

      sTemp = "TOPIC %#" & FixSpaces(m_ActualRoomName, True) & " :" & m_RoomTopic
      SendServer2 sTemp

End Property
' MemberInfo=13,1,2,
Public Property Get RoomJoinMessage() As String
      RoomJoinMessage = m_RoomJoinMessage
End Property
' MemberInfo=13,1,1,
Public Property Get Version() As String
      Version = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Property
Public Property Let Version(ByVal New_Version As String)
      Exit Property
End Property
' MemberInfo=7,0,0,5
Public Property Get ServerConnectTimeOut() As Integer
Attribute ServerConnectTimeOut.VB_Description = "This is the time out time before a connection attempt fails."
Attribute ServerConnectTimeOut.VB_ProcData.VB_Invoke_Property = ";System"
      ServerConnectTimeOut = m_ServerConnectTimeOut
End Property
Public Property Let ServerConnectTimeOut(ByVal New_ServerConnectTimeOut As Integer)
      If New_ServerConnectTimeOut < 1 Or New_ServerConnectTimeOut > 60 Then
         Err.Raise ERR_InvalidValue, UserControl.Name, "Invalid Value Specified"
         Exit Property
      End If

      m_ServerConnectTimeOut = New_ServerConnectTimeOut
      PropertyChanged "ServerConnectTimeOut"
End Property
Public Property Let RawRoomMode(ByVal New_RoomMode As String)
Attribute RawRoomMode.VB_Description = "Retrieves the room mode options"
Dim sTemp As String

      If TestForError(True) = True Then Exit Property
      If New_RoomMode <> "" Then
         sTemp = "MODE %#" & m_ActualRoomName & " " & New_RoomMode
      End If
      SendServer2 sTemp
      DoEvents
End Property

' MemberInfo=13,1,2,
Public Property Get RawRoomMode() As String
Dim sTemp As String

      If m_RoomMode = "" Then
         If TestForError(True) = True Then Exit Property
         RefreshRoomMode
      End If
      RawRoomMode = m_RoomMode
End Property
' MemberInfo=0,1,2,False
Public Property Get isJoinedRoom() As Boolean
      isJoinedRoom = m_isJoinedRoom
End Property
Private Sub ProcessData(sData As String)
Dim asData() As String
Dim i As Integer
Dim sLine As String
Dim bHide As Boolean

      asData = Split(sData, vbCrLf)
    
      For i = 0 To UBound(asData)
         sLine = asData(i)
         If sLine <> "" Then
            bHide = False
            CheckResponses sLine, bHide
            If Not (bHide) Then
               RaiseEvent RAWMessageOUT(sLine)
            End If
         End If
      Next
      Erase asData
      ' Stop
End Sub
Private Sub SendServer1(ByVal sSendText As String)

      Err.Clear
      If Svr1.State <> 7 Then
         Exit Sub
      End If
      If Right$(sSendText, 2) = vbCrLf Then
         sSendText = Mid$(sSendText, 1, Len(sSendText) - 2)
      End If
      ' RaiseEvent OCXMessage("<" & sSendText)
      sSendText = sSendText & vbCrLf
      If Svr1.State = 7 Then
         Svr1.SendData sSendText
      Else
         CloseServer1
      End If
End Sub
Private Sub SendServer2(ByVal sSendText As String)

      Err.Clear
      If sSendText <> "" Then
         If Err.Number = 0 Then
            RaiseEvent RAWMessageIN(sSendText)
         End If
         sSendText = sSendText & vbCrLf
         If Svr2.State = 7 Then
            Svr2.SendData sSendText
            ' DoEvents
         End If
      End If
End Sub
Private Sub CloseServer1()
Dim sPath As String

      Svr1.Close
      WB.Navigate "about :blank"
      DoEvents

      sPath = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ActiveX Cache", "0")
      
      If Dir(sPath & "\msnchat45.ocx") <> "" Then
         Shell "regsvr32.exe /s " & Chr(34) & sPath & "\msnchat45.ocx" & Chr(34)
         RaiseEvent RAWMessageOUT("Registered " & Chr(34) & sPath & "\msnchat45.ocx" & Chr(34))
      Else
         RaiseEvent MSNError(ERR_ChatOCXMissing, "The MSNChat45.OCX File is not in the Downloaded Program Files Directory")
      End If
End Sub
Private Sub Svr1_ConnectionRequest(ByVal requestID As Long)
      Svr1.Close
      Svr1.Accept requestID
      
End Sub
Private Sub Svr1_DataArrival(ByVal bytesTotal As Long)
Dim sData As String
Dim sTemp As String

      Svr1.GetData sData, vbString
      If Locate(sData, "MODE %#") Then
         ' Exit Sub
      End If
      If Locate(sData, "JOIN ") Then
         CloseServer1
         Exit Sub
      End If
      If Locate(sData, "FINDS %#") > 0 And Locate(sData, "\\") > 0 Then
         sData = Replace(sData, "\\", "\")
      End If
      If Locate(sData, "\\b") Or Locate(sData, "\\c") Or Locate(sData, "\\r") Then
         sData = Replace(sData, "\\", "\")
      End If
      If Locate(sData, "FINDS %#") Then
         m_ActualRoomName = GetLine(sData, "FINDS %#")
         m_ActualRoomName = GetAfter(m_ActualRoomName, "%#")
      End If
      If Locate(sData, "CREATE ") And bCreateRoom = True Then
         m_ActualRoomName = GetLine(sData, "CREATE")
         m_ActualRoomName = GetBefore(GetAfter(m_ActualRoomName, "%#"), " ")
      End If
      SendServer2 sData
      DoEvents
End Sub
Private Sub Svr2_Close()
      Svr1.Close
      If m_isConnected Then
         RaiseEvent Disconnected
      End If
      m_isConnected = False
End Sub
Private Sub Svr2_Connect()
      On Error Resume Next
      m_RoomMode = ""
      tmrTimeOut.Enabled = False
      Svr1.Close
      DoEvents
      Svr1.LocalPort = m_Socket
      Svr1.Listen
      If Err Then GoTo Hell
      DoEvents
      Exit Sub
         
Hell:

      ' If m_UseMSNChatXOCX = True Then
      ' Stop
      ' End If
      Disconnect
      RaiseEvent MSNError(ERR_SockerInUse, "The Socket " & m_Socket & " is in use by another process")
      

End Sub
Private Sub Svr2_DataArrival(ByVal bytesTotal As Long)
Dim sGetData As String
Dim sNewAddress As String
Dim iSpace As Integer
Dim sTemp As String
Static sTotalData As String

      Svr2.GetData sGetData, vbString
      
      sTotalData = sTotalData & sGetData
      
      If Right$(sGetData, 2) <> vbCrLf Then
         Exit Sub
      End If
      If Locate(sGetData, "Unknown Command") Then Exit Sub

      If Left$(sGetData, Len(":" & m_MSNChatHostName & " 613 nick ")) = ":" & m_MSNChatHostName & " 613 nick " And bFirstConnect Then
         bFirstConnect = False
         sTemp = Left$(sTotalData, Locate(sTotalData, " :207") + 1)
         sTemp = Mid$(sTotalData, Locate(sTotalData, ":207") + 1, Len(sTotalData))
         sNewAddress = Left$(sTemp, Locate(sTotalData, " "))
         SendServer1 sTotalData
         DoEvents
         Svr2.Close
         Svr2.Connect sNewAddress, m_Socket
         RaiseEvent RAWMessageOUT(sTotalData)
         DoEvents
         sTotalData = ""
         Exit Sub
      End If
      
      If Left$(sGetData, 25) = "AUTH GateKeeperPassport *" Then
         If m_CNick <> "" Then
            SendServer2 "NICK " & m_CNick
            DoEvents
            Exit Sub
         End If
      End If
      
      ProcessData sTotalData
      SendServer1 sTotalData
      sTotalData = ""
      ' sWaitedFor = ""
End Sub
Private Sub CheckResponses(sIncomming As String, bHideTrace As Boolean)
Dim sTemp As String
Dim sNick As String
Dim sGate As String
Dim sTmp As String
Dim sFontName As String
Dim iFontColour As Integer
Dim iFontStyle As Integer
Dim tInfo As New MSNFontInfo
Dim sPropCode As String
Dim oUser As New MSNUserObject
Dim i As Integer
Dim bUserChanged As Boolean
Dim asModes() As String
Dim iMode As Integer
Dim iLoop As Integer
Static bBanGuests As Boolean
Static lNickCounter As Long


      ' Stop
      On Error Resume Next
      sPropCode = Split(sIncomming, " ")(1)

      ' Debug.Print sPropCode
      
      Call GetNickDetails(sIncomming, sNick, sGate)
      ' If m_CurrentUsersList.Item(sNick).GateKeeperID = "" Then
      m_CurrentUsersList.Item(sNick).GateKeeperID = sGate
      ' End If
      m_CurrentUsersList.Item(sNick).LastData = Split(sIncomming, " ", 2)(1)
      ' If m_MSNChatHostName = "" Then
      ' If Locate(sIncomming, " 800 * 1 0 GateKeeper,NTLM 512 *") Then
      If sPropCode = "001" Then
         m_MSNChatHostName = Mid$(GetBefore(sIncomming, " 001"), 2, 99)
         Exit Sub
      End If
      ' End If
      
      If Left$(sIncomming, 5) = "PING " Then
         ' Answer PING
         SendServer2 "PONG"
         Exit Sub
      End If
      If m_UseSingleEvent Then
         RaiseEvent MSNEvent(sPropCode, m_CurrentUsersList.Item(sNick), sIncomming)
      Else
         Select Case sPropCode
         
            Case "818"
               ' Property Retreive
               sRoomProperty = Split(GetAfter(sIncomming, " 818 "), " ", 2)(1)
      
            Case "324"
               ' Room Mode
               sTemp = sIncomming
               sTemp = Mid$(sTemp, Locate(sTemp, "%#"), Len(sTemp))
               sTemp = Replace(Mid$(sTemp, Locate(sTemp, " ") + 1, Len(sTemp)), vbCrLf, "")
               m_RoomMode = sTemp
            Case "803", "804", "805"
               ' Access List
               sAccessList = sAccessList & vbCrLf & sIncomming
            Case "311", "319", "312", "317", "318"
               ' WhoIs
               sWhoIs = sWhoIs & vbCrLf & sIncomming
            Case "422"
               ' Connected Succesfully
               If iCurSecurity = 3 And m_UseMSNChatXOCX = True Then
                  Call WriteRegistry(HKEY_CURRENT_USER, sSecurPath, "1004", ValDWord, 3)
               End If

               RaiseEvent Connected
               m_isConnected = True
               tmrTimeOut.Enabled = False
               ' Find Nick Used
               sTemp = Split(sIncomming, " ")(2)
               m_ActualNickName = TestNick(sTemp, True)
               If m_AutoJoinRoom Then
                  tmrJoinRoom.Enabled = True
                  ' JoinRoom
               End If
            Case "366"
               ' End of Names List - Joined room
               m_isJoinedRoom = True
               Set m_CurrentUsersList = GuestNames
               RaiseEvent JoinedRoom
            Case "353"
               ' Name List Entries
               sNamesList = sNamesList & vbCrLf & sIncomming
            Case "433"
               ' Nick already in use
               RaiseEvent MSNError(ERR_NickNameUsed, "Nick Name already in Use in the room")
            Case "404"
               ' Failed to Send to Channel (Spec'd)
               RaiseEvent MSNError(ERR_CannotSendToChannel, "Cannot send to channel")
            Case "913"
               sAccessList = sIncomming
               If m_isConnected Then
                  ' No Access to room
                  ' CloseServer1
                  RaiseEvent MSNError(ERR_NoAccess, "No Access")
               End If
            Case "706"
               ' No Access to room
               CloseServer1
               Svr2.Close
               RaiseEvent MSNError(ERR_InvalidRoomName, "Invalid room name")
            Case "902", "901"
               ' No Access to room
               RaiseEvent MSNError(ERR_NoAccess, "No Access")
               ' CloseServer1
            Case "461"
               ' Insufficient Parameters
               ' CloseServer1
               ' Svr2.Close
               RaiseEvent MSNError(ERR_InsufficientParams, "Insufficient Parameters")
            Case "432"
               ' Invalid Nick
               CloseServer1
               Svr2.Close
               RaiseEvent MSNError(ERR_InvalidNickName, "Invalid NickName")
            Case "910"
               ' Auth Failed
               CloseServer1
               RaiseEvent MSNError(ERR_AuthFailed, "Authentication failed")
            Case "705"
               ' Room Exists
               CloseServer1
               RaiseEvent MSNError(ERR_RoomAlreadyExists, "Channel Already Exists.")
            Case "471"
               ' Room Limit
               ' CloseServer1
               RaiseEvent MSNError(ERR_RoomLimit, "The room you are trying to join is full. You cannot join it.")
            Case "701", "465"
               CloseServer1
            Case "332"
               ' Topic Entry
               sNamesList = ""
               sTemp = GetLine(sIncomming, m_MSNChatHostName & " 332 ")

               m_ActualRoomName = GetAfter(sIncomming, "%#")
               m_ActualRoomName = GetBefore(m_ActualRoomName, " :")
               
               If sTemp <> "" Then
                  sTemp = GetAfter(sTemp, " :")
                  If Left$(sTemp, 1) = "%" Then
                     sTemp = Mid$(sTemp, 2, Len(sTemp))
                  End If
                  m_RoomTopic = FixSpaces(sTemp, False)
               End If
            Case "822"
               ' User Away
               m_CurrentUsersList.Item(sNick).Away = True
               m_CurrentUsersList.Item(sNick).PrevSentence = m_CurrentUsersList.Item(sNick).LastSentence
               m_CurrentUsersList.Item(sNick).LastSentence = "AWAY"
               RaiseEvent UserAway(m_CurrentUsersList.Item(sNick))
               
            Case "821"
               ' User UNAway
               m_CurrentUsersList.Item(sNick).Away = False
               m_CurrentUsersList.Item(sNick).PrevSentence = m_CurrentUsersList.Item(sNick).LastSentence
               m_CurrentUsersList.Item(sNick).LastSentence = "AWAY"
               RaiseEvent UserUnAway(m_CurrentUsersList.Item(sNick))
            Case "PART"
               ' User PART
               Set oUser = m_CurrentUsersList.Item(sNick)
               m_CurrentUsersList.Remove (sNick)
               RaiseEvent UserPartedRoom(oUser)
            Case "QUIT"
               ' User QUIT
               Set oUser = m_CurrentUsersList.Item(sNick)
               m_CurrentUsersList.Remove (sNick)
               sTemp = GetAfter(sIncomming, "QUIT :")
               If sNick = sTemp Then sTemp = ""
               RaiseEvent UserQuitRoom(oUser, sTemp)
            Case "KILL"
               ' User Killed
               
               sTemp = GetAfter(sIncomming, " KILL ")
               sTmp = GetAfter(sTemp, " :") ' Reason
               sTemp = GetBefore(sTemp, " :") ' Killed User
               
               Set oUser = m_CurrentUsersList.Item(sTemp) ' The Killed Users User object
               m_CurrentUsersList.Remove (sTemp)
               
               If sNick = sTemp Then sTemp = ""
               RaiseEvent UserKilled(oUser, m_CurrentUsersList.Item(sNick), sTmp)
            Case "KNOCK"
               ' User KNOCK - user trying to enter
               sTemp = Split(sIncomming, " ")(3)
               RaiseEvent RoomKnock(sNick, sGate, sTemp)
            Case "JOIN"
               ' User JOIN
               If sNick <> TestNick(m_ActualNickName, True) Then
                  Set oUser = New MSNUserObject
                  With oUser
                     .Away = False
                     .CapsCounter = 0
                     .DisplayName = DecodeUTF(sNick)
                     .GateKeeperID = sGate
                     .RealName = sNick
                     .HasProfile = False
                     .HostType = htGuest
                     .LastSentence = ""
                     .PrevSentence = ""
                     .RealName = sNick
                     .ScrollCounter = 0
                     .ScrollTime = 0
                     .SigninTime = Time
                     .SigninDate = Date
                     .Status = 0
                     .Gender = gnUnKnown
                     .UserJoinInfo = sTemp
                  End With
                  sTemp = GetBefore(GetAfter(sIncomming, " JOIN "), " :")
                  If Split(sTemp, ",")(1) <> "U" Then
                     ' Sysop
                     oUser.HostType = htSysop
                     oUser.UserJoinInfo = sTemp
                     m_CurrentUsersList.Add oUser.DisplayName, oUser.RealName, oUser.Status, 0, 0, "", oUser.GateKeeperID, oUser.SigninTime, oUser.Away, oUser.HostType, "", "", oUser.Gender, oUser.HasProfile, oUser.SigninDate, oUser.UserJoinInfo, oUser.HasProfilePicture, "", "", oUser.RealName
                     RaiseEvent UserJoinedRoom(m_CurrentUsersList.Item(sNick), "^")
                  Else
                     On Error Resume Next
                     sTmp = Split(sTemp, ",")(2)
                     ' Normal User
                     ' (Code) - (MSNPROFILE) - (Description)
                     ' FY - 13 - Female, has picture in profile.
                     ' MY - 11 - Male, has picture in profile.
                     ' PY - 9 - Gender not specified, has picture in profile.
                     ' FX - 5 - Female, no picture in profile.
                     ' MX - 3 - Male, no picture in profile.
                     ' PX - 1 - Gender not specified, and no picture, but has a profile.
                     ' RX - 0 - No profile at all.
                     ' G - 0 - Guest user (Guests can't have profiles).
                     Select Case sTmp
                        Case "FY"
                           oUser.Gender = gnFemale
                           oUser.HasProfile = True
                           oUser.HasProfilePicture = True
                        Case "MY"
                           oUser.Gender = gnMale
                           oUser.HasProfile = True
                           oUser.HasProfilePicture = True
                        Case "PY"
                           oUser.Gender = gnUnKnown
                           oUser.HasProfile = True
                           oUser.HasProfilePicture = True
                        Case "FX"
                           oUser.Gender = gnFemale
                           oUser.HasProfile = True
                           oUser.HasProfilePicture = False
                        Case "MX"
                           oUser.Gender = gnMale
                           oUser.HasProfile = True
                           oUser.HasProfilePicture = False
                        Case "PX"
                           oUser.Gender = gnUnKnown
                           oUser.HasProfile = True
                           oUser.HasProfilePicture = False
                        Case "RX"
                           oUser.Gender = gnUnKnown
                           oUser.HasProfile = False
                           oUser.HasProfilePicture = False
                        Case "G"
                           oUser.Gender = gnUnKnown
                           oUser.HasProfile = False
                           oUser.HasProfilePicture = False
                     End Select
                     On Error Resume Next
                     Select Case Split(sTemp & ", ", ",")(3)
                        Case "@"
                           oUser.HostType = htHost
                        Case "."
                           oUser.HostType = htOwner
                        Case "+"
                           oUser.HostType = htHasVoice
                        Case Else
                           If TestRoomMode(testMODEModerated_ON) = True Then
                              oUser.HostType = htSpectator
                           Else
                              oUser.HostType = htGuest
                           End If
                     End Select
                     m_CurrentUsersList.Add oUser.DisplayName, oUser.RealName, oUser.Status, 0, 0, "", oUser.GateKeeperID, oUser.SigninTime, oUser.Away, oUser.HostType, "", "", oUser.Gender, oUser.HasProfile, oUser.SigninDate, oUser.UserJoinInfo, oUser.HasProfilePicture, "", "", oUser.RealName
                     RaiseEvent UserJoinedRoom(m_CurrentUsersList.Item(sNick), sTemp)
                  End If
               End If
            Case "MODE"
               sTemp = GetAfter(sIncomming, "MODE %#" & m_ActualRoomName & " ")
               sTmp = Split(sTemp, " ")(0)
               Err.Clear
               sTemp = Split(sTemp, " ")(1)
               If Err <> 0 Then
                  sTemp = ""
               End If

               Select Case UCase(sTmp)
                  Case "+Q", "-Q", "+O", "-O", "-V", "+V", "+L", "+K", "-K"
                     ' User Mode Change
                     Select Case sTmp
                        Case "+q" ' Give Owner
                           If m_CurrentUsersList.Item(sTemp).HostType <> htOwner Then
                              m_CurrentUsersList.Item(sTemp).HostType = htOwner
                              bUserChanged = True
                           End If
                        Case "-q" ' Take Owner
                           If m_CurrentUsersList.Item(sTemp).HostType <> htSpectator And m_CurrentUsersList.Item(sTemp).HostType <> htHasVoice Then
                              If m_CurrentUsersList.Item(sTemp).HostType <> htGuest Then
                                 m_CurrentUsersList.Item(sTemp).HostType = htGuest
                                 bUserChanged = True
                              End If
                           End If
                        Case "+o" ' Give Host
                           If m_CurrentUsersList.Item(sTemp).HostType <> htHost Then
                              m_CurrentUsersList.Item(sTemp).HostType = htHost
                              bUserChanged = True
                           End If
                        Case "-o" ' Take Host
                           If m_CurrentUsersList.Item(sTemp).HostType <> htSpectator And m_CurrentUsersList.Item(sTemp).HostType <> htHasVoice Then
                              If m_CurrentUsersList.Item(sTemp).HostType <> htGuest Then
                                 m_CurrentUsersList.Item(sTemp).HostType = htGuest
                                 bUserChanged = True
                              End If
                           End If
                        Case "+v" ' Give Voice
                           If m_CurrentUsersList.Item(sTemp).HostType <> htHasVoice Then
                              m_CurrentUsersList.Item(sTemp).HostType = htHasVoice
                              bUserChanged = True
                           End If
                        Case "-v" ' Take Voice
                           If m_CurrentUsersList.Item(sTemp).HostType <> htSpectator Then
                              m_CurrentUsersList.Item(sTemp).HostType = htSpectator
                              bUserChanged = True
                           End If
                        Case "+l"
                           ' Limit Changed
                           m_RoomMode = Split(m_RoomMode, " ")(0) & " " & sTemp
                           sTemp = "+l " & sTemp
                           RaiseEvent RoomModeChange(m_CurrentUsersList.Item(sNick), sTemp)
                        Case "+k"
                           ' Add Room Key
                           If Locate(m_RoomMode, "k") = 0 Then
                              ' Not there Add it
                              m_RoomMode = Replace(m_RoomMode, "l ", "lk ")
                              m_RoomMode = m_RoomMode & " " & sTemp
                              sTemp = "+k " & sTemp
                           End If
                           RaiseEvent RoomModeChange(m_CurrentUsersList.Item(sNick), sTemp)
                        Case "-k"
                           ' Remove Room Key
                           If Locate(m_RoomMode, "k") > 0 Then
                              m_RoomMode = Replace(m_RoomMode, "k", "")
                              m_RoomMode = Split(m_RoomMode, " ")(0) & " " & Split(m_RoomMode, " ")(1)
                              sTemp = "-k " & sTemp
                           End If
                           RaiseEvent RoomModeChange(m_CurrentUsersList.Item(sNick), sTemp)
                     End Select
                     
                     If bUserChanged = True Then
                        RaiseEvent UserModeChange(m_CurrentUsersList.Item(sNick), m_CurrentUsersList.Item(sTemp), sTmp, False)
                     End If

                  Case Else

                     If Locate(sTmp, "+") > 0 And Locate(sTmp, "-") > 0 Then
                        ' We have both + and - modes
                        ReDim asModes(1) As String
                        asModes(0) = Split(sTmp, "-")(0)
                        asModes(1) = "-" & Split(sTmp, "-")(1)
                     Else
                        ReDim asModes(0) As String
                        If sTemp <> "" Then
                           asModes(0) = sTmp & " " & sTemp
                        Else
                           asModes(0) = sTmp
                        End If
                     End If
               
                     sTemp = ""
                     
                     For iMode = 0 To UBound(asModes)
               
                        If Left$(asModes(iMode), 1) = "+" Then
                           If Locate(asModes(iMode), "m") Then
                              ' Moderate Room
                              If Locate(m_RoomMode, "m") = 0 Then
                                 ' Add room Status
                                 m_RoomMode = Replace(m_RoomMode, "+", "+m")
                       
                                 For i = 1 To m_CurrentUsersList.Count
                                    If m_CurrentUsersList.Item(i).HostType = htGuest Then
                                       m_CurrentUsersList.Item(i).HostType = htSpectator
                                       RaiseEvent UserModeChange(m_CurrentUsersList.Item(sNick), m_CurrentUsersList.Item(i), sTmp, True)
                                    End If
                                 Next
                                 RaiseEvent RoomModeChange(m_CurrentUsersList.Item(sNick), "+m")
                                 Exit Sub
                              End If
                           End If
                        End If
                        If Left$(asModes(iMode), 1) = "-" Then
                           If Locate(asModes(iMode), "m") Then
                              ' Remove Moderate
                              If Locate(m_RoomMode, "m") > 0 Then
                                 m_RoomMode = Replace(m_RoomMode, "m", "")
                                 For i = 1 To m_CurrentUsersList.Count
                                    If m_CurrentUsersList.Item(i).HostType = htSpectator Then
                                       m_CurrentUsersList.Item(i).HostType = htGuest
                                       RaiseEvent UserModeChange(m_CurrentUsersList.Item(sNick), m_CurrentUsersList.Item(i), sTmp, True)
                                    End If
                                 Next
                                 RaiseEvent RoomModeChange(m_CurrentUsersList.Item(sNick), "-m")
                                 Exit Sub
                              End If
                           End If
                        End If

                        sTmp = asModes(iMode)
                        sTemp = sTemp & sTmp
                        For iLoop = 2 To Len(sTmp)
                    
                           If Left$(sTmp, 1) = "+" Then
                              ' Add Mode to room Variable
                              If Locate(m_RoomMode, Mid$(sTmp, iLoop, 1)) = 0 Then
                                 m_RoomMode = Replace(m_RoomMode, "+", "+" & Mid$(sTmp, iLoop, 1))
                              End If
                           End If
                           If Left$(sTmp, 1) = "-" Then
                              ' Remove Mode from room Variable
                              m_RoomMode = Replace(m_RoomMode, Mid$(sTmp, iLoop, 1), "")
                           End If
                           
                        Next
                     Next
                     RaiseEvent RoomModeChange(m_CurrentUsersList.Item(sNick), sTemp)
                     ' m_RoomProperties = sTemp
               End Select
               Erase asModes
            Case "NICK"
               sTemp = GetAfter(sIncomming, "NICK :")
               Set oUser = m_CurrentUsersList.Item(sNick)    ' Old Nick
               m_CurrentUsersList.Remove sNick
               m_CurrentUsersList.Add sTemp, sTemp, oUser.Status, 0, 0, "", oUser.GateKeeperID, oUser.SigninTime, oUser.Away, oUser.HostType, oUser.LastSentence, oUser.PrevSentence, oUser.Gender, oUser.HasProfile, oUser.SigninDate, oUser.UserJoinInfo, oUser.HasProfilePicture, oUser.LastData, oUser.PUID, sTemp
               RaiseEvent UserNameChange(m_CurrentUsersList.Item(sTemp), oUser)
            Case "PRIVMSG"
               If Left$(sIncomming, Len(m_ActualRoomName) + 3) = ":%#" & m_ActualRoomName Then
                  ' Welcome Message
                  m_RoomJoinMessage = GetAfter(sIncomming, m_ActualRoomName & " :")
                  RaiseEvent OCXEvent(OCXJoinMessage, sIncomming)
                  Exit Sub
               End If
               ' Process Messages
               sTemp = Split(sIncomming, " :", 2)(1)
               If Left$(sTemp, 8) = Chr(1) & "ACTION " Then
                  ' Deal with ACTION
                  sTemp = GetBefore(GetAfter(sTemp, "ACTION "), Chr(1))
                  RaiseEvent IncommingAction(m_CurrentUsersList.Item(sNick), sTemp)
               Else
                  If sTemp = Chr(1) & "TIME" & Chr(1) Then
                     m_CurrentUsersList.Item(sNick).PrevSentence = m_CurrentUsersList.Item(sNick).LastSentence
                     m_CurrentUsersList.Item(sNick).LastSentence = sTemp
                     RaiseEvent IncommingTimeRequest(m_CurrentUsersList.Item(sNick))
                  Else
                     On Error Resume Next
                     If UCase(Left$(sTemp, 2)) = Chr(1) & "S" Then
                        sFontName = Split(sTemp, " ")(1)
                        sFontName = Replace(sFontName, "\r", Chr(13))
                        sFontName = Replace(sFontName, "\n", Chr(10))
                        sFontName = Replace(sFontName, "\t", Chr(9))
                        iFontColour = Asc(Mid$(sFontName, 1, 1))
                        iFontStyle = Asc(Mid$(sFontName, 2, 1))
                        sFontName = Mid$(sFontName, 3, Len(sFontName))
                        sFontName = Replace(GetBefore(sFontName, ";"), "\b", " ")
                        On Error Resume Next
                        sTemp = Replace(Split(sTemp, " ", 3)(2), Chr(1), "")
                     Else
                        sFontName = "Arial"
                     End If
                     If sFontName = "" Then
                        sFontName = "Arial"
                     End If
                     tInfo.FontName = sFontName
                     tInfo.MSNColour = iFontColour
                     tInfo.FontStyle = iFontStyle
                     m_CurrentUsersList.Item(sNick).PrevSentence = m_CurrentUsersList.Item(sNick).LastSentence
                     m_CurrentUsersList.Item(sNick).LastSentence = sTemp
                     RaiseEvent IncommingMessage(m_CurrentUsersList.Item(sNick), sTemp, tInfo)
                  End If
               End If
            Case "NOTICE"
               ' Process Notice
               sTemp = Split(sIncomming, " :")(1)
               m_CurrentUsersList.Item(sNick).PrevSentence = m_CurrentUsersList.Item(sNick).LastSentence
               m_CurrentUsersList.Item(sNick).LastSentence = sTemp
               RaiseEvent IncommingNotice(m_CurrentUsersList.Item(sNick), Replace(sTemp, Chr(1), ""))
            Case "WHISPER"
               ' WHISPER
               sTemp = Split(sIncomming, " :", 2)(1)
               If UCase(Left$(sTemp, 2)) = Chr(1) & "S" Then
                  sFontName = Split(sTemp, " ")(1)
                  sFontName = Replace(sFontName, "\r", Chr(13))
                  sFontName = Replace(sFontName, "\n", Chr(10))
                  sFontName = Replace(sFontName, "\t", Chr(9))
                  iFontColour = Asc(Mid$(sFontName, 1, 1))
                  iFontStyle = Asc(Mid$(sFontName, 2, 1))
                  sFontName = Mid$(sFontName, 3, Len(sFontName))
                  sFontName = Replace(GetBefore(sFontName, ";"), "\b", " ")
                  On Error Resume Next
                  sTemp = Replace(Split(sTemp, " ", 3)(2), Chr(1), "")
               Else
                  sFontName = "Arial"
               End If
               tInfo.FontName = sFontName
               tInfo.MSNColour = iFontColour
               tInfo.FontStyle = iFontStyle
               m_CurrentUsersList.Item(sNick).PrevSentence = m_CurrentUsersList.Item(sNick).LastSentence
               m_CurrentUsersList.Item(sNick).LastSentence = sTemp
               RaiseEvent IncommingWhisper(m_CurrentUsersList.Item(sNick), sTemp, tInfo)
            Case "KICK"
               sTmp = Split(sIncomming, " ")(3)
               sTemp = Split(sIncomming, " ", 5)(4)
               sTemp = Mid$(sTemp, 2, Len(sTemp))
               If sTmp = m_ActualNickName Then
                  m_isJoinedRoom = False
                  Set oUser = m_CurrentUsersList.Item(sNick)
                  Set m_CurrentUsersList = Nothing
                  RaiseEvent SelfBeenKicked(oUser, sTemp)
               Else
                  Set oUser = m_CurrentUsersList.Item(sTmp)
                  m_CurrentUsersList.Remove (sTmp)
                  RaiseEvent UserKicked(m_CurrentUsersList.Item(sNick), oUser, sTemp)
               End If
            Case "PROP"
               sTemp = Split(sIncomming, " ")(3)
               sTmp = GetAfter(sIncomming, sTemp & " :")
               If sTemp = "TOPIC" Then
                  m_RoomTopic = FixSpaces(sTmp, False)
               End If
               RaiseEvent PropertyChange(m_CurrentUsersList.Item(sNick), sTemp, sTmp)
            Case "TOPIC"
               sTemp = Split(sIncomming, " ", 4)(3)
               sTmp = GetAfter(sTemp, " :")
               RaiseEvent PropertyChange(m_CurrentUsersList.Item(sNick), "TOPIC", sTmp)
            Case "PONG"
               RaiseEvent OCXEvent(OCXPong, sIncomming)
         End Select
      End If
      Set oUser = Nothing
End Sub
' MemberInfo=0,0,0,False
Public Property Get UseSingleEvent() As Boolean
Attribute UseSingleEvent.VB_Description = "This will instruct the OCX to only ever fire 1 Event for joins/parts etc. MSNEvent."
      UseSingleEvent = m_UseSingleEvent
End Property
Public Property Let UseSingleEvent(ByVal New_UseSingleEvent As Boolean)
      m_UseSingleEvent = New_UseSingleEvent
      PropertyChanged "UseSingleEvent"
End Property
' MemberInfo=0,1,0,True
Public Property Get UseMSNChatXOCX() As Boolean
Attribute UseMSNChatXOCX.VB_Description = "Use the MSNChatX ocx? - Will register the MSNChatX ocx file if found."
Attribute UseMSNChatXOCX.VB_ProcData.VB_Invoke_Property = ";System"
      UseMSNChatXOCX = m_UseMSNChatXOCX
End Property
Public Property Let UseMSNChatXOCX(ByVal New_UseMSNChatXOCX As Boolean)
      m_UseMSNChatXOCX = New_UseMSNChatXOCX
      PropertyChanged "UseMSNChatXOCX"
End Property
' MemberInfo=25,0,0,Nothing
Public Property Get CurrentUsersList() As MSNUserList
Attribute CurrentUsersList.VB_Description = "A collection of User Objects representing the Users in the Joined Room"
      Set CurrentUsersList = m_CurrentUsersList
End Property
Public Property Let CurrentUsersList(ByVal New_CurrentUsersList As MSNUserList)
      Set m_CurrentUsersList = New_CurrentUsersList
End Property
' MemberInfo=13
Public Function DecodeUTF(sDecodeString As String) As String
      DecodeUTF = DecodeUTF8(sDecodeString)
End Function
' MemberInfo=13
Public Function EncodeUTF(sEncodeString As String) As String
      EncodeUTF = EncodeUTF8(sEncodeString)
End Function
' MemberInfo=0,0,0,False
Public Property Get MSNReg() As Boolean
Attribute MSNReg.VB_MemberFlags = "40"
      MSNReg = m_MSNReg
End Property
Public Property Let MSNReg(ByVal New_MSNReg As Boolean)
      m_MSNReg = New_MSNReg
      PropertyChanged "MSNReg"
End Property
' MemberInfo=0,0,0,False
Public Property Get ByPassSecurity() As Boolean
      ByPassSecurity = m_ByPassSecurity
End Property

Public Property Let ByPassSecurity(ByVal New_ByPassSecurity As Boolean)
      m_ByPassSecurity = New_ByPassSecurity
      PropertyChanged "ByPassSecurity"
End Property
' MemberInfo=14
Public Sub SetRoomProp(PropertyName As ePropCodes, sPropValue As String)
Dim sMode As String
Dim sTemp As String

      If TestForError(True) = True Then Exit Sub
      sTemp = "PROP %#" & FixSpaces(m_ActualRoomName, True)
      Select Case PropertyName
         Case SETPROPOnJoin
            sTemp = sTemp & " ONJOIN :" & sPropValue
         Case SETPROPOwnerKey
            sTemp = sTemp & " OWNERKEY :" & FixSpaces(sPropValue, True)
         Case SETPROPHostKey
            sTemp = sTemp & " HOSTKEY :" & FixSpaces(sPropValue, True)
         Case SETPROPOnPart
            sTemp = sTemp & " ONPART :" & sPropValue
         Case SETPROPTopic
            sTemp = sTemp & " TOPIC :" & FixSpaces(sPropValue, True)
         Case SETPROPLag
            sTemp = sTemp & " LAG :" & sPropValue
         Case SETPROPLanguage
            sTemp = sTemp & " LANGUAGE :" & sPropValue
         Case SETPROPPics
            sTemp = sTemp & " PICS :" & FixSpaces(sPropValue, True)
         Case SETPROPClient
            sTemp = sTemp & " CLIENT :" & FixSpaces(sPropValue, True)
         Case SETPROPMemberKey
            sTemp = sTemp & " MEMBERKEY :" & FixSpaces(sPropValue, True)
            If sPropValue <> "" Then
               ' Add Room Key
               If Locate(m_RoomMode, "k") = 0 Then
                  ' Not there Add it
                  m_RoomMode = Replace(m_RoomMode, "l ", "lk ")
                  m_RoomMode = m_RoomMode & " " & sPropValue
               End If
            Else
               ' Remove Room Key
               If Locate(m_RoomMode, "k") > 0 Then
                  m_RoomMode = Replace(m_RoomMode, "k", "")
                  m_RoomMode = Split(m_RoomMode, " ")(0) & " " & Split(m_RoomMode, " ")(1)
               End If
            End If
      End Select
      SendServer2 sTemp
End Sub
Public Function GetRoomProp(PropertyName As ePropCodesRead) As String
Dim sMode As String
Dim sTemp As String
Dim i As Long

      If TestForError(True) = True Then Exit Function
      sTemp = "PROP %#" & FixSpaces(m_ActualRoomName, True)
      Select Case PropertyName
         Case SETPROPOnJoin
            sMode = "ONJOIN"
         Case SETPROPOnPart
            sMode = "ONPART"
         Case SETPROPTopic
            sMode = "TOPIC"
         Case SETPROPLag
            sMode = "LAG"
         Case SETPROPLanguage
            sMode = "LANGUAGE"
         Case SETPROPPics
            sMode = "PICS"
         Case GETPROPName
            sMode = "NAME"
         Case GETPROPCreation
            sMode = "CREATION"
         Case GETPROPClient
            sMode = "CLIENT"
         Case GETPROPSubject
            sMode = "SUBJECT"
      End Select
      sTemp = sTemp & " " & sMode
      SendServer2 sTemp

      i = 0
      Do
         i = i + 1
         If i > 300000 Then
            RaiseEvent MSNError(ERR_FailedRetrieveProperty, "Property was not retrieved in a timely manner")
            Exit Function
         End If
         If sRoomProperty <> "" Then
            Exit Do
         End If
         DoEvents
      Loop
      
      If UCase(Split(sRoomProperty, " ")(4)) <> UCase(sMode) Then
         ' Not what was requested
         RaiseEvent MSNError(ERR_FailedRetrieveProperty, "Property was not retrieved in a timely manner")
         Exit Function
      End If
      GetRoomProp = Mid$(Split(sRoomProperty, " ")(5), 2, Len(sRoomProperty))

End Function

Public Sub SetRoomMode(RoomMode As MSNRoomModes, Optional Setting As String)
Dim sMode As String
Dim sTemp As String

      If TestForError(True) = True Then Exit Sub
      Select Case RoomMode
         Case MSNMODESecret_ON
            sMode = "+s"
         Case MSNMODESecret_OFF
            sMode = "-s"
         Case MSNMODEInviteOnly_ON
            sMode = "+i"
         Case MSNMODEInviteOnly_OFF
            sMode = "-i"
         Case MSNMODEPrivate_ON
            sMode = "+p"
         Case MSNMODEPrivate_OFF
            sMode = "-p"
         Case MSNMODEGuestBans_ON
            sTemp = "ACCESS %#" & m_ActualRoomName & " ADD DENY >* "
            If Setting <> "" Then
               sTemp = sTemp & Setting
            End If
         Case MSNMODEGuestBans_OFF
            sTemp = "ACCESS %#" & m_ActualRoomName & " DELETE DENY >*"
         Case MSNMODENoGuestWhispers_on
            sMode = "+W"
         Case MSNMODENoGuestWhispers_OFF
            sMode = "-W"
         Case MSNMODEOnlyHostWhispers_ON
            sMode = "+w"
         Case MSNMODEOnlyHostWhispers_OFF
            sMode = "-w"
         Case MSNMODERoomLimit
            sMode = "+l " & Setting
         Case MSNMODEKnock_ON
            sMode = "+u"
         Case MSNMODEKnock_OFF
            sMode = "-u"
         Case MSNMODEModerated_On
            sMode = "+m"
         Case MSNMODEModerated_Off
            sMode = "-m"
         Case MSNMODEHiddenRoom_ON
            sMode = "+h"
         Case MSNMODEHiddenRoom_OFF
            sMode = "-h"
         Case MSNMODEOnlyHostChangeTopic_ON
            sMode = "+t"
         Case MSNMODEOnlyHostChangeTopic_OFF
            sMode = "-t"
      End Select
      If sTemp = "" Then
         sTemp = "MODE %#" & m_ActualRoomName & " " & sMode
      End If
      SendServer2 sTemp
      DoEvents
      
      If UCase(Left(sTemp, 2)) = "+L" Then
         ' Limit Changed
         m_RoomMode = Split(m_RoomMode, " ")(0) & " " & Split(sTemp, " ")(1)
      Else
         If Left$(sTemp, 1) = "+" Then
            ' Add Mode to room Variable
            m_RoomMode = Replace(m_RoomMode, "+", "+" & Mid$(sTemp, 2, 2))
         End If
         If Left$(sTemp, 1) = "-" Then
            ' Remove Mode from room Variable
            m_RoomMode = Replace(m_RoomMode, Mid$(sTemp, 2, 2), "")
         End If
      End If
End Sub
' MemberInfo=14
Public Sub CreateRoomJoin(CreateModes As MSNRoomCreationSettings)
      bCreateRoom = True
      Set m_CreateModes = CreateModes
      StartConnection
End Sub
' MemberInfo=14
Public Function GETPUID(RealNickName As String) As String
Dim i As Long

      If TestForError(True) = True Then Exit Function

      If m_CurrentUsersList.Item(RealNickName).PUID = "" Then
         ' Get PUID
         sRoomProperty = ""
         SendServer2 "PROP " & RealNickName & " PUID"
         i = 0
         Do
            i = i + 1
            If i > 300000 Then
               RaiseEvent MSNError(ERR_FailedRetrievePUID, "PUID was not retrieved in a timely manner")
               Exit Function
            End If
            If sRoomProperty <> "" Then
               Exit Do
            End If
            DoEvents
         Loop
      
         If UCase(Split(sRoomProperty, " ")(0)) <> UCase(RealNickName) Then
            ' Not what was requested
            RaiseEvent MSNError(ERR_FailedRetrievePUID, "PUID was not retrieved in a timely manner")
            Exit Function
         End If
         GETPUID = Mid$(Split(sRoomProperty, " ")(2), 2, Len(sRoomProperty))
      
         m_CurrentUsersList.Item(RealNickName).PUID = GETPUID
      End If
End Function
' MemberInfo=0
Public Sub SetUserOp(RealName As String, OpType As eHostType)
Dim sTemp As String
Dim sOp As String * 2

      If TestForError(True) = True Then Exit Sub

      sOp = ""
      
      Select Case OpType
         Case htGuest
            sOp = "-o"
         Case htHasVoice
            sOp = "+v"
         Case htHost
            sOp = "+o"
         Case htOwner
            sOp = "+q"
         Case htSpectator
            sOp = "-v"
      End Select
      ' If TestRoomMode(testMODEModerated_ON) Then
      ' If sOp = "-o" Then sOp = "+v"
      ' End If
      If sOp <> "  " Then
         sTemp = "MODE %#" & FixSpaces(m_ActualRoomName, True) & " " & sOp & " " & RealName
         SendServer2 sTemp
      End If
End Sub
' MemberInfo=7,0,0,0
Public Property Get ProfileIcon() As Integer
      ProfileIcon = m_ProfileIcon
End Property
Public Property Let ProfileIcon(ByVal New_ProfileIcon As Integer)
      m_ProfileIcon = New_ProfileIcon
      PropertyChanged "ProfileIcon"
End Property
Private Function TestForError(bTestJoined As Boolean) As Boolean
      If m_isConnected = False Then
         RaiseEvent MSNError(ERR_NotConected, "You Have not connected to the Server")
         TestForError = True
      Else
         If bTestJoined = True Then
            If m_isJoinedRoom = False And m_OverRide = False Then
               RaiseEvent MSNError(ERR_NotJoinedRoom, "You have not Joined a Room")
               TestForError = True
            End If
         End If
      End If
End Function
' MemberInfo=13,0,0,
Public Property Get CNick() As String
Attribute CNick.VB_MemberFlags = "40"
      CNick = m_CNick
End Property
Public Property Let CNick(ByVal New_CNick As String)
      m_CNick = New_CNick
      PropertyChanged "CNick"
End Property
'MemberInfo=0,0,0,False
Public Property Get OverRide() As Boolean
Attribute OverRide.VB_MemberFlags = "40"
    OverRide = m_OverRide
End Property
Public Property Let OverRide(ByVal New_OverRide As Boolean)
    m_OverRide = New_OverRide
    PropertyChanged "OverRide"
End Property
