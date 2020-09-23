VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MSNUserObject"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MSNRoomCreationSettings"
Option Explicit

Public Enum eHostType
   htOwner = 1
   htHost = 2
   htGuest = 0
   htSysop = 3
   htSpectator = 4
   htHasVoice = 5
End Enum

Public Enum eGender
   gnMale = 1
   gnFemale = 2
   gnUnKnown = 0
End Enum

Private mvarDisplayName As String
Private mvarRealName As String
Private mvarStatus As Integer
Private mvarCapsCounter As Integer
Private mvarScrollCounter As Integer
Private mvarScrollTime As String
Private mvarGateKeeperID As String
Private mvarSigninTime As String
Private mvarAway As Boolean
Private mvarHostType As eHostType
Private mvarLastSentence As String
Private mvarGender As eGender
Private mvarHasProfile As Boolean
Private mvarSigninDate As String
Private mvarUserJoinInfo As String
Private mvarHasProfilePicture As Boolean
Private mvarLastData As String
Private mvarPUID As String
Private mvarPrevSentence As String
Public Property Let PrevSentence(ByVal vData As String)
    mvarPrevSentence = vData
End Property
Public Property Get PrevSentence() As String
    PrevSentence = mvarPrevSentence
End Property
Public Property Let PUID(ByVal vData As String)
   mvarPUID = vData
End Property
Public Property Get PUID() As String
   PUID = mvarPUID
End Property
Public Property Let LastData(ByVal vData As String)
   mvarLastData = vData
End Property
Public Property Get LastData() As String
   LastData = mvarLastData
End Property
Public Property Let HasProfilePicture(ByVal vData As Boolean)
   mvarHasProfilePicture = vData
End Property
Public Property Get HasProfilePicture() As Boolean
   HasProfilePicture = mvarHasProfilePicture
End Property
Public Property Let UserJoinInfo(ByVal vData As String)
   mvarUserJoinInfo = vData
End Property
Public Property Get UserJoinInfo() As String
   UserJoinInfo = mvarUserJoinInfo
End Property
Public Property Let SigninDate(ByVal vData As String)
   mvarSigninDate = vData
End Property
Public Property Get SigninDate() As String
   SigninDate = mvarSigninDate
End Property
Public Property Let HasProfile(ByVal vData As Boolean)
   mvarHasProfile = vData
End Property
Public Property Get HasProfile() As Boolean
   HasProfile = mvarHasProfile
End Property
Public Property Let Gender(ByVal vData As eGender)
   mvarGender = vData
End Property
Public Property Get Gender() As eGender
   Gender = mvarGender
End Property
Public Property Let LastSentence(ByVal vData As String)
   mvarLastSentence = vData
End Property
Public Property Get LastSentence() As String
   LastSentence = mvarLastSentence
End Property
Public Property Let HostType(ByVal vData As eHostType)
   mvarHostType = vData
End Property
Public Property Get HostType() As eHostType
   HostType = mvarHostType
End Property
Public Property Let Away(ByVal vData As Boolean)
   mvarAway = vData
End Property
Public Property Get Away() As Boolean
   Away = mvarAway
End Property
Public Property Let SigninTime(ByVal vData As String)
   mvarSigninTime = vData
End Property
Public Property Get SigninTime() As String
   SigninTime = mvarSigninTime
End Property
Public Property Let GateKeeperID(ByVal vData As String)
   mvarGateKeeperID = vData
End Property
Public Property Get GateKeeperID() As String
   GateKeeperID = mvarGateKeeperID
End Property
Public Property Let ScrollTime(ByVal vData As String)
   mvarScrollTime = vData
End Property
Public Property Get ScrollTime() As String
   ScrollTime = mvarScrollTime
End Property
Public Property Let ScrollCounter(ByVal vData As Integer)
   mvarScrollCounter = vData
End Property
Public Property Get ScrollCounter() As Integer
   ScrollCounter = mvarScrollCounter
End Property
Public Property Let CapsCounter(ByVal vData As Integer)
   mvarCapsCounter = vData
End Property
Public Property Get CapsCounter() As Integer
   CapsCounter = mvarCapsCounter
End Property
Public Property Let Status(ByVal vData As Integer)
   mvarStatus = vData
End Property
Public Property Get Status() As Integer
   Status = mvarStatus
End Property
Public Property Let RealName(ByVal vData As String)
   mvarRealName = vData
End Property
Public Property Get RealName() As String
   RealName = mvarRealName
End Property
Public Property Let DisplayName(ByVal vData As String)
   mvarDisplayName = vData
End Property
Public Property Get DisplayName() As String
   DisplayName = mvarDisplayName
End Property
