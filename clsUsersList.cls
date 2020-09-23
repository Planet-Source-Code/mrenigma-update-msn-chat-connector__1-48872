VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MSNUserList"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsUser"
Attribute VB_Ext_KEY = "Member0" ,"clsUser"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(DisplayName As String, RealName As String, Status As Integer, CapsCounter As Integer, ScrollCounter As Integer, ScrollTime As String, GateKeeperID As String, SigninTime As String, Away As Boolean, HostType As eHostType, LastSentence As String, PrevSentence As String, Gender As eGender, HasProfile As Boolean, SigninDate As String, UserJoinInfo As String, HasProfilePicture As Boolean, LastData As String, PUID As String, Optional sKey As String) As MSNUserObject
      ' create a new object
Dim objNewMember As MSNUserObject

      Set objNewMember = New MSNUserObject

      ' set the properties passed into the method
      objNewMember.DisplayName = DisplayName
      objNewMember.RealName = RealName
      objNewMember.Status = Status
      objNewMember.CapsCounter = CapsCounter
      objNewMember.ScrollCounter = ScrollCounter
      objNewMember.ScrollTime = ScrollTime
      objNewMember.GateKeeperID = GateKeeperID
      objNewMember.Away = Away
      objNewMember.HostType = HostType
      objNewMember.LastSentence = LastSentence
      objNewMember.PrevSentence = PrevSentence
      objNewMember.Gender = Gender
      objNewMember.SigninTime = SigninTime
      objNewMember.SigninDate = SigninDate
      objNewMember.UserJoinInfo = UserJoinInfo
      objNewMember.HasProfile = HasProfile
      objNewMember.HasProfilePicture = HasProfilePicture
      objNewMember.LastData = LastData
      objNewMember.PUID = PUID
      If Len(sKey) = 0 Then
         mCol.Add objNewMember
      Else
         mCol.Add objNewMember, sKey
      End If

      ' return the object created
      Set Add = objNewMember
      Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As MSNUserObject
      Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
      Count = mCol.Count
End Property
Public Sub Remove(vntIndexKey As Variant)
      mCol.Remove vntIndexKey
End Sub
Public Property Get NewEnum() As IUnknown
      Set NewEnum = mCol.[_NewEnum]
End Property
Private Sub Class_Initialize()
      ' creates the collection when this class is created
      Set mCol = New Collection
End Sub
Private Sub Class_Terminate()
      ' destroys collection when this class is terminated
      Set mCol = Nothing
End Sub

