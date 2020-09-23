VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MSNRoomCreationSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarstgFeature As String
Private mvarstgTopic As String
Private mvarstgWelcomeMsg As String
Private mvarstgCatagory As String
Private mvarstgLocale As String
Private mvarstgLanguage As String
Public Property Let stgLanguage(ByVal vData As String)
      mvarstgLanguage = vData
End Property
Public Property Get stgLanguage() As String
      stgLanguage = mvarstgLanguage
End Property
Public Property Let stgLocale(ByVal vData As String)
      mvarstgLocale = vData
End Property
Public Property Get stgLocale() As String
      stgLocale = mvarstgLocale
End Property
Public Property Let stgCatagory(ByVal vData As String)
      mvarstgCatagory = vData
End Property
Public Property Get stgCatagory() As String
      stgCatagory = mvarstgCatagory
End Property
Public Property Let stgWelcomeMsg(ByVal vData As String)
      mvarstgWelcomeMsg = vData
End Property
Public Property Get stgWelcomeMsg() As String
      stgWelcomeMsg = mvarstgWelcomeMsg
End Property
Public Property Let stgTopic(ByVal vData As String)
      mvarstgTopic = vData
End Property
Public Property Get stgTopic() As String
      stgTopic = mvarstgTopic
End Property
Public Property Let stgFeature(ByVal vData As String)
      mvarstgFeature = vData
End Property
Public Property Get stgFeature() As String
      stgFeature = mvarstgFeature
End Property
