VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MSNFontInfo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum eCols
   CBlack = 1
   CWhite = 2
   CMaroon = 3
   CGreen = 4
   CNavy = 5
   COlive = 6
   CPurple = 7
   CTeal = 8
   CSilver = 9
   CGray = 10
   CRed = 11
   CLime = 12
   CBlue = 13
   CYellow = 14
   CFuschia = 15
   CAqua = 16
End Enum
Public Enum eFontStyle
   stNormal = 1
   stBold = 2
   stItalic = 3
   stBoldItalic = 4
End Enum

Private Const sCols  As String = "0,15,4,2,1,6,5,3,7,8,12,10,9,14,13,11"

Private mvarFontName As String
Private mvarQBColour As Integer
Private mvarMSNColour As eCols
Private mvarFontStyle As eFontStyle

Public Property Let FontStyle(ByVal vData As eFontStyle)
    mvarFontStyle = vData
End Property
Public Property Get FontStyle() As eFontStyle
    FontStyle = mvarFontStyle
End Property
Public Property Let MSNColour(ByVal vData As eCols)
    mvarMSNColour = vData
End Property
Public Property Get MSNColour() As eCols
    MSNColour = mvarMSNColour
End Property
Public Property Get GetQBColour() As Integer
    GetQBColour = Split(sCols, ",")(mvarMSNColour - 1)
End Property
Public Property Let FontName(ByVal vData As String)
    mvarFontName = vData
End Property
Public Property Get FontName() As String
    FontName = mvarFontName
End Property

