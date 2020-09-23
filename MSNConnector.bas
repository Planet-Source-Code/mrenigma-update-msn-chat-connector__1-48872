Attribute VB_Name = "Module1"
Option Explicit
Global Const sRoot As String = "SOFTWARE\Classes\CLSID"
Global Const sPassRoot As String = "SOFTWARE\MICROSOFT\MSNCHAT"
Global Const sSubKeyV4 As String = "{29c13b62-b9f7-4cd3-8cef-0a58a1a99441}"
Global Const sKey1  As String = "{E113C6A6-D44A-4639-A40E-3B6DE32A1A40}"
' Private Const ERROR_FILE_NOT_FOUND = 2&
' Private Const ERROR_PATH_NOT_FOUND = 3&
' Private Const ERROR_BAD_FORMAT = 11&
' Private Const SE_ERR_ACCESSDENIED = 5        ' access denied
' Private Const SE_ERR_ASSOCINCOMPLETE = 27
' Private Const SE_ERR_DDEBUSY = 30
' Private Const SE_ERR_DDEFAIL = 29
' Private Const SE_ERR_DDETIMEOUT = 28
' Private Const SE_ERR_DLLNOTFOUND = 32
' Private Const SE_ERR_FNF = 2                ' file not found
' Private Const SE_ERR_NOASSOC = 31
' Private Const SE_ERR_PNF = 3                ' path not found
' Private Const SE_ERR_OOM = 8                ' out of memory
' Private Const SE_ERR_SHARE = 26
Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long
Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Global fso As New FileSystemObject

Public Function BuildHTML(sRoomName As String, _
      sNick As String, _
      sServerIP As String, _
      sServerPORT As String, _
      sChatOCXVersion As String, _
      sChatCLASSID As String, _
      bUsePassPort As Boolean, _
      sCookie As String, _
      sTicket As String, _
      sProfile As String, _
      sFileName As String, _
      sCabFilePath As String, _
      Optional bHex As Boolean, _
      Optional ProfileIcon As Integer = 0, _
      Optional CreateSettings As MSNRoomCreationSettings) As String
     
Dim sFile As String
Dim sTemp As String
Dim sServerAddress As String
Dim sRoomLabel As String
Dim iFile As Integer
Dim sOCXInfo As String
Dim sClassID As String
Dim sPN As String
Dim sVL As String
Dim sEnd As String

      sPN = "<PARAM NAME=" & Chr(34)
      sVL = Chr(34) & " VALUE=" & Chr(34)
      sEnd = Chr(34) & ">" & vbNewLine
      sRoomLabel = "RoomName"
      If bHex Then
         sRoomLabel = "HexRoomName"
      End If
        
      sServerAddress = sServerIP & ":" & Trim(sServerPORT)
      If sCabFilePath = "" Then
         sOCXInfo = "http://fdl.msn.com/public/chat/msnchat45.cab#Version=" & sChatOCXVersion
      Else
         sOCXInfo = sCabFilePath & "#Version=" & sChatOCXVersion
      End If
      sFile = "<OBJECT ID=""ChatFrame"" CLASSID=""CLSID:" & sChatCLASSID & """ width=""100%"" height=""100%"" CODEBASE=""" & sOCXInfo & sEnd
      sFile = sFile & sPN & sRoomLabel & sVL & FixSpaces(sRoomName, True) & sEnd
      sFile = sFile & sPN & "Server" & sVL & sServerAddress & sEnd
      sFile = sFile & sPN & "BaseURL" & sVL & "http://chat.msn.com/" & sEnd
      
      If bUsePassPort Then
         sFile = sFile & sPN & "NickName" & sVL & "PASSPORT" & sEnd
         sFile = sFile & sPN & "MSNREGCookie" & sVL & Replace(sCookie, vbCrLf, "") & sEnd
         sFile = sFile & sPN & "PassportTicket" & sVL & Replace(sTicket, vbCrLf, "") & sEnd
         sFile = sFile & sPN & "PassportProfile" & sVL & Replace(sProfile, vbCrLf, "") & sEnd
         sFile = sFile & sPN & "MSNPROFILE" & sVL & ProfileIcon & sEnd
      Else
         sFile = sFile & sPN & "NickName" & sVL & sNick & sEnd
      End If
      
      If TypeName(CreateSettings) = "Nothing" Then
         sFile = sFile & sPN & "ChatMode" & sVL & "2" & sEnd
         sFile = sFile & sPN & "Category" & sVL & "UL" & sEnd
         sFile = sFile & sPN & "Locale" & sVL & "EN-GB" & sEnd
         sFile = sFile & sPN & "Topic" & sVL & "." & sEnd
         sFile = sFile & sPN & "WelcomeMsg" & sVL & "." & sEnd
      Else
         sFile = sFile & sPN & "ChatMode" & sVL & "1" & sEnd
         sFile = sFile & sPN & "Category" & sVL & CreateSettings.stgCatagory & sEnd
         sFile = sFile & sPN & "Locale" & sVL & CreateSettings.stgLocale & sEnd
         sFile = sFile & sPN & "Topic" & sVL & CreateSettings.stgTopic & sEnd
         sFile = sFile & sPN & "WelcomeMsg" & sVL & CreateSettings.stgWelcomeMsg & sEnd
         If CreateSettings.stgFeature <> "" Then
            sFile = sFile & sPN & "Feature" & sVL & CreateSettings.stgFeature & sEnd
         End If
      End If
      ' Stop
      iFile = FreeFile
        
      Open sFileName For Output As #iFile
        
      Print #iFile, sFile
      Close #iFile
      BuildHTML = sFile

End Function
Public Function FixSpaces(sText As String, bNoSpaces As Boolean) As String
      If bNoSpaces Then
         FixSpaces = Replace(Replace(sText, " ", "\b"), ",", "\c")
      Else
         FixSpaces = Replace(Replace(sText, "\b", " "), "\c", ",")
      End If
End Function
Function Locate(sData As String, sFind As String) As Long
      Locate = InStr(1, sData, sFind, vbBinaryCompare)
End Function
Public Function GetLine(sData As String, sFind As String) As String
Dim asTemp() As String
Dim i As Integer
        
      asTemp = Split(sData, vbCrLf)
      For i = 0 To UBound(asTemp)
         If Locate(asTemp(i), sFind) Then
            GetLine = asTemp(i)
         End If
      Next
End Function
Public Function GetAfter(sData As String, sFind As String) As String
      On Error Resume Next
      GetAfter = Mid$(sData, Locate(UCase(sData), UCase(sFind)) + Len(sFind), Len(sData))
End Function
Public Function GetBefore(sData As String, sFind As String) As String
      On Error Resume Next
      GetBefore = Left$(sData, Locate(sData, sFind) - 1)
End Function
Public Function TestNick(sNickName As String, Optional bRemove As Boolean) As String
      TestNick = sNickName
      If bRemove Then

         TestNick = Replace(TestNick, "Guest_", ">")
         TestNick = Replace(TestNick, "(Host)", "")
      Else
         If Left$(sNickName, 1) = ">" Then
            TestNick = "Guest_" & Mid$(sNickName, 2, Len(sNickName))
         End If
      End If
End Function
Public Sub GetNickDetails(sData As String, sNick As String, sGate As String)
      sNick = GetBefore(sData, "!")
      sNick = Mid$(sNick, 2, Len(sNick))
      sGate = GetBefore(sData, "@")
      sGate = GetAfter(sGate, "!")
End Sub

Public Sub TextGotFocus(oTextBox As TextBox)
      oTextBox.SelStart = 0
      oTextBox.SelLength = Len(oTextBox.Text)

End Sub
Public Sub TextLostFocus(oTextBox As TextBox)
      oTextBox.SelLength = 0

End Sub


Public Function Encrypt1(sString As String) As String
Dim S As String
Dim i As Integer
Dim sHex As String
Dim sNewHex As String
Dim iDec As Integer
Dim sTmp As String

      ' This method will convert each character in the string
      ' into hex then swap the hex around then turn it back to
      ' decimal and turn it into a char
      
      For i = 1 To Len(sString)
         ' Turn Char into a hex string
       
         sHex = Hex$(Asc(Mid$(sString, i, 1)))
         
         If Len(sHex) = 1 Then
            ' Lets pad the string with zeros
            sHex = "0" & sHex
         End If

         ' Swap Hex awound for example
         ' 6E becomes E6

         sNewHex = Right$(sHex, 1) & Left$(sHex, 1)
         
         ' Convert the new hex into decimal
         
         iDec = Val("&H" & sNewHex)
         
         If iDec > 0 Then
            ' now add the char value to the new string
            
            sTmp = sTmp & Chr(iDec)
            If Len(sTmp) > 50 Then
               ' This increase performance on large strings
               S = S & sTmp
               sTmp = ""
            End If
                        
         End If
      Next
      S = S & sTmp
      Encrypt1 = S
End Function


Public Function Convert2Hex(sString As String) As String
Dim i As Integer
Dim sChr As String * 1
Dim iDec As Integer
      
      For i = 1 To Len(sString)
         sChr = Mid$(sString, i, 1)
         iDec = Asc(sChr)
         Convert2Hex = Convert2Hex & Hex(iDec)
      Next
End Function

Public Function ConvertHex(sHex As String) As String
Dim i As Integer
Dim sChr As String * 1
Dim iDec As Integer

      For i = 1 To Len(sHex) Step 2
         sChr = Chr(Val("&H" & Mid$(sHex, i, 2)))
         ConvertHex = ConvertHex & sChr
      Next
End Function
Public Function ConvertHexURL(SUrl As String) As String
Dim sChr As String * 1
Dim iDec As Integer
Dim iPos As Integer
Dim sHex As String

      Do
         iPos = InStr(1, SUrl, "%")
         If iPos Then
            sHex = Mid$(SUrl, iPos + 1, 2)
            sChr = Chr(Val("&H" & sHex))
            SUrl = Left$(SUrl, iPos - 1) & sChr & Mid$(SUrl, iPos + 3, Len(SUrl))
         End If
      Loop While iPos > 0
      ConvertHexURL = SUrl
End Function
