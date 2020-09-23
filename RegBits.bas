Attribute VB_Name = "zbasRegBits"
'
' Created by E.Spencer (elliot@spnc.demon.co.uk) - This code is public domain.
'
Option Explicit
Public Const ERROR_FILE_NOT_FOUND = 2&
Public Const ERROR_PATH_NOT_FOUND = 3&
Public Const ERROR_BAD_FORMAT = 11&
Public Const SE_ERR_ACCESSDENIED = 5        ' access denied
Public Const SE_ERR_ASSOCINCOMPLETE = 27
Public Const SE_ERR_DDEBUSY = 30
Public Const SE_ERR_DDEFAIL = 29
Public Const SE_ERR_DDETIMEOUT = 28
Public Const SE_ERR_DLLNOTFOUND = 32
Public Const SE_ERR_FNF = 2                ' file not found
Public Const SE_ERR_NOASSOC = 31
Public Const SE_ERR_PNF = 3                ' path not found
Public Const SE_ERR_OOM = 8                ' out of memory
Public Const SE_ERR_SHARE = 26

Public Enum EShellShowConstants
   essSW_HIDE = 0
   essSW_MAXIMIZE = 3
   essSW_MINIMIZE = 6
   essSW_SHOWMAXIMIZED = 3
   essSW_SHOWMINIMIZED = 2
   essSW_SHOWNORMAL = 1
   essSW_SHOWNOACTIVATE = 4
   essSW_SHOWNA = 8
   essSW_SHOWMINNOACTIVE = 7
   essSW_SHOWDEFAULT = 10
   essSW_RESTORE = 9
   essSW_SHOW = 5
End Enum

' Security Mask constants
Public Const READ_CONTROL As Variant = &H20000
Public Const SYNCHRONIZE As Variant = &H100000
Public Const STANDARD_RIGHTS_ALL As Variant = &H1F0000
Public Const STANDARD_RIGHTS_READ As Variant = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE As Variant = READ_CONTROL
Public Const KEY_QUERY_VALUE As Variant = &H1
Public Const KEY_SET_VALUE As Variant = &H2
Public Const KEY_CREATE_SUB_KEY As Variant = &H4
Public Const KEY_ENUMERATE_SUB_KEYS As Variant = &H8
Public Const KEY_NOTIFY As Variant = &H10
Public Const KEY_CREATE_LINK As Variant = &H20
Public Const KEY_ALL_ACCESS As Variant = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const KEY_READ As Variant = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE As Variant = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_WRITE As Variant = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
' Possible registry data types
Public Enum InTypes_enum
   ValNull = 0
   ValString = 1
   ValXString = 2
   ValBinary = 3
   ValDWord = 4
   ValLink = 6
   ValMultiString = 7
   ValResList = 8
End Enum
' Temporary Stack Storage Variables
Private Temp As Long
Dim TempEx As Long
Dim TempExA As String
Private TempExB&, TempExC%

' Handle And Other Storage Variables
Private hHnd As Long
Private KeyPath As String
Private hDepth As Long

' Variable To Hold Last Error
Public RegLastError As Long

' Registry value type definitions
Public Const REG_NONE As Long = 0
Public Const REG_SZ As Long = 1
Public Const REG_EXPAND_SZ As Long = 2
Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4
Public Const REG_LINK As Long = 6
Public Const REG_MULTI_SZ As Long = 7
Public Const REG_RESOURCE_LIST As Long = 8
' Registry section definitions
Public Enum hKeyNames
   HKEY_CLASSES_ROOT = &H80000000
   HKEY_CURRENT_USER = &H80000001
   HKEY_LOCAL_MACHINE = &H80000002
   HKEY_USERS = &H80000003
   HKEY_PERFORMANCE_DATA = &H80000004
   HKEY_CURRENT_CONFIG = &H80000005
   HKEY_DYN_DATA = &H80000006
End Enum

' Codes returned by Reg API calls
Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_INVALID_PARAMETER = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259
' Registry API functions used in this module (there are more of them)
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Public Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long

' This routine allows you to get values from anywhere in the Registry, it currently
' only handles string, double word and binary values. Binary values are returned as
' hex strings.
'
' Example
' Text1.Text = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\", "DefaultUserName")
'
Public Function ReadRegistry(ByVal Group As Long, ByVal Section As String, ByVal Key As String) As String
Dim lResult As Long, lKeyValue As Long, lDataTypeValue As Long, lValueLength As Long, sValue As String, td As Double
Dim TStr1 As String, TStr2 As String
Dim i As Integer
      On Error Resume Next
      lResult = RegOpenKey(Group, Section, lKeyValue)
      sValue = Space$(2048)
      lValueLength = Len(sValue)
      lResult = RegQueryValueEx(lKeyValue, Key, 0&, lDataTypeValue, sValue, lValueLength)
      If (lResult = 0) And (Err.Number = 0) Then
         If lDataTypeValue = REG_DWORD Then
            td = Asc(Mid$(sValue, 1, 1)) + &H100& * Asc(Mid$(sValue, 2, 1)) + &H10000 * Asc(Mid$(sValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid$(sValue, 4, 1)))
            sValue = Format$(td, "000")
         End If
         If lDataTypeValue = REG_BINARY Then
            ' Return a binary field as a hex string (2 chars per byte)
            TStr2 = ""
            For i = 1 To lValueLength
               TStr1 = Hex(Asc(Mid$(sValue, i, 1)))
               If Len(TStr1) = 1 Then TStr1 = "0" & TStr1
               TStr2 = TStr2 + TStr1
            Next
            sValue = TStr2
         Else
            sValue = Left$(sValue, lValueLength - 1)
         End If
      Else
         sValue = "Not Found"
      End If
      lResult = RegCloseKey(lKeyValue)
      ReadRegistry = sValue
End Function

' This routine allows you to write values into the entire Registry, it currently
' only handles string and double word values.
'
' Example
' WriteRegistry HKEY_CURRENT_USER, "SOFTWARE\My Name\My App\", "NewSubKey", ValString, "NewValueHere"
' WriteRegistry HKEY_CURRENT_USER, "SOFTWARE\My Name\My App\", "NewSubKey", ValDWord, "31"
'
Public Sub WriteRegistry(ByVal Group As Long, ByVal Section As String, ByVal Key As String, ByVal ValType As InTypes_enum, ByVal Value As Variant)
Dim lResult As Long
Dim lKeyValue As Long
Dim InLen As Long
Dim lNewVal As Long
Dim sNewVal As String
Dim i As Integer
Dim lDataSize As Integer
Dim ByteArray() As Byte

      On Error Resume Next
      lResult = RegCreateKey(Group, Section, lKeyValue)
      If ValType = ValDWord Then
         lNewVal = CLng(Value)
         InLen = 4
         lResult = RegSetValueExLong(lKeyValue, Key, 0&, ValType, lNewVal, InLen)
      Else
         ' Fixes empty string bug - spotted by Marcus Jansson
         If ValType = ValString Then Value = Value + Chr(0)
         If ValType = ValBinary Then
            InLen = Len(Value)
            ReDim ByteArray(InLen) As Byte
            For i = 1 To InLen
               ByteArray(i) = Asc(Mid$(Value, i, 1))
            Next
            lResult = RegSetValueExB(lKeyValue, Key, 0&, REG_BINARY, ByteArray(1), InLen)
         Else
            sNewVal = Value
            InLen = Len(sNewVal)
            lResult = RegSetValueExString(lKeyValue, Key, 0&, 1&, sNewVal, InLen)
         End If
      End If
      lResult = RegFlushKey(lKeyValue)
      lResult = RegCloseKey(lKeyValue)
End Sub
Private Function RegCheckError(ByRef ErrorValue As Long) As Boolean

        If ((ErrorValue < 8) And (ErrorValue > 1)) Or _
        (ErrorValue = 87) Or (ErrorValue = 259) Then _
        RegCheckError = -1 Else RegCheckError = 0

End Function


Public Function DeleteRegistry(ByVal hKey As Long, ByVal Section As String, ByVal Value As String) As Boolean

      ' Combine The Key And SubKey Paths
   
      ' Open The Key For Operations
      Temp& = RegOpenKey(hKey, Section$, hHnd&)
    
      ' Process Returned Information
      If RegCheckError(Temp&) Then GoTo DeleteValueError
    
      ' Delete Existing Value From Key
      Temp& = RegDeleteValue(hHnd&, Value)
    
      ' Process Returned Information
      If RegCheckError(Temp&) Then GoTo DeleteValueError
    
      ' Close Handle To Key
      Temp& = RegCloseKey(hHnd&)
    
      ' Operation Was Successful
      DeleteRegistry = -1

      ' Exit Function With Passed Value
      Exit Function

DeleteValueError:
    
      ' Store Error In Variable
      RegLastError = Temp&
    
      ' Operation Was Not Successful
      DeleteRegistry = 0
    
      ' Close Handle To Key
      Temp& = RegCloseKey(hHnd&)
    
End Function



Public Function ShellEx( _
      ByVal sFile As String, _
      Optional ByVal eShowCmd As EShellShowConstants = essSW_SHOWDEFAULT, _
      Optional ByVal sParameters As String = "", _
      Optional ByVal sDefaultDir As String = "", _
      Optional sOperation As String = "open", _
      Optional Owner As Long = 0) As Boolean
        
Dim lR As Long
Dim lErr As Long, sErr As Long
      If (InStr(UCase$(sFile), ".EXE") <> 0) Then
         eShowCmd = 0
      End If
      On Error Resume Next
      If (sParameters = "") And (sDefaultDir = "") Then
         lR = ShellExecuteForExplore(Owner, sOperation, sFile, 0, 0, essSW_SHOWNORMAL)
      Else
         lR = ShellExecute(Owner, sOperation, sFile, sParameters, sDefaultDir, eShowCmd)
      End If
      If (lR < 0) Or (lR > 32) Then
         ShellEx = True
      Else
         ' raise an appropriate error:
         lErr = vbObjectError + 1048 + lR
         Select Case lR
            Case 0
               lErr = 7: sErr = "Out of memory"
            Case ERROR_FILE_NOT_FOUND
               lErr = 53: sErr = "File not found"
            Case ERROR_PATH_NOT_FOUND
               lErr = 76: sErr = "Path not found"
            Case ERROR_BAD_FORMAT
               sErr = "The executable file is invalid or corrupt"
            Case SE_ERR_ACCESSDENIED
               lErr = 75: sErr = "Path/file access error"
            Case SE_ERR_ASSOCINCOMPLETE
               sErr = "This file type does not have a valid file association."
            Case SE_ERR_DDEBUSY
               lErr = 285: sErr = "The file could not be opened because the target application is busy. Please try again in a moment."
            Case SE_ERR_DDEFAIL
               lErr = 285: sErr = "The file could not be opened because the DDE transaction failed. Please try again in a moment."
            Case SE_ERR_DDETIMEOUT
               lErr = 286: sErr = "The file could not be opened due to time out. Please try again in a moment."
            Case SE_ERR_DLLNOTFOUND
               lErr = 48: sErr = "The specified dynamic-link library was not found."
            Case SE_ERR_FNF
               lErr = 53: sErr = "File not found"
            Case SE_ERR_NOASSOC
               sErr = "No application is associated with this file type."
            Case SE_ERR_OOM
               lErr = 7: sErr = "Out of memory"
            Case SE_ERR_PNF
               lErr = 76: sErr = "Path not found"
            Case SE_ERR_SHARE
               lErr = 75: sErr = "A sharing violation occurred."
            Case Else
               sErr = "An error occurred occurred whilst trying to open or print the selected file."
         End Select
                
         Err.Raise lErr, , App.EXEName & ".GShell", sErr
         ShellEx = False
      End If

End Function

