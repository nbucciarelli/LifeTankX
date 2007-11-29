Attribute VB_Name = "basRegistry"
Option Explicit

' --------------------------------------------------------------
' Update the Windows registry.
' Written by Kenneth Ives                        kenaso@home.com
' NT tested by Brett Gerhardi       Brett.Gerhardi@trinite.co.uk
'
' Perform the four basic functions on the Windows registry.
'           Add
'           Change
'           Delete
'           Query
'
' Important:   If you treat all key data strings as being
'              case sensitive, you should never have a problem.
'              Always backup your registry files (System.dat
'              and User.dat) before performing any type of
'              modifications
'
' Software developers vary on where they want to update the
' registry with their particular information.  The most common
' are in HKEY_lOCAL_MACHINE or HKEY_CURRENT_USER.
'
' This BAS module handles all of my needs for string and
' basic numeric updates in the Windows registry.
'
' Brett found that NT users must delete each major key
' separately.  See bottom of TEST routine for an example.
' --------------------------------------------------------------

' --------------------------------------------------------------
' Private variables
' --------------------------------------------------------------
  Private m_lngRetVal As Long
  
' --------------------------------------------------------------
' Constants required for values in the keys
' --------------------------------------------------------------
  Private Const REG_NONE As Long = 0                  ' No value type
  Private Const REG_SZ As Long = 1                    ' nul terminated string
  Private Const REG_EXPAND_SZ As Long = 2             ' nul terminated string w/enviornment var
  Private Const REG_BINARY As Long = 3                ' Free form binary
  Private Const REG_DWORD As Long = 4                 ' 32-bit number
  Private Const REG_DWORD_LITTLE_ENDIAN As Long = 4   ' 32-bit number (same as REG_DWORD)
  Private Const REG_DWORD_BIG_ENDIAN As Long = 5      ' 32-bit number
  Private Const REG_LINK As Long = 6                  ' Symbolic Link (unicode)
  Private Const REG_MULTI_SZ As Long = 7              ' Multiple Unicode strings
  Private Const REG_RESOURCE_LIST As Long = 8         ' Resource list in the resource map
  Private Const REG_FULL_RESOURCE_DESCRIPTOR As Long = 9 ' Resource list in the hardware description
  Private Const REG_RESOURCE_REQUIREMENTS_LIST As Long = 10

' --------------------------------------------------------------
' Registry Specific Access Rights
' --------------------------------------------------------------
  Private Const KEY_QUERY_VALUE As Long = &H1
  Private Const KEY_SET_VALUE As Long = &H2
  Private Const KEY_CREATE_SUB_KEY As Long = &H4
  Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
  Private Const KEY_NOTIFY As Long = &H10
  Private Const KEY_CREATE_LINK As Long = &H20
  Private Const KEY_ALL_ACCESS As Long = &H3F

' --------------------------------------------------------------
' Constants required for key locations in the registry
' --------------------------------------------------------------
  Public Const HKEY_CLASSES_ROOT As Long = &H80000000
  Public Const HKEY_CURRENT_USER As Long = &H80000001
  Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
  Public Const HKEY_USERS As Long = &H80000003
  Public Const HKEY_PERFORMANCE_DATA As Long = &H80000004
  Public Const HKEY_CURRENT_CONFIG As Long = &H80000005
  Public Const HKEY_DYN_DATA As Long = &H80000006

' --------------------------------------------------------------
' Constants required for return values (Error code checking)
' --------------------------------------------------------------
  Private Const ERROR_SUCCESS As Long = 0
  Private Const ERROR_ACCESS_DENIED As Long = 5
  Private Const ERROR_NO_MORE_ITEMS As Long = 259

' --------------------------------------------------------------
' Open/Create constants
' --------------------------------------------------------------
  Private Const REG_OPTION_NON_VOLATILE As Long = 0
  Private Const REG_OPTION_VOLATILE As Long = &H1

' --------------------------------------------------------------
' Declarations required to access the Windows registry
' --------------------------------------------------------------
  Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal lngRootKey As Long) As Long
  
  Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
            (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  
  Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
            (ByVal lngRootKey As Long, ByVal lpSubKey As String) As Long
  
  Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
            (ByVal lngRootKey As Long, ByVal lpValueName As String) As Long
  
  Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
            (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  
  Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
            (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
             lpType As Long, lpData As Any, lpcbData As Long) As Long
  
  Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
            (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
             ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Function regDelete_Sub_Key(ByVal lngRootKey As Long, _
                                  ByVal strRegKeyPath As String, _
                                  ByVal strRegSubKey As String)
    
' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   Function for removing a sub key.
'
' Parameters:
'           lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                  HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'    strRegKeyPath - is name of the key path you wish to traverse.
'     strRegSubKey - is the name of the key which will be removed.
'
' Syntax:
'    regDelete_Sub_Key HKEY_CURRENT_USER, _
                  "Software\AAA-Registry Test\Products", "StringTestData"
'
' Removes the sub key "StringTestData"
' --------------------------------------------------------------
    
' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim lngKeyHandle As Long
  
' --------------------------------------------------------------
' Make sure the key exist before trying to delete it
' --------------------------------------------------------------
  If regDoes_Key_Exist(lngRootKey, strRegKeyPath) Then
  
      ' Get the key handle
      m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
      
      ' Delete the sub key.  If it does not exist, then ignore it.
      m_lngRetVal = RegDeleteValue(lngKeyHandle, strRegSubKey)
  
      ' Always close the handle in the registry.  We do not want to
      ' corrupt the registry.
      m_lngRetVal = RegCloseKey(lngKeyHandle)
  End If
  
End Function

Public Function regDoes_Key_Exist(ByVal lngRootKey As Long, _
                                  ByVal strRegKeyPath As String) As Boolean
    
' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   Function to see if a key does exist
'
' Parameters:
'           lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                  HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'    strRegKeyPath - is name of the key path you want to test
'
' Syntax:
'    strKeyQuery = regQuery_A_Key(HKEY_CURRENT_USER, _
'                       "Software\AAA-Registry Test\Products")
'
' Returns the value of TRUE or FALSE
' --------------------------------------------------------------
    
' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim lngKeyHandle As Long

' --------------------------------------------------------------
' Initialize variables
' --------------------------------------------------------------
  lngKeyHandle = 0
  
' --------------------------------------------------------------
' Query the key path
' --------------------------------------------------------------
  m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  
' --------------------------------------------------------------
' If no key handle was found then there is no key.  Leave here.
' --------------------------------------------------------------
  If lngKeyHandle = 0 Then
      regDoes_Key_Exist = False
  Else
      regDoes_Key_Exist = True
  End If
  
' --------------------------------------------------------------
' Always close the handle in the registry.  We do not want to
' corrupt these files.
' --------------------------------------------------------------
  m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Function

Public Function regQuery_A_Key(ByVal lngRootKey As Long, _
                               ByVal strRegKeyPath As String, _
                               ByVal strRegSubKey As String) As Variant
' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   Function for querying a sub key value.
'
' Parameters:
'           lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                  HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'    strRegKeyPath - is name of the key path you wish to traverse.
'     strRegSubKey - is the name of the key which will be queryed.
'
' Syntax:
'    strKeyQuery = regQuery_A_Key(HKEY_CURRENT_USER, _
'                       "Software\AAA-Registry Test\Products", _
                        "StringTestData")
'
' Returns the key value of "StringTestData"
' --------------------------------------------------------------
    
' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim intPosition As Integer
  Dim lngKeyHandle As Long
  Dim lngDataType As Long
  Dim lngBufferSize As Long
  Dim lngBuffer As Long
  Dim strBuffer As String

' --------------------------------------------------------------
' Initialize variables
' --------------------------------------------------------------
  lngKeyHandle = 0
  lngBufferSize = 0
  
' --------------------------------------------------------------
' Query the key path
' --------------------------------------------------------------
  m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  
' --------------------------------------------------------------
' If no key handle was found then there is no key.  Leave here.
' --------------------------------------------------------------
  If lngKeyHandle = 0 Then
      regQuery_A_Key = ""
      m_lngRetVal = RegCloseKey(lngKeyHandle)   ' always close the handle
      Exit Function
  End If
  
' --------------------------------------------------------------
' Query the registry and determine the data type.
' --------------------------------------------------------------
  m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, _
                         lngDataType, ByVal 0&, lngBufferSize)
  
' --------------------------------------------------------------
' If no key handle was found then there is no key.  Leave.
' --------------------------------------------------------------
  If lngKeyHandle = 0 Then
      regQuery_A_Key = ""
      m_lngRetVal = RegCloseKey(lngKeyHandle)   ' always close the handle
      Exit Function
  End If
  
' --------------------------------------------------------------
' Make the API call to query the registry based on the type
' of data.
' --------------------------------------------------------------
  Select Case lngDataType
         Case REG_SZ:       ' String data (most common)
              ' Preload the receiving buffer area
              strBuffer = Space(lngBufferSize)
      
              m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, 0&, _
                                     ByVal strBuffer, lngBufferSize)
              
              ' If NOT a successful call then leave
              If m_lngRetVal <> ERROR_SUCCESS Then
                  regQuery_A_Key = ""
              Else
                  ' Strip out the string data
                  intPosition = InStr(1, strBuffer, Chr(0))  ' look for the first null char
                  If intPosition > 0 Then
                      ' if we found one, then save everything up to that point
                      regQuery_A_Key = Left(strBuffer, intPosition - 1)
                  Else
                      ' did not find one.  Save everything.
                      regQuery_A_Key = strBuffer
                  End If
              End If
              
         Case REG_DWORD:    ' Numeric data (Integer)
              m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                     lngBuffer, 4&)  ' 4& = 4-byte word (long integer)
              
              ' If NOT a successful call then leave
              If m_lngRetVal <> ERROR_SUCCESS Then
                  regQuery_A_Key = ""
              Else
                  ' Save the captured data
                  regQuery_A_Key = lngBuffer
              End If
         
         Case Else:    ' unknown
              regQuery_A_Key = ""
  End Select
  
' --------------------------------------------------------------
' Always close the handle in the registry.  We do not want to
' corrupt these files.
' --------------------------------------------------------------
  m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Function
Public Sub regCreate_Key_Value(ByVal lngRootKey As Long, ByVal strRegKeyPath As String, _
                               ByVal strRegSubKey As String, varRegData As Variant)
    
' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   Function for saving string data.
'
' Parameters:
'           lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                  HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'    strRegKeyPath - is name of the key path you wish to traverse.
'     strRegSubKey - is the name of the key which will be updated.
'       varRegData - Update data.
'
' Syntax:
'    regCreate_Key_Value HKEY_CURRENT_USER, _
'                      "Software\AAA-Registry Test\Products", _
'                      "StringTestData", "22 Jun 1999"
'
' Saves the key value of "22 Jun 1999" to sub key "StringTestData"
' --------------------------------------------------------------
    
' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim lngKeyHandle As Long
  Dim lngDataType As Long
  Dim lngKeyValue As Long
  Dim strKeyValue As String
  
' --------------------------------------------------------------
' Determine the type of data to be updated
' --------------------------------------------------------------
  If IsNumeric(varRegData) Then
      lngDataType = REG_DWORD
  Else
      lngDataType = REG_SZ
  End If
  
' --------------------------------------------------------------
' Query the key path
' --------------------------------------------------------------
  m_lngRetVal = RegCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)
    
' --------------------------------------------------------------
' Update the sub key based on the data type
' --------------------------------------------------------------
  Select Case lngDataType
         Case REG_SZ:       ' String data
              strKeyValue = Trim(varRegData) & Chr(0)     ' null terminated
              m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                          ByVal strKeyValue, Len(strKeyValue))
                                   
         Case REG_DWORD:    ' numeric data
              lngKeyValue = CLng(varRegData)
              m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                          lngKeyValue, 4&)  ' 4& = 4-byte word (long integer)
                                   
  End Select
  
' --------------------------------------------------------------
' Always close the handle in the registry.  We do not want to
' corrupt these files.
' --------------------------------------------------------------
  m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Sub
Public Function regCreate_A_Key(ByVal lngRootKey As Long, ByVal strRegKeyPath As String)

' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   This function will create a new key
'
' Parameters:
'          lngRootKey  - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                  HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'   strRegKeyPath  - is name of the key you wish to create.
'                  to make sub keys, continue to make this
'                  call with each new level.  MS says you
'                  can do this in one call; however, the
'                  best laid plans of mice and men ...
'
' Syntax:
'   regCreate_A_Key HKEY_CURRENT_USER, "Software\AAA-Registry Test"
'   regCreate_A_Key HKEY_CURRENT_USER, "Software\AAA-Registry Test\Products"
' --------------------------------------------------------------

' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim lngKeyHandle As Long
  
' --------------------------------------------------------------
' Create the key.  If it already exist, ignore it.
' --------------------------------------------------------------
  m_lngRetVal = RegCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)

' --------------------------------------------------------------
' Always close the handle in the registry.  We do not want to
' corrupt these files.
' --------------------------------------------------------------
  m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Function
Public Function regDelete_A_Key(ByVal lngRootKey As Long, _
                                ByVal strRegKeyPath As String, _
                                ByVal strRegKeyName As String) As Boolean
    
' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   Function for removing a complete key.
'
' Parameters:
'           lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                        HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'    strRegKeyPath - is name of the key path you wish to traverse.
'   strRegKeyValue - is the name of the key which will be removed.
'
' Returns a True or False on completion.
'
' Syntax:
'    regDelete_A_Key HKEY_CURRENT_USER, "Software", "AAA-Registry Test"
'
' Removes the key "AAA-Registry Test" and all of its sub keys.
' --------------------------------------------------------------
    
' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim lngKeyHandle As Long
  
' --------------------------------------------------------------
' Preset to a failed delete
' --------------------------------------------------------------
  regDelete_A_Key = False
  
' --------------------------------------------------------------
' Make sure the key exist before trying to delete it
' --------------------------------------------------------------
  If regDoes_Key_Exist(lngRootKey, strRegKeyPath) Then
  
      ' Get the key handle
      m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
      
      ' Delete the key
      m_lngRetVal = RegDeleteKey(lngKeyHandle, strRegKeyName)
      
      ' If the value returned is equal zero then we have succeeded
      If m_lngRetVal = 0 Then regDelete_A_Key = True
      
      ' Always close the handle in the registry.  We do not want to
      ' corrupt the registry.
      m_lngRetVal = RegCloseKey(lngKeyHandle)
  End If
  
End Function

Sub Test()

' --------------------------------------------------------------
' Test Windows registry basic functions.
' Written by Kenneth Ives                     kenaso@home.com
'
' Rename this to "Main".  Press F8 to step thru the code.   You
' will be able to stop at will and execute Regedit.exe to see
' the results.  Or, you can press F5 and this test procedure
' has its own stops built in.
'
' Perform the four basic functions on the Windows registry.
'           Add
'           Change
'           Delete
'           Query
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry (System.dat and
'                User.dat) before performing any type of updates.
'
' Rename this procedure back to TEST so as not to intefere if
' this BAS module is used in another application.
' --------------------------------------------------------------
    
' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim lngRootKey As Long
  Dim strKeyQuery As Variant    ' we are not sure what type of
                              ' data will be returned
  
' --------------------------------------------------------------
' Initialize variables
' --------------------------------------------------------------
  strKeyQuery = vbNullString
  lngRootKey = HKEY_CURRENT_USER
  
' --------------------------------------------------------------
' See if the key already exist.  If the key does not exist, we
' will create one.  Some people want to automatically create a
' key if it does not exist.  This philosophy can be dangerous.
' Querying the registry is one function and updating is another.
' --------------------------------------------------------------
  If Not regDoes_Key_Exist(lngRootKey, "Software\AAA-Registry Test") Then
      ' create the main key and the first sub key
      regCreate_A_Key lngRootKey, "Software\AAA-Registry Test"
      regCreate_A_Key lngRootKey, "Software\AAA-Registry Test\Products"
  End If
      
' --------------------------------------------------------------
' see if the next sub key exist.
' --------------------------------------------------------------
  If Not regDoes_Key_Exist(lngRootKey, "Software\AAA-Registry Test\Products") Then
      ' create the first sub key
      regCreate_A_Key lngRootKey, "Software\AAA-Registry Test\Products"
  End If
      
' --------------------------------------------------------------
' Create a string type sub key
' --------------------------------------------------------------
  regCreate_Key_Value lngRootKey, "Software\AAA-Registry Test\Products", _
                      "StringTestData", "22 SEP 1999"
                      
' --------------------------------------------------------------
' Create a numeric type sub key
' --------------------------------------------------------------
  regCreate_Key_Value lngRootKey, "Software\AAA-Registry Test\Products", _
                      "NumericTestData", 1234567890
      
' --------------------------------------------------------------
' See if we have successfully created the key.  The value of
' of the sub key will be returned.  strKeyQuery is a variant
' because we do not know if the data being returned is string
' or numeric.  Once it is returned then we can manipulate it.
' --------------------------------------------------------------
  strKeyQuery = regQuery_A_Key(lngRootKey, "Software\AAA-Registry Test\Products", "StringTestData")
  strKeyQuery = regQuery_A_Key(lngRootKey, "Software\AAA-Registry Test\Products", "NumericTestData")
  
' --------------------------------------------------------------
' Stop processing here.
' Execute Regedit.exe and verify that all the keys have
' been added to the registry.
' Press F5 or F8 to continue.
' --------------------------------------------------------------
  Stop
  
' --------------------------------------------------------------
' Change the value of the sub key, "StringTestData", from
' "22 SEP 1999" to "September 22, 1999"
' --------------------------------------------------------------
  regCreate_Key_Value lngRootKey, "Software\AAA-Registry Test\Products", _
                      "StringTestData", "September 22, 1999"

' --------------------------------------------------------------
' See if the sub key has been updated
' --------------------------------------------------------------
  strKeyQuery = regQuery_A_Key(lngRootKey, "Software\AAA-Registry Test\Products", "StringTestData")
  
' --------------------------------------------------------------
' Stop processing here.
' Execute Regedit.exe and verify that the sub key has
' been updated in the registry.
' Press F5 or F8 to continue.
' --------------------------------------------------------------
  Stop
  
' --------------------------------------------------------------
' Delete the sub key, "NumericTestData", only.
' --------------------------------------------------------------
  regDelete_Sub_Key lngRootKey, "Software\AAA-Registry Test\Products", "NumericTestData"
  
' --------------------------------------------------------------
' Stop processing here.
' Execute Regedit.exe and verify the sub key ("NumericTestData")
' has been removed from the registry.
' Press F5 or F8 to continue.
' --------------------------------------------------------------
  Stop
  
' --------------------------------------------------------------
' Remove the complete key from the registry.  You do not want
' to remove the "Software" key.  NT users must remove each
' major key component separately as shown below.  Windows 95/98
' users can do this in one step by using the second line only.
' --------------------------------------------------------------
  If Not regDelete_A_Key(lngRootKey, "Software\AAA-Registry Test", "Products") Then
      MsgBox "Failed to delete requested subkey!    ", vbOKOnly + vbExclamation, "Registry Key Delete"
      GoTo Normal_Exit:
  End If
  
  If Not regDelete_A_Key(lngRootKey, "Software", "AAA-Registry Test") Then
      MsgBox "Failed to delete requested main key!    ", vbOKOnly + vbExclamation, "Registry Key Delete"
      GoTo Normal_Exit:
  End If
    
    
Normal_Exit:

End Sub
