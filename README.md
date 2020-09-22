<div align="center">

## Advanced INI Functions


</div>

### Description

This code is the first posted on Planet-Source-Code to do more than just ADD and GET INFO from INI files. You can now also DELETE Keys, Values, EVEN SECTIONS. I am working to create functions to rename sections and keys also. Enjoy this code! Fully explained!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[David Peace](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david-peace.md)
**Level**          |Beginner
**User Rating**    |4.3 (30 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/david-peace-advanced-ini-functions__1-5411/archive/master.zip)

### API Declarations

```
Option Explicit
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
```


### Source Code

```
Function GetKeyVal(ByVal INIFileLoc As String, ByVal Section As String, ByVal Key As String)
'This Function retrieves information from an INI File
'INIFileLoc = The location of the INI File (ex. "C:\Windows\INIFile.ini")
'Section = Section where the Key is held
'Key = The Key of which you want to retrieve information
'Checking to see if the INI File specified exists
If Dir(INIFileLoc) = "" Then MsgBox "File Not Found: " & INIFileLoc & vbCrLf & "Please refer to code in function 'GetKeyVal'", vbExclamation, "INI File Not Found": Exit Function
'If INI File exists then proceed to Get Key Value
Dim RetVal As String, Worked As Integer
RetVal = String$(255, 0)
Worked = GetPrivateProfileString(Section, Key, "", RetVal, Len(RetVal), INIFileLoc)
If Worked = 0 Then
  GetINI = ""
Else
  GetINI = Left(RetVal, InStr(RetVal, Chr(0)) - 1)
End If
End Function
Function AddToINI(ByVal INIFileLoc As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
'This Function adds a Section, Key, or Value to an INI file
'Also used to CREATE NEW INI FILE
'INIFileLoc = The location of the INI File (ex. "C:\Windows\INIFile.ini")
'Section = The name of the referred to Section or newly created Section (ex. "New Section 1")
'Key = The name of the referred to Key or newly created Key (ex. "New Key 1")
'Value = The value to hold in the given Key (ex. "New Info Held")
'Checking to see if the INI File specified exists
If Dir(INIFileLoc) = "" Then MsgBox "File Not Found: " & INIFileLoc & vbCrLf & "Please refer to code in function 'AddToINI'", vbExclamation, "INI File Not Found": Exit Function
'If INI File exists then proceed to Add the information to the INI File
WritePrivateProfileString Section, Key, Value, INIFileLoc
End Function
Function DeleteSection(ByVal INIFileLoc As String, ByVal Section As String)
'This Function Deletes a specified Section from an INI file
'INIFileLoc = The location of the INI File (ex. "C:\Windows\INIFile.ini")
'Section = The name of the Section you wish to remove (ex. "Section Number 1")
'Checking to see if the INI File specified exists
If Dir(INIFileLoc) = "" Then MsgBox "File Not Found: " & INIFileLoc & vbCrLf & "Please refer to code in function 'DeleteSection'", vbExclamation, "INI File Not Found": Exit Function
'If INI File exists then proceed to delete Section
WritePrivateProfileString Section, vbNullString, vbNullString, INIFileLoc
'NOTE: vbNullString is the coding in which to delete a Section, or Key
End Function
Function DeleteKey(ByVal INIFileLoc As String, ByVal Section As String, ByVal Key As String)
'This Function Deletes a Key in a specified Section from an INI file
'INIFileLoc = The location of the INI File (ex. "C:\Windows\INIFile.ini")
'Section = The name of the Section in which the Key to be deleted is held (ex. "Section Number 1")
'Key = The name of the Key you wish to remove (ex. "Key Number 5")
'Checking to see if the INI File specified exists
If Dir(INIFileLoc) = "" Then MsgBox "File Not Found: " & INIFileLoc & vbCrLf & "Please refer to code in function 'DeleteKey'", vbExclamation, "INI File Not Found": Exit Function
'If INI File exists then proceed to delete Key
WritePrivateProfileString Section, Key, vbNullString, INIFileLoc
'NOTE: vbNullString is the coding in which to delete a Section, or Key
End Function
Function DeleteKeyValue(ByVal INIFileLoc As String, ByVal Section As String, ByVal Key As String)
'This Function deletes the value in a specified Key from an INI file
'INIFileLoc = The location of the INI File (ex. "C:\Windows\INIFile.ini")
'Section = The name of the Section in which the Key is held (ex. "Section Number 1")
'Key = The name of the Key you wish to remove the value from (ex. "Key Number 5")
'Checking to see if the INI File specified exists
If Dir(INIFileLoc) = "" Then MsgBox "File Not Found: " & INIFileLoc & vbCrLf & "Please refer to code in function 'DeleteKeyValue'", vbExclamation, "INI File Not Found": Exit Function
'If INI File exists then proceed to delete Key Value
WritePrivateProfileString Section, Key, "", INIFileLoc
' "" = is a short way of saying Nothing
End Function
Function RenameSection()
'Coming Soon
End Function
Function RenameKey()
'Coming Soon
End Function
```

