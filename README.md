<div align="center">

## Add32Font, Add16Font,AddNTFont


</div>

### Description

How to install a font in WIN16/WIN32
 
### More Info
 
First copy the file to c:\windows\system (in Win 3.1 and Win NT) or to

c:\windows\fonts in Win 95 and call AddFont16 or AddFont32 from the

following code with the name of the font file; e.g. to install arial.ttf,

copy arial.ttf to \windows\system and then call AddFont16("arial.ttf")


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Unknown
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-add32font-add16font-addntfont__1-255/archive/master.zip)

### API Declarations

```
#If Win16 Then
    Private Declare Function CreateScalableFontResource% Lib "GDI"
(ByVal fHidden%, ByVal lpszResourceFile$, ByVal lpszFontFile$, ByVal
lpszCurrentPath$)
    Private Declare Function AddFontResource Lib "GDI" (ByVal
lpFilename As Any) As Integer
    Private Declare Function WriteProfileString Lib "Kernel" (ByVal
lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As
String) As Integer
    Private Declare Function SendMessage Lib "User" (ByVal hWnd As
Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As
Long
    Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer As
String, ByVal nSize As Integer) As Integer
    Private Const HWND_BROADCAST As Integer = &HFFFF
    Private Const WM_FONTCHANGE As Integer = &H1D
  #End If
  #If Win32 Then
    '32-bit declares
    Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
      ' Maintenance string for PSS usage
    End Type
    Private Declare Function PostMessage Lib "user32" _
      Alias "PostMessageA" (ByVal hWnd As Long, ByVal _
      wMsg As Long, ByVal wParam As Long, ByVal _
      lParam As Long) As Long
    Private Declare Function AddFontResource Lib "gdi32" _
      Alias "AddFontResourceA" (ByVal lpFilename As _
      String) As Long
    Private Declare Function CreateScalableFontResource _
      Lib "gdi32" Alias "CreateScalableFontResourceA" _
      (ByVal fHidden As Long, ByVal lpszResourceFile _
      As String, ByVal lpszFontFile As String, ByVal _
      lpszCurrentPath As String) As Long
    Private Declare Function RemoveFontResource Lib _
      "gdi32" Alias "RemoveFontResourceA" (ByVal _
      lpFilename As String) As Long
    Private Declare Function GetWindowsDirectory Lib _
      "kernel32" Alias "GetWindowsDirectoryA" (ByVal _
      lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function GetSystemDirectory Lib _
      "kernel32" Alias "GetWindowsDirectoryA" (ByVal _
      lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function RegSetValueEx Lib _
      "advapi32.dll" Alias "RegSetValueExA" (ByVal _
      hKey As Long, ByVal lpValueName As String, _
      ByVal Reserved As Long, ByVal dwType As Long, _
      lpData As Any, ByVal cbData As Long) As Long
    Private Declare Function RegOpenKey Lib _
      "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey _
      As Long, ByVal lpSubKey As String, phkResult _
      As Long) As Long
    Private Declare Function RegCloseKey Lib _
      "advapi32.dll" (ByVal hKey As Long) As Long
    Private Declare Function RegDeleteValue Lib _
      "advapi32.dll" Alias "RegDeleteValueA" (ByVal _
      hKey As Long, ByVal lpValueName As String) As Long
    Private Declare Function GetVersionEx Lib "kernel32" _
    Alias "GetVersionExA" (lpVersionInformation As _
    OSVERSIONINFO) As Long
    ' dwPlatformId defines:
    Private Const VER_PLATFORM_WIN32_NT = 2
    Private Const HWND_BROADCAST = &HFFFF&
    Private Const WM_FONTCHANGE = &H1D
    Private Const MAX_PATH = 260
    Private Const HKEY_LOCAL_MACHINE = &H80000002
    Private Const REG_SZ = 1  ' Unicode null terminated string
  #End If
```


### Source Code

```
Private Sub Add32Font(Filename As String)
  #If Win32 Then
    Dim lResult As Long
    Dim strFontPath As String, strFontname As String
    Dim hKey As Long
    'This is the font name and path
    strFontPath = Space$(MAX_PATH)
    strFontname = Filename
    If NT Then
      'Windows NT - Call and get the path to the
      '\windows\system directory
      lResult = GetWindowsDirectory(strFontPath, _
        MAX_PATH)
      If lResult <> 0 Then Mid$(strFontPath, _
        lResult + 1, 1) = "\"
      strFontPath = RTrim$(strFontPath)
    Else
      'Win95 - Call and get the path to the
      '\windows\fonts directory
      lResult = GetWindowsDirectory(strFontPath, _
        MAX_PATH)
      If lResult <> 0 Then Mid$(strFontPath, _
        lResult + 1) = "\fonts\"
      strFontPath = RTrim$(strFontPath)
    End If
    'This Actually adds the font to the system's available
    'fonts for this windows session
    lResult = AddFontResource(strFontPath + strFontname)
    ' If lResult = 0 Then MsgBox "Error Occured " & _
      "Calling AddFontResource"
    'Write the registry value to permanently install the
    'font
    lResult = RegOpenKey(HKEY_LOCAL_MACHINE, _
      "software\microsoft\windows\currentversion\" & _
      "fonts", hKey)
    lResult = RegSetValueEx(hKey, "Proscape Font " & strFontname & _
      " (TrueType)", 0, REG_SZ, ByVal strFontname, _
      Len(strFontname))
    lResult = RegCloseKey(hKey)
    'This call broadcasts a message to let all top-level
    'windows know that a font change has occured so they
    'can reload their font list
    lResult = PostMessage(HWND_BROADCAST, WM_FONTCHANGE, _
      0, 0)
    ' MsgBox "Font Added!"
  #End If
End Sub
Private Function NT() As Boolean
  #If Win32 Then
    Dim lResult As Long
    Dim vi As OSVERSIONINFO
    vi.dwOSVersionInfoSize = Len(vi)
    lResult = GetVersionEx(vi)
    If vi.dwPlatformId And VER_PLATFORM_WIN32_NT Then
      NT = True
    Else
      NT = False
    End If
  #End If
End Function
Public Sub Add16Font(Filename As String)
  #If Win16 Then
    On Error Resume Next
    Dim sName As String, sFont As String, sDir As String, I As Integer
Dim r as Long
    ' Windows' System directory
    sDir = GetWinSysDir()
    ' Name of font resource file
    I = InStr(Filename, ".")
    If I > 0 Then
      sFont = Left(Filename, I - 1) + ".fot"
    Else
      sFont = Filename + ".fot"
    End If
    sFont = sDir & "\" & sFont
    Kill sDir & "\" & sFont
    sName = "Font " & Filename & " (True Type)"
    r = CreateScalableFontResource%(0, sFont, Filename, sDir)  '
Create the font resource file
    r = AddFontResource(sFont)                  ' Add
resource to Windows font table
    r = WriteProfileString("Fonts", sName, sFont)        ' Make
changes to WIN.INI to reflect new font
    r = SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0, 0&)    ' Let
other applications know of the change:
  #End If
End Sub
Function GetWinSysDir() As String
  #If Win16 Then
    ' returns Windows System directory
    Dim Buffer As String * 254, r As Integer, sDir As String
    r = GetSystemDirectory(Buffer, 254)
    sDir = Left(Buffer, r)
    If Right(sDir, 1) = "\" Then sDir = Left(sDir, Len(sDir) - 1)
    GetWinSysDir = sDir
  #End If
End Function
Function GetWinDir() As String
  #If Win32 Then
    ' returns Windows directory
    Dim Buffer As String * 254, r As Long, sDir As String
    r = GetWindowsDirectory(Buffer, 254)
    sDir = Left(Buffer, r)
    If Right(sDir, 1) = "\" Then sDir = Left(sDir, Len(sDir) - 1)
    GetWinDir = sDir
  #End If
End Function
Public Function Reverse(Text As String) As String
  On Error Resume Next
  Dim I%, mx%, result$
  mx = Len(Text)
  For I = mx To 1 Step -1
    result = result + Mid$(Text, I, 1)
  Next
  Reverse = result
End Function
```

