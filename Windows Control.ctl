VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl Windows 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   435
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   420
   ScaleWidth      =   435
   ToolboxBitmap   =   "Windows Control.ctx":0000
   Begin MSComDlg.CommonDialog comdlgs 
      Left            =   2040
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   1800
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Files|*.*"
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   0
      Picture         =   "Windows Control.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "Windows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Const MIIM_ID = &H2
Private Const MIIM_TYPE = &H10
Private Const MFT_STRING = &H0&
Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" ( _
ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10

Private Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Type SHITEMID
    SHItem As Long
    itemID() As Byte
End Type
Private Type ITEMIDLIST
    shellID As SHITEMID
End Type

Const SF_DESKTOP = &H0
Const SF_PROGRAMS = &H2
Const SF_MYDOCS = &H5
Const SF_FAVORITES = &H6     ' 98+
Const SF_STARTUP = &H7
Const SF_RECENT = &H8
Const SF_SENDTO = &H9
Const SF_STARTMENU = &HB
Const SF_MYMUSIC = &HD       ' Me+
Const SF_DESKTOP2 = &H10
Const SF_NETHOOD = &H13
Const SF_FONTS = &H14
Const SF_SHELLNEW = &H15
Const SF_STARTUP2 = &H18
Const SF_ALLUSERSDESK = &H19
Const SF_APPDATA = &H1A
Const SF_PRINTHOOD = &H1B
Const SF_APPDATA2 = &H1C
Const SF_TEMPINETFILES = &H20
Const SF_COOKIES = &H21
Const SF_HISTORY = &H22
Const SF_ALLUSERSAPPDATA = &H23
Const SF_WINDOWS = &H24
Const SF_WINSYSTEM = &H25
Const SF_PROGFILES = &H26
Const SF_MYPICS = &H27       ' Me+
Const SF_USERDIR = &H28
Const SF_WINSYSTEM2 = &H29
Const SF_COMMON = &H2B


Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwnd As Long, ByVal folderid As Long, shidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal shidl As Long, ByVal shPath As String) As Long


Private Type LUID
         UsedPart As Long
         IgnoredForNowHigh32BitPart As Long
      End Type

      Private Type TOKEN_PRIVILEGES
        PrivilegeCount As Long
        TheLuid As LUID
        Attributes As Long
      End Type

      Private Const EWX_SHUTDOWN As Long = 1
      Private Const EWX_FORCE As Long = 4
      Private Const EWX_REBOOT = 2

      Private Declare Function ExitWindowsEx Lib "user32" (ByVal _
           dwOptions As Long, ByVal dwReserved As Long) As Long

      Private Declare Function GetCurrentProcess Lib "Kernel32" () As Long
      Private Declare Function OpenProcessToken Lib "advapi32" (ByVal _
         ProcessHandle As Long, _
         ByVal DesiredAccess As Long, TokenHandle As Long) As Long
      Private Declare Function LookupPrivilegeValue Lib "advapi32" _
         Alias "LookupPrivilegeValueA" _
         (ByVal lpSystemName As String, ByVal lpName As String, lpLuid _
         As LUID) As Long
      Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
         (ByVal TokenHandle As Long, _
         ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES _
         , ByVal BufferLength As Long, _
      PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
      
      Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
      
      Private Declare Function FindWindowEx Lib "user32" _
    Alias "FindWindowExA" (ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long
   Private Const SPI_SCREENSAVERRUNNING = 97
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
    ByVal nCmdShow As Long) As Long
      
      
      Private Declare Function GetTickCount Lib "Kernel32" () As Long
     
      
      Dim result



Private Sub UserControl_Resize()
UserControl.Height = 420
UserControl.Width = 435
End Sub



Sub OpenCD()
Dim lngreturn As String
lngreturn = mciSendString("set CDAudio door open", strReturn, 127, 0)
End Sub

Sub OpenPath(Path As String)
Dim lSuccess As Long
'-- Shell to default browser
lSuccess = ShellExecute(fH, "Open", Path, 0&, 0&, 10)
End Sub

Sub OpenBrowse()

comdlg.ShowOpen


Dim lSuccess As Long
'-- Shell to default browser
lSuccess = ShellExecute(fH, "Open", comdlg.FileName, 0&, 0&, 10)
End Sub

Sub OpenBrowser(URL As String)
Dim lSuccess As Long
'-- Shell to default browser
lSuccess = ShellExecute(fH, "Open", URL, 0&, 0&, 10)
End Sub

Sub OpenExeCopy()
Dim lSuccess As Long
'-- Shell to default browser
lSuccess = ShellExecute(fH, "Open", App.Path & "\" & App.EXEName, 0&, 0&, 10)
End Sub

Sub GetSysDir()
Dim SYSTEMFOLDER As String * 256
Dim sys As String * 256
GetSystemDirectory SYSTEMFOLDER, 256
sys = Left(SYSTEMFOLDER, InStr(SYSTEMFOLDER, Chr(0)) - 1)

Dim lSuccess As Long
'-- Shell to default browser
lSuccess = ShellExecute(fH, "Open", sys, 0&, 0&, 10)
End Sub

Private Function getSpecialFolder(whichFolder As Long) As String
    Dim Path As String * 256
    Dim myid As ITEMIDLIST
    Dim rval As Long

    If IsMissing(useForm) Then
    rval = SHGetSpecialFolderLocation(UserControl.hwnd, whichFolder, myid)
    Else
    rval = SHGetSpecialFolderLocation(UserControl.hwnd, whichFolder, myid)
    End If
    
    If rval = 0 Then ' If success
      rval = SHGetPathFromIDList(ByVal myid.shellID.SHItem, ByVal Path)
        If rval Then ' If True
        getSpecialFolder = Left(Path, InStr(Path, Chr(0)) - 1)
        End If
    End If
    
End Function

Sub GetProgFiles()
Dim lSuccess As Long
'-- Shell to default browser
lSuccess = ShellExecute(fH, "Open", getSpecialFolder(&H26), 0&, 0&, 10)
End Sub

Sub GetFontDir()

Dim lSuccess As Long
'-- Shell to default browser
lSuccess = ShellExecute(fH, "Open", getSpecialFolder(&H14), 0&, 0&, 10)
End Sub

Sub GetStartMenu()


Dim lSuccess As Long
'-- Shell to default browser
lSuccess = ShellExecute(fH, "Open", getSpecialFolder(&HB), 0&, 0&, 10)

End Sub



 Private Sub AdjustToken()
         Const TOKEN_ADJUST_PRIVILEGES = &H20
         Const TOKEN_QUERY = &H8
         Const SE_PRIVILEGE_ENABLED = &H2
         Dim hdlProcessHandle As Long
         Dim hdlTokenHandle As Long
         Dim tmpLuid As LUID
         Dim tkp As TOKEN_PRIVILEGES
         Dim tkpNewButIgnored As TOKEN_PRIVILEGES
         Dim lBufferNeeded As Long

         hdlProcessHandle = GetCurrentProcess()
         OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
            TOKEN_QUERY), hdlTokenHandle

      ' Get the LUID for shutdown privilege.
         LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid

         tkp.PrivilegeCount = 1    ' One privilege to set
         tkp.TheLuid = tmpLuid
         tkp.Attributes = SE_PRIVILEGE_ENABLED

     ' Enable the shutdown privilege in the access token of this process.
         AdjustTokenPrivileges hdlTokenHandle, False, _
         tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded

     End Sub
 
 
Sub ShutDownWin()
 AdjustToken
 ExitWindowsEx (EWX_SHUTDOWN), &HFFFF
End Sub

Public Sub RestartWin()
  AdjustToken
  ExitWindowsEx (EWX_REBOOT), &HFFFF
End Sub

Public Sub LogOffWin()
  AdjustToken
  ExitWindowsEx (EWX_FORCE), &HFFFF

End Sub

Sub GetNotepad()

result = Shell("notepad.exe", vbNormalFocus)

End Sub

Sub GetCtrlPanel()

result = Shell("rundll32.exe shell32.dll,Control_RunDLL", 5)
End Sub

Sub GetCalculator()
result = Shell("calc", 5)
End Sub



Sub HowLongSession()
MsgBox "This windows session has been going for " & Format(GetTickCount / 60000, "0") & " minutes.", vbOKOnly + vbInformation, "Windows Session Length"
End Sub

Sub AboutWinCtrl()
MsgBox "Windows Control. Brought to you by Tom Hendriks Software. Takes the pain out of Windows functions!", vbOKOnly + vbQuestion, "About Windows Control"
End Sub

Sub ShowAboutBox(Text As String, Product As String)
MsgBox Text, vbOKOnly + vbInformation, Product
End Sub
