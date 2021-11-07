VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   2475
   ClientTop       =   2775
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "TASKBAR.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2535
   ScaleWidth      =   4710
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Exit"
      Height          =   372
      Left            =   3300
      TabIndex        =   4
      Top             =   1680
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Command1"
      Height          =   432
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1452
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2280
      Top             =   1440
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Command4"
      Height          =   432
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1452
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Command3"
      Height          =   432
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1452
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Command2"
      Height          =   432
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1452
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4140
      Top             =   1920
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TASKBAR.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TASKBAR.frx":04E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TASKBAR.frx":06BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TASKBAR.frx":0898
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   3120
      TabIndex        =   8
      Top             =   900
      Width           =   852
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Icon:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2460
      TabIndex        =   7
      Top             =   900
      Width           =   492
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2460
      TabIndex        =   6
      Top             =   540
      Width           =   2052
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mouse Event:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2460
      TabIndex        =   5
      Top             =   180
      Width           =   1512
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type OsVersionInfo
    dwVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatform As Long
    szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx& Lib "kernel32.dll" Alias "GetVersionExA" (lpStruct As OsVersionInfo)

Private OsVers As OsVersionInfo

Private lTempLong&
Dim Msghook1 As Msghook


' --------------------------------------------------------
' Used with GetVersion and GetWinFlags
' --------------------------------------------------------
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2

'----------------------------------------------------------
'This sample application implements callbacks via the MsgHook
'message handling control.  Windows will generate a User
'Message with an lParam that identifies a mouse event, such
'as WM_MOUSEMOVE, and the wParam will contain the icon number.
'If the application installs more than one icon, each must be
'given a unique ID number.
'-----------------------------------------------------------

'-----------------------------------------------------------
'The 10 available mouse events:
'-----------------------------------------------------------

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209

Dim TaskBr As New CTaskBar
Private iIconUsed&
Private lIconAdded&
Private Sub Command1_Click()

'-----------------------------------------------------
'This button adds a new icon to the tray.
'-----------------------------------------------------
Dim hIcon&
Dim sTip$
    '---------------------------------------------------
    'Select one of the icons from the ImageList.  The
    'ImageList may have 16x16, 32x32, or 48x48 icons, but
    'if it is a bitmap source rather than an icon source
    'this program will fail.  32x32 and larger are scaled
    'down to 16x16, so 16x16 will look the best.
    '---------------------------------------------------
    
    '---------------------------------------------------
    'In a "real-world" application the icon would
    'not be randomly generated, and we would keep track
    'of which icon had which ID number.
    '---------------------------------------------------
    
    iIconUsed = iIconUsed + 1
    If iIconUsed = 5 Then
        iIconUsed = 1
    End If
    hIcon = ImageList1.ListImages(iIconUsed).Picture

    '------------------------------------------------------------
    'Select an ID number for use during callbacks.
    '------------------------------------------------------------
    lIconAdded = lIconAdded + 1
    sTip = "This is icon number " & CStr(lIconAdded)
    TaskBr.AddIcon lIconAdded, sTip, hIcon
End Sub

Private Sub Command2_Click()
    '--------------------------------------------------------
    'This routine changes the text a user will see when placing
    'the mouse over a tray icon.
    '--------------------------------------------------------
    Dim sTip$
    Dim lID&
    lID = Val(InputBox("Icon ID Number", , ""))
    If lID = 0 Then
        Exit Sub
    End If
    sTip = InputBox("New text", , "This is a test")
    If Len(sTip) > 63 Then sTip = Left$(sTip, 63)
    TaskBr.ChangeMessage lID, sTip

End Sub

Private Sub Command3_Click()
    '--------------------------------------------------
    'This routine demonstrates changing the icon under
    'program control.
    'Icon modification would routinely be used as a status
    'indicator.  For example, the Win95 Dial-Up utility
    'places a modem with read and send indicators that cycle
    'between red and green.
    '--------------------------------------------------
    Dim lID&
    Dim hIcon&
    lID = Val(InputBox("Icon ID Number", , ""))
    If lID = 0 Then
        Exit Sub
    End If
    iIconUsed = iIconUsed + 1
    If iIconUsed = 5 Then
        iIconUsed = 1
    End If
    hIcon = ImageList1.ListImages(iIconUsed).Picture
    TaskBr.ChangeIcon lID, hIcon

End Sub

Private Sub Command4_Click()
    '-------------------------------------------
    'Deletes an Icon
    '-------------------------------------------
    Dim lID&
    lID = Val(InputBox("Icon ID Number", , ""))
    If lID = 0 Then
        Exit Sub
    End If
    TaskBr.DeleteIcon lID
End Sub

Private Sub Command5_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim lVerNum&
    Dim iVerWord%
    Dim iVersNum%
    Dim iTrueVers%
    '-----------------------------------------------------
    'First we find out what Windows is running.  This will
    'not work with NT 3.5 and earlier.
    '-----------------------------------------------------
    OsVers.dwVersionInfoSize = 148&
    lTempLong = GetVersionEx(OsVers)
    Select Case OsVers.dwPlatform
        Case VER_PLATFORM_WIN32_NT
            iTrueVers = OsVers.dwMajorVersion * 100 + OsVers.dwMinorVersion
            Select Case iTrueVers
                Case Is < 351
                    MsgBox "This program will not work on NT versions earlier then 3.51"
                    Unload Form1
                    Exit Sub
                Case 351
                    'Program has not been tested thoroughly with NT 3.51.  Should
                    'not crash, but definitely requires NewShell to work.
            End Select
        Case VER_PLATFORM_WIN32_WINDOWS
            'Windows 95 - we're OK
        Case Else 'Shouldn't happen
            MsgBox "This program is intended only for use with 32-bit Windows versions."
            Unload Form1
    End Select
     
    '---------------------------------------------------------
    'Check for presence of Taskbar.  The user may have a different
    'Shell that doesn't support one.
    '---------------------------------------------------------
    If TaskBr.AppBarExists <> 1 Then
        MsgBox "There is no tray currently available"
        Unload Form1
        Exit Sub
    End If
    '----------------------------------------------------------
    'Enable the MsgHook control to receive callbacks from the taskbar
    '----------------------------------------------------------
    
    Msghook1.HwndHook = Me.hWnd
    Msghook1.Message(TaskBr.Message) = True
    
    '-------------------------------------------------------------
    'The MsgHook control used with this sample is MSGHOO32.OCX.
    'Originally supplied with the Waite Group's "Visual Basic 4 HOW-TO"
    'book, but now freeware.  Many thanks to the author of MSGHOO32,
    'Zane Thomas, Mabry Software, and Waite Group Press.  You will need
    'to obtain this control prior to running the sample.  It can be found
    'on Compuserve in the VBPJFO forum; from Zane's web page at
    'http://activexpert.com/msghook.htm: from the author's web page at
    'http://www.redshift.com/~arcatpet/vb.html.
    '--------------------------------------------------------------

    Command1.Caption = "Add Icon"
    Command2.Caption = "Tooltip Text"
    Command3.Caption = "Change Icon"
    Command4.Caption = "Delete Icon"
    
    TaskBr.hWnd = Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    TaskBr.RemoveAllIcons (lIconAdded)
End Sub


Private Sub MsgHook1_Message(ByVal msg&, ByVal wparam&, ByVal lparam&, result&)
    '--------------------------------------------------------------------------
    'Using the TaskBar Tray for any purpose other than signaling requires that
    'we be able to receive messages generated by the system when there are mouse
    'events associated with the Notification Area icons.  The only way to receive these
    'messages in VB4/32 is with a message handling OCX.  There is not one
    'supplies with this example at this time.  All of the code is present to
    'use MSGHOO32.OCX, which is included with the Waite Group's Visual Basic 4
    'How To book.  To use that control, place an instance of it on this form and
    'remove the commenting in the Form_Load event.  To use another control
    'such as MsgBlaster.OCX you will need to alther the code to meet the parameters
    'of that code.
    '---------------------------------------------------------------------------
    
    '--------------------------------------------------------------------------
    'These are all of the possible messages sent to our program from the
    'TaskBar Notification Area icons.
    '--------------------------------------------------------------------------
    Select Case msg
      Case TaskBr.Message
        Label4 = wparam
        Select Case lparam
            Case WM_MOUSEMOVE
                Label2 = "MOUSEMOVE"
            Case WM_RBUTTONDBLCLK
                Label2 = "RBUTTONDBLCLK"
            Case WM_RBUTTONDOWN
                Label2 = "RBUTTONDOWN"
            Case WM_RBUTTONUP
                Label2 = "RBUTTONUP"
            Case WM_MBUTTONDBLCLK
                Label2 = "MBUTTONDBLCLK"
            Case WM_MBUTTONDOWN
                Label2 = "MBUTTONDOWN"
            Case WM_MBUTTONUP
                Label2 = "MBUTTONUP"
            Case WM_LBUTTONDBLCLK
                Label2 = "LBUTTONDBLCLK"
            Case WM_LBUTTONDOWN
                Label2 = "LBUTTONDOWN"
            Case WM_LBUTTONUP
                Label2 = "LBUTTONUP"
        End Select
        Timer1.Enabled = True
  End Select

End Sub

Private Sub Timer1_Timer()
    '-------------------------------------
    'Used to blank the message labels after
    'a short period.
    '-------------------------------------
    Label2 = ""
    Label4 = ""
    Timer1.Enabled = False
End Sub

