VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00000000&
   ClientHeight    =   7290
   ClientLeft      =   165
   ClientTop       =   15
   ClientWidth     =   9465
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   2235
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   9465
      TabIndex        =   0
      Top             =   5055
      Visible         =   0   'False
      Width           =   9465
      Begin VB.PictureBox picDest 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1305
         Left            =   5145
         ScaleHeight     =   83
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   192
         TabIndex        =   2
         Top             =   105
         Width           =   2940
      End
      Begin VB.PictureBox picSource 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Height          =   480
         Left            =   540
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   1
         Top             =   180
         Width           =   640
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2535
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   4755
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11060
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Date Created : April 5, 2004
'Naga City Philippines

'This program demonstrates how to make a background (centered,stretched mode) on
'an MDI form.
'Double click on the Form's background to browse images and change the current
'background image.
'Right-Clicking on the background toggles the mode either centered or stretched.

Option Explicit

Public connected As Boolean

Private Sub MDIForm_Load()
    
    On Error Resume Next
    
    Me.Width = Screen.TwipsPerPixelX * 800 ' Thanks to Luke H.  Email: luke@bemroseconsulting.com, for all the years I really ignored the object Screen
    Me.Height = Screen.TwipsPerPixelY * 600 ' Thanks to Luke H.  Email: luke@bemroseconsulting.com
    
        
    connected = True
    
    currentImageFile = GetSetting(App.EXEName, "Settings", "FormPic", App.Path & "\Rex.bmp")
    
    picSource.Picture = LoadPicture(currentImageFile)
    mdibg_mode = CBool(GetSetting(App.EXEName, "Settings", "MDIBGMode", "True"))
    Call CreateFormPic(Me.hWnd, picSource, picDest, mdibg_mode)
    
End Sub

Private Sub MDIForm_DblClick()
    On Error GoTo ExitNow
    
    CommonDialog1.InitDir = GetSetting(App.EXEName, "Settings", "FormPicDir", App.Path & "\Images")
    
    CommonDialog1.Filter = "All files (*.*)|*.*"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.DefaultExt = "bmp"
    CommonDialog1.DialogTitle = "Select image file"
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen
    currentImageFile = CommonDialog1.FileName
    SaveSetting App.EXEName, "Settings", "FormPicDir", Mid(CommonDialog1.FileName, 1, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle))
    SaveSetting App.EXEName, "Settings", "FormPic", currentImageFile
    
    picSource.Picture = LoadPicture(currentImageFile)
    
    Call CreateFormPic(Me.hWnd, picSource, picDest, mdibg_mode)
    
ExitNow:

End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    If Button = 2 Then 'right click
        mdibg_mode = Not (mdibg_mode) 'toggle it
        SaveSetting App.EXEName, "Settings", "MDIBGMode", CStr(mdibg_mode) 'save new setting
        Call CreateFormPic(Me.hWnd, picSource, picDest, mdibg_mode) 'paint it new
    End If
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If Not connected Then Exit Sub

    Select Case UnloadMode
        Case vbFormControlMenu
            ' Form is being closed by user.
            If MsgBox("Close this program now?", vbQuestion + vbOKCancel, "Confirm Close") = vbCancel Then
                Cancel = 1
            End If
        Case vbFormCode
            If MsgBox("Close this program now?", vbQuestion + vbOKCancel, "Confirm Close") = vbCancel Then
                Cancel = 1
            End If
    End Select

End Sub

Private Sub MDIForm_Resize()
    Call CreateFormPic(Me.hWnd, picSource, picDest, mdibg_mode)
End Sub

Private Sub mnuWindowArrangeIcons_Click()
  Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
  Me.Arrange vbCascade
End Sub

Private Sub mnuWindowTileHorizontal_Click()
  Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileVertical_Click()
  Me.Arrange vbTileVertical
End Sub

Private Sub mnuExit_Click()
     Unload Me
End Sub
