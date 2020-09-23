VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmICL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon library extracter"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   Icon            =   "frmICL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7665
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   3885
      Begin MSComDlg.CommonDialog cdlSave 
         Left            =   3060
         Top             =   1170
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CheckBox chSel 
         Caption         =   "Selected only"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   6795
         Width           =   2850
      End
      Begin VB.Frame Frame2 
         Caption         =   "Extract the icons into:"
         Height          =   4785
         Left            =   180
         TabIndex        =   10
         Top             =   1440
         Width           =   3480
         Begin VB.DirListBox Dir1 
            Height          =   3915
            Left            =   180
            TabIndex        =   12
            Top             =   675
            Width           =   3075
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   180
            TabIndex        =   11
            Top             =   315
            Width           =   3075
         End
      End
      Begin VB.CommandButton btClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   180
         TabIndex        =   5
         Top             =   7155
         Width           =   3480
      End
      Begin VB.CommandButton btOpen 
         Caption         =   "&Select icon library"
         Height          =   375
         Left            =   180
         TabIndex        =   4
         Top             =   270
         Width           =   3435
      End
      Begin VB.CommandButton btExtract 
         Caption         =   "Extract icons"
         Enabled         =   0   'False
         Height          =   465
         Left            =   180
         TabIndex        =   3
         Top             =   6300
         Width           =   3480
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   492
         Left            =   225
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   2
         Top             =   2610
         Visible         =   0   'False
         Width           =   492
      End
      Begin MSComctlLib.ImageList IML 
         Left            =   765
         Top             =   2520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   14805982
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CDL 
         Left            =   1350
         Top             =   2610
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "icl"
         DialogTitle     =   "Open ICL Library"
         Filter          =   "Icon Libraries (*.icl;*.ni)|*.icl;*.ni;*.il|Icons (*.ico)|*.ico|Executables (*.exe;*.dll)|*.exe;*.dll|All files|*.*"
      End
      Begin VB.Label lbLib 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Please, select an icon library."
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         TabIndex        =   8
         Top             =   720
         Width           =   3750
      End
      Begin VB.Label lbInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   1170
         Width           =   3885
      End
   End
   Begin MSComctlLib.ListView lvICONS 
      Height          =   7215
      Left            =   4050
      TabIndex        =   0
      Top             =   135
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   12726
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ProgressBar PROG 
      Height          =   330
      Left            =   4050
      TabIndex        =   6
      Top             =   7380
      Visible         =   0   'False
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lbStat 
      Caption         =   "Ready."
      Height          =   240
      Left            =   4095
      TabIndex        =   13
      Top             =   7425
      Width           =   5550
   End
   Begin VB.Label Label1 
      Caption         =   "LOADING...PLEASE WAIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1680
      Left            =   5625
      TabIndex        =   9
      Top             =   2835
      Width           =   3975
   End
   Begin VB.Menu mnSelect 
      Caption         =   "Sel"
      Visible         =   0   'False
      Begin VB.Menu mnSaveAs 
         Caption         =   "Save as ICO file"
      End
      Begin VB.Menu mnSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnExtract 
         Caption         =   "Extract selected icons"
      End
      Begin VB.Menu mnExtractAll 
         Caption         =   "Extract all icons"
      End
   End
End
Attribute VB_Name = "frmICL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************       BF Icon Extractor      **************************
'*                      * Written by Chavdar Yordanov  *                        *
'*         (part of the code taken from www.vbaccelerator.com)                  *
'* Previews and extracts icons from DLLs, ICL libraries, EXEs and other formats *
'*                                                                              *
'*           Currently supports only 32 x 32 x 256 color icons :-(              *
'*                        but is quite useful anyway.                           *
'*        Please, help with suggestions at chavdar_jordanov@yahoo.com           *
'********************************************************************************

Option Explicit

Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

Dim sLibraryName As String          'The name of selected icon library
    
Dim oIcon As cFileIcon              'Icon object which processes the icon images
Dim IconName() As String            'Array containing the icon names
Dim TransparentColor As Long        'Transparency color value

Dim IconCount                       'Number of icons in the library
Dim hModule                         'Handle needed for the API calls

Dim Iconh&
Dim X&

Private Type tDeviceImage
   iSizeX As Long
   iSizeY As Long
   cDepth As Long
   cPal As cPalette
End Type

Private m_tDeviceImage As tDeviceImage

Private Sub btClose_Click()
    Set oIcon = Nothing  'destroy the object
    Unload Me
End Sub

Private Sub btExtract_Click()
    ExtractIcons chSel.Value = 1
End Sub

Private Sub btOpen_Click()
    On Error GoTo 100
    CDL.CancelError = True
    CDL.ShowOpen
    sLibraryName = CDL.Filename
    lbLib = sLibraryName
    ReadLibrary
10
    Exit Sub
100
    Resume 10
End Sub


'------ Gets the icons from the specified library ------
'-      and shows them into the List View control      -
Sub ReadLibrary()
    Dim i As Long
    Dim sAPILibName As String
    
    On Error GoTo 100
    sAPILibName = sLibraryName + Chr$(0)          'Filename for the API call
    IconCount = ExtractIcon(hModule, sAPILibName, -1) 'get the number of icons in the library
    
    lvICONS.Icons = Nothing                       'Detach the ImageList from the ListView control
    IML.ListImages.Clear                          'Clear any previously loaded icons
    lvICONS.ListItems.Clear
    
    If IconCount > 0 Then
        lbStat = "Loading library..."
        lbInfo.Caption = "This file contains " + CStr(IconCount) + " icon/s." 'shows number of icons on label
        PROG.Max = IconCount
        PROG.Visible = True
        lvICONS.Visible = False
        GetIconNames
        For i = 1 To IconCount
            Set Pic.Picture = LoadPicture("")                'Clear the Picture box
            Iconh = ExtractIcon(hModule, sAPILibName, i - 1) 'Extracts the first icon
            X& = DrawIcon(Pic.hdc, 0, 0, Iconh)              'Draws the icon into the Picture box
            IML.ListImages.Add i, , Pic.Image                'Adds it to the Image List
            lvICONS.Icons = IML
            lvICONS.ListItems.Add , , IconName(i), i        'And shows it in the List View
            PROG.Value = i
            DoEvents
        Next i
        lvICONS.Visible = True
        PROG.Visible = False
        lbStat = "Ready."
    Else
        lbInfo = "This file does not contain icons."
    End If
    btExtract.Enabled = IconCount > 0
10
    Exit Sub
100
    MsgBox Err.Description, vbCritical
    Resume 10
End Sub

'---------- Extracts icons from a library into specified dir ------
Sub ExtractIcons(bSelectedOnly As Boolean)
    Dim i, j
    On Error GoTo 100
    Pic.Visible = True
    PROG.Visible = True
    lbStat = "Extracting icons..."
    For j = 1 To IconCount
        If bSelectedOnly Then  'check if the icon is selected
            If lvICONS.ListItems(j).Selected = False Then GoTo SkipIcon
        End If
        SaveIconToFile j, ToPath(Dir1.Path) + IconName(j) + ".ico"  'and save it
        PROG.Value = j
        DoEvents
SkipIcon:
    Next j
    Pic.Visible = False
    PROG.Visible = False
    lbStat = "Ready."
10
    Exit Sub
100
    MsgBox Err.Description, vbCritical
    Resume 10
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Set oIcon = New cFileIcon
    Drive1.Drive = Left(App.Path, 1) 'point to the application dir
    Dir1.Path = App.Path
    CDL.InitDir = App.Path
    hModule = Me.hWnd
    Pic.AutoRedraw = True       'important!
    With m_tDeviceImage         'prepare the icon format
      .iSizeX = 32
      .iSizeY = 32
      .cDepth = 256
      Set .cPal = New cPalette
        .cPal.CreateWebSafe
    End With
    TransparentColor = IML.MaskColor 'Get the default transparent color
    Me.Show
End Sub

'---------- Extracts the icon names from an ICL library ---------
'- or just puts numbers if the library is of any other type     -
Private Sub GetIconNames()
    Dim S As String, FN As Long
    Dim x1 As Long, i As Long
    Dim Cnt As Long
    Dim Z As Long
    ReDim IconName(1 To IconCount)
    If Right(sLibraryName, 3) = "icl" Then
        Cnt = 0
        FN = FreeFile                            'In order to get the names from the ICL library
        Open sLibraryName For Binary As #FN
        S = Space(LOF(FN))
        Get #FN, , S                             'we shall load it into a string variable first
        Close #FN
        x1 = InStr(1, S, "ICL", vbBinaryCompare) 'Look for the letters 'ICL'
        If x1 = 0 Then GoTo PutNumberedNames     'if not found then use numbers instead
        x1 = x1 + 3                              'Names in the ICL library follow letters ICL
        Do                                       'and are separated by a single byte containing
                                                 'the length of the following file name
                                                 
            Z = Asc(Mid(S, x1, 1))               'GET THE LENGTH OF THE FILE NAME
            If Z = 0 Then Exit Do                'IF IT IS ZERO, WE HAVE REACHED THE LAST NAME
            Cnt = Cnt + 1
            IconName(Cnt) = Mid(S, x1 + 1, Z)    'GET THE NAME
            x1 = x1 + Z + 1                      'FIND THE NEXT POSITION
        Loop
        S = ""                                   'Delete the string
    ElseIf Right(sLibraryName, 3) = "ico" Then
        
        For i = Len(sLibraryName) - 5 To 1 Step -1
            If Mid(sLibraryName, i, 1) = "\" Then
                IconName(1) = Mid(sLibraryName, i + 1, Len(sLibraryName) - i - 4)
                Exit For
            End If
        Next i
    Else
PutNumberedNames:
        
        For i = 1 To IconCount
            IconName(i) = "Icon" + Format(i, "0000")
        Next i
    End If
End Sub

Private Sub lvICONS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnSelect
    End If
End Sub

Private Sub lvICONS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    mnSaveAs.Enabled = False
    mnExtract.Enabled = False
    For i = 1 To lvICONS.ListItems.Count
        If lvICONS.ListItems(i).Selected Then
            mnSaveAs.Enabled = True
            mnExtract.Enabled = True
            Exit For
        End If
    Next i
    
End Sub

Private Sub mnExtract_Click()
    ExtractIcons True
End Sub

Private Sub mnExtractAll_Click()
    ExtractIcons False
End Sub

'------ Saves a single icon into a specified location ---------
Private Sub mnSaveAs_Click()
    Dim sSaveName As String
    Dim IconIndex As Long
    IconIndex = lvICONS.SelectedItem.Index
    On Error GoTo 100
    With cdlSave
        .CancelError = True
        .Filter = "*.ico"
        .DefaultExt = "ico"
        .Filename = IconName(IconIndex)
        .InitDir = Dir1.Path
        .ShowSave
        sSaveName = .Filename                 'choose a name for the icon file
    End With
    
    On Error GoTo 200
    SaveIconToFile IconIndex, sSaveName       'and save it to the disk
10
    Exit Sub
100
    Resume 10
200
    MsgBox Err.Description, vbCritical
    Resume 10
End Sub

'------------- Saves an icon to the hard disk ---------------
Private Sub SaveIconToFile(ByVal Index As Long, ByVal SaveName As String)
    Set Pic.Picture = IML.ListImages(Index).Picture  'Paint the selected icon to the picture box
    oIcon.AddImage m_tDeviceImage.iSizeX, m_tDeviceImage.iSizeY, m_tDeviceImage.cDepth  'Create an icon template
    m_tDeviceImage.cPal.SetPaletteToIcon oIcon, 1
    oIcon.SetIconFromBitmap Pic.hdc, 1, 0, 0, True, TransparentColor  'Create an icon object in memory
    oIcon.SaveIcon SaveName                                           'and save it to the disk
    oIcon.RemoveImage 1                                               'Clear the icon
End Sub

'---- Assures that a path has a trailing \ slash ----
Function ToPath(ByVal S As String) As String
    If Right(S, 1) <> "\" Then S = S & "\"
    ToPath = S
End Function

