VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "- Set BackGround Image For IE/Explorer -"
   ClientHeight    =   4530
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4080
      Top             =   4080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Image"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Del Shortcut"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      ToolTipText     =   "Click to remove program shortcut"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create Shortcut"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Creating a shortcut will allow you to quickly open this program through Internet Explorers ""Links"" Menu"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   120
      Pattern         =   "*.bmp"
      TabIndex        =   3
      Top             =   1920
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3315
      Left            =   2640
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3135
   End
   Begin VB.Menu mnuScroll 
      Caption         =   "sdfds"
      Visible         =   0   'False
      Begin VB.Menu mnuAutoScroll 
         Caption         =   "Auto Scroll"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'What this code does is set a bitmap (only works with bitmaps) image to be the
'background in all Explorer/My Computer/Internet Explorer windows. It works just by
'change a registry value, so it was very easy to code. You can drag and drop bitmaps
'from explorer and it will load them and change the dirlistbox to the dir it
'was dragged from. Ive added lots of fancy stuff but the main functionality is 3 or so lines

'Ive included a few textures which i found on my hardrive. The wood one looks
'quite good, or you could always just use a pic of a girl in a bikini ;)

'BTW Ive only tested it one my home win98 computer, so I dont know if it'll work for any other versions.
'You must have rights to write to the registry, other wise it wont work

Dim strBitmap As String
Dim strURL As String
Dim strDir As String
Dim strData As String
Dim strDragDir As String
Dim strFileName As String
Dim strCurrent As String

Dim strCheck As String
Dim intCheck As Integer
Dim intSlash As String


Private Sub Command1_Click()
On Error GoTo err

If File1.FileName = "" Then
    MsgBox "Please select a bitmap file first", vbExclamation, App.Title
    'no files been selected
    Exit Sub
Else
    strExtension = File1.FileName
    length = Len(strExtension)
    where = InStr(strExtension, ".") 'find where the . is
    strExtension = Right$(strExtension, length - where) 'chops string to only the letters after the "." eg JPG GIF etc
    strExtension = LCase(strExtension) 'changes string to lower case
    If strExtension = "jpg" Or strExtension = "gif" Or strExtension = "jpeg" Or strExtension = "jpe" Then
    'checks to see what the extension of the file is, We want to trap all jpg's and gif's
            If MsgBox("The selected file is not a bitmap. Do you wish to convert it to a bitmap?", vbYesNo, App.Title) = vbYes Then
                ConvertToBMP (File1.path & "\" & File1.FileName) 'calls function which saves selected picture to a bitmap
                strBitmap = File1.path & "\" & File1.List(File1.ListIndex)
                Call UpdateKey(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Toolbar", "BackBitmap", strBitmap)
                'this updates the registry key which specifies what file explorer uses for its background (skin)
                'thats all you have to change to make it work!
                MsgBox "Skin set." & vbCrLf & "You will notice the change when you  a open a new browser window.", vbInformation, App.Title

            Else
                Exit Sub
            End If
    Else
        strBitmap = File1.path & "\" & File1.FileName
        Call UpdateKey(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Toolbar", "BackBitmap", strBitmap)
        'this updates the registry key which specifies what file explorer uses for its background (skin)
        'thats all you have to change to make it work! **This one is only called when the originally selected file is a bitmap
        MsgBox "Skin set." & vbCrLf & "You will notice the change when you  a open a new browser window.", vbInformation, App.Title
    End If
End If
Exit Sub
err:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub Command2_Click()
strDir = WinDir(False) ' cant presume windows is installed in "C:\windows", this function finds where windows is installed
FileCopy App.path & "\" & App.EXEName & ".exe", strDir & "\Favorites\Links\Skin.exe"
'copies a copy of the exe to a windows dir (windows\Favorites\Links\) which will create a shortcut you can see in IE

'NOTE: if your running this in the VB IDE the above line will crash.
'You have to be using the exe (or at least have the exe compiled in same dir as project) for it to work

strMsg = "Shortcut has been created. " & vbCrLf
strMsg = strMsg & "To use the shortcut, right click on your toolbar in Internet Explorer and make sure 'Links' are turned on." & vbCrLf
strMsg = strMsg & "You should now be able to see the shortcut, 'Skin.exe'. Click on it and it will open this program"
MsgBox strMsg, vbInformation, App.Title 'build a string to message

End Sub

Private Sub Command3_Click()
On Error GoTo err
strDir = WinDir(False) ' cant presume windows is installed in "C:\windows"
Kill strDir & "\Favorites\Links\Skin.exe" 'delete exe from links dir
MsgBox "Shortcut removed", vbInformation, App.Title
Exit Sub
err:
MsgBox "No shortcut to remove", vbExclamation, App.Title 'error is reached if file isnt found
End Sub

Private Sub Command4_Click()
frmBatch.Show
frmBatch.Dir1.path = Dir1.path
End Sub

Private Sub Dir1_Change()
File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
On Error GoTo err
Dir1.path = Drive1.Drive
Exit Sub
err:
MsgBox err.Description, vbCritical, App.Title 'need to trap this error
End Sub

Private Sub File1_Click()
Dim strExtension As String
On Error GoTo err:
Image1.Picture = LoadPicture(File1.path & "\" & File1.FileName) 'loads the file clicked in the file list box into the image control
Command1.ToolTipText = "Click to set " & File1.path & "\" & File1.FileName & " to be your new skin"
'make the tooltip say which file is going to be set.. not very useful
Label1.Caption = File1.path & "\" & File1.FileName
Exit Sub
err:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 46:
    If MsgBox("Are you sure you want to delete this file?", vbYesNo, App.Title) = vbYes Then
        Kill (File1.path & "\" & File1.FileName) 'they pressed del, so kill (delete) the file
        File1.Refresh
    End If
End Select
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then 'clicked right mouse button
     PopupMenu mnuScroll
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 And Timer1.Enabled = True Then
    Timer1.Enabled = False
    mnuAutoScroll.Checked = False
 End If
End Sub



Private Sub Form_Load()
On Error GoTo err:
Dim strOrigPic As String
File1.Pattern = "*.jpeg;*.jpe;*.gif;*.jpg;*.bmp"
'MsgBox "Make sure you close all open Explorer/IE Windows before changing skin!", vbExclamation, App.Title
'<-- uncomment above line to tell users to explorer windows
strCurrent = GetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Toolbar", "BackBitmap")
'get the path of the current backgorund/skin
Label1.Caption = strCurrent
intSlash = LastSlash(Label1.Caption) 'find out where the last backslash is using a function i wrote
Dir1.path = Left$(strCurrent, intSlash) 'chops everything from where the last backslash is found; this returns the path, and chops of the filename
Image1.Picture = LoadPicture(strCurrent)
strOrigPic = Replace(strCurrent, Dir1.path, "") 'chop the path from the string and your left with just the filename
strOrigPic = Replace(strOrigPic, "\", "") 'get rid of backslashes
For i = 0 To File1.ListCount - 1
    If File1.List(i) = strOrigPic Then 'goes through all the files in the filelistbox to see if the they match the filename pulled from the registry
        File1.ListIndex = i 'if filename in listbox is same as one from registry, select it
    End If
Next i

Exit Sub
err:
If err.Number = 53 Then 'file cant be found
    MsgBox "Unable to find the file specified in the registry. I may have been deleted or moved", vbCritical, App.Title
Else 'any other error
    MsgBox "Unable to load current background."
End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1.Picture = 0 Then 'no picture is loaded, so tell 'em what to do
    Image1.ToolTipText = "Locate a bitmap file on your computer using the navigation tools to the left. You can also drag and drop bitmaps from explorer"
Else
    Image1.ToolTipText = "" 'genius's must of worked out how to load files =)
End If

End Sub

Private Sub Image1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
strData = Data.Files(1) 'puts the file name being dragged into program into a variable
strCheck = LCase(Data.Files(1)) 'changes file name to lower case in case file was .BMP instead of .bmp
intCheck = InStr(strCheck, ".bmp") 'checks to see if file name contains ".bmp" thus being a bitmap
If intCheck <> 0 Then ' .bmp was found in name of dragged file
    intSlash = LastSlash(strData) 'calls function LastSlash, which returns where the last backslash in a string is
    strFileName = Left$(strData, intSlash) 'read above
    strDragDir = Replace(Data.Files(1), strFileName, "")
    Dir1.path = strDragDir 'sets dir path to same path as file being dragged
    Image1.Picture = LoadPicture(Data.Files(1)) 'loads dragged bitmap
Else
    MsgBox "You must drag valid bitmap files", vbCritical, App.Title 'bmp wasnt found in filename being dragged
End If
End Sub

Private Sub Label1_Change()
Label1.Caption = LCase(Label1.Caption) 'I hate capitals, its rude =)
End Sub

Private Sub mnuAutoScroll_Click()
   If Timer1.Enabled = True Then
        Timer1.Enabled = False 'turn auto scrolling off
        mnuAutoScroll.Checked = False
    Else
        Timer1.Enabled = True 'turn auto scrolling on
        mnuAutoScroll.Checked = True
    End If
End Sub

Private Sub Timer1_Timer()
'On Error GoTo Err:
File1.ListIndex = File1.ListIndex + 1 'select next item in filelistbox
'maybe someones to lazy to press arrow keys/click mouse so they might use this

End Sub

Function ConvertToBMP(strFileName)

Picture1.Picture = LoadPicture(strFileName) 'loads the file clicked in the file list box into the image control
strFileName = LCase(strFileName)
strFileName = Replace(strFileName, ".jpg", "")
strFileName = Replace(strFileName, ".gif", "")
strFileName = strFileName & ".bmp" ' get rid of the extension and change it to bmp instead
SavePicture Picture1.Picture, strFileName 'save the contents on picture1 to disk, as strFileName
File1.Refresh
For i = 0 To File1.ListCount - 1
    strMatch = File1.path & "\" & File1.List(i)
    strMatch = LCase(strMatch)
    If strMatch = strFileName Then 'goes thru the filelistbox until the file selected is equal to the name of the
    'bitmap saved above. Stops when it matches them, then selects that item
        File1.ListIndex = i
        Exit Function
    End If
Next i
End Function

