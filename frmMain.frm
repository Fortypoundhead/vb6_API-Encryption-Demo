VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FortyPoundHead.com Encrypt/Decrypt Demo"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "&Decrypt"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "&Encrypt"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.FileListBox filFileList 
      Height          =   4575
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.DirListBox dirDirectory 
      Height          =   4140
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.DriveListBox drvDriveList 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDecrypt_Click()
    
    Dim strSelectedFile As String
    
    strSelectedFile = dirDirectory.Path & "\" & filFileList.List(filFileList.ListIndex)
    
    ret = Decrypt(strSelectedFile)
    
End Sub

Private Sub cmdEncrypt_Click()

    Dim strSelectedFile As String
    Dim ret As Long

    strSelectedFile = dirDirectory.Path & "\" & filFileList.List(filFileList.ListIndex)
    
    ret = Encrypt(strSelectedFile)
   
End Sub

Private Sub Command1_Click()

    Unload Me
    
End Sub

Private Sub dirDirectory_Change()

    ' The directory has changed, so update the file list
    
    filFileList.Path = dirDirectory.Path
    
End Sub

Private Sub drvDriveList_Change()

    ' The selected drive has changed, so update the directory path
    
    dirDirectory.Path = drvDriveList.Drive
    
End Sub
