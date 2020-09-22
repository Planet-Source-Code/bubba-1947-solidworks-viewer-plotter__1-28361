VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBatchPlot 
   Caption         =   "Solidworks 2001 Batch Plotting"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPreviewAll 
      Caption         =   "Preview All"
      Height          =   375
      Left            =   7560
      TabIndex        =   18
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picHidden 
      Height          =   495
      Left            =   8040
      ScaleHeight     =   435
      ScaleMode       =   0  'User
      ScaleWidth      =   69
      TabIndex        =   17
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About..."
      Height          =   375
      Left            =   7560
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   6060
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15769
            MinWidth        =   15769
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstPrinters 
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   5760
      Width           =   6015
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit Batch"
      Height          =   495
      Left            =   7560
      TabIndex        =   13
      Top             =   5520
      Width           =   1215
   End
   Begin VB.ListBox lstCopies 
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemItem 
      Caption         =   "Remove Item"
      Height          =   375
      Left            =   7560
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ListBox lstToPrint 
      Height          =   2400
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   7335
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   3720
      MultiSelect     =   1  'Simple
      TabIndex        =   5
      Top             =   240
      Width           =   3735
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   3495
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdAddAll 
      Caption         =   "Add All"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddSel 
      Caption         =   "Add Selected"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Options..."
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image imgHidden 
      Height          =   480
      Left            =   8040
      Picture         =   "frmBatchPlot.frx":0000
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSmall 
      Height          =   975
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "No. Copies"
      Height          =   255
      Left            =   7560
      TabIndex        =   12
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Active Printer:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Solidworks Files:"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   0
      Width           =   2055
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8760
      Y1              =   3120
      Y2              =   3120
   End
End
Attribute VB_Name = "frmBatchPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit ':( Line inserted

Private Sub cmdSubmit_Click()
     SW_Printing
     
End Sub

Private Sub Form_Load()
     
     File1.Pattern = "*.sldprt;*.sldasm;*.slddrw" ';*.txt;*.ini;*.cfg;*.log;*.rtf;*.doc;*.xls;*.pdf;*.ppt;*.dwg;*.bmp;*.jpg;*.gif"
     
     ' Retrieve path to TEMP
     strTemp = String$(100, Chr$(0))
     GetTempPath 100, strTemp
     strTemp = Left$(strTemp, InStr(strTemp, Chr$(0)) - 1)
     
End Sub

Private Sub cmdAbout_Click()

     MsgBox " Add your own customized about..."
     
End Sub

Private Sub cmdAddAll_Click()

  Dim i As Integer

     For i = 0 To (File1.ListCount - 1)
          lstToPrint.AddItem Dir1.Path + "\" + File1.List(i)
     Next i
   
     cmdPreviewAll.Visible = True
     
End Sub

Private Sub cmdAddSel_Click()

  Dim i As Integer
     
     For i = 0 To (File1.ListCount - 1)
          If File1.Selected(i) = True Then
               Me.lstToPrint.AddItem Dir1.Path + "\" + File1.List(i)
          End If
     Next i
   
     For i = 0 To (File1.ListCount - 1)
          If File1.Selected(i) = True Then
               File1.Selected(i) = False
          End If
     Next i
     
     cmdPreviewAll.Visible = True
     
End Sub

Private Sub cmdClearAll_Click()
     
     lstToPrint.Clear
     cmdPreviewAll.Visible = False
     
End Sub

Private Sub cmdPreviewAll_Click()

  'Clear out memory

     Unload frmPreviewBig
     Set frmPreviewBig = Nothing
     
     Load frmPreviewBig
     frmPreviewBig.Show

End Sub

Private Sub Dir1_Change()

     File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()

     Dir1.Path = Drive1.Drive

End Sub

Private Sub File1_Click()
   
     Slide_Show Dir1.Path + "\" + File1.List(File1.ListIndex)

End Sub

Private Sub Form_Unload(Cancel As Integer)
     End
End Sub

Private Sub imgSmall_Click()

     Load frmPreviewBig
     frmPreviewBig.imgBig.Picture = imgSmall.Picture
     frmPreviewBig.Show
     
End Sub


