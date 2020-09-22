VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmPreviewBig 
   Caption         =   "Preview Big"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8040
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.UpDown UD1 
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   423
      _Version        =   393216
      Orientation     =   1
      Enabled         =   -1  'True
   End
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   6255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8055
      ExtentX         =   14208
      ExtentY         =   11033
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Image imgBig 
      BorderStyle     =   1  'Fixed Single
      Height          =   6255
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmPreviewBig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

  Dim i As Integer
     
     Resize_Controls
     
     If frmBatchPlot.lstToPrint.ListCount > 0 Then
          i = 0
          frmBatchPlot.lstToPrint.Selected(i) = True
          Slide_Show frmBatchPlot.lstToPrint.List(i)
          
     End If

End Sub

Private Sub Form_Resize()

     Resize_Controls

End Sub

Private Sub imgBig_DblClick()

     Unload Me

End Sub

Private Sub Resize_Controls()

     UD1.Width = Me.Width / 4
     imgBig.Height = Me.Height - 50
     imgBig.Width = Me.Width - 50
     imgBig.Refresh

End Sub

Private Sub UD1_DownClick()

  Dim i As Integer
     
     If frmBatchPlot.lstToPrint.ListCount > 0 Then
          For i = 0 To frmBatchPlot.lstToPrint.ListCount - 1
               If frmBatchPlot.lstToPrint.Selected(i) = True Then
                    frmBatchPlot.lstToPrint.Selected(i) = False
                    Exit For
               End If
          Next i
          
          i = i - 1
          If i = -1 Then
               i = frmBatchPlot.lstToPrint.ListCount - 1
          End If
          frmBatchPlot.lstToPrint.Selected(i) = True
          Slide_Show frmBatchPlot.lstToPrint.List(i)
     End If
     
End Sub

Private Sub UD1_UpClick()

  Dim i As Integer
     
     If frmBatchPlot.lstToPrint.ListCount > 0 Then
          For i = 0 To frmBatchPlot.lstToPrint.ListCount - 1
               If frmBatchPlot.lstToPrint.Selected(i) = True Then
                    frmBatchPlot.lstToPrint.Selected(i) = False
                    Exit For
               End If
          Next i
          
          If i = frmBatchPlot.lstToPrint.ListCount - 1 Then
               i = -1
          End If
          
          i = i + 1
          frmBatchPlot.lstToPrint.Selected(i) = True
          Slide_Show frmBatchPlot.lstToPrint.List(i)
     End If
     
End Sub


