Attribute VB_Name = "modBatchPlotMain"
Option Explicit

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public strTemp As String
Public Const vbSizeD = 25
Public Function File_Extension(ByVal sFilename As String) As String

  Dim i    As Integer
  Dim j    As Integer

     j = 0
     For i = Len(sFilename) To 1 Step -1
          If Asc(Mid$(sFilename, i, 1)) = 46 Then
               File_Extension = Mid$(sFilename, i + 1)
               Exit For
            Else
               j = j + 1
          End If
     Next i

End Function

Public Sub Slide_Show(sfile As String)
     
     Select Case UCase$(File_Extension(sfile))
       Case "SLDPRT"
       frmPreviewBig.imgBig.Visible = True
          frmPreviewBig.web1.Visible = False
          Slide_Show_SW sfile
       Case "SLDDRW"
       frmPreviewBig.imgBig.Visible = True
       frmPreviewBig.web1.Visible = False
          Slide_Show_SW sfile
       Case "SLDASM"
       frmPreviewBig.imgBig.Visible = True
              frmPreviewBig.web1.Visible = False
          Slide_Show_SW sfile
       Case "DOC"
    
       frmPreviewBig.imgBig.Visible = False
        frmPreviewBig.web1.Visible = True
          frmBatchPlot.imgSmall.Picture = frmBatchPlot.imgHidden.Picture
          Slide_Show_Web sfile
       Case "PDF"
       frmPreviewBig.imgBig.Visible = False
       frmPreviewBig.web1.Visible = True
          frmBatchPlot.imgSmall.Picture = frmBatchPlot.imgHidden.Picture
          Slide_Show_Web sfile
       Case "INI"
       frmPreviewBig.imgBig.Visible = False
       frmPreviewBig.web1.Visible = True
          frmBatchPlot.imgSmall.Picture = frmBatchPlot.imgHidden.Picture
          Slide_Show_Web sfile
       Case "CFG"
       frmPreviewBig.imgBig.Visible = False
       frmPreviewBig.web1.Visible = True
          frmBatchPlot.imgSmall.Picture = frmBatchPlot.imgHidden.Picture
          Slide_Show_Web sfile
       Case "XLS"
       frmPreviewBig.imgBig.Visible = False
       frmPreviewBig.web1.Visible = True
          frmBatchPlot.imgSmall.Picture = frmBatchPlot.imgHidden.Picture
          Slide_Show_Web sfile
       Case "TXT"
       frmPreviewBig.imgBig.Visible = False
       frmPreviewBig.web1.Visible = True
          frmBatchPlot.imgSmall.Picture = frmBatchPlot.imgHidden.Picture
          Slide_Show_Web sfile
       Case "RTF"
          frmPreviewBig.imgBig.Visible = False
          frmPreviewBig.web1.Visible = True
          frmBatchPlot.imgSmall.Picture = frmBatchPlot.imgHidden.Picture
          Slide_Show_Web sfile
       Case "LOG"
          frmPreviewBig.imgBig.Visible = False
          frmPreviewBig.web1.Visible = True
          frmBatchPlot.imgSmall.Picture = frmBatchPlot.imgHidden.Picture
          Slide_Show_Web sfile
             
       Case "JPEG"
       frmPreviewBig.web1.Visible = False
         frmPreviewBig.imgBig.Visible = True
          Slide_Show_Image sfile
       Case "JPG"
       frmPreviewBig.web1.Visible = False
       frmPreviewBig.imgBig.Visible = True
          Slide_Show_Image sfile
       Case "BMP"
       frmPreviewBig.web1.Visible = False
       frmPreviewBig.imgBig.Visible = True
          Slide_Show_Image sfile
       Case "GIF"
       frmPreviewBig.web1.Visible = False
       frmPreviewBig.imgBig.Visible = True
          Slide_Show_Image sfile
       Case "ICO"
          
       Case "CUR"
          
       Case "DWG"
          
       Case "PPT"
          frmBatchPlot.imgSmall.Picture = frmBatchPlot.imgHidden.Picture
          Slide_Show_Web sfile
          
     End Select
     
End Sub

Sub Slide_Show_Image(sfile As String)

     DoEvents
     frmBatchPlot.imgSmall.Picture = LoadPicture(sfile)
     frmPreviewBig.imgBig.Picture = LoadPicture(sfile)

End Sub

Sub Slide_Show_Web(sfile As String)

     frmPreviewBig.imgBig.Visible = False
     frmPreviewBig.web1.Visible = True
     frmPreviewBig.web1.Navigate2 sfile
     
End Sub

Sub Slide_Show_SW(sfile As String)

  Dim sldFilename As String
  Dim tmpFilename As String
    
     ' get handle for the extract bitmap function in sdm.dll
  Dim BmpSaver As New SDMLib.smBitMap
    
     sldFilename = sfile
     tmpFilename = strTemp + "tempsdlbmp.bmp"
    
     If BmpSaver.extractBitMap2File(sldFilename, tmpFilename) Then

       Else
    
          ' to get the real size of the picture
          frmBatchPlot.picHidden.Picture = LoadPicture(tmpFilename)
          frmBatchPlot.imgSmall.Picture = frmBatchPlot.picHidden.Picture
          frmPreviewBig.imgBig.Picture = frmBatchPlot.picHidden.Picture
     End If
    
End Sub

Sub SW_Printing()
     
     Dim SWApp As Object
     Dim SWModel As New ModelDoc
     Dim i As Integer
     Dim JunkLong As Long
     Dim JunkDouble As Double
     
      ' Attach to exising SolidWorks session or start one
     Set SWApp = CreateObject("SldWorks.Application.9")
     SWApp.Visible = True
     For i = 0 To frmBatchPlot.lstToPrint.ListCount - 1
          Set SWModel = SWApp.OpenDoc4(frmBatchPlot.lstToPrint.List(i), Get_SW_Doc_Type(frmBatchPlot.lstToPrint.List(i)), _
                                   swOpenDocOptions_ViewOnly, "", JunkLong)
          SWModel.PrintSetup(swPrintOrientation) = vbPRORLandscape
          SWModel.PrintSetup(swprintpapersize) = vbSizeD
          SWModel.PrintDirect
          SWApp.CloseDoc frmBatchPlot.lstToPrint.List(i)
          Set SWModel = Nothing
     Next
     Set SWApp = Nothing
End Sub

Function Get_SW_Doc_Type(sFilename As String) As Integer
     
     If UCase(File_Extension(sFilename)) = "SLDPRT" Then
          Get_SW_Doc_Type = 1
     ElseIf UCase(File_Extension(sFilename)) = "SLDASM" Then
          Get_SW_Doc_Type = 2
     ElseIf UCase(File_Extension(sFilename)) = "SLDDRW" Then
          Get_SW_Doc_Type = 3
     End If
     
End Function
