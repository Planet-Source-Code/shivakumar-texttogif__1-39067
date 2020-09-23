VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImageConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim WithEvents cGif As clsGif
Attribute cGif.VB_VarHelpID = -1
Private myImagePath As String

Private Sub Class_Initialize()

    frmInvisible.Visible = False
  
End Sub

Public Function convertToGif(imgFont As String, imgFontItalic As Boolean, imgSize As Integer, imgForeColor As Long, imgBackColor As Long, imgWidth As String, imgHeight As String, imgText As String, imgName As String, imgPath As String) As Boolean

On Error GoTo Err_Handler
   
   initBmpProperties imgFont, imgFontItalic, imgSize, imgForeColor, imgBackColor, imgWidth, imgHeight
   
   Dim myAlpha As String
   Dim errFlag As Boolean
   
   errFlag = False
   myImagePath = imgPath & "\"
   
   frmInvisible.myPicture.Cls
   myAlpha = imgText
   frmInvisible.myPicture.Width = frmInvisible.myPicture.TextWidth(myAlpha)
   frmInvisible.myPicture.Height = frmInvisible.myPicture.TextHeight(myAlpha)
   frmInvisible.myPicture.Print myAlpha
   
   DoEvents
   SavePicture frmInvisible.myPicture.Image, getBmp(myAlpha)
   frmInvisible.myPicture.Picture = LoadPicture(myImagePath & "Font_" & imgText & ".bmp")
   frmInvisible.myPicture.Picture = frmInvisible.myPicture.Image
   frmInvisible.myPicture.Refresh
   
   Set cGif = New clsGif
   errFlag = cGif.SaveGIF(frmInvisible.myPicture.Picture, myImagePath & imgName & ".gif", frmInvisible.myPicture.hDc, False, frmInvisible.myPicture.Point(0, 0))
      
   convertToGif = errFlag
   clearBmp myAlpha
   
   Set cGif = Nothing
   
   Exit Function
   
Err_Handler:
        Set cGif = Nothing
        convertToGif = errFlag
        
End Function

Private Sub initBmpProperties(imgFont As String, imgFontItalic As Boolean, imgSize As Integer, imgForeColor As Long, imgBackColor As Long, imgWidth As String, imgHeight As String)

   With frmInvisible.myPicture
   
      .AutoSize = True
      .AutoRedraw = True
      .FontTransparent = True
      .BorderStyle = 0
      
      .Font.Name = imgFont
      .FontItalic = imgFontItalic
      .Font.Size = imgSize
      .ForeColor = imgForeColor
      .BackColor = imgBackColor
      .Width = .TextWidth(imgWidth)
      .Height = .TextHeight(imgHeight)
      
   End With
   
End Sub

Private Function getBmp(imgText As String) As String

   getBmp = myImagePath & "Font_" & imgText & ".bmp"
   
End Function

Private Sub clearBmp(imgText As String)

  Dim objFso As Object
  Dim strFilePath As String
  
  Set objFso = CreateObject("Scripting.FileSystemObject")
  strFilePath = myImagePath & "Font_" & imgText & ".bmp"
  
  If (objFso.FolderExists(myImagePath)) Then
  
        If (objFso.FileExists(strFilePath)) Then
                objFso.DeleteFile strFilePath, True
        End If
  
  End If
  
  Set objFso = Nothing
  
End Sub

