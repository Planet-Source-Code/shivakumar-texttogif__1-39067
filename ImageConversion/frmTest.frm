VERSION 5.00
Begin VB.Form frmImg 
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   5280
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   2880
      ScaleHeight     =   915
      ScaleWidth      =   2235
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ConvertToImage"
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "frmImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim obj As Object
    Dim flag As Boolean

    If Len(Trim(Text1.Text)) > 0 Then
    
        frmImg.MousePointer = 11
        Set obj = CreateObject("prjImageConversion.clsImageConversion")
        flag = obj.convertToGif("Arial", False, 25, &HFF&, &H80000006, "20", "20", Text1.Text, "Font_" & Text1.Text, App.Path)
        frmImg.MousePointer = 0
        
        If (flag) Then
            Picture1.Picture = LoadPicture(App.Path & "\" & "Font_" & Text1.Text & ".gif")
        End If
        
        Set obj = Nothing
        
    Else
        MsgBox "Enter the text to convert in GIF format"
    End If
      
End Sub

