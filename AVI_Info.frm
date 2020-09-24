VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form AVI_Info 
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtData 
      Height          =   6495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   600
      Width           =   6855
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "AVI_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AVII As AviStructure

Private Sub cmdOpen_Click()
Dialog.Filter = "Avi File|*.avi"
Dialog.FileName = ""
Dialog.ShowOpen
If Dialog.FileName <> "" Then
    txtFileName.Text = Dialog.FileName
    AviInfoRead Dialog.FileName, AVII
    ShowAVIdata
End If
End Sub


Private Sub ShowAVIdata()
txtData.Text = vbNewLine



With AVII
    AddLine "File Size:", .FileSize
        AddLine ""
    AddLine "Streams (in header)", .Streams
    AddLine "Streams Found", .StreamsFound
    AddLine "Lenght", .Lenght
    AddLine "TotalFrames", .TotalFrames
    AddLine "Resolution (in header)", .Width & " x " & .Height
    'AddLine "FPS", .FPS
    AddLine "Flag", .Flag
        
        
    Dim i As Byte
    For i = LBound(.StreamInfo) To UBound(.StreamInfo)
            AddLine ""
            AddLine ""
            AddLine "Stream " & i + 1, "="
            AddLine ""
        AddLine "Type", .StreamInfo(i).Type
        AddLine "Codec", .StreamInfo(i).Codec
        AddLine "Flag", .StreamInfo(i).Flag
        AddLine "Lenght", .StreamInfo(i).Lenght
        AddLine "Quality", .StreamInfo(i).Quality
        AddLine "VideoSize (kB)", .VideoSize / 1000
        AddLine "AudioSize (kB)", .AudioSize / 1000
        AddLine "Quality", .StreamInfo(i).Quality
        AddLine "bps", .StreamInfo(i).bPS
                AddLine ""
    Select Case .StreamInfo(i).Type
    Case "vids":
        AddLine "Codec Code", .StreamInfo(i).Video_Codec
        AddLine "Resolution", .StreamInfo(i).Video_Width & " x " & .StreamInfo(i).Video_Height
        AddLine "FPS", .StreamInfo(i).Video_FPS
        AddLine "BPP", .StreamInfo(i).Video_bPP
    Case "auds":
        AddLine "Codec Code", .StreamInfo(i).Audio_Codec
        AddLine "Channels", .StreamInfo(i).Audio_Channels
        AddLine "SamplePerSec", .StreamInfo(i).Audio_SamplePerSec
    End Select
    Next
            AddLine ""
            AddLine ""
            AddLine "Tags", "="
            AddLine ""
        AddLine "Name", .Tags.Name, ""
        AddLine "Artist", .Tags.Artist, ""
        AddLine "Copyright", .Tags.Copyright, ""
        AddLine "Product", .Tags.Product, ""
        AddLine "Creation Date", .Tags.Creation_Date, ""
        AddLine "Genre", .Tags.Genre, ""
        AddLine "Subject", .Tags.Subject, ""
        AddLine "Keywords", .Tags.Keywords, ""
        AddLine "Comments", .Tags.Comments, ""
        AddLine "Software", .Tags.Software, ""
        AddLine "Technician", .Tags.Technician, ""
        AddLine "Digitizing Date", .Tags.Digitizing_Date, ""
        AddLine "Source Form", .Tags.Source_Form, ""
        AddLine "Medium", .Tags.Medium, ""
        AddLine "Source", .Tags.Source, ""
        AddLine "Archival Location", .Tags.Archival_Location, ""
        AddLine "Commissioned by", .Tags.Commissioned_by, ""
        AddLine "Engineer", .Tags.Engineer, ""
        AddLine "Cropped", .Tags.Cropped, ""
        AddLine "Sharpness", .Tags.Sharpness, ""
        AddLine "Dimensions", .Tags.Dimensions, ""
        AddLine "Lightness", .Tags.Lightness, ""
        AddLine "Dots Per Inch", .Tags.Dots_Per_Inch, ""
        AddLine "Palette Setting", .Tags.Palette_Setting, ""
    
End With

End Sub

Private Sub AddLine(x As String, Optional Y As Variant = "", Optional SkipIfYIs As String = "///")
If Y = SkipIfYIs Then Exit Sub
If Y = "=" Then
    txtData.Text = txtData.Text & " ==== " & x & " ===============================================" & vbNewLine
ElseIf x = "" And Y = "" Then
    txtData.Text = txtData.Text & vbNewLine
Else
    txtData.Text = txtData.Text & "  " & x & ": " & Y & vbNewLine
End If
End Sub
