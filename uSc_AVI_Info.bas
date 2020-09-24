Attribute VB_Name = "uSc_AVI_Info"
'********************************************************
'*  Copyright (C) uScream 2003 - All Rights Reserved    *
'*                                                      *
'*  Contact: uscream@vip.hr                             *
'*                                                      *
'*  CHANGE HISTORY:                                     *
'*      05.07.2003. - v 1.0                             *
'*      19.07.2003. - v 1.1                             *
'*                                                      *
'*  Thanks to Philippe Duby                             *
'*                                                      *
'********************************************************
Option Explicit

Public Enum AVIFlag
    AVIF_HASINDEX = &H10
    AVIF_MUSTUSEINDEX = &H20
    AVIF_ISINTERLEAVED = &H100
    AVIF_TRUSTCKTYPE = &H800
    AVIF_WASCAPTUREFILE = &H10000
    AVIF_COPYRIGHTED = &H20000
End Enum

Public Enum AVISFlag
    AVISF_DISABLED = &H1
    AVISF_VIDEO_PALCHANGES = &H10000
End Enum

Private Type RECT '16 bytes
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type RGBQUAD '4 bytes
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type BITMAPINFO '44
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

Private Type tWAVEFORMATEX '18 bytes
        wFormatTag As Integer
        nChannels As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
        wBitsPerSample As Integer
        cbSize As Integer
End Type

Private Type WAVEFORMAT '14 bytes
        wFormatTag As Integer
        nChannels As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
End Type

Private Type MainAVIHeader '44 bytes
    dwMicroSecPerFrame As Long
    dwMaxBytesPerSec As Long
    dwReserved1 As Long
    dwFlags As Long
    dwTotalFrames As Long
    dwInitialFrames As Long
    dwStreams As Long
    dwSuggestedBufferSize As Long
    dwWidth As Long
    dwHeight As Long
    dwReserved(4) As Long
End Type

Private Type FOURCC '4 bytes
  ch0 As Byte
  ch1 As Byte
  ch2 As Byte
  ch3 As Byte
End Type

Private Type AVIStreamHeader '64 bytes
    fccType As FOURCC
    fccHandler As FOURCC
    dwFlags As Long
    wPriority As Integer
    wLanguage As Integer
    dwInitialFrames As Long
    dwScale As Long
    dwRate As Long
    dwStart As Long
    dwLength As Long
    dwSuggestedBufferSize As Long
    dwQuality As Long
    dwSampleSize As Long
    rcFrame As RECT
End Type

'#########################################
Public Type TagsType
    Name As String
    Artist As String
    Copyright As String
    Product As String
    Creation_Date As String
    Genre As String
    Subject As String
    Keywords As String
    Comments As String
    Software As String
    Technician As String
    Digitizing_Date As String
    Source_Form As String
    Medium As String
    Source As String
    Archival_Location As String
    Commissioned_by As String
    Engineer As String
    Cropped As String
    Sharpness As String
    Dimensions As String
    Lightness As String
    Dots_Per_Inch As String
    Palette_Setting As String
End Type

Public Type StreamInfoType
    Type As String * 4
    Flag As AVISFlag 'Flags
    Lenght As Long 'str1Length
    Quality As Long ''str1Quality
    Codec As String
    bPS As Long
        
    Video_Codec As String * 4 'Handler
    Video_bPP As Long 'videoBitCount
    Video_FPS As Single
    Video_Width As Long
    Video_Height As Long
    
    Audio_Codec As Integer
    Audio_Channels As Integer
    Audio_SamplePerSec As Long

End Type

Public Type AviStructure
    FileName As String
    FileSize As Long
    Lenght As Long '->Prvi Stream
    TotalFrames As Long '    dwTotalFrames
    Streams As Long '    dwStreams
    StreamsFound As Long
    Flag As AVIFlag '    dwFlags
    Width As Long '    dwWidth
    Height As Long '    dwHeight
    StreamInfo() As StreamInfoType
    VideoSize As Long
    AudioSize As Long
    Tags As TagsType
End Type

Dim AVIHeader As MainAVIHeader
Dim StreamHeader As AVIStreamHeader
Dim VideoHeader As BITMAPINFO
Dim AudioHeader As tWAVEFORMATEX

Dim AIlocal As AviStructure

Dim Dummy As Byte
Dim Buffer() As Byte

Dim HeaderSize As Long

Dim AudioCodecConst As String
Dim VideoCodecConst As String

Public Function AviInfoRead(ByVal FileName As String, AI As AviStructure) As Byte

    ReadCodecsConst
    HeaderSize = 0
    Dim AEmpty As AviStructure
    AI = AEmpty
    AIlocal = AEmpty
    AIlocal.StreamsFound = 0
        
    If FileName = "" Or Dir(FileName) = "" Then GoTo ReadErrH
    
    Open FileName For Binary As #1
    
    AIlocal.FileName = FileName
    AIlocal.FileSize = LOF(1)
'############################################
    
    Dim strPosition As String * 4
    Dim ccPosition As FOURCC

    Do
        ccPosition.ch0 = ccPosition.ch1
        ccPosition.ch1 = ccPosition.ch2
        ccPosition.ch2 = ccPosition.ch3
        Get #1, , ccPosition.ch3
        strPosition = FourCCtoStr(ccPosition)
        Select Case strPosition
            Case "avih": ReadAVIHeader
            Case "strh": ReadStreamHeader
            Case "INFO": ReadTags
        End Select
        
        HeaderSize = HeaderSize + 1
    Loop While strPosition <> "movi"
    
    DoAdditionalOps

'############################################
    Close #1
    'WriteFoundData
    AI = AIlocal
    AviInfoRead = 1
    Exit Function
ReadErrH:
    Err.Clear
    AviInfoRead = 0
End Function

Private Sub ReadAVIHeader()
    MoveBytes 4
    'ReDim Buffer(43)
    Get #1, , AVIHeader
    'CopyMemory AVIHeader, Buffer(0), 44
    
    '########################
    
    AIlocal.TotalFrames = AVIHeader.dwTotalFrames
    AIlocal.Streams = AVIHeader.dwStreams
    AIlocal.Flag = AVIHeader.dwFlags
    AIlocal.Width = AVIHeader.dwWidth
    AIlocal.Height = AVIHeader.dwHeight
    
    HeaderSize = HeaderSize + 44
End Sub


Private Sub ReadStreamHeader()
    ReDim Preserve AIlocal.StreamInfo(AIlocal.StreamsFound)
    
    MoveBytes 4
    Get #1, , StreamHeader
    'CopyMemory StreamHeader, Buffer(0), 64
    
    '########################
    
    AIlocal.StreamInfo(AIlocal.StreamsFound).Type = FourCCtoStr(StreamHeader.fccType)
    
        If AIlocal.StreamInfo(AIlocal.StreamsFound).Type <> "auds" _
        And AIlocal.StreamInfo(AIlocal.StreamsFound).Type <> "vids" Then
            ReDim Preserve AIlocal.StreamInfo(AIlocal.StreamsFound - 1)
            Exit Sub
        End If
    AIlocal.StreamInfo(AIlocal.StreamsFound).Flag = StreamHeader.dwFlags
    AIlocal.StreamInfo(AIlocal.StreamsFound).Lenght = StreamHeader.dwLength / StreamHeader.dwRate * StreamHeader.dwScale
    AIlocal.StreamInfo(AIlocal.StreamsFound).Quality = StreamHeader.dwQuality / 100
        'ovo bash i ne valja
    If Not (AIlocal.StreamsFound And AIlocal.Lenght) Then AIlocal.Lenght = StreamHeader.dwLength / StreamHeader.dwRate * StreamHeader.dwScale
    
    
    Select Case AIlocal.StreamInfo(AIlocal.StreamsFound).Type
    Case "vids": ReadVideoHeader
    Case "auds": ReadAudioHeader
    End Select
    
    AIlocal.StreamsFound = AIlocal.StreamsFound + 1
    
    HeaderSize = HeaderSize + 64
End Sub

Private Sub ReadVideoHeader()
    'ReDim Buffer(39)
    Get #1, , VideoHeader
    'CopyMemory VideoHeader, Buffer(0), 40
    
    '########################
    
    AIlocal.StreamInfo(AIlocal.StreamsFound).Video_bPP = VideoHeader.bmiHeader.biBitCount
    AIlocal.StreamInfo(AIlocal.StreamsFound).Video_Height = VideoHeader.bmiHeader.biHeight
    AIlocal.StreamInfo(AIlocal.StreamsFound).Video_Width = VideoHeader.bmiHeader.biWidth
    
    AIlocal.StreamInfo(AIlocal.StreamsFound).Codec = VideoCodecFromCode(StreamHeader.fccHandler)
    AIlocal.StreamInfo(AIlocal.StreamsFound).Video_FPS = StreamHeader.dwRate / StreamHeader.dwScale
    AIlocal.StreamInfo(AIlocal.StreamsFound).Video_Codec = FourCCtoStr(StreamHeader.fccHandler)
    
    HeaderSize = HeaderSize + 40
End Sub

Private Sub ReadAudioHeader()
    'ReDim Buffer(17)
    Get #1, , AudioHeader
    'CopyMemory AudioHeader, Buffer(0), 18
    
    
    '########################
    
    AIlocal.StreamInfo(AIlocal.StreamsFound).Audio_Codec = AudioHeader.wFormatTag
    AIlocal.StreamInfo(AIlocal.StreamsFound).Audio_Channels = AudioHeader.nChannels
    AIlocal.StreamInfo(AIlocal.StreamsFound).Audio_SamplePerSec = AudioHeader.nSamplesPerSec
    
    AIlocal.StreamInfo(AIlocal.StreamsFound).bPS = AudioHeader.nAvgBytesPerSec * 8
    AIlocal.StreamInfo(AIlocal.StreamsFound).Codec = AudioCodecFromCode(AudioHeader.wFormatTag)
    
    'AIlocal.StreamInfo.Codec = AudioHeader.wFormatTag
    
    HeaderSize = HeaderSize + 18
End Sub

Private Sub ReadTags()
    
    Dim strPosition As String * 4
    Dim ccPosition As FOURCC

    Do
        ccPosition.ch0 = ccPosition.ch1
        ccPosition.ch1 = ccPosition.ch2
        ccPosition.ch2 = ccPosition.ch3
        Get #1, , ccPosition.ch3
        strPosition = FourCCtoStr(ccPosition)
        With AIlocal.Tags
        Select Case strPosition
            
                Case "INAM": ReadSingleTag .Name
                Case "IART": ReadSingleTag .Artist
                Case "ICOP": ReadSingleTag .Copyright
                Case "IPRD": ReadSingleTag .Product
                Case "ICRD": ReadSingleTag .Creation_Date
                Case "IGNR": ReadSingleTag .Genre
                Case "ISBJ": ReadSingleTag .Subject
                Case "IKEY": ReadSingleTag .Keywords
                Case "ICMT": ReadSingleTag .Comments
                Case "ISFT": ReadSingleTag .Software
                Case "ITCH": ReadSingleTag .Technician
                Case "IDIT": ReadSingleTag .Digitizing_Date
                Case "ISRF": ReadSingleTag .Source_Form
                Case "IMED": ReadSingleTag .Medium
                Case "ISRC": ReadSingleTag .Source
                Case "IARL": ReadSingleTag .Archival_Location
                Case "ICMS": ReadSingleTag .Commissioned_by
                Case "IENG": ReadSingleTag .Engineer
                Case "ICRP": ReadSingleTag .Cropped
                Case "ISHP": ReadSingleTag .Sharpness
                Case "IDIM": ReadSingleTag .Dimensions
                Case "ILGT": ReadSingleTag .Lightness
                Case "IDPI": ReadSingleTag .Dots_Per_Inch
                Case "IPLT": ReadSingleTag .Palette_Setting
            
        End Select
        End With
        HeaderSize = HeaderSize + 1
    Loop While strPosition <> "JUNK" 'Not (ccPosition.ch0 = 0 And ccPosition.ch1 = 0 And ccPosition.ch2 = 0 And ccPosition.ch3 = 0)

End Sub

Private Sub ReadSingleTag(Tag As String)
MoveBytes 4
 
Do
    Get #1, , Dummy
    If Dummy > 31 Then Tag = Tag & Chr(Dummy)
    HeaderSize = HeaderSize + 1
Loop While Dummy <> 0
    
End Sub

Private Sub DoAdditionalOps()
    With AIlocal
    .AudioSize = 0
    Dim i As Byte
    For i = LBound(.StreamInfo) To UBound(.StreamInfo)
'        If .StreamInfo(i).Type = "auds" Then
'            .AudioSize = .AudioSize + (CDbl(.StreamInfo(i).bPS) * .StreamInfo(i).Lenght) / 8
'        End If ' => For some files lenght of audio stream is not correct!!!
        If .StreamInfo(i).Type = "auds" Then
            .AudioSize = .AudioSize + (CDbl(.StreamInfo(i).bPS) * .StreamInfo(0).Lenght) / 8
        End If
    Next
    For i = LBound(.StreamInfo) To UBound(.StreamInfo)
'        If .StreamInfo(i).Type = "vids" Then
'            .VideoSize = .FileSize - HeaderSize - .AudioSize
'            .StreamInfo(i).bPS = .VideoSize / .StreamInfo(i).Lenght * 8
'        End If
        If .StreamInfo(i).Type = "vids" Then
            .VideoSize = .FileSize - HeaderSize - .AudioSize
            .StreamInfo(i).bPS = .VideoSize / .StreamInfo(0).Lenght * 8
        End If
    Next
    End With
End Sub

Private Sub MoveBytes(Optional ByVal BNum As Long = 1)
    Dim i As Long
    For i = 1 To BNum
        Get #1, , Dummy
    Next
End Sub

Private Sub GoToByte(SString As String)
    Dim PosCurrent As FOURCC
    Dim PosStop As FOURCC
    
    PosStop.ch0 = Asc(Mid(SString, 1, 1))
    PosStop.ch1 = Asc(Mid(SString, 2, 1))
    PosStop.ch2 = Asc(Mid(SString, 3, 1))
    PosStop.ch3 = Asc(Mid(SString, 4, 1))
    
    Do
        PosCurrent.ch0 = PosCurrent.ch1
        PosCurrent.ch1 = PosCurrent.ch2
        PosCurrent.ch2 = PosCurrent.ch3
        Get #1, , PosCurrent.ch3
    Loop While (PosStop.ch0 <> PosCurrent.ch0 Or PosStop.ch1 <> PosCurrent.ch1 Or PosStop.ch2 <> PosCurrent.ch2 Or PosStop.ch3 <> PosCurrent.ch3)
End Sub


Private Function FourCCtoStr(FCC As FOURCC) As String
    FourCCtoStr = Chr(FCC.ch0) & Chr(FCC.ch1) & Chr(FCC.ch2) & Chr(FCC.ch3)
End Function

Private Sub ReadCodecsConst()
    If AudioCodecConst = "" Then
        Read_Text_File AudioCodecConst, App.Path & "\audio.txt"
    End If
    If VideoCodecConst = "" Then
        Read_Text_File VideoCodecConst, App.Path & "\video.txt"
    End If
End Sub

Private Function DajString(HTMLcode As String, strA As String, strB As String) As String
  On Error GoTo NedajString
  DajString = HTMLcode
  DajString = Right(DajString, Len(DajString) - (InStr(1, DajString, strA, vbTextCompare) + Len(strA)) + 1)
  DajString = Left(DajString, InStr(1, DajString, strB, vbTextCompare) - 1)
  Exit Function
NedajString:
  DajString = ""
  Err.Clear
End Function

Private Function Read_Text_File(TextString As String, ByVal FileName As String) As Byte

On Error GoTo Read_Text_File_Err

  Dim fs, F
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set F = fs.OpenTextFile(FileName, 1, -2)
  TextString = F.readall
  F.Close
  
  Read_Text_File = 1
  Exit Function
  
Read_Text_File_Err:
    Err.Clear
    Read_Text_File = 0
End Function

Private Function AudioCodecFromCode(Code As Integer) As String
    Dim HSTR As String
    HSTR = "0x" & Right("0000" & Hex(Code), 4) & "="
    AudioCodecFromCode = DajString(AudioCodecConst, vbNewLine & HSTR, vbNewLine)
End Function

Private Function VideoCodecFromCode(Code As FOURCC) As String
    Dim HSTR As String
    HSTR = FourCCtoStr(Code) & "="
    VideoCodecFromCode = DajString(VideoCodecConst, vbNewLine & HSTR, vbNewLine)
End Function

