VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AudioEncoder"
   ClientHeight    =   2490
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   2985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1500
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   660
      List            =   "Form1.frx":002E
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1320
      Width           =   675
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0071
      Left            =   0
      List            =   "Form1.frx":0093
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   900
      Width           =   2955
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Text            =   "D:\MP3\001\Move Over Darling.wav"
      Top             =   240
      Width           =   2955
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Encode"
      Height          =   375
      Left            =   1500
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Info"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Bitrate"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   1380
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Quality"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   660
      Width           =   2955
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00.00%"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2220
      Width           =   2955
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Wav File"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2955
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This project and modifications to LAME dll was made by Arto Rusanen
' http://www.4dsoftware.8m.com


' Credits...

' LAME was originally developed by Mike Cheng (www.uq.net.au/~zzmcheng).
' Now maintained by Mark Taylor (www.sulaco.org/mp3).

' You can find LAME and its source from
' http://www.mp3dev.org

' Thanks to orginal maker of Exception Handler.. Sorry I can't remember
' where I orginally find it...

Option Explicit

Private Cancelled As Boolean

Private Sub Command3_Click()
  Cancelled = True
End Sub

Private Sub Form_Load()
  ' Lets initialize Exception Handler
  ' VB doesn't crash so often when I debugged code...
  
  SetUnhandledExceptionFilter AddressOf MyExceptionFilter

  Combo1.ListIndex = 1
  Combo2.ListIndex = 9
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Remove exception filter...
  SetUnhandledExceptionFilter 0
End Sub


Private Sub Command1_Click()
  ' First test... Let's try to get info out of DLL
   
  Dim TempInfo As PBE_VERSION
  beVersion = TempInfo
  MsgBox "Version " & TempInfo.byDLLMajorVersion & "." & TempInfo.byMinorVersion & vbCrLf & _
          "Homepage: " & StripNulls(TempInfo.zHomepage) & vbCrLf & _
          "Date: " & TempInfo.byDay & "." & TempInfo.byMonth & "." & TempInfo.wYear, vbInformation, _
          "LAME Information"
          

  ' Good.. it worked... Easy wasn't it...
End Sub


Private Sub Command2_Click()
  ' Now.. This one is going to be little bit complicated....
  
  'on error GoTo ErrHandler
  ChDrive App.Path
  ChDir App.Path
  
  ' Check that file exists...
  If Dir(Text1) = "" Then
    MsgBox "File not found...", vbCritical, "Critical Error"
    Exit Sub
  End If
  
  Cancelled = False
  Command3.Visible = True
  
  ' Fill beConfig structure....
  Dim beConfig As PBE_CONFIG
  beConfig.dwConfig = BE_CONFIG_LAME
  
  With beConfig.format.LHV1
    '// this are the default settings for testcase.wav
    .dwStructVersion = 1
    .dwStructSize = Len(beConfig)
    .dwSampleRate = 44100         '// INPUT FREQUENCY
    .dwReSampleRate = 0           '// DON"T RESAMPLE
    .nMode = BE_MP3_MODE_JSTEREO  '// OUTPUT IN STREO
    '.dwBitrate = 128             '// MINIMUM BIT RATE
    '.nPreset = LQP_HIGH_QUALITY  '// QUALITY PRESET SETTING
    .dwMpegVersion = MPEG1        '// MPEG VERSION (I or II)
    .dwPsyModel = 0               '// USE DEFAULT PSYCHOACOUSTIC MODEL
    .dwEmphasis = 0               '// NO EMPHASIS TURNED ON
    .bOriginal = True             '// SET ORIGINAL FLAG
    .bNoRes = True                '// No Bit resorvoir
    
    Select Case Combo1.ListIndex
      Case 0: .nPreset = LQP_LOW_QUALITY
      Case 1: .nPreset = LQP_NORMAL_QUALITY
      Case 2: .nPreset = LQP_HIGH_QUALITY
      Case 3: .nPreset = LQP_VOICE_QUALITY
      Case 4: .nPreset = LQP_PHONE
      Case 5: .nPreset = LQP_RADIO
      Case 6: .nPreset = LQP_TAPE
      Case 7: .nPreset = LQP_HIFI
      Case 8: .nPreset = LQP_CD
      Case 9: .nPreset = LQP_STUDIO
      Case Else
        MsgBox "You didn't select quality..."
        Exit Sub
    End Select
    If Combo2.Text <> "" Then
      .dwBitrate = Val(Combo2.Text)
    Else
      MsgBox "You didn't select Bitrate..."
    End If
  End With

  Dim error As Long
  Dim dwSamples As Long, dwMP3Buffer As Long, hbeStream As Long
  
  ' Init MP3 Stream
  error = beInitStream(VarPtr(beConfig), VarPtr(dwSamples), VarPtr(dwMP3Buffer), VarPtr(hbeStream))
    
  '// Check result
  If error <> BE_ERR_SUCCESSFUL Then
    Err.Raise error, "Lame", GetErrorString(error)
  End If
  
  
  ' Open Files...
  Dim toRead As Long, toWrite As Long
  Dim Done As Long
  Dim length As Long
  
  length = FileLen(Text1)
  
  Dim ReadFile As clsFileIo
  Set ReadFile = New clsFileIo
  ReadFile.OpenFile Text1
  
  Dim WriteFile As clsFileIo
  Set WriteFile = New clsFileIo
  WriteFile.OpenFile ChangeExt(Text1, "mp3")
  
  
  ' Allocate memory for buffers... :)
  Dim WavPtr1 As Long
  Dim WavPtr2 As Long
  Dim MP3Ptr1 As Long
  Dim MP3Ptr2 As Long
  WavPtr1 = GlobalAlloc(&H40, dwSamples * 2)
  WavPtr2 = GlobalLock(WavPtr1)
  MP3Ptr1 = GlobalAlloc(&H40, dwMP3Buffer)
  MP3Ptr2 = GlobalLock(MP3Ptr1)
  
  'Skip WAV header
  Dim Temp(1 To 44) As Byte
  Call ReadFile.ReadBytes(VarPtr(Temp(1)), 44)
  
  ' And here we go....
  Do While Done < length
    '//set up how much to readinto the buffer
    If Done + dwSamples * 2 < length Then
      toRead = dwSamples * 2
    Else
      toRead = length - Done
    End If
    
    ' Read into buffer
    Call ReadFile.ReadBytes(WavPtr2, toRead)
    
    Done = Done + toRead
    toRead = toRead / 2
    
    ' Encode buffer
    error = beEncodeChunk(hbeStream, toRead, WavPtr2, MP3Ptr2, VarPtr(toWrite))
    
    ' Check result...
    If error <> BE_ERR_SUCCESSFUL Then
      Call beCloseStream(hbeStream)
      Err.Raise error, "Lame", GetErrorString(error)
    End If
    
    ' Write Buffer...
    If toWrite > 0 Then
      Call WriteFile.WriteBytes(MP3Ptr2, toWrite)
    End If
    
    ' Report status to user....
    Status.Caption = format(Done / length * 100, "00.00") & "%"
    If Cancelled = True Then Exit Do
    DoEvents
  Loop
  
  ' Deinitialize stream and write last bytes to MP3
  error = beDeinitStream(hbeStream, MP3Ptr2, VarPtr(toWrite))

  If toWrite > 0 Then
    WriteFile.WriteBytes MP3Ptr2, toWrite
  End If
  
  
  ' Clear buffers....
  GlobalFree MP3Ptr2
  GlobalFree WavPtr2
  
  ' Close files
  Call WriteFile.CloseFile
  Call ReadFile.CloseFile
  
  Set WriteFile = Nothing
  Set ReadFile = Nothing
  
  ' Close stream
  Call beCloseStream(hbeStream)
  
  Command3.Visible = False
  
  ' WriteVBRHeader (if we use variable bitrate...)
  'Call beWriteVBRHeader(ChangeExt(Text1, "mp3"))

  Exit Sub
  
ErrHandler:
  ' Damn.. Something went wrong and this one should tell what...
  ' At time of debugin.. This one was place where code ended all the time...
  ' But now it should never come into this point....
  
  MsgBox Err.Description, vbCritical, "Critical error..."
  If WavPtr2 Then GlobalFree WavPtr2
  If MP3Ptr2 Then GlobalFree MP3Ptr2
  
  WriteFile.CloseFile
  ReadFile.CloseFile
  Set WriteFile = Nothing
  Set ReadFile = Nothing
  
  Err.Clear
  Command3.Visible = False
  Exit Sub
End Sub


