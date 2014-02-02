VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFMOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ******************************
' *** DO NOT TOUCH THIS FILE ***
' ******************************

' This file holds the bare minimum of procedures to run FMOD with in Eclipse Reborn.
' Please do not mess with this file if you have no understanding of how FMOD works.

Private bInit_Audio As Boolean
Private curSong As String

Private songHandle As Long
Private streamHandle As Long
Private SoundHandle As Long

Private CurrentChannel As Long

Private Const FMOD_VERSION = 3.75

Private Enum FMOD_ERRORS
    FMOD_ERR_NONE             ' No errors
    FMOD_ERR_BUSY             ' Cannot call this command after FSOUND_Init.  Call FSOUND_Close first.
    FMOD_ERR_UNINITIALIZED    ' This command failed because FSOUND_Init was not called
    FMOD_ERR_INIT             ' Error initializing output device.
    FMOD_ERR_ALLOCATED        ' Error initializing output device, but more specifically, the output device is already in use and cannot be reused.
    FMOD_ERR_PLAY             ' Playing the sound failed.
    FMOD_ERR_OUTPUT_FORMAT    ' Soundcard does not support the features needed for this soundsystem (16bit stereo output)
    FMOD_ERR_COOPERATIVELEVEL ' Error setting cooperative level for hardware.
    FMOD_ERR_CREATEBUFFER     ' Error creating hardware sound buffer.
    FMOD_ERR_FILE_NOTFOUND    ' File not found
    FMOD_ERR_FILE_FORMAT      ' Unknown file format
    FMOD_ERR_FILE_BAD         ' Error loading file
    FMOD_ERR_MEMORY           ' Not enough memory
    FMOD_ERR_VERSION          ' The version number of this file format is not supported
    FMOD_ERR_INVALID_PARAM    ' An invalid parameter was passed to this function
    FMOD_ERR_NO_EAX           ' Tried to use an EAX command on a non EAX enabled channel or output.
    FMOD_ERR_CHANNEL_ALLOC    ' Failed to allocate a new channel
    FMOD_ERR_RECORD           ' Recording is not supported on this machine
    FMOD_ERR_MEDIAPLAYER      ' Windows Media Player not installed so cannot play wma or use internet streaming.
    FMOD_ERR_CDDEVICE         ' An error occured trying to open the specified CD device
End Enum

Private Enum FSOUND_MODES
    FSOUND_LOOP_OFF = &H1             ' For non looping samples.
    FSOUND_LOOP_NORMAL = &H2          ' For forward looping samples.
    FSOUND_LOOP_BIDI = &H4            ' For bidirectional looping samples.  (no effect if in hardware).
    FSOUND_8BITS = &H8                ' For 8 bit samples.
    FSOUND_16BITS = &H10              ' For 16 bit samples.
    FSOUND_MONO = &H20                ' For mono samples.
    FSOUND_STEREO = &H40              ' For stereo samples.
    FSOUND_UNSIGNED = &H80            ' For source data containing unsigned samples.
    FSOUND_SIGNED = &H100             ' For source data containing signed data.
    FSOUND_DELTA = &H200              ' For source data stored as delta values.
    FSOUND_IT214 = &H400              ' For source data stored using IT214 compression.
    FSOUND_IT215 = &H800              ' For source data stored using IT215 compression.
    FSOUND_HW3D = &H1000              ' Attempts to make samples use 3d hardware acceleration. (if the card supports it)
    FSOUND_2D = &H2000                ' Ignores any 3d processing.  overrides FSOUND_HW3D.  Located in software.
    FSOUND_STREAMABLE = &H4000        ' For realtime streamable samples.  If you dont supply this sound may come out corrupted.
    FSOUND_LOADMEMORY = &H8000        ' For FSOUND_Sample_Load - name will be interpreted as a pointer to data
    FSOUND_LOADRAW = &H10000          ' For FSOUND_Sample_Load/FSOUND_Stream_Open - will ignore file format and treat as raw pcm.
    FSOUND_MPEGACCURATE = &H20000     ' For FSOUND_Stream_Open - scans MP2/MP3 (VBR also) for accurate FSOUND_Stream_GetLengthMs/FSOUND_Stream_SetTime.
    FSOUND_FORCEMONO = &H40000        ' For forcing stereo streams and samples to be mono - needed with FSOUND_HW3D - incurs speed hit
    FSOUND_HW2D = &H80000             ' 2d hardware sounds.  allows hardware specific effects
    FSOUND_ENABLEFX = &H100000        ' Allows DX8 FX to be played back on a sound.  Requires DirectX 8 - Note these sounds cant be played more than once, or have a changing frequency
    FSOUND_MPEGHALFRATE = &H200000    ' For FMODCE only - decodes mpeg streams using a lower quality decode, but faster execution
    FSOUND_XADPCM = &H400000          ' For XBOX only - Describes a user sample that its contents are compressed as XADPCM
    FSOUND_VAG = &H800000             ' For PS2 only - Describes a user sample that its contents are compressed as Sony VAG format.
    FSOUND_NONBLOCKING = &H1000000    ' For FSOUND_Stream_Open - Causes stream to open in the background and not block the foreground app - stream plays only when ready.
    FSOUND_GCADPCM = &H2000000        ' For Gamecube only - Contents are compressed as Gamecube DSP-ADPCM format
    FSOUND_MULTICHANNEL = &H4000000   ' For PS2 only - Contents are interleaved into a multi-channel (more than stereo) format
    FSOUND_USECORE0 = &H8000000       ' For PS2 only - Sample/Stream is forced to use hardware voices 00-23
    FSOUND_USECORE1 = &H10000000      ' For PS2 only - Sample/Stream is forced to use hardware voices 24-47
    FSOUND_LOADMEMORYIOP = &H20000000 ' For PS2 only - "name" will be interpreted as a pointer to data for streaming and samples.  The address provided will be an IOP address
        FSOUND_IGNORETAGS = &H40000000    ' Skips id3v2 etc tag checks when opening a stream, to reduce seek/read overhead when opening files (helps with CD performance)
    FSOUND_STREAM_NET = &H80000000    ' Specifies an internet stream
    
    FSOUND_NORMAL = FSOUND_16BITS Or FSOUND_SIGNED Or FSOUND_MONO
End Enum

Private Enum FSOUND_INITMODES
    FSOUND_INIT_USEDEFAULTMIDISYNTH = &H1       'Causes MIDI playback to force software decoding.
    FSOUND_INIT_GLOBALFOCUS = &H2               'For DirectSound output - sound is not muted when window is out of focus.
    FSOUND_INIT_ENABLESYSTEMCHANNELFX = &H4     'For DirectSound output - Allows FSOUND_FX api to be used on global software mixer output!
    FSOUND_INIT_ACCURATEVULEVELS = &H8          'This latency adjusts FSOUND_GetCurrentLevels, but incurs a small cpu and memory hit
    FSOUND_INIT_PS2_DISABLECORE0REVERB = &H10   'PS2 only - Disable reverb on CORE 0 to regain SRAM
    FSOUND_INIT_PS2_DISABLECORE1REVERB = &H20   'PS2 only - Disable reverb on CORE 1 to regain SRAM
    FSOUND_INIT_PS2_SWAPDMACORES = &H40         'PS2 only - By default FMOD uses DMA CH0 for mixing, CH1 for uploads, this flag swaps them around
    FSOUND_INIT_DONTLATENCYADJUST = &H80        'Callbacks are not latency adjusted, and are called at mix time.  Also information functions are immediate
    FSOUND_INIT_GC_INITLIBS = &H100             'Gamecube only - Initializes GC audio libraries
    FSOUND_INIT_STREAM_FROM_MAIN_THREAD = &H200 'Turns off fmod streamer thread, and makes streaming update from FSOUND_Update called by the user
    FSOUND_INIT_PS2_USEVOLUMERAMPING = &H400    'PS2 only   - Turns on volume ramping system to remove hardware clicks.
    FSOUND_INIT_DSOUND_DEFERRED = &H800         'Win32 only - For DirectSound output.  3D commands are batched together and executed at FSOUND_Update.
    FSOUND_INIT_DSOUND_HRTF_LIGHT = &H1000      'Win32 only - For DirectSound output.  FSOUND_HW3D buffers use a slightly higher quality algorithm when 3d hardware acceleration is not present.
    FSOUND_INIT_DSOUND_HRTF_FULL = &H2000       'Win32 only - For DirectSound output.  FSOUND_HW3D buffers use full quality 3d playback when 3d hardware acceleration is not present.
    FSOUND_INIT_XBOX_REMOVEHEADROOM = &H4000    'XBox only - By default directsound attenuates all sound by 6db to avoid clipping/distortion.  CAUTION.  If you use this flag you are responsible for the final mix to make sure clipping / distortion doesn't happen.
    FSOUND_INIT_PSP_SILENCEONUNDERRUN = &H8000  'PSP only - If streams skip / stutter when device is powered on, either increase stream buffersize, or use this flag instead to play silence while the UMD is recovering.
End Enum

Private Declare Function FSOUND_Stream_Close Lib "fmod.dll" Alias "_FSOUND_Stream_Close@4" (ByVal Stream As Long) As Byte
Private Declare Function FSOUND_Stream_Stop Lib "fmod.dll" Alias "_FSOUND_Stream_Stop@4" (ByVal Stream As Long) As Byte
Private Declare Function FSOUND_Stream_Play Lib "fmod.dll" Alias "_FSOUND_Stream_Play@8" (ByVal channel As Long, ByVal Stream As Long) As Long
Private Declare Function FSOUND_Stream_Open Lib "fmod.dll" Alias "_FSOUND_Stream_Open@16" (ByVal FileName As String, ByVal Mode As FSOUND_MODES, ByVal offset As Long, ByVal Length As Long) As Long
Private Declare Function FMUSIC_SetMasterVolume Lib "fmod.dll" Alias "_FMUSIC_SetMasterVolume@8" (ByVal module As Long, ByVal Volume As Long) As Byte
Private Declare Function FSOUND_Init Lib "fmod.dll" Alias "_FSOUND_Init@12" (ByVal mixrate As Long, ByVal maxchannels As Long, ByVal Flags As FSOUND_INITMODES) As Byte
Private Declare Function FSOUND_Close Lib "fmod.dll" Alias "_FSOUND_Close@0" () As Long
Private Declare Function FSOUND_GetError Lib "fmod.dll" Alias "_FSOUND_GetError@0" () As FMOD_ERRORS
Private Declare Function FSOUND_SetVolume Lib "fmod.dll" Alias "_FSOUND_SetVolume@8" (ByVal channel As Long, ByVal Vol As Long) As Byte
Private Declare Function FMUSIC_LoadSong Lib "fmod.dll" Alias "_FMUSIC_LoadSong@4" (ByVal name As String) As Long
Private Declare Function FMUSIC_FreeSong Lib "fmod.dll" Alias "_FMUSIC_FreeSong@4" (ByVal module As Long) As Byte
Private Declare Function FMUSIC_PlaySong Lib "fmod.dll" Alias "_FMUSIC_PlaySong@4" (ByVal module As Long) As Byte
Private Declare Function FMUSIC_StopSong Lib "fmod.dll" Alias "_FMUSIC_StopSong@4" (ByVal module As Long) As Byte
Private Declare Function ConvCStringToVBString Lib "kernel32" Alias "lstrcpyA" (ByVal lpsz As String, ByVal pt As Long) As Long ' Notice the As Long return value replacing the As String given by the API Viewer.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'************
'* FUNCTIONS (Added by Adion)
'************
'Usage: myerrorstring = FSOUND_GetErrorString(FSOUND_GetError)
Private Function FSOUND_GetErrorString(ByVal errorcode As Long) As String
    Select Case errorcode
        Case FMOD_ERR_NONE:             FSOUND_GetErrorString = "No errors"
        Case FMOD_ERR_BUSY:             FSOUND_GetErrorString = "Cannot call this command after FSOUND_Init.  Call FSOUND_Close first."
        Case FMOD_ERR_UNINITIALIZED:    FSOUND_GetErrorString = "This command failed because FSOUND_Init was not called"
        Case FMOD_ERR_PLAY:             FSOUND_GetErrorString = "Playing the sound failed."
        Case FMOD_ERR_INIT:             FSOUND_GetErrorString = "Error initializing output device."
        Case FMOD_ERR_ALLOCATED:        FSOUND_GetErrorString = "The output device is already in use and cannot be reused."
        Case FMOD_ERR_OUTPUT_FORMAT:    FSOUND_GetErrorString = "Soundcard does not support the features needed for this soundsystem (16bit stereo output)"
        Case FMOD_ERR_COOPERATIVELEVEL: FSOUND_GetErrorString = "Error setting cooperative level for hardware."
        Case FMOD_ERR_CREATEBUFFER:     FSOUND_GetErrorString = "Error creating hardware sound buffer."
        Case FMOD_ERR_FILE_NOTFOUND:    FSOUND_GetErrorString = "File not found"
        Case FMOD_ERR_FILE_FORMAT:      FSOUND_GetErrorString = "Unknown file format"
        Case FMOD_ERR_FILE_BAD:         FSOUND_GetErrorString = "Error loading file"
        Case FMOD_ERR_MEMORY:           FSOUND_GetErrorString = "Not enough memory "
        Case FMOD_ERR_VERSION:          FSOUND_GetErrorString = "The version number of this file format is not supported"
        Case FMOD_ERR_INVALID_PARAM:    FSOUND_GetErrorString = "An invalid parameter was passed to this function"
        Case FMOD_ERR_NO_EAX:           FSOUND_GetErrorString = "Tried to use an EAX command on a non EAX enabled channel or output."
        Case FMOD_ERR_CHANNEL_ALLOC:    FSOUND_GetErrorString = "Failed to allocate a new channel"
        Case FMOD_ERR_RECORD:           FSOUND_GetErrorString = "Recording is not supported on this machine"
        Case FMOD_ERR_MEDIAPLAYER:      FSOUND_GetErrorString = "Required Mediaplayer codec is not installed"
        Case FMOD_ERR_CDDEVICE:         FSOUND_GetErrorString = "An error occured trying to open the specified CD device"
        Case Else:                      FSOUND_GetErrorString = "Unknown error"
    End Select
End Function

Public Function init() As Boolean
Dim result As Boolean

    On Error GoTo errorhandler
    
    If App.LogMode = 0 Then Exit Function
    
    ' init music engine
    result = FSOUND_Init(44100, 32, FSOUND_INIT_USEDEFAULTMIDISYNTH)
    If Not result Then GoTo errorhandler
    
    CurrentChannel = 1
    
    ' return positive
    bInit_Audio = True
    Exit Function
    
errorhandler:
    MsgBox "Error"
    bInit_Audio = False
End Function

Public Sub Destroy()
    ' destroy music engine
    Music_Stop
    FSOUND_Close
    bInit_Audio = False
    curSong = vbNullString
    CurrentChannel = 1
End Sub

Public Sub Music_Play(ByVal song As String)
    If Not bInit_Audio Then Exit Sub
    If Options.Music = 0 Then Exit Sub
    
    ' exit out early if we have the system turned off
    If Options.Music = 0 Then Exit Sub
    
    ' does it exist?
    If Not FileExist(App.path & MUSIC_PATH & song) Then Exit Sub
    
    ' don't re-start currently playing songs
    If curSong = song Then Exit Sub
    
    ' stop the existing music
    Music_Stop
    
    ' find the extension
    Select Case Right$(song, 4)
        Case ".mid", ".s3m", ".mod"
            ' open the song
            songHandle = FMUSIC_LoadSong(App.path & MUSIC_PATH & song)
            ' play it
            FMUSIC_PlaySong songHandle
            ' set volume
            FMUSIC_SetMasterVolume songHandle, Options.Volume
            
        Case ".wav", ".mp3", ".ogg", ".wma"
            ' open the stream
            streamHandle = FSOUND_Stream_Open(App.path & MUSIC_PATH & song, FSOUND_LOOP_NORMAL, 0, 0)
            ' play it
            FSOUND_Stream_Play 0, streamHandle
            ' set volume
            FSOUND_SetVolume streamHandle, Options.Volume
        Case Else
            Exit Sub
    End Select
    
    ' new current song
    curSong = song
End Sub

Public Sub Music_Stop()
    If Not streamHandle = 0 Then
        ' stop stream
        FSOUND_Stream_Stop streamHandle
        ' destroy
        FSOUND_Stream_Close streamHandle
        streamHandle = 0
    End If
    
    If Not songHandle = 0 Then
        ' stop song
        FMUSIC_StopSong songHandle
        ' destroy
        FMUSIC_FreeSong songHandle
        songHandle = 0
    End If
    
    ' no music
    curSong = vbNullString
End Sub

Public Sub Sound_Play(ByVal Sound As String, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1)
    If Not bInit_Audio Then Exit Sub
    If Options.Sound = 0 Then Exit Sub
    Dim Volume As Long
    
    ' exit out early if we have the system turned off
    If Options.Sound = 0 Then Exit Sub
    
    ' does it exist?
    If Not FileExist(App.path & SOUND_PATH & Sound) Then Exit Sub
    
    ' Is it a file we can play?
    If Right$(Sound, 4) <> ".wav" And Right$(Sound, 4) <> ".mp3" And Right$(Sound, 4) <> ".ogg" And Right$(Sound, 4) <> ".wma" Then Exit Sub
    
    ' open the stream
    SoundHandle = FSOUND_Stream_Open(App.path & SOUND_PATH & Sound, FSOUND_LOOP_OFF, 0, 0)
    ' set volume
    Volume = CalculateSoundVolume(x, y)
    FSOUND_SetVolume SoundHandle, Volume
    ' play it
    FSOUND_Stream_Play CurrentChannel, SoundHandle
    
    CurrentChannel = CurrentChannel + 1
    If CurrentChannel > 32 Then CurrentChannel = 1
End Sub

Function CalculateSoundVolume(SoundX As Long, SoundY As Long) As Byte
Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long, distance As Double
    If InGame = False Then CalculateSoundVolume = Options.Volume: Exit Function
    If Not bInit_Audio Then Exit Function
    If Options.Sound = 0 Then Exit Function
    
    If SoundX > -1 Or SoundY > -1 Then
        X1 = (myChar.x * 32) + myChar.xOffset
        Y1 = (myChar.y * 32) + myChar.yOffset
        X2 = (SoundX * 32) + 16
        Y2 = (SoundY * 32) + 16
        If ((X1 - X2) ^ 2) + ((Y1 - Y2) ^ 2) < 0 Then
            distance = Sqr(((X1 - X2) ^ 2) + ((Y1 - Y2) ^ 2) * -1)
        Else
            distance = Sqr(((X1 - X2) ^ 2) + ((Y1 - Y2) ^ 2))
        End If
        If distance >= 256 Then
            CalculateSoundVolume = 0
        Else
            CalculateSoundVolume = Options.Volume - ((distance / 256) * Options.Volume)
            If CalculateSoundVolume = 0 Then CalculateSoundVolume = 1
        End If
    Else
        CalculateSoundVolume = Options.Volume
    End If
End Function
