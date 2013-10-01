VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIpImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileA Lib "kernel32" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpFileSizeHigh As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function StringFromGUID2 Lib "ole32.dll" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long

Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As IUnknown, Image As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal width As Long, ByVal height As Long, ByVal Stride As Long, ByVal PixelFormat As Long, scan0 As Any, BITMAP As Long) As Long
Private Declare Function GdipGetImageBounds Lib "GdiPlus.dll" (ByVal nImage As Long, srcRect As RECTF, srcUnit As Long) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As Long) As Long
Private Declare Function GdipGetImageRawFormat Lib "gdiplus" (ByVal hImage As Long, ByVal GUID As Long) As Long
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal hImage As Long, PixelFormat As Long) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As Long
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal Stream As IUnknown, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipBitmapLockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mRect As RECTI, ByVal mFlags As Long, ByVal mPixelFormat As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" (ByVal hbm As Long, ByVal hpal As Long, ByRef pbitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromGraphics Lib "GdiPlus.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal ptarget As Long, ByRef pbitmap As Long) As Long
Private Declare Function GdipGetImagePaletteSize Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef psize As Long) As Long
Private Declare Function GdipGetImagePalette Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef palette As Any, ByVal psize As Long) As Long

Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, ByRef pBitmapInfo As Any, ByVal un As Long, ByRef lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Type BITMAPINFOHEADER   ' structure used within icon image data
    biSize As Long
    biWidth As Long
    biHeight As Long            ' always doubled for icons/cursors
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As Long
End Type
Private Type RECTF
    nLeft As Single
    nTop As Single
    nWidth As Single
    nHeight As Single
End Type
Private Type RECTI
    nLeft As Long
    nTop As Long
    nWidth As Long
    nHeight As Long
End Type
Private Type ImageCodecInfo
    ClassID(0 To 3)   As Long
    FormatID(0 To 3)  As Long
    CodecName         As Long
    DllName           As Long
    FormatDescription As Long
    FilenameExtension As Long
    MimeType          As Long
    Flags             As Long
    Version           As Long
    SigCount          As Long
    SigSize           As Long
    SigPattern        As Long
    SigMask           As Long
End Type
Private Type BitmapData
    width As Long
    height As Long
    Stride As Long
    PixelFormat As Long
    Scan0Ptr As Long
    ReservedPtr As Long
End Type

Public Enum ImageTypeConstants
    imageUnknown = 0
    imageBitmap = 1
    imagePNG = 7
End Enum
Public Enum LockModeConstants
    ImageLockModeRead = &H1
    ImageLockModeWrite = &H2
    ImageLockModeUserInputBuf = &H4
End Enum
Public Enum ImageColorFormatConstants
    PixelFormat32bppARGB = &H26200A 'Specifies that the format is 32 bits per pixel; 8 bits each are used for the alpha, red, green, and blue components.
    PixelFormat32bppPARGB = &HE200B 'Specifies that the format is 32 bits per pixel; 8 bits each are used for the alpha, red, green, and blue components. The red, green, and blue components are premultiplied according to the alpha component.
    PixelFormat32bppRGB = &H22009 'Specifies that the format is 32 bits per pixel; 8 bits each are used for the red, green, and blue components. The remaining 8 bits are not used.
    PixelFormat1bppIndexed = &H30101 'Specifies that the format is 1 bit per pixel, indexed.
    PixelFormat4bppIndexed = &H30402 'Specifies that the format is 4 bits per pixel, indexed.
    PixelFormat8bppIndexed = &H30803 'Specifies that the format is 8 bits per pixel, indexed.
    PixelFormat16bppGrayScale = &H101004 'Specifies that the format is 16 bits per pixel, grayscale.
    PixelFormat16bppRGB555 = &H21005 'Specifies that the format is 16 bits per pixel; 5 bits each are used for the red, green, and blue components. The remaining bit is not used.
    PixelFormat16bppRGB565 = &H21006 'Specifies that the format is 16 bits per pixel; 5 bits are used for the red component, 6 bits are used for the green component, and 5 bits are used for the blue component.
    PixelFormat48bppRGB = &H10300C 'Specifies that the format is 48 bits per pixel; 16 bits each are used for the red, green, and blue components.
    PixelFormat24bppRGB = &H21808 'Specifies that the format is 24 bits per pixel; 8 bits each are used for the red, green, and blue components.
End Enum

Private Enum pvCleanUpEnum  ' used locally
    cuDestroyAll = -1&          ' destroy all objects related to current image
    cuDestroyClone = 1&         ' destroy clone only, if it exists (GIF/TIFF images only)
    cuResetAttrs = 2&           ' reset attributes only
    cuDestroyImage = 4&         ' destroy source image
End Enum

Private Const png_Signature1 As Long = 1196314761   ' PNG signature is 8 bytes
Private Const png_Signature2 As Long = 169478669
Private Const bmp_Signature As Integer = &H4D42&    ' BMP signature is 2 bytes
Private Const UnitPixel As Long = 2&                ' GDI+ constant


Private m_Image(0 To 1) As Long             ' GDI+ image handle
Private m_Token As cGDIpToken               ' reference to GDI Token class
Private m_Size As RECTF                     ' Image dimensions
' There are some bugs with GDI+ and here is one. If applying attributes, flipping, cloning an image, and more,
' one can no longer select an image within a multi-image format; the result is always the 1st frame/page.
' So to workaround this bug, we create a "clone" that is actually a PNG from the current frame/page.
' Then that "clone" is used to render with mirroring/attributes. The routines herein remmove the
' clone when no longer needed.  A similar issue is with MetaFiles, they can be flipped, but cannot be
' rendered with attributes, so WMF/EMF will be converted to Bitmap for rendering.
Private m_SourceIndex As Long               ' ref to original image or PNG clone of TIFF,GIF,WMF,EMF as needed
' m_Image & m_Source usage for clarification
' Image Type        m_Image(0)  m_Source(0)     m_Image(1)  m_Source(1)
'   PNG             yes         never           never       never
'   JPG             yes         never           never       never
'   BMP             yes         never           never       never
'   GIF             yes         yes**1          Cloned**2   Cloned**2
'   TIFF            yes         yes**1          Cloned**2   Cloned**2
'   ICO/CUR         yes         yes**1          never       never
'   WMF/EMF         yes         yes**1          never       never
' Notes
' PNG, JPG, BMP, 1-page TIFFs are always converted to 1,4,8,24,32pARGB bpp bitmaps
' ICO, CUR, WMF, EMF are converted to 32bpp pARGB internally only
' (**1) Original source data/bytes are maintained so that multiple images from the resource can be extracted
'       and so that original format can be returned on demand
'       Note that the CUR/ICO/WMF/EMF source may be either stdPicture or IIStream depending on how the source was loaded
'       Exception: Single page TIFFs do not cache original source data
'       Note: ICO/CUR are cached for ease. One can easily add a BitmapToIcon routine to create the appropriate icon/cursor format
' (**2) When multiple images exist in resource, a clone (PNG format) is created if to be rendered with attributes.
'       Any created clone is destroyed when the attributes are reset

' Return the GDI+ image handle
Public Property Get Handle() As Long
    Handle = m_Image(m_SourceIndex)
End Property

' Return the GDI+ image height
Public Property Get height() As Long
    height = m_Size.nHeight
End Property

Public Function LoadPicture_FileName(FileName As String, TokenClass As cGDIpToken, Optional ClearAttributes As Boolean = True) As Boolean

    ' Unicode compatible if the FileName passes contains unicode characters
    ' Routine simply opens the file, reads the data, and passes to the LoadPicture_Stream function
    If TokenClass Is Nothing Then Exit Function
    If TokenClass.Token = 0& Then Exit Function
    
    Const GENERIC_READ As Long = &H80000000
    Const OPEN_EXISTING = &H3
    Const FILE_SHARE_READ = &H1
    Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
    Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
    Const FILE_ATTRIBUTE_READONLY As Long = &H1
    Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
    Const FILE_ATTRIBUTE_NORMAL = &H80&
    Const INVALID_HANDLE_VALUE As Long = -1&
    
    Dim Flags As Long, hFile As Long, bData() As Byte
    
    On Error GoTo ExitRoutine
    Flags = FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_NORMAL _
            Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM
    If IsWindowUnicode(GetDesktopWindow()) = 0& Then
        hFile = CreateFileA(FileName, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, Flags, 0&)
    Else
        hFile = CreateFileW(StrPtr(FileName), GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, Flags, 0&)
    End If
    
    If Not (hFile = INVALID_HANDLE_VALUE Or hFile = 0&) Then
        Flags = GetFileSize(hFile, 0&)
        ReDim bData(0 To Flags - 1&)
        ReadFile hFile, bData(0), Flags, Flags, ByVal 0&
        CloseHandle hFile
        If Flags > UBound(bData) Then LoadPicture_FileName = LoadPicture_Stream(bData(), TokenClass, ClearAttributes)
    End If
    
ExitRoutine:
End Function

Public Function LoadPicture_FromNothing(width As Long, height As Long, graphics As Long, TokenClass As cGDIpToken) As Boolean

    ' Unicode compatible if the FileName passes contains unicode characters
    ' Routine simply opens the file, reads the data, and passes to the LoadPicture_Stream function
    If TokenClass Is Nothing Then Exit Function
    If TokenClass.Token = 0& Then Exit Function
    
    Dim hImage As Long
    
    hImage = pvCreateSourcelessImage2(hImage, width, height, 1&, TokenClass, graphics)
                    If hImage Then
                        m_Image(0) = hImage
                        m_Size.nWidth = width
                        m_Size.nHeight = height
                        LoadPicture_FromNothing = True
                    End If
                    
                    
                    Exit Function
    
ExitRoutine:
End Function

Public Function LoadPicture_Stream(Stream() As Byte, TokenClass As cGDIpToken, Optional ByVal ClearAttributes As Boolean = True) As Boolean
    
    ' must pass a 1D, zero-bound array
    
    If TokenClass Is Nothing Then Exit Function
    If TokenClass.Token = 0& Then Exit Function
    
    Dim hImage As Long, IStream As IUnknown
    Dim lClearFlag As pvCleanUpEnum
    Dim lSize As Long, lValue As Long, lCount As Long, lLength As Long
    Dim arrPtr As Long, Index As Long, lType As Long
    
    On Error GoTo ExitRoutine
    lValue = Not Stream()
    Debug.Assert App.hInstance
    If lValue = -1& Then Exit Function ' empty array
    If LBound(Stream) <> 0 Then Exit Function ' must be zero bound; just lazy on my part

    If m_Token Is Nothing Then
        Set m_Token = TokenClass
        TokenClass.AddUser Me
    End If
    If ClearAttributes Then lClearFlag = cuDestroyAll Else lClearFlag = cuDestroyAll Xor cuResetAttrs

    ' /// Testing for icons/cursors. GDI+ can't seem to load cursors and has many issues with various types of icons
    ' See the pvIconTo32bpp routine for more info
    lLength = UBound(Stream) + 1&
    If lLength > 5& Then
        On Error Resume Next
        CopyMemory lType, Stream(2), 2&
        If lType = 1 Or lType = 2 Then ' validate: must be 1 or 2 to be an icon/cursor
            CopyMemory lCount, Stream(4), 2& ' count of icon/cursor in resource
            If lCount > 0& Then
                If lCount * 16& + 6& <= lLength Then  ' minimial bytes required for now
                    For Index = 0& To lCount - 1&
                        arrPtr = 14& + Index * 16&
                        CopyMemory lSize, Stream(arrPtr), 4&        ' bytes in resource
                        CopyMemory lValue, Stream(arrPtr + 4&), 4&  ' offset in resource
                        If lValue < 1& Or lSize < 1& Then Exit For  ' valid entries?
                        If lValue + lSize > lLength Then Exit For   ' enough bytes in resource?
                        ' let's validate the header size, bitcount, width, height
                        CopyMemory lSize, Stream(lValue), 4&
                        If lSize = png_Signature1 Then
                            CopyMemory lSize, Stream(lValue + 4&), 4&
                            If lSize <> png_Signature2 Then Exit For    ' invalid header size
                        Else
                            If lSize < 40& Then Exit For                ' invalid header size
                            CopyMemory lSize, Stream(lValue + 4&), 4&
                            If lSize < 1& Then Exit For                 ' invalid width
                            CopyMemory lSize, Stream(lValue + 8&), 4&
                            If lSize < 1& Then Exit For                 ' invalid height
                            CopyMemory lSize, Stream(lValue + 14&), 2&
                            Select Case (lSize And &HFFFF&)             ' validate bit count
                                Case 1, 2, 4, 8 ' supported
                                Case 16, 24, 32 ' supported
                                Case Else       ' not supported or not an icon format
                                    Exit For
                            End Select
                        End If
                        If Err Then
                            Err.Clear                               ' overflow?
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
        On Error GoTo ExitRoutine
    End If
    
    ' see if this is a 32bit bitmap & process if it is
    If pvProcessAlphaBitmap(Stream(), 0&, 0&, lClearFlag) = True Then
        LoadPicture_Stream = True
    Else
        Set IStream = pvStreamFromArray(VarPtr(Stream(0)), UBound(Stream) + 1&)
        If Not IStream Is Nothing Then
            If GdipLoadImageFromStream(IStream, hImage) = 0& Then
                lType = pvGetImageType(hImage)                      ' what type image do we have?
                Call GdipGetImagePixelFormat(hImage, lValue)    ' get original pixel format
                hImage = pvCreateSourcelessImage(hImage, 0&, 0&, 0) ' convert to sourceless
                If hImage = 0& Then Exit Function
                Call pvCleanUp(lClearFlag)                      ' clear existing image
                m_Image(m_SourceIndex) = hImage
                Call GdipGetImageBounds(hImage, m_Size, UnitPixel)
                LoadPicture_Stream = True
            End If
        End If
    End If
    
ExitRoutine:
End Function

Public Function SaveAsPNG(outArray() As Byte) As Boolean

    ' saves image as a PNG
    ' NOTE: If source is a multi-frame/page image, only the current frame/page will be saved
    
    Dim uEncCLSID(0 To 3) As Long, IIStream As IUnknown, hImage As Long
    Const MimeType As String = "image/png"
    
    If m_Image(m_SourceIndex) Then
        If pvGetEncoderClsID(MimeType, uEncCLSID) <> -1& Then
            Set IIStream = pvStreamFromArray(0&, 0&)
            If Not IIStream Is Nothing Then
                hImage = m_Image(m_SourceIndex)
                If hImage Then
                    If GdipSaveImageToStream(hImage, IIStream, uEncCLSID(0), ByVal 0&) = 0& Then
                        SaveAsPNG = pvStreamToArray(ObjPtr(IIStream), outArray())
                    End If
                End If
            End If
        End If
    End If
End Function

' Returns the image width
Public Property Get width() As Long
    width = m_Size.nWidth
End Property

Private Sub pvCleanUp(Mode As pvCleanUpEnum)
    ' m_Image & m_Source usage for clarification
    ' Image Type        m_Image(0)  m_Source(0)     m_Image(1)  m_Source(1)
    '   PNG             yes         never           never       never
    '   JPG             yes         never           never       never
    '   BMP             yes         never           never       never
    '   GIF             yes         yes**1          Cloned**2   Cloned**2
    '   TIFF            yes         yes**1          Cloned**2   Cloned**2
    '   ICO/CUR         yes         yes**1          never       never
    '   WMF/EMF         yes         yes**1          never       never
    ' Notes
    ' PNG, JPG, BMP, 1-page TIFFs are always converted to 1,4,8,24,32pARGB bpp bitmaps
    ' ICO, CUR, WMF, EMF are converted to 32bpp pARGB internally only
    ' (**1) Original source data/bytes are maintained so that multiple images from the resource can be extracted
    '       and so that original format can be returned on demand
    '       Note that the CUR/ICO/WMF/EMF source may be either stdPicture or IIStream depending on how the source was loaded
    '       Exception: Single page TIFFs do not cache original source data
    '       Note: ICO/CUR are cached for ease. One can easily add a BitmapToIcon routine to create the appropriate icon/cursor format
    ' (**2) When multiple images exist in resource, a clone (PNG format) is created if to be rendered with attributes.
    '       Any created clone is destroyed when the attributes are reset
    
    If (Mode And cuResetAttrs) Then
        Mode = Mode Or cuDestroyClone   ' get rid of clone if needed
    End If
    
    If (Mode And cuDestroyClone) Then   ' destroy clone if it exists
        If m_Image(1) Then
            GdipDisposeImage m_Image(1)
            m_Image(1) = 0&
            m_SourceIndex = 0&
        End If
    End If
    
    If (Mode And cuDestroyImage) Then   ' destroy core image
        If m_Image(0) Then
            GdipDisposeImage m_Image(0)
            m_Image(0) = 0&
        End If
        m_Size.nHeight = 0!             ' reset its properties
        m_Size.nWidth = 0!
    End If
    
End Sub

Private Function pvCreateSourcelessImage2(hImage As Long, width As Long, height As Long, flipFlag As Long, TokenClass As cGDIpToken, hGraphics1 As Long) As Long

    ' function creates a stand-alone GDI+ image from a linked image
    ' The routine here follows the instructions given at this MS KB article link
    '       http://support.microsoft.com/kb/814675
    Dim tSize As RECTF, tSizeI As RECTI, BHI As BITMAPINFO
    Dim tBMPsrc As BitmapData
    Dim cRender As cGDIpRenderer, pal() As Byte
    Dim oldImage As Long, newImage As Long, hObj As Long
    Dim dDC As Long, Depth As Long, hGraphics As Long, dibPtr As Long
    
    If width = 0& Or height = 0& Then               ' get size of image
        Call GdipGetImageBounds(hImage, tSize, UnitPixel)
        width = tSize.nWidth: height = tSize.nHeight
    End If
    Call GdipGetImagePixelFormat(hImage, Depth)
    Select Case Depth
        Case PixelFormat1bppIndexed, PixelFormat4bppIndexed, PixelFormat8bppIndexed
            With BHI.bmiHeader
                If Depth = PixelFormat1bppIndexed Then
                    .biBitCount = 1
                ElseIf Depth = PixelFormat4bppIndexed Then
                    .biBitCount = 4
                Else
                    .biBitCount = 8
                End If
                Call GdipGetImagePaletteSize(hImage, .biClrImportant)
                ReDim pal(0 To .biClrImportant + 7&)
                GdipGetImagePalette hImage, pal(0), .biClrImportant
                CopyMemory BHI.bmiColors(0), pal(8), .biClrImportant
                .biClrImportant = .biClrImportant \ 4&
                .biClrUsed = 2 ^ .biBitCount
                Erase pal()
            End With
        Case PixelFormat24bppRGB, PixelFormat16bppGrayScale, PixelFormat16bppRGB555, PixelFormat16bppRGB565, PixelFormat48bppRGB
            BHI.bmiHeader.biBitCount = 24
                
        Case Else
            Set cRender = New cGDIpRenderer
            cRender.AttachTokenClass TokenClass
            oldImage = m_Image(m_SourceIndex)                   ' using the cGDIpRenderer class
            m_Image(m_SourceIndex) = hImage                     ' which requires this class to reflect correct image handle
            hGraphics = hGraphics1 ' create a generic graphics object
            If hGraphics Then                                   ' then create new GDI+ bitmap
                If GdipCreateBitmapFromGraphics(width, height, hGraphics, newImage) = 0& Then
                    cRender.DestroyHGraphics hGraphics
                    Call GdipGetImageBounds(hImage, tSize, UnitPixel)
                    m_Image(m_SourceIndex) = newImage           ' create graphics object around new bitmap
                    hGraphics = cRender.CreateGraphicsFromImageClass(Me)
                    If hGraphics Then                           ' draw the passed bitmap onto the new bitmap & clean up
                        Call cRender.RenderToHGraphics(hImage, hGraphics, 0&, 0&, width, height, tSize.nLeft, tSize.nTop, tSize.nWidth, tSize.nHeight)
                        cRender.DestroyHGraphics hGraphics
                        If flipFlag Then GdipImageRotateFlip newImage, flipFlag
                        pvCreateSourcelessImage2 = newImage
                    End If
                Else
                    cRender.DestroyHGraphics hGraphics
                End If
            End If
            m_Image(m_SourceIndex) = oldImage
    End Select
    
    If BHI.bmiHeader.biBitCount Then                            ' handle paletted & 24bit bitmaps here
        tSizeI.nHeight = height: tSizeI.nWidth = width
        With BHI.bmiHeader
            .biHeight = height
            .biPlanes = 1
            .biSize = 40
            .biWidth = width
            If .biBitCount = 24 Then
                If GdipBitmapLockBits(hImage, tSizeI, ImageLockModeRead, PixelFormat24bppRGB, tBMPsrc) Then .biBitCount = 0& 'flag meaning failure
            Else
                If GdipBitmapLockBits(hImage, tSizeI, ImageLockModeRead, Depth, tBMPsrc) Then .biBitCount = 0&
            End If
        End With
        If BHI.bmiHeader.biBitCount Then
            dDC = GetDC(GetDesktopWindow())
            hObj = CreateDIBSection(dDC, BHI, 0&, dibPtr, 0&, 0&)
            ReleaseDC GetDesktopWindow(), dDC
            If hObj Then
                CopyMemory ByVal dibPtr, ByVal tBMPsrc.Scan0Ptr, tBMPsrc.Stride * height
                GdipBitmapUnlockBits newImage, tBMPsrc
                GdipCreateBitmapFromHBITMAP hObj, 0&, newImage
                DeleteObject hObj
                GdipImageRotateFlip newImage, (flipFlag Xor 6&)
                pvCreateSourcelessImage2 = newImage
            End If
        End If
    End If
    GdipDisposeImage hImage
    
End Function
Private Function pvCreateSourcelessImage(hImage As Long, width As Long, height As Long, flipFlag As Long) As Long

    ' function creates a stand-alone GDI+ image from a linked image
    ' The routine here follows the instructions given at this MS KB article link
    '       http://support.microsoft.com/kb/814675
    Dim tSize As RECTF, tSizeI As RECTI, BHI As BITMAPINFO
    Dim tBMPsrc As BitmapData
    Dim cRender As cGDIpRenderer, pal() As Byte
    Dim oldImage As Long, newImage As Long, hObj As Long
    Dim dDC As Long, Depth As Long, hGraphics As Long, dibPtr As Long
    
    If width = 0& Or height = 0& Then               ' get size of image
        Call GdipGetImageBounds(hImage, tSize, UnitPixel)
        width = tSize.nWidth: height = tSize.nHeight
    End If
    Call GdipGetImagePixelFormat(hImage, Depth)
    Select Case Depth
        Case PixelFormat1bppIndexed, PixelFormat4bppIndexed, PixelFormat8bppIndexed
            With BHI.bmiHeader
                If Depth = PixelFormat1bppIndexed Then
                    .biBitCount = 1
                ElseIf Depth = PixelFormat4bppIndexed Then
                    .biBitCount = 4
                Else
                    .biBitCount = 8
                End If
                Call GdipGetImagePaletteSize(hImage, .biClrImportant)
                ReDim pal(0 To .biClrImportant + 7&)
                GdipGetImagePalette hImage, pal(0), .biClrImportant
                CopyMemory BHI.bmiColors(0), pal(8), .biClrImportant
                .biClrImportant = .biClrImportant \ 4&
                .biClrUsed = 2 ^ .biBitCount
                Erase pal()
            End With
        Case PixelFormat24bppRGB, PixelFormat16bppGrayScale, PixelFormat16bppRGB555, PixelFormat16bppRGB565, PixelFormat48bppRGB
            BHI.bmiHeader.biBitCount = 24
                
        Case Else
            Set cRender = New cGDIpRenderer
            cRender.AttachTokenClass m_Token
            oldImage = m_Image(m_SourceIndex)                   ' using the cGDIpRenderer class
            m_Image(m_SourceIndex) = hImage                     ' which requires this class to reflect correct image handle
            hGraphics = cRender.CreateGraphicsFromImageClass(Me) ' create a generic graphics object
            If hGraphics Then                                   ' then create new GDI+ bitmap
                If GdipCreateBitmapFromGraphics(width, height, hGraphics, newImage) = 0& Then
                    cRender.DestroyHGraphics hGraphics
                    Call GdipGetImageBounds(hImage, tSize, UnitPixel)
                    m_Image(m_SourceIndex) = newImage           ' create graphics object around new bitmap
                    hGraphics = cRender.CreateGraphicsFromImageClass(Me)
                    If hGraphics Then                           ' draw the passed bitmap onto the new bitmap & clean up
                        Call cRender.RenderToHGraphics(hImage, hGraphics, 0&, 0&, width, height, tSize.nLeft, tSize.nTop, tSize.nWidth, tSize.nHeight)
                        cRender.DestroyHGraphics hGraphics
                        If flipFlag Then GdipImageRotateFlip newImage, flipFlag
                        pvCreateSourcelessImage = newImage
                    End If
                Else
                    cRender.DestroyHGraphics hGraphics
                End If
            End If
            m_Image(m_SourceIndex) = oldImage
    End Select
    
    If BHI.bmiHeader.biBitCount Then                            ' handle paletted & 24bit bitmaps here
        tSizeI.nHeight = height: tSizeI.nWidth = width
        With BHI.bmiHeader
            .biHeight = height
            .biPlanes = 1
            .biSize = 40
            .biWidth = width
            If .biBitCount = 24 Then
                If GdipBitmapLockBits(hImage, tSizeI, ImageLockModeRead, PixelFormat24bppRGB, tBMPsrc) Then .biBitCount = 0& 'flag meaning failure
            Else
                If GdipBitmapLockBits(hImage, tSizeI, ImageLockModeRead, Depth, tBMPsrc) Then .biBitCount = 0&
            End If
        End With
        If BHI.bmiHeader.biBitCount Then
            dDC = GetDC(GetDesktopWindow())
            hObj = CreateDIBSection(dDC, BHI, 0&, dibPtr, 0&, 0&)
            ReleaseDC GetDesktopWindow(), dDC
            If hObj Then
                CopyMemory ByVal dibPtr, ByVal tBMPsrc.Scan0Ptr, tBMPsrc.Stride * height
                GdipBitmapUnlockBits newImage, tBMPsrc
                GdipCreateBitmapFromHBITMAP hObj, 0&, newImage
                DeleteObject hObj
                GdipImageRotateFlip newImage, (flipFlag Xor 6&)
                pvCreateSourcelessImage = newImage
            End If
        End If
    End If
    GdipDisposeImage hImage
    
End Function

'   -.-.-.-.-.-.-.-.-.-.-.-.-.-.-
'   pvGetEncdoerClsID Routine
'   -.-.-.-.-.-.-.-.-.-.-.-.-.-.-
' Routine is a helper function for the SavePicture routine
Private Function pvGetEncoderClsID(strMimeType As String, ClassID() As Long) As Long
  
  Dim Num      As Long
  Dim SizeA    As Long
  Dim SizeB    As Long
  Dim LIdx     As Long
  Dim ICI()    As ImageCodecInfo
  Dim buffer() As Byte, sMime As String
    
    '-- Get the encoder array size
    Call GdipGetImageEncodersSize(Num, SizeA)
    If (SizeA = 0&) Then Exit Function ' Failed!
    
    '-- Allocate room for the arrays dynamically
    ReDim ICI(1 To Num) As ImageCodecInfo
    ReDim buffer(1 To SizeA) As Byte
    
    '-- Get the array and string data
    Call GdipGetImageEncoders(Num, SizeA, buffer(1))
    '-- Copy the class headers
    Call CopyMemory(ICI(1), buffer(1), (Len(ICI(1)) * Num))
    
    SizeA = Len(strMimeType)
    sMime = String$(SizeA, vbNullChar)
    '-- Loop through all the codecs
    For LIdx = Num To 1& Step -1&
        '-- Must convert the pointer into a usable string
        With ICI(LIdx)
            SizeB = lstrlenW(ByVal .MimeType)
            If SizeA = SizeB Then
                Call CopyMemory(ByVal StrPtr(sMime), ByVal .MimeType, SizeA * 2&)
                If sMime = strMimeType Then
                    CopyMemory ClassID(0), .ClassID(0), 16&
                    Exit For
                End If
            End If
        End With
    Next LIdx
    pvGetEncoderClsID = LIdx
End Function

Private Function pvGetImageType(hImage As Long) As ImageTypeConstants
    
    ' http://com.it-berater.org/gdiplus/noframes/GdiPlus_constants.htm
    Const ImageFormatBMP As String = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
    Const ImageFormatPNG As String = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
    Dim GUID(0 To 3) As Long, sGUID As String, lRet As Long
    
    If GdipGetImageRawFormat(hImage, VarPtr(GUID(0))) = 0& Then
        sGUID = String$(40, vbNullChar)
        lRet = StringFromGUID2(VarPtr(GUID(0)), StrPtr(sGUID), 40&)
        Select Case Left$(sGUID, lRet - 1&)
            Case ImageFormatPNG: pvGetImageType = imagePNG
            Case ImageFormatBMP: pvGetImageType = imageBitmap
            Case Else: pvGetImageType = imageUnknown
        End Select
    Else
        pvGetImageType = imageUnknown
    End If

End Function

Private Function pvProcessAlphaBitmap(inArray() As Byte, width As Long, height As Long, resetFlag As pvCleanUpEnum) As Boolean

    ' This is a non-foolproof method of determining whether image uses ARGB or pARGB alpha channel

    Dim x As Long, y As Long, scanWidth As Long, hDib As Long, dibPtr As Long, dDC As Long
    Dim B As Long, lColorFormat As Long, hImage As Long, hdrSize As Long
    Dim nullAlphas As Long, fullAlphas As Long
    
    If width = 0& And height = 0& Then      ' called from LoadPicture_Stream
        CopyMemory B, inArray(0), 2&        ' magic number
        If B <> bmp_Signature Then Exit Function  ' not a bitmap
        CopyMemory x, inArray(28), 2&       ' bit depth
        If x <> 32 Then Exit Function       '  not processed here
        CopyMemory B, inArray(26), 2&       ' validate planes (must be 1)
        If B <> 1& Then Exit Function
        CopyMemory width, inArray(18), 4&
        CopyMemory height, inArray(22), 4&  ' validate sizes
        If width < 1 Or height < 1 Then Exit Function
        CopyMemory hdrSize, inArray(14), 4&
        ' validate image size compared file size
        B = width * height * 4 + hdrSize + 13
        If B > UBound(inArray) Then Exit Function
        hdrSize = hdrSize + 14&
    Else
        dibPtr = VarPtr(inArray(0))         ' called from LoadPicture_DIBhandle
    End If
    
    scanWidth = width * 4&                  ' bytes per row
    For y = 0& To height - 1                ' loop thru image, abort early if possible
        x = hdrSize + y * scanWidth
        For x = x + 3& To x + scanWidth - 1& Step 4&
            If inArray(x) = 0 Then
                nullAlphas = nullAlphas + 1& ' track zero alpha bytes
            ElseIf inArray(x) = 255 Then
                fullAlphas = fullAlphas + 1& ' track fully opaque bytes
            Else
                For B = x - 3& To x - 1&
                    If inArray(B) > inArray(x) Then
                        ' pARGB can never have a R,G,B value > alpha value so this must be ARGB
                        lColorFormat = PixelFormat32bppARGB ' ARGB format; done checking
                        x = x + scanWidth   ' exit X loop
                        y = height          ' exit Y loop
                        Exit For
                    End If
                Next
            End If
        Next
    Next
    ' decision time
    B = width * height ' total number of bytes/pixels
    ' if the image is not 100% zero alpha or 100% full alpha then it uses the alpha channel
    If ((nullAlphas = B) Or (fullAlphas = B)) Then  ' 32bpp with no alpha in play
        lColorFormat = PixelFormat32bppRGB
    Else
        ' if we didn't detect ARGB formatting then assume pARGB formatting
        ' This logic isn't 100% foolproof but exceptions to this logic would be extremely rare
        If lColorFormat = 0& Then lColorFormat = PixelFormat32bppPARGB
    End If
    ' because GDI+ won't load alphachannel bitmaps we will create a DIB and have GDI+ load from it
    If dibPtr = 0& Then
        dDC = GetDC(GetDesktopWindow())
        hDib = CreateDIBSection(dDC, inArray(14), 0&, dibPtr, 0&, 0&)
        ReleaseDC GetDesktopWindow(), dDC
        If hDib Then
            CopyMemory ByVal dibPtr, inArray(hdrSize), scanWidth * height
        Else
            Exit Function
        End If
    End If
    If lColorFormat = PixelFormat32bppRGB Then
        GdipCreateBitmapFromHBITMAP hDib, 0&, hImage
    Else
        Call GdipCreateBitmapFromScan0(width, height, scanWidth, lColorFormat, ByVal dibPtr, hImage)
        If hImage Then hImage = pvCreateSourcelessImage(hImage, width, height, 0&)
    End If
    If hDib Then DeleteObject hDib
    If hImage Then
        Call pvCleanUp(resetFlag)
        m_Image(m_SourceIndex) = hImage
        Call GdipGetImageBounds(hImage, m_Size, UnitPixel)
        pvProcessAlphaBitmap = True
    End If

End Function

Private Function pvStreamFromArray(ArrayPtr As Long, Length As Long) As stdole.IUnknown
    
    ' Purpose: Create an IStream-compatible IUnknown interface containing the
    ' passed byte aray. This IUnknown interface can be passed to GDI+ functions
    ' that expect an IStream interface -- neat hack
    
    On Error GoTo HandleError
    Dim o_hMem As Long
    Dim o_lpMem  As Long
     
    If ArrayPtr = 0& Then
        CreateStreamOnHGlobal 0&, 1&, pvStreamFromArray
    ElseIf Length <> 0& Then
        o_hMem = GlobalAlloc(&H2&, Length)
        If o_hMem <> 0 Then
            o_lpMem = GlobalLock(o_hMem)
            If o_lpMem <> 0 Then
                CopyMemory ByVal o_lpMem, ByVal ArrayPtr, Length
                Call GlobalUnlock(o_hMem)
                Call CreateStreamOnHGlobal(o_hMem, 1&, pvStreamFromArray)
            End If
        End If
    End If
    
HandleError:
End Function

Private Function pvStreamToArray(hStream As Long, arrayBytes() As Byte) As Boolean

    ' Return the array contained in an IUnknown interface (stream)
    
    Dim o_hMem As Long, o_lpMem As Long
    Dim o_lngByteCount As Long
    
    If hStream Then
        If GetHGlobalFromStream(ByVal hStream, o_hMem) = 0 Then
            o_lngByteCount = GlobalSize(o_hMem)
            If o_lngByteCount > 0 Then
                o_lpMem = GlobalLock(o_hMem)
                If o_lpMem <> 0 Then
                    ReDim arrayBytes(0 To o_lngByteCount - 1)
                    CopyMemory arrayBytes(0), ByVal o_lpMem, o_lngByteCount
                    GlobalUnlock o_hMem
                    pvStreamToArray = True
                End If
            End If
        End If
    End If
    
End Function

Private Sub Class_Terminate()
    Call pvCleanUp(cuDestroyAll)
    If Not m_Token Is Nothing Then
        m_Token.RemoveUser Me
        Set m_Token = Nothing
    End If
End Sub
