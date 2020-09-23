Attribute VB_Name = "ProgressInTask"
'Function to Show Progress In TaskBar

'1.- Creates A Valid Picture From a ProgressBar HDC That Later is converted to an Icon using a ImageList
'2.- Shows The new Icon Every Time The User-Selected ProgressBar Is Updated

'Author: Mario Flores G
'E-mail: sistec_de_juarez@hotmail.com

'CD JUAREZ CHIHUAHUA MEXICO


Option Explicit


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lppictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean



'=====================================================
'THE NOTIFYICONDATA STRUCTURE
Private Type NOTIFYICONDATA        'The NOTIFYICONDATA structure Contains information that
    cbSize           As Long       'the system needs to process taskbar status area messages.
    hwnd             As Long
    uId              As Long
    uFlags           As Long
    uCallBackMessage As Long
    hIcon            As Long
    szTip            As String * 64
End Type
'=====================================================

Private nId As NOTIFYICONDATA

Private Const NIM_ADD     As Long = &H0     'Add to Tray
Private Const NIM_MODIFY  As Long = &H1     'Modify Details (Icon Progress)
Private Const NIM_DELETE  As Long = &H2     'Remove From Tray
Private Const NIF_ICON    As Long = &H2     'Icon
Private Const NIF_TIP     As Long = &H4     'TooTipText

    
'=====================================================
'THE ICONINFO STRUCTURE
Private Type ICONINFO
    fIcon           As Long        'The ICONINFO structure contains information about an icon or a cursor.
    xHotspot        As Long
    yHotspot        As Long
    hBMMask         As Long
    hBMColor        As Long
End Type
'=====================================================

'=====================================================
'THE PICTDESC STRUCTURE
Private Type PICTDESC
    Size                As Long    'The PICTDESC structure contains parameters to create a picture
    Type                As Long    'object through the OleCreatePictureIndirect function.
    hBmpOrIcon          As Long
    hPal                As Long
End Type
'=====================================================

Private m_Progress_HDC As Long       ' <<---  Reference The Selected Tray Progress


Private Sub nIdStructure(ByRef nWnd As Form, ByVal ToolTip As String)
   
   With nId 'with system tray
            .cbSize = Len(nId)
            .hwnd = nWnd.hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP
            .hIcon = nWnd.Icon                          'use form's icon in tray (Progress Bar)
            .szTip = ToolTip & vbNullChar               'tooltip text is the Progress Bar Percent
    End With
    
End Sub

Public Sub AddTotray(nWnd As Form, ByVal ToolTip As String)
   
    Call nIdStructure(nWnd, ToolTip)
    Shell_NotifyIcon NIM_ADD, nId 'add to tray

End Sub

Public Sub RefreshTray(ByRef nWnd As Form, ByVal ToolTip As String)
  
    Call nIdStructure(nWnd, ToolTip)
    Shell_NotifyIcon NIM_MODIFY, nId 'Refresh the tray (Gives Progress Animation)

End Sub

Public Sub RemoveTray()
  
  Shell_NotifyIcon NIM_DELETE, nId 'remove from tray

End Sub

Public Property Get Progress_HDC() As Long
   Progress_HDC = m_Progress_HDC
End Property

Public Property Let Progress_HDC(ByVal cProgress_HDC As Long)
    m_Progress_HDC = cProgress_HDC
End Property

Public Sub BuildTheIcon(ByRef Forma As Form, ByVal Width As Long, ByVal Height As Long, ByVal m_ThDC As Long)
      Forma.ImgList.ListImages.Clear
      Forma.ImgList.ListImages.Add , , cImage(Width, Height, m_ThDC)
      Set Forma.Icon = Forma.ImgList.ListImages.Item(1).ExtractIcon
End Sub

Public Function BitmapToPicture( _
            ByVal hBmp As Long, _
            Optional ByVal hPal As Long = 0) As IPicture
'--- Returns a VB picture object containing the specified bitmap.
    Dim oNewPic         As Picture
    Dim lppictDesc      As PICTDESC
    Dim aGuid(0 To 3)   As Long
    
    '--- fill struct
    With lppictDesc
        .Size = Len(lppictDesc)
        .Type = vbPicTypeBitmap
        .hBmpOrIcon = hBmp
        .hPal = hPal
    End With
    '--- Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGuid(0) = &H7BF80980
    aGuid(1) = &H101ABF32
    aGuid(2) = &HAA00BB8B
    aGuid(3) = &HAB0C3000
    '--- Create picture from bitmap handle
    OleCreatePictureIndirect lppictDesc, aGuid(0), True, oNewPic
    '--- success
    Set BitmapToPicture = oNewPic
End Function

'Purpose: Returns a <b>StdPicture</b> objects which contains current image of the <b>cMemDC</b> object.
Public Function cImage(ByVal Width As Long, ByVal Height As Long, ByVal HDC As Long) As StdPicture
    Dim hdcPaint            As Long
    Dim hbmPaint            As Long
    Dim hbmPaintOrig        As Long
    Dim hpalPaintOrig       As Long
    
    '--- prepare
    hdcPaint = CreateCompatibleDC(HDC)
        hbmPaint = CreateCompatibleBitmap(HDC, Width, Height)
        hbmPaintOrig = SelectObject(hdcPaint, hbmPaint)
    '--- bitblit
        Call BitBlt(hdcPaint, 0, 0, Width, Height, HDC, 0, 0, vbMergeCopy)
      
    '--- deselect
    Call SelectObject(hdcPaint, hbmPaintOrig)
    Call DeleteDC(hdcPaint)
    '--- get image
    Set cImage = BitmapToPicture(hbmPaint, 0)
End Function

