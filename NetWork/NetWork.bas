Attribute VB_Name = "NetWork"

Option Explicit

Private Type tyNETRESOURCE
   dwScope As Long
   dwType As Long
   dwDisplayType As Long
   dwUsage As Long
   lpLocalName As Long
   lpRemoteName As Long
   lpComment As Long
   lpProvider As Long
End Type

'Decl of lib files
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias _
  "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, _
  ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long

Private Declare Function WNetEnumResource Lib "mpr.dll" Alias _
  "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, _
  ByVal lpBuffer As Long, lpBufferSize As Long) As Long

Private Declare Function WNetCloseEnum Lib "mpr.dll" _
   (ByVal hEnum As Long) As Long

Private Const RESOURCE_CONNECTED = &H1
Private Const RESOURCE_GLOBALNET = &H2
Private Const RESOURCE_REMEMBERED = &H3

Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCETYPE_DISK = &H1
Private Const RESOURCETYPE_PRINT = &H2
Private Const RESOURCETYPE_UNKNOWN = &HFFFF

Private Const RESOURCEUSAGE_CONNECTABLE = &H1
Private Const RESOURCEUSAGE_CONTAINER = &H2
Private Const RESOURCEUSAGE_RESERVED = &H80000000

Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Declare Function GlobalAlloc Lib "KERNEL32" _
  (ByVal wFlags As Long, ByVal dwBytes As Long) As Long

Private Declare Function GlobalFree Lib "KERNEL32" _
  (ByVal hMem As Long) As Long

Private Declare Sub CopyMemory Lib "KERNEL32" Alias _
  "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, _
   ByVal cbCopy As Long)
   
Private Declare Function CopyPointer2String Lib _
  "KERNEL32" Alias "lstrcpyA" (ByVal NewString As _
  String, ByVal OldString As Long) As Long

'Shows Network name and all computers on the network in a list box

Public Function ShowNetWorkList(list As Object) As Boolean
    Dim hEnum As Long, lpBuff As Long, NRES As tyNETRESOURCE
    Dim cbBuff As Long, cCount As Long
    Dim lBuf As Long, res As Long, i As Long

    On Error Resume Next
    list.Clear
    If Err.Number > 0 Then Exit Function
    
    On Error GoTo ErrorHandler
    
    NRES.lpRemoteName = 0
    cbBuff = 10000
    cCount = &HFFFFFFFF
    
    res = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, 0, NRES, hEnum)

    If res = 0 Then
       'Create a buffer
       lpBuff = GlobalAlloc(GPTR, cbBuff)
       res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)
        If res = 0 Then
            lBuf = lpBuff
            For i = 1 To cCount
                CopyMemory NRES, ByVal lBuf, LenB(NRES)
                list.AddItem "Network Name " & ConvPointerToString(NRES.lpRemoteName)
                ShowNetWorkList2 NRES, list
                lBuf = lBuf + LenB(NRES)
            Next i
        End If
        ShowNetWorkList = True
ErrorHandler:
       On Error Resume Next
       If lpBuff <> 0 Then GlobalFree (lpBuff)
       WNetCloseEnum (hEnum)
    End If
End Function

Private Function ConvPointerToString(p As Long) As String
   Dim s As String
   s = String(255, Chr$(0))
   CopyPointer2String s, p
   ConvPointerToString = Left(s, InStr(s, Chr$(0)) - 1)
End Function

Private Sub ShowNetWorkList2(NRES As tyNETRESOURCE, list As Object)
   Dim hEnum As Long, lpBuff As Long
   Dim cbBuff As Long, cCount As Long
   Dim p As Long, res As Long, i As Long

   'Setup the tyNETRESOURCE input structure.
   cbBuff = 10000
   cCount = &HFFFFFFFF

   res = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, 0, NRES, hEnum)
   If res = 0 Then
      lpBuff = GlobalAlloc(GPTR, cbBuff)
      res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)
      If res = 0 Then
         p = lpBuff
         For i = 1 To cCount
            CopyMemory NRES, ByVal p, LenB(NRES)
            list.AddItem "Network Computer #" & i & " " & ConvPointerToString(NRES.lpRemoteName)
            p = p + LenB(NRES)
         Next i
      End If
      If lpBuff <> 0 Then GlobalFree (lpBuff)
      WNetCloseEnum (hEnum)  'Close the enumeration
   End If
End Sub
