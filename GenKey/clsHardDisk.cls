VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHardDisk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias _
"GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer _
As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal _
lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
(ByVal nDrive As String) As Long


Const DRIVE_UNKNOWN = 0
Const DRIVE_REMOVABLE = 2
Const DRIVE_FIXED = 3
Const DRIVE_REMOTE = 4
Const DRIVE_CDROM = 5
Const DRIVE_RAMDISK = 6


Private Sub getSerialHD(ByVal drv As String, volfsname As String, _
voltype As String, parttype As String, snumber As String)
Dim fsflags As Long
Dim serial As Long
Dim maxlen As Long
Dim fsname As String * 255
Dim volname As String * 255
Dim drive As String

If Right(drv, Len(drv)) <> "\" Then drive = drv & "\"
Select Case GetDriveType(drive)
    Case DRIVE_UNKNOWN: voltype = "Undetermined"
    Case DRIVE_REMOVABLE: voltype = "Removable Drive"
    Case DRIVE_FIXED: voltype = "Fixed Drive"
    Case DRIVE_REMOTE: voltype = "Network Drive"
    Case DRIVE_CDROM: voltype = "CD-ROM Drive"
    Case DRIVE_RAMDISK: voltype = "RAM Disk"
    Case Else: voltype = ""
End Select
If GetDriveType(drive) = DRIVE_FIXED Then
    If GetVolumeInformation(drive, volname, 255, serial, maxlen, fsflags, fsname, 255) Then
        volfsname = Left(volname, InStr(volname, Chr(0)) - 1)
        parttype = Left(fsname, InStr(fsname, Chr(0)) - 1)
        Check = Format(Hex(serial), "00000000")
        snumber = Left$(Check, 4) + "-" + Right$(Check, 4)
    Else
        snumber = ""
        volfsname = ""
        parttype = ""
    End If
Else
    snumber = ""
    volfsname = ""
    parttype = ""
End If
End Sub
Public Function getSerialNumber() As String
Dim diskname As String
Dim disk() As String
Dim disknames As String * 255
Dim nDisk As Long
Dim voltype As String
Dim volfsname As String
Dim parttype As String
Dim currentListItem As Long
Dim snumber As String

nDisk = GetLogicalDrives()
errd = GetLogicalDriveStrings(255, disknames)
If errd > 0 Then
    disk = Split(disknames, Chr(0), nDisk)
    currentListItem = 1
    For i = 0 To UBound(disk)
        If Len(disk(i)) > 0 Then
            If Asc(UCase(Left(disk(i), 1))) > Asc("A") Then
                getSerialHD Left(disk(i), 2), volfsname, voltype, parttype, snumber
                If snumber <> "" Then
                'Obtuvo serial
                     getSerialNumber = snumber
                     Exit Function
                End If
            End If
        End If
    Next
End If
    getSerialNumber = "0000-0000"
End Function
