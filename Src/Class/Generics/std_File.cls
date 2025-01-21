VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private ErrorCatalog(1, 99) As Variant
Private FSO As New FileSystemObject
Private p_Handler As New std_Error

Public Enum std_FileDrive
    Availablespace = 0
    Driveletter = 1
    Drivetype = 2
    Filesystem = 3
    Freespace = 4
    Isready = 5
    Path = 6
    Rootfolder = 7
    Serialnumber = 8
    Sharename = 9
    Totalsize = 10
    Volumename = 11
End Enum

Public Enum std_FileFolder
    Attributes = 0
    Datecreated = 1
    Datelastaccessed = 2
    Datelastmodified = 3
    Drive = 4
    Files = 5
    Isrootfolder = 6
    Name = 7
    Parentfolder = 8
    Path = 9
    Shortname = 10
    Shortpath = 11
    Size = 12
    Subfolders = 13
    Typee = 14
End Enum

Public Enum std_FileFile
    Attributes = 0
    Datecreated = 1
    Datelastaccessed = 2
    Datelastmodified = 3
    Drive = 4
    Name = 5
    Openastextstream = 6
    Parentfolder = 7
    Path = 8
    Shortname = 9
    Shortpath = 10
    Size = 11
    Typee = 12
    Atendofline = 13
    Atendofstream = 14
    Column = 15
    Line = 16
    Read = 17
    Readall = 18
    Readline = 19
End Enum

Public Enum std_FileType
    All = 0
    Drive = 1
    Folder = 2
    File = 3
End Enum

' Public Procedures

    Public Property Let Handler(n_Handler As std_Error)
        Set p_Handler = n_Handler
    End Property

    Public Property Get IS_ERROR() As Boolean
        IS_ERROR = p_Handler.IS_ERROR
    End Property

    Public Function Create(Optional n_Handler As std_Error = Nothing) As std_File
        Set Create = New std_File
        If n_Handler Is Nothing Then
            Create.Handler = std_Error
        Else
            Create.Handler = n_Handler
        End If
    End Function

    Public Function GetDriveValue(SearchValue As Variant, SearchProperty As std_FileDrive, ReturnProperty As String, Optional ThrowError As Boolean = True) As Variant

        Dim Drive As Scripting.Drive
        Dim Found As Boolean
        Dim ErrorCode As Long
        For Each Drive In FSO.Drives
            Select Case SearchProperty
                Case std_FileDrive.Availablespace: If SearchValue = Drive.AvailableSpace Then Found = True: Exit For
                Case std_FileDrive.Driveletter   : If SearchValue = Drive.DriveLetter    Then Found = True: Exit For
                Case std_FileDrive.Drivetype     : If SearchValue = Drive.DriveType      Then Found = True: Exit For
                Case std_FileDrive.Filesystem    : If SearchValue = Drive.FileSystem     Then Found = True: Exit For
                Case std_FileDrive.Freespace     : If SearchValue = Drive.FreeSpace      Then Found = True: Exit For
                Case std_FileDrive.Isready       : If SearchValue = Drive.IsReady        Then Found = True: Exit For
                Case std_FileDrive.Path          : If SearchValue = Drive.Path           Then Found = True: Exit For
                Case std_FileDrive.Rootfolder    : If SearchValue = Drive.RootFolder     Then Found = True: Exit For
                Case std_FileDrive.Serialnumber  : If SearchValue = Drive.SerialNumber   Then Found = True: Exit For
                Case std_FileDrive.Sharename     : If SearchValue = Drive.ShareName      Then Found = True: Exit For
                Case std_FileDrive.Totalsize     : If SearchValue = Drive.TotalSize      Then Found = True: Exit For
                Case std_FileDrive.Volumename    : If SearchValue = Drive.VolumeName     Then Found = True: Exit For
                Case Else
                    ErrorCode = 4: GoTo Error
            End Select
        Next Drive
        If Found = True Then
            Select Case UCase(ReturnProperty)
                Case std_FileDrive.Availablespace: GetDriveValue = Drive.AvailableSpace
                Case std_FileDrive.Driveletter   : GetDriveValue = Drive.DriveLetter
                Case std_FileDrive.Drivetype     : GetDriveValue = Drive.DriveType
                Case std_FileDrive.Filesystem    : GetDriveValue = Drive.FileSystem
                Case std_FileDrive.Freespace     : GetDriveValue = Drive.FreeSpace
                Case std_FileDrive.Isready       : GetDriveValue = Drive.IsReady
                Case std_FileDrive.Path          : GetDriveValue = Drive.Path
                Case std_FileDrive.Rootfolder    : GetDriveValue = Drive.RootFolder
                Case std_FileDrive.Serialnumber  : GetDriveValue = Drive.SerialNumber
                Case std_FileDrive.Sharename     : GetDriveValue = Drive.ShareName
                Case std_FileDrive.Totalsize     : GetDriveValue = Drive.TotalSize
                Case std_FileDrive.Volumename    : GetDriveValue = Drive.VolumeName
                Case Else
                    ErrorCode = 4: GoTo Error
            End Select
            Else
                ErrorCode = 2: GoTo Error
        End If
        Exit Function

        Error:
        GetDriveValue = p_Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, SearchValue, SearchProperty, ReturnProperty)

    End Function

    Public Function GetFolderValue(FolderPath As String, ReturnProperty As std_FileFolder, Optional ThrowError As Boolean = True) As Variant

        Dim Folder As Scripting.Folder
        Dim ErrorCode As Long
        On Error GoTo Error
        ErrorCode = 9
        Set Folder = FSO.GetFolder(FolderPath)
        Select Case ReturnProperty
            Case std_FileFolder.Attributes       : GetFolderValue = Folder.Attributes
            Case std_FileFolder.Datecreated      : GetFolderValue = Folder.DateCreated
            Case std_FileFolder.Datelastaccessed : GetFolderValue = Folder.DateLastAccessed
            Case std_FileFolder.Datelastmodified : GetFolderValue = Folder.DateLastModified
            Case std_FileFolder.Drive            : GetFolderValue = Folder.Drive
            Case std_FileFolder.Files            : GetFolderValue = Folder.Files
            Case std_FileFolder.Isrootfolder     : GetFolderValue = Folder.IsRootFolder
            Case std_FileFolder.Name             : GetFolderValue = Folder.Name
            Case std_FileFolder.Parentfolder     : GetFolderValue = Folder.ParentFolder
            Case std_FileFolder.Path             : GetFolderValue = Folder.Path
            Case std_FileFolder.Shortname        : GetFolderValue = Folder.ShortName
            Case std_FileFolder.Shortpath        : GetFolderValue = Folder.ShortPath
            Case std_FileFolder.Size             : GetFolderValue = Folder.Size
            Case std_FileFolder.Subfolders       : GetFolderValue = Folder.SubFolders
            Case std_FileFolder.Typee             : GetFolderValue = Folder.Type
            Case Else
                ErrorCode = 11: GoTo Error
        End Select
        Exit Function

        Error:
        GetFolderValue = p_Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, FolderPath, ReturnProperty)
            
    End Function

    Public Function GetFileValue(FilePath As String, ReturnProperty As std_FileFile, Optional ReturnCharacters As Long = 0, Optional ThrowError As Boolean = True) As Variant

        Dim File As Scripting.File
        Dim ErrorCode As Long
        On Error GoTo Error
        ErrorCode = 16
        Set File = FSO.GetFile(FilePath)
        Select Case ReturnProperty
            Case std_FileFile.Attributes       : GetFileValue = File.Attributes
            Case std_FileFile.Datecreated      : GetFileValue = File.DateCreated
            Case std_FileFile.Datelastaccessed : GetFileValue = File.DateLastAccessed
            Case std_FileFile.Datelastmodified : GetFileValue = File.DateLastModified
            Case std_FileFile.Drive            : GetFileValue = File.Drive
            Case std_FileFile.Name             : GetFileValue = File.Name
            Case std_FileFile.Openastextstream : GetFileValue = File.OpenAsTextStream
            Case std_FileFile.Parentfolder     : GetFileValue = File.ParentFolder
            Case std_FileFile.Path             : GetFileValue = File.Path
            Case std_FileFile.Shortname        : GetFileValue = File.ShortName
            Case std_FileFile.Shortpath        : GetFileValue = File.ShortPath
            Case std_FileFile.Size             : GetFileValue = File.Size
            Case std_FileFile.Typee            : GetFileValue = File.Type
            Case std_FileFile.Atendofline      : GetFileValue = File.OpenAsTextStream().AtEndOfLine
            Case std_FileFile.Atendofstream    : GetFileValue = File.OpenAsTextStream().AtEndOfStream
            Case std_FileFile.Column           : GetFileValue = File.OpenAsTextStream().Column
            Case std_FileFile.Line             : GetFileValue = File.OpenAsTextStream().Line
            Case std_FileFile.Read             : GetFileValue = File.OpenAsTextStream().Read(ReturnCharacters)
            Case std_FileFile.Readall          : GetFileValue = File.OpenAsTextStream().ReadAll
            Case std_FileFile.Readline         : GetFileValue = File.OpenAsTextStream().ReadLine()
            Case Else
                ErrorCode = 18: GoTo Error
        End Select
        Exit Function

        Error:
        GetFileValue = p_Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, FilePath, ReturnProperty)
            
    End Function

    Public Sub GetAllFileValues(FolderPath As String, ReturnProperty As std_FileFile, ByRef ReturnArray() As Variant, Optional ReturnCharacters As Long = 0)
        
        Dim Folder As Scripting.Folder
        Dim SubFolder As Scripting.Folder
        Dim File As Scripting.File

        Set Folder = FSO.GetFolder(FolderPath)
        For Each File in Folder.Files
            If ReturnArray(UBound(ReturnArray)) <> Empty Then ReDim Preserve ReturnArray(Ubound(ReturnArray) + 1)
            ReturnArray(UBound(ReturnArray)) = GetFileValue(File.Path, ReturnProperty, ReturnCharacters)
        Next File
        For Each SubFolder In Folder.SubFolders
            Call GetAllFileValues(SubFolder.Path, ReturnProperty, ReturnArray, ReturnCharacters)
        Next SubFolder

    End Sub

    Public Function Exists(Path As String, Optional ExistsAs As std_FileType = std_FileType.All, Optional ShouldExist As Boolean = True, Optional ThrowError As Boolean = True) As Boolean

        Dim Returns As Boolean
        Dim ErrorCode As Long

        Select Case ExistsAs
            Case std_FileType.All
                ErrorCode = 1
                If Returns = False Then Returns = FSO.DriveExists(Path)
                ErrorCode = 8
                If Returns = False Then Returns = FSO.FolderExists(Path)
                ErrorCode = 15
                If Returns = False Then Returns = FSO.FileExists(Path)
            Case std_FileType.Drive
                ErrorCode = 1
                If Returns = False Then Returns = FSO.DriveExists(Path)
            Case std_FileType.Folder
                ErrorCode = 8
                If Returns = False Then Returns = FSO.FolderExists(Path)
            Case std_FileType.File
                ErrorCode = 15
                If Returns = False Then Returns = FSO.FileExists(Path)
            Case Else
                ErrorCode = 29: GoTo Error
        End Select
        If ShouldExist Then
            If Returns = False Then ErrorCode = ErrorCode + 1: GoTo Error
        Else
            If Returns Then
                If ThrowError Then GoTo Error
            End If
        End If
        Exit Function

        Error:
        Exists = p_Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, Path, ExistsAs, ShouldExist)

    End Function

    Public Function CreateFolderObj(Path As String, Name As String, Optional Unique As Boolean = True, Optional ThrowError As Boolean = True) As Scripting.Folder

        Dim Found As Boolean
        Dim Number As Long
        On Error GoTo Error
        Number = 0
        Repeat:
        Found = False
        Found = FSO.FolderExists(Path & "\" & Name)
        If Found Then
            If Unique Then
                GoTo Error
            Else
                Dim Position As Long
                Position = InStr(1, Name, ".") - 1
                If Position = -1 Then Position = Len(Name)
                Name = Mid(Name, 1, Position) + CStr(Number)
                Number = Number + 1
                GoTo Repeat
            End If
        Else
            MkDir Path & "\" & Name
            Set CreateFolderObj = FSO.GetFolder(Path & "\" & Name)
        End If
        Exit Function

        Error:
        If ThrowError Then
            Call p_Handler.Handle(ErrorCatalog, 8, ThrowError, Path, Name, Unique, ThrowError)
        End If
        Set CreateFolderObj = Nothing

    End Function

    Public Function CreateFolder(Path As String, Name As String, Optional Unique As Boolean = True, Optional ThrowError As Boolean = True) As Boolean
        Dim Temp As Scripting.Folder
        Set Temp = CreateFolderObj(Path, Name, Unique, ThrowError)
        If Temp Is Nothing Then CreateFolder = p_Handler.IS_ERROR
    End Function

    Public Function CreateFileObj(FolderPath As String, FileName As String, Optional Unique As Boolean = True, Optional ThrowError As Boolean = True) As Scripting.File

        Dim Found As Boolean
        Dim Number As Long
        Dim Stream As TextStream
        On Error GoTo Error
        Number = 0
        Repeat:
        Found = False
        Found = FSO.FileExists(FolderPath & "\" & FileName)
        If Found Then
            If Unique Then
                GoTo Error
            Else
                Dim Position As Long
                Position = InStr(1, FileName, ".") - 1
                If Position = -1 Then Position = Len(FileName)
                FileName = Mid(FileName, 1, Position) + CStr(Number) & Mid(FileName, Position + 1, Len(FileName) - Position - 1)
                Number = Number + 1
                GoTo Repeat
            End If
        Else
            Set Stream = FSO.CreateTextFile(FolderPath & "\" & FileName)
            Set CreateFileObj = FSO.GetFile(FolderPath & "\" & FileName)
        End If
        Exit Function

        Error:
        If ThrowError Then
            Call p_Handler.Handle(ErrorCatalog, 15, FolderPath, ThrowError, FileName, Unique, ThrowError)
        End If
        Set CreateFileObj = Nothing

    End Function

    Public Function CreateFile(FolderPath As String, FileName As String, Optional Unique As Boolean = True, Optional ThrowError As Boolean = True) As Boolean
        Dim Temp As Scripting.File
        Set Temp = CreateFileObj(FolderPath, FileName, Unique, ThrowError)
        If Temp Is Nothing Then CreateFile = p_Handler.IS_ERROR
    End Function
'


' Private Procedures
    Private Sub Class_Initialize()
        ProtInit
    End Sub 

    Private Sub Class_Terminate()
    End Sub


    Private Sub ProtInit()
        ErrorCatalog(0, 0000) = 0002: ErrorCatalog(1, 0000) = "std_File"
        ErrorCatalog(0, 0001) = 1000: ErrorCatalog(1, 0001) = "Drive exists"
        ErrorCatalog(0, 0002) = 1000: ErrorCatalog(1, 0002) = "Drive doesnt exist"
        ErrorCatalog(0, 0003) = 1000: ErrorCatalog(1, 0003) = "Drive Property exists"
        ErrorCatalog(0, 0004) = 1000: ErrorCatalog(1, 0004) = "Drive Property doesnt exist"
        ErrorCatalog(0, 0005) = 1000: ErrorCatalog(1, 0005) = "Drive Property equals"
        ErrorCatalog(0, 0006) = 1000: ErrorCatalog(1, 0006) = "Drive Property does not equal"
        ErrorCatalog(0, 0007) = 1000: ErrorCatalog(1, 0007) = "Could not change Drive Properties"
        ErrorCatalog(0, 0008) = 1000: ErrorCatalog(1, 0008) = "Folder exists"
        ErrorCatalog(0, 0009) = 1000: ErrorCatalog(1, 0009) = "Folder doesnt exist"
        ErrorCatalog(0, 0010) = 1000: ErrorCatalog(1, 0010) = "Folder Property exists"
        ErrorCatalog(0, 0011) = 1000: ErrorCatalog(1, 0011) = "Folder Property doesnt exist"
        ErrorCatalog(0, 0012) = 1000: ErrorCatalog(1, 0012) = "Folder Property equals"
        ErrorCatalog(0, 0013) = 1000: ErrorCatalog(1, 0013) = "Folder Property does not equal"
        ErrorCatalog(0, 0014) = 1000: ErrorCatalog(1, 0014) = "Could not change Folder Properties"
        ErrorCatalog(0, 0015) = 1000: ErrorCatalog(1, 0015) = "File exists"
        ErrorCatalog(0, 0016) = 1000: ErrorCatalog(1, 0016) = "File doesnt exist"
        ErrorCatalog(0, 0017) = 1000: ErrorCatalog(1, 0017) = "File Property exists"
        ErrorCatalog(0, 0018) = 1000: ErrorCatalog(1, 0018) = "File Property doesnt exist"
        ErrorCatalog(0, 0019) = 1000: ErrorCatalog(1, 0019) = "File Property equals"
        ErrorCatalog(0, 0020) = 1000: ErrorCatalog(1, 0020) = "File Property does not equal"
        ErrorCatalog(0, 0021) = 1000: ErrorCatalog(1, 0021) = "Could not change File Properties"
        ErrorCatalog(0, 0022) = 1000: ErrorCatalog(1, 0022) = "Textstram exists"
        ErrorCatalog(0, 0023) = 1000: ErrorCatalog(1, 0023) = "Textstram doesnt exist"
        ErrorCatalog(0, 0024) = 1000: ErrorCatalog(1, 0024) = "Textstram Property exists"
        ErrorCatalog(0, 0025) = 1000: ErrorCatalog(1, 0025) = "Textstram Property doesnt exist"
        ErrorCatalog(0, 0026) = 1000: ErrorCatalog(1, 0026) = "Textstram Property equals"
        ErrorCatalog(0, 0027) = 1000: ErrorCatalog(1, 0027) = "Textstram Property does not equal"
        ErrorCatalog(0, 0028) = 1000: ErrorCatalog(1, 0028) = "Could not change Textstram Properties"
        ErrorCatalog(0, 0029) = 1000: ErrorCatalog(1, 0029) = "value invalid"
        ErrorCatalog(0, 0030) = 1000: ErrorCatalog(1, 0030) = "PLACEHOLDER"
        ErrorCatalog(0, 0031) = 1000: ErrorCatalog(1, 0031) = "PLACEHOLDER"
        ErrorCatalog(0, 0032) = 1000: ErrorCatalog(1, 0032) = "PLACEHOLDER"
        ErrorCatalog(0, 0033) = 1000: ErrorCatalog(1, 0033) = "PLACEHOLDER"
        ErrorCatalog(0, 0034) = 1000: ErrorCatalog(1, 0034) = "PLACEHOLDER"
        ErrorCatalog(0, 0035) = 1000: ErrorCatalog(1, 0035) = "PLACEHOLDER"
        ErrorCatalog(0, 0036) = 1000: ErrorCatalog(1, 0036) = "PLACEHOLDER"
        ErrorCatalog(0, 0037) = 1000: ErrorCatalog(1, 0037) = "PLACEHOLDER"
        ErrorCatalog(0, 0038) = 1000: ErrorCatalog(1, 0038) = "PLACEHOLDER"
        ErrorCatalog(0, 0039) = 1000: ErrorCatalog(1, 0039) = "PLACEHOLDER"
        ErrorCatalog(0, 0040) = 1000: ErrorCatalog(1, 0040) = "PLACEHOLDER"
        ErrorCatalog(0, 0041) = 1000: ErrorCatalog(1, 0041) = "PLACEHOLDER"
        ErrorCatalog(0, 0042) = 1000: ErrorCatalog(1, 0042) = "PLACEHOLDER"
        ErrorCatalog(0, 0043) = 1000: ErrorCatalog(1, 0043) = "PLACEHOLDER"
        ErrorCatalog(0, 0044) = 1000: ErrorCatalog(1, 0044) = "PLACEHOLDER"
        ErrorCatalog(0, 0045) = 1000: ErrorCatalog(1, 0045) = "PLACEHOLDER"
        ErrorCatalog(0, 0046) = 1000: ErrorCatalog(1, 0046) = "PLACEHOLDER"
        ErrorCatalog(0, 0047) = 1000: ErrorCatalog(1, 0047) = "PLACEHOLDER"
        ErrorCatalog(0, 0048) = 1000: ErrorCatalog(1, 0048) = "PLACEHOLDER"
        ErrorCatalog(0, 0049) = 1000: ErrorCatalog(1, 0049) = "PLACEHOLDER"
        ErrorCatalog(0, 0050) = 1000: ErrorCatalog(1, 0050) = "PLACEHOLDER"
        ErrorCatalog(0, 0051) = 1000: ErrorCatalog(1, 0051) = "PLACEHOLDER"
        ErrorCatalog(0, 0052) = 1000: ErrorCatalog(1, 0052) = "PLACEHOLDER"
        ErrorCatalog(0, 0053) = 1000: ErrorCatalog(1, 0053) = "PLACEHOLDER"
        ErrorCatalog(0, 0054) = 1000: ErrorCatalog(1, 0054) = "PLACEHOLDER"
        ErrorCatalog(0, 0055) = 1000: ErrorCatalog(1, 0055) = "PLACEHOLDER"
        ErrorCatalog(0, 0056) = 1000: ErrorCatalog(1, 0056) = "PLACEHOLDER"
        ErrorCatalog(0, 0057) = 1000: ErrorCatalog(1, 0057) = "PLACEHOLDER"
        ErrorCatalog(0, 0058) = 1000: ErrorCatalog(1, 0058) = "PLACEHOLDER"
        ErrorCatalog(0, 0059) = 1000: ErrorCatalog(1, 0059) = "PLACEHOLDER"
        ErrorCatalog(0, 0060) = 1000: ErrorCatalog(1, 0060) = "PLACEHOLDER"
        ErrorCatalog(0, 0061) = 1000: ErrorCatalog(1, 0061) = "PLACEHOLDER"
        ErrorCatalog(0, 0062) = 1000: ErrorCatalog(1, 0062) = "PLACEHOLDER"
        ErrorCatalog(0, 0063) = 1000: ErrorCatalog(1, 0063) = "PLACEHOLDER"
        ErrorCatalog(0, 0064) = 1000: ErrorCatalog(1, 0064) = "PLACEHOLDER"
        ErrorCatalog(0, 0065) = 1000: ErrorCatalog(1, 0065) = "PLACEHOLDER"
        ErrorCatalog(0, 0066) = 1000: ErrorCatalog(1, 0066) = "PLACEHOLDER"
        ErrorCatalog(0, 0067) = 1000: ErrorCatalog(1, 0067) = "PLACEHOLDER"
        ErrorCatalog(0, 0068) = 1000: ErrorCatalog(1, 0068) = "PLACEHOLDER"
        ErrorCatalog(0, 0069) = 1000: ErrorCatalog(1, 0069) = "PLACEHOLDER"
        ErrorCatalog(0, 0070) = 1000: ErrorCatalog(1, 0070) = "PLACEHOLDER"
        ErrorCatalog(0, 0071) = 1000: ErrorCatalog(1, 0071) = "PLACEHOLDER"
        ErrorCatalog(0, 0072) = 1000: ErrorCatalog(1, 0072) = "PLACEHOLDER"
        ErrorCatalog(0, 0073) = 1000: ErrorCatalog(1, 0073) = "PLACEHOLDER"
        ErrorCatalog(0, 0074) = 1000: ErrorCatalog(1, 0074) = "PLACEHOLDER"
        ErrorCatalog(0, 0075) = 1000: ErrorCatalog(1, 0075) = "PLACEHOLDER"
        ErrorCatalog(0, 0076) = 1000: ErrorCatalog(1, 0076) = "PLACEHOLDER"
        ErrorCatalog(0, 0077) = 1000: ErrorCatalog(1, 0077) = "PLACEHOLDER"
        ErrorCatalog(0, 0078) = 1000: ErrorCatalog(1, 0078) = "PLACEHOLDER"
        ErrorCatalog(0, 0079) = 1000: ErrorCatalog(1, 0079) = "PLACEHOLDER"
        ErrorCatalog(0, 0080) = 1000: ErrorCatalog(1, 0080) = "PLACEHOLDER"
        ErrorCatalog(0, 0081) = 1000: ErrorCatalog(1, 0081) = "PLACEHOLDER"
        ErrorCatalog(0, 0082) = 1000: ErrorCatalog(1, 0082) = "PLACEHOLDER"
        ErrorCatalog(0, 0083) = 1000: ErrorCatalog(1, 0083) = "PLACEHOLDER"
        ErrorCatalog(0, 0084) = 1000: ErrorCatalog(1, 0084) = "PLACEHOLDER"
        ErrorCatalog(0, 0085) = 1000: ErrorCatalog(1, 0085) = "PLACEHOLDER"
        ErrorCatalog(0, 0086) = 1000: ErrorCatalog(1, 0086) = "PLACEHOLDER"
        ErrorCatalog(0, 0087) = 1000: ErrorCatalog(1, 0087) = "PLACEHOLDER"
        ErrorCatalog(0, 0088) = 1000: ErrorCatalog(1, 0088) = "PLACEHOLDER"
        ErrorCatalog(0, 0089) = 1000: ErrorCatalog(1, 0089) = "PLACEHOLDER"
        ErrorCatalog(0, 0090) = 1000: ErrorCatalog(1, 0090) = "PLACEHOLDER"
        ErrorCatalog(0, 0091) = 1000: ErrorCatalog(1, 0091) = "PLACEHOLDER"
        ErrorCatalog(0, 0092) = 1000: ErrorCatalog(1, 0092) = "PLACEHOLDER"
        ErrorCatalog(0, 0093) = 1000: ErrorCatalog(1, 0093) = "PLACEHOLDER"
        ErrorCatalog(0, 0094) = 1000: ErrorCatalog(1, 0094) = "PLACEHOLDER"
        ErrorCatalog(0, 0095) = 1000: ErrorCatalog(1, 0095) = "PLACEHOLDER"
        ErrorCatalog(0, 0096) = 1000: ErrorCatalog(1, 0096) = "PLACEHOLDER"
        ErrorCatalog(0, 0097) = 1000: ErrorCatalog(1, 0097) = "PLACEHOLDER"
        ErrorCatalog(0, 0098) = 1000: ErrorCatalog(1, 0098) = "PLACEHOLDER"
        ErrorCatalog(0, 0099) = 1000: ErrorCatalog(1, 0099) = "PLACEHOLDER"
    End Sub
'