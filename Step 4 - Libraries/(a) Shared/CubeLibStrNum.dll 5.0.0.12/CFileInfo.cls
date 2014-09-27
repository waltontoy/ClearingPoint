VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarFilename As String 'local copy
Private mvarDirectory As String 'local copy
Private mvarFullFilePath As String 'local copy
Private mvarFileType As String 'local copy
Private mvarVerNotUsedVB As Long 'local copy
Private mvarVerMajor As Long 'local copy
Private mvarVerMinor As Long 'local copy
Private mvarVerRevision As Long 'local copy

Private mvarFileSizeKB As Long 'local copy

Private mvarFileAttributes As Long 'local copy
Private mvarFileTypeLong As Long 'local copy
Private mvarCreationDate As Date 'local copy
Private mvarLastModifyTime As Date 'local copy
Private mvarLastAccessTime As Date 'local copy
Private mvarCompanyName As String 'local copy
Private mvarFileDescription As String 'local copy
Private mvarFileVersion As String 'local copy
Private mvarInternalName As String 'local copy
Private mvarLegalCopyright As String 'local copy
Private mvarOriginalFilename As String 'local copy
Private mvarProductName As String 'local copy
Private mvarProductVersion As String 'local copy
'local variable(s) to hold property value(s)
Private mvarIsDirectory As Boolean 'local copy
Private mvarIsArchive As Boolean 'local copy
Private mvarIsCompressed As Boolean 'local copy
Private mvarIsHidden As Boolean 'local copy
Private mvarIsNormal As Boolean 'local copy
Private mvarIsReadOnly As Boolean 'local copy
Private mvarIsSystem As Boolean 'local copy
Private mvarIsTemporary As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarFileSizeDesc As String 'local copy
Public Property Let FileSizeDesc(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileSizeDesc = 5
    mvarFileSizeDesc = vData
End Property


Public Property Get FileSizeDesc() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileSizeDesc
    FileSizeDesc = mvarFileSizeDesc
End Property



Public Property Let IsTemporary(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsTemporary = 5
    mvarIsTemporary = vData
End Property


Public Property Get IsTemporary() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsTemporary
    IsTemporary = mvarIsTemporary
End Property



Public Property Let IsSystem(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsSystem = 5
    mvarIsSystem = vData
End Property


Public Property Get IsSystem() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsSystem
    IsSystem = mvarIsSystem
End Property



Public Property Let IsReadOnly(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsReadOnly = 5
    mvarIsReadOnly = vData
End Property


Public Property Get IsReadOnly() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsReadOnly
    IsReadOnly = mvarIsReadOnly
End Property



Public Property Let IsNormal(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsNormal = 5
    mvarIsNormal = vData
End Property


Public Property Get IsNormal() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsNormal
    IsNormal = mvarIsNormal
End Property



Public Property Let IsHidden(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsHidden = 5
    mvarIsHidden = vData
End Property


Public Property Get IsHidden() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsHidden
    IsHidden = mvarIsHidden
End Property



Public Property Let IsCompressed(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsCompressed = 5
    mvarIsCompressed = vData
End Property


Public Property Get IsCompressed() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsCompressed
    IsCompressed = mvarIsCompressed
End Property



Public Property Let IsArchive(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsArchive = 5
    mvarIsArchive = vData
End Property


Public Property Get IsArchive() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsArchive
    IsArchive = mvarIsArchive
End Property



Public Property Let IsDirectory(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsDirectory = 5
    mvarIsDirectory = vData
End Property


Public Property Get IsDirectory() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsDirectory
    IsDirectory = mvarIsDirectory
End Property



Public Property Let ProductVersion(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProductVersion = 5
    mvarProductVersion = vData
End Property


Public Property Get ProductVersion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductVersion
    ProductVersion = mvarProductVersion
End Property



Public Property Let ProductName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProductName = 5
    mvarProductName = vData
End Property


Public Property Get ProductName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductName
    ProductName = mvarProductName
End Property



Public Property Let OriginalFilename(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OriginalFilename = 5
    mvarOriginalFilename = vData
End Property


Public Property Get OriginalFilename() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OriginalFilename
    OriginalFilename = mvarOriginalFilename
End Property



Public Property Let LegalCopyright(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LegalCopyright = 5
    mvarLegalCopyright = vData
End Property


Public Property Get LegalCopyright() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LegalCopyright
    LegalCopyright = mvarLegalCopyright
End Property



Public Property Let InternalName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.InternalName = 5
    mvarInternalName = vData
End Property


Public Property Get InternalName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.InternalName
    InternalName = mvarInternalName
End Property



Public Property Let FileVersion(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileVersion = 5
    mvarFileVersion = vData
End Property


Public Property Get FileVersion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileVersion
    FileVersion = mvarFileVersion
End Property



Public Property Let FileDescription(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileDescription = 5
    mvarFileDescription = vData
End Property


Public Property Get FileDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileDescription
    FileDescription = mvarFileDescription
End Property



Public Property Let CompanyName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CompanyName = 5
    mvarCompanyName = vData
End Property


Public Property Get CompanyName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CompanyName
    CompanyName = mvarCompanyName
End Property



Public Property Let LastAccessTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastAccessTime = 5
    mvarLastAccessTime = vData
End Property


Public Property Get LastAccessTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastAccessTime
    LastAccessTime = mvarLastAccessTime
End Property



Public Property Let LastModifyTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastModifyTime = 5
    mvarLastModifyTime = vData
End Property


Public Property Get LastModifyTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastModifyTime
    LastModifyTime = mvarLastModifyTime
End Property



Public Property Let CreationDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CreationDate = 5
    mvarCreationDate = vData
End Property


Public Property Get CreationDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CreationDate
    CreationDate = mvarCreationDate
End Property



Public Property Let FileTypeLong(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileTypeLong = 5
    mvarFileTypeLong = vData
End Property


Public Property Get FileTypeLong() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileTypeLong
    FileTypeLong = mvarFileTypeLong
End Property



Public Property Let FileAttributes(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileAttributes = 5
    mvarFileAttributes = vData
End Property


Public Property Get FileAttributes() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileAttributes
    FileAttributes = mvarFileAttributes
End Property



Public Property Let FileSizeKB(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileSize = 5
    mvarFileSizeKB = vData
End Property


Public Property Get FileSizeKB() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileSize
    FileSizeKB = mvarFileSizeKB
End Property



Public Property Let VerRevision(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VerRevision = 5
    mvarVerRevision = vData
End Property


Public Property Get VerRevision() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VerRevision
    VerRevision = mvarVerRevision
End Property



Public Property Let VerMinor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VerMinor = 5
    mvarVerMinor = vData
End Property


Public Property Get VerMinor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VerMinor
    VerMinor = mvarVerMinor
End Property



Public Property Let VerMajor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VerMajor = 5
    mvarVerMajor = vData
End Property


Public Property Get VerMajor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VerMajor
    VerMajor = mvarVerMajor
End Property



Public Property Let VerNotUsedVB(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VerNotUsedVB = 5
    mvarVerNotUsedVB = vData
End Property


Public Property Get VerNotUsedVB() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VerNotUsedVB
    VerNotUsedVB = mvarVerNotUsedVB
End Property



Public Property Let FileType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileType = 5
    mvarFileType = vData
End Property


Public Property Get FileType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileType
    FileType = mvarFileType
End Property



Public Property Let FullFilePath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FullFilePath = 5
    mvarFullFilePath = vData
End Property


Public Property Get FullFilePath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FullFilePath
    FullFilePath = mvarFullFilePath
End Property



Public Property Let Directory(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Directory = 5
    mvarDirectory = vData
End Property


Public Property Get Directory() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Directory
    Directory = mvarDirectory
End Property



Public Property Let Filename(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Filename = 5
    mvarFilename = vData
End Property


Public Property Get Filename() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Filename
    Filename = mvarFilename
End Property



