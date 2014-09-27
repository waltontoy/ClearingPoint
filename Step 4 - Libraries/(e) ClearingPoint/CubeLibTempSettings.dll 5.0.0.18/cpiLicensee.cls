VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiLicensee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarLic_ID As Long
Private mvarLic_Name As String
Private mvarLic_Address As String
Private mvarLic_City As String
Private mvarLic_PostalCode As String
Private mvarLic_Country As String
Private mvarLic_Phone As String
Private mvarLic_Fax As String
Private mvarLic_Email As String
Private mvarLic_LegalInfo As String
Private mvarLic_Currency As String
Private mvarLic_Database As String
Private mvarLic_UseEXACTID As Boolean
Private mvarLic_Language As String
Private mvarLic_Logo As Variant
Private mvarLic_Logosize As String
Private mvarLic_LogoProperties As String
Private mvarLic_Website As String
Private mvarLic_Key As String
Private mvarLic_User As String

Public Property Let Lic_ID(ByVal vData As Long)
    
    mvarLic_ID = vData

End Property

Public Property Get Lic_ID() As Long
    
    Lic_ID = mvarLic_ID

End Property

Public Property Let Lic_Name(ByVal vData As String)
    
    mvarLic_Name = vData

End Property

Public Property Get Lic_Name() As String
    
    Lic_Name = mvarLic_Name

End Property

Public Property Let Lic_Address(ByVal vData As String)
    
    mvarLic_Address = vData

End Property

Public Property Get Lic_Address() As String
    
    Lic_Address = mvarLic_Address

End Property

Public Property Let Lic_City(ByVal vData As String)
    
    mvarLic_City = vData

End Property

Public Property Get Lic_City() As String
    
    Lic_City = mvarLic_City

End Property

Public Property Let Lic_PostalCode(ByVal vData As String)
    
    mvarLic_PostalCode = vData

End Property

Public Property Get Lic_PostalCode() As String
    
    Lic_PostalCode = mvarLic_PostalCode

End Property

Public Property Let Lic_Country(ByVal vData As String)
    
    mvarLic_Country = vData

End Property

Public Property Get Lic_Country() As String
    
    Lic_Country = mvarLic_Country

End Property

Public Property Let Lic_Phone(ByVal vData As String)
    
    mvarLic_Phone = vData

End Property

Public Property Get Lic_Phone() As String
    
    Lic_Phone = mvarLic_Phone

End Property

Public Property Let Lic_Fax(ByVal vData As String)
    
    mvarLic_Fax = vData

End Property

Public Property Get Lic_Fax() As String
    
    Lic_Fax = mvarLic_Fax

End Property

Public Property Let Lic_Email(ByVal vData As String)
    
    mvarLic_Email = vData

End Property

Public Property Get Lic_Email() As String
    
    Lic_Email = mvarLic_Email

End Property

Public Property Get Lic_LegalInfo() As String
    
    Lic_LegalInfo = mvarLic_LegalInfo

End Property

Public Property Let Lic_LegalInfo(ByVal vData As String)
    
    mvarLic_LegalInfo = vData

End Property

Public Property Get Lic_Currency() As String
    
    Lic_Currency = mvarLic_Currency

End Property

Public Property Let Lic_Currency(ByVal vData As String)
    
    mvarLic_Currency = vData

End Property

Public Property Get Lic_Database() As String
    
    Lic_Database = mvarLic_Database

End Property

Public Property Let Lic_Database(ByVal vData As String)
    
    mvarLic_Database = vData

End Property

Public Property Get Lic_UseEXACTID() As Boolean
    
    Lic_UseEXACTID = mvarLic_UseEXACTID

End Property

Public Property Let Lic_UseEXACTID(ByVal vData As Boolean)
    
    mvarLic_UseEXACTID = vData

End Property

Public Property Get Lic_Language() As String
    
    Lic_Language = mvarLic_Language

End Property

Public Property Let Lic_Language(ByVal vData As String)
    
    mvarLic_Language = vData

End Property

Public Property Get Lic_Logo() As Variant
    
    Lic_Logo = mvarLic_Logo

End Property

Public Property Let Lic_Logo(ByVal vData As Variant)
    
    mvarLic_Logo = vData

End Property

Public Property Get Lic_Logosize() As String
    
    Lic_Logosize = mvarLic_Logosize

End Property

Public Property Let Lic_Logosize(ByVal vData As String)
    
    mvarLic_Logosize = vData

End Property

Public Property Get Lic_LogoProperties() As String
    
    Lic_LogoProperties = mvarLic_LogoProperties

End Property

Public Property Let Lic_LogoProperties(ByVal vData As String)
    
    mvarLic_LogoProperties = vData

End Property

Public Property Get Lic_Website() As String
    
    Lic_Website = mvarLic_Website

End Property

Public Property Let Lic_Website(ByVal vData As String)
    
    mvarLic_Website = vData

End Property

Public Property Get Lic_Key() As String
    
    Lic_Key = mvarLic_Key

End Property

Public Property Let Lic_Key(ByVal vData As String)
    
    mvarLic_Key = vData

End Property

Public Property Get Lic_User() As String
    
    Lic_User = mvarLic_User

End Property

Public Property Let Lic_User(ByVal vData As String)
    
    mvarLic_User = vData

End Property


