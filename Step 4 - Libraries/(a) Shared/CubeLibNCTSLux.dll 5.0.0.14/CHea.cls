VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHea"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Private mvarDateTimeOfPreparation As CDateTimeElement 'local copy

Private mvarDocNumHEA5  As String
Private mvarCanDecHEA93  As String
Private mvarDatOfCanReqHEA147  As String
Private mvarCanIniByCusHEA94  As String
Private mvarDatOfCanDecHEA146  As String
Private mvarCanJusHEA248  As String
Private mvarCanJusHEA248LNG  As String
Private mvarCanReaHEA250  As String
Private mvarCanReaHEA250LNG  As String
Private mvarRefNumHEA4  As String
Private mvarTypOfDecHEA24  As String
Private mvarCouOfDesCodHEA30  As String
Private mvarAgrLocOfGooCodHEA38  As String
Private mvarAgrLocOfGooHEA39  As String
Private mvarAgrLocOfGooHEA39LNG  As String
Private mvarAutLocOfGooCodHEA41  As String
Private mvarPlaOfLoaCodHEA46  As String
Private mvarCouOfDisCodHEA55  As String
Private mvarCusSubPlaHEA66  As String
Private mvarInlTraModHEA75  As String
Private mvarTraModAtBorHEA76  As String
Private mvarIdeOfMeaOfTraAtDHEA78  As String
Private mvarIdeOfMeaOfTraAtDHEA78LNG  As String
Private mvarNatOfMeaOfTraAtDHEA80  As String
Private mvarIdeOfMeaOfTraCroHEA85  As String
Private mvarIdeOfMeaOfTraCroHEA85LNG  As String
Private mvarNatOfMeaOfTraCroHEA87  As String
Private mvarTypOfMeaOfTraCroHEA88  As String
Private mvarConInHEA96  As String
Private mvarDiaLanIndAtDepHEA254  As String
Private mvarNCTSAccDocHEA601LNG  As String
Private mvarNumOfLoaLisHEA304  As String
Private mvarTotNumOfIteHEA305  As String
Private mvarTotNumOfPacHEA306  As String
Private mvarTotGroMasHEA307  As String
Private mvarDecDatHEA383  As String
Private mvarDecPlaHEA394  As String
Private mvarDecPlaHEA394LNG  As String
Private mvarDecRejDatHEA159  As String
Private mvarDecRejReaHEA252  As String
Private mvarDecRejReaHEA252LNG  As String
Private mvarAccDotHEA158  As String
Private mvarNCTRetCopHEA104  As String
Private mvarAccDatHEA158  As String
Private mvarIssDatHEA186  As String
Private mvarWriOffDat  As String
Private mvarDatOfConNotHEA148  As String
Private mvarOriMesIdeMES22 As String


'LET
Public Property Let DocNumHEA5(ByVal vData As String)
    mvarDocNumHEA5 = vData
End Property
Public Property Let CanDecHEA93(ByVal vData As String)
    mvarCanDecHEA93 = vData
End Property
Public Property Let DatOfCanReqHEA147(ByVal vData As String)
    mvarDatOfCanReqHEA147 = vData
End Property
Public Property Let CanIniByCusHEA94(ByVal vData As String)
    mvarCanIniByCusHEA94 = vData
End Property
Public Property Let DatOfCanDecHEA146(ByVal vData As String)
    mvarDatOfCanDecHEA146 = vData
End Property
Public Property Let CanJusHEA248(ByVal vData As String)
    mvarCanJusHEA248 = vData
End Property
Public Property Let CanJusHEA248LNG(ByVal vData As String)
    mvarCanJusHEA248LNG = vData
End Property
Public Property Let CanReaHEA250(ByVal vData As String)
    mvarCanReaHEA250 = vData
End Property
Public Property Let CanReaHEA250LNG(ByVal vData As String)
    mvarCanReaHEA250LNG = vData
End Property
Public Property Let RefNumHEA4(ByVal vData As String)
    mvarRefNumHEA4 = vData
End Property
Public Property Let TypOfDecHEA24(ByVal vData As String)
    mvarTypOfDecHEA24 = vData
End Property
Public Property Let CouOfDesCodHEA30(ByVal vData As String)
    mvarCouOfDesCodHEA30 = vData
End Property
Public Property Let AgrLocOfGooCodHEA38(ByVal vData As String)
    mvarAgrLocOfGooCodHEA38 = vData
End Property
Public Property Let AgrLocOfGooHEA39(ByVal vData As String)
    mvarAgrLocOfGooHEA39 = vData
End Property
Public Property Let AgrLocOfGooHEA39LNG(ByVal vData As String)
    mvarAgrLocOfGooHEA39LNG = vData
End Property
Public Property Let AutLocOfGooCodHEA41(ByVal vData As String)
    mvarAutLocOfGooCodHEA41 = vData
End Property
Public Property Let PlaOfLoaCodHEA46(ByVal vData As String)
    mvarPlaOfLoaCodHEA46 = vData
End Property
Public Property Let CouOfDisCodHEA55(ByVal vData As String)
    mvarCouOfDisCodHEA55 = vData
End Property
Public Property Let CusSubPlaHEA66(ByVal vData As String)
    mvarCusSubPlaHEA66 = vData
End Property
Public Property Let InlTraModHEA75(ByVal vData As String)
    mvarInlTraModHEA75 = vData
End Property
Public Property Let TraModAtBorHEA76(ByVal vData As String)
    mvarTraModAtBorHEA76 = vData
End Property
Public Property Let IdeOfMeaOfTraAtDHEA78(ByVal vData As String)
    mvarIdeOfMeaOfTraAtDHEA78 = vData
End Property
Public Property Let IdeOfMeaOfTraAtDHEA78LNG(ByVal vData As String)
    mvarIdeOfMeaOfTraAtDHEA78LNG = vData
End Property
Public Property Let NatOfMeaOfTraAtDHEA80(ByVal vData As String)
    mvarNatOfMeaOfTraAtDHEA80 = vData
End Property
Public Property Let IdeOfMeaOfTraCroHEA85(ByVal vData As String)
    mvarIdeOfMeaOfTraCroHEA85 = vData
End Property
Public Property Let IdeOfMeaOfTraCroHEA85LNG(ByVal vData As String)
    mvarIdeOfMeaOfTraCroHEA85LNG = vData
End Property
Public Property Let NatOfMeaOfTraCroHEA87(ByVal vData As String)
    mvarNatOfMeaOfTraCroHEA87 = vData
End Property
Public Property Let TypOfMeaOfTraCroHEA88(ByVal vData As String)
    mvarTypOfMeaOfTraCroHEA88 = vData
End Property
Public Property Let ConInHEA96(ByVal vData As String)
    mvarConInHEA96 = vData
End Property
Public Property Let DiaLanIndAtDepHEA254(ByVal vData As String)
    mvarDiaLanIndAtDepHEA254 = vData
End Property
Public Property Let NCTSAccDocHEA601LNG(ByVal vData As String)
    mvarNCTSAccDocHEA601LNG = vData
End Property
Public Property Let NumOfLoaLisHEA304(ByVal vData As String)
    mvarNumOfLoaLisHEA304 = vData
End Property
Public Property Let TotNumOfIteHEA305(ByVal vData As String)
    mvarTotNumOfIteHEA305 = vData
End Property
Public Property Let TotNumOfPacHEA306(ByVal vData As String)
    mvarTotNumOfPacHEA306 = vData
End Property
Public Property Let TotGroMasHEA307(ByVal vData As String)
    mvarTotGroMasHEA307 = vData
End Property
Public Property Let DecDatHEA383(ByVal vData As String)
    mvarDecDatHEA383 = vData
End Property
Public Property Let DecPlaHEA394(ByVal vData As String)
    mvarDecPlaHEA394 = vData
End Property
Public Property Let DecPlaHEA394LNG(ByVal vData As String)
    mvarDecPlaHEA394LNG = vData
End Property
Public Property Let DecRejDatHEA159(ByVal vData As String)
    mvarDecRejDatHEA159 = vData
End Property
Public Property Let DecRejReaHEA252(ByVal vData As String)
    mvarDecRejReaHEA252 = vData
End Property
Public Property Let DecRejReaHEA252LNG(ByVal vData As String)
    mvarDecRejReaHEA252LNG = vData
End Property
Public Property Let AccDotHEA158(ByVal vData As String)
    mvarAccDotHEA158 = vData
End Property
Public Property Let NCTRetCopHEA104(ByVal vData As String)
    mvarNCTRetCopHEA104 = vData
End Property
Public Property Let AccDatHEA158(ByVal vData As String)
    mvarAccDatHEA158 = vData
End Property
Public Property Let IssDatHEA186(ByVal vData As String)
    mvarIssDatHEA186 = vData
End Property
Public Property Let WriOffDat(ByVal vData As String)
    mvarWriOffDat = vData
End Property
Public Property Let DatOfConNotHEA148(ByVal vData As String)
    mvarDatOfConNotHEA148 = vData
End Property
Public Property Let OriMesIdeMES22(ByVal vData As String)
    mvarOriMesIdeMES22 = vData
End Property



'GET
Public Property Get DocNumHEA5() As String
    DocNumHEA5 = mvarDocNumHEA5
End Property
Public Property Get CanDecHEA93() As String
    CanDecHEA93 = mvarCanDecHEA93
End Property
Public Property Get DatOfCanReqHEA147() As String
    DatOfCanReqHEA147 = mvarDatOfCanReqHEA147
End Property
Public Property Get CanIniByCusHEA94() As String
    CanIniByCusHEA94 = mvarCanIniByCusHEA94
End Property
Public Property Get DatOfCanDecHEA146() As String
    DatOfCanDecHEA146 = mvarDatOfCanDecHEA146
End Property
Public Property Get CanJusHEA248() As String
    CanJusHEA248 = mvarCanJusHEA248
End Property
Public Property Get CanJusHEA248LNG() As String
    CanJusHEA248LNG = mvarCanJusHEA248LNG
End Property
Public Property Get CanReaHEA250() As String
    CanReaHEA250 = mvarCanReaHEA250
End Property
Public Property Get CanReaHEA250LNG() As String
    CanReaHEA250LNG = mvarCanReaHEA250LNG
End Property
Public Property Get RefNumHEA4() As String
    RefNumHEA4 = mvarRefNumHEA4
End Property
Public Property Get TypOfDecHEA24() As String
    TypOfDecHEA24 = mvarTypOfDecHEA24
End Property
Public Property Get CouOfDesCodHEA30() As String
    CouOfDesCodHEA30 = mvarCouOfDesCodHEA30
End Property
Public Property Get AgrLocOfGooCodHEA38() As String
    AgrLocOfGooCodHEA38 = mvarAgrLocOfGooCodHEA38
End Property
Public Property Get AgrLocOfGooHEA39() As String
    AgrLocOfGooHEA39 = mvarAgrLocOfGooHEA39
End Property
Public Property Get AgrLocOfGooHEA39LNG() As String
    AgrLocOfGooHEA39LNG = mvarAgrLocOfGooHEA39LNG
End Property
Public Property Get AutLocOfGooCodHEA41() As String
    AutLocOfGooCodHEA41 = mvarAutLocOfGooCodHEA41
End Property
Public Property Get PlaOfLoaCodHEA46() As String
    PlaOfLoaCodHEA46 = mvarPlaOfLoaCodHEA46
End Property
Public Property Get CouOfDisCodHEA55() As String
    CouOfDisCodHEA55 = mvarCouOfDisCodHEA55
End Property
Public Property Get CusSubPlaHEA66() As String
    CusSubPlaHEA66 = mvarCusSubPlaHEA66
End Property
Public Property Get InlTraModHEA75() As String
    InlTraModHEA75 = mvarInlTraModHEA75
End Property
Public Property Get TraModAtBorHEA76() As String
    TraModAtBorHEA76 = mvarTraModAtBorHEA76
End Property
Public Property Get IdeOfMeaOfTraAtDHEA78() As String
    IdeOfMeaOfTraAtDHEA78 = mvarIdeOfMeaOfTraAtDHEA78
End Property
Public Property Get IdeOfMeaOfTraAtDHEA78LNG() As String
    IdeOfMeaOfTraAtDHEA78LNG = mvarIdeOfMeaOfTraAtDHEA78LNG
End Property
Public Property Get NatOfMeaOfTraAtDHEA80() As String
    NatOfMeaOfTraAtDHEA80 = mvarNatOfMeaOfTraAtDHEA80
End Property
Public Property Get IdeOfMeaOfTraCroHEA85() As String
    IdeOfMeaOfTraCroHEA85 = mvarIdeOfMeaOfTraCroHEA85
End Property
Public Property Get IdeOfMeaOfTraCroHEA85LNG() As String
    IdeOfMeaOfTraCroHEA85LNG = mvarIdeOfMeaOfTraCroHEA85LNG
End Property
Public Property Get NatOfMeaOfTraCroHEA87() As String
    NatOfMeaOfTraCroHEA87 = mvarNatOfMeaOfTraCroHEA87
End Property
Public Property Get TypOfMeaOfTraCroHEA88() As String
    TypOfMeaOfTraCroHEA88 = mvarTypOfMeaOfTraCroHEA88
End Property
Public Property Get ConInHEA96() As String
    ConInHEA96 = mvarConInHEA96
End Property
Public Property Get DiaLanIndAtDepHEA254() As String
    DiaLanIndAtDepHEA254 = mvarDiaLanIndAtDepHEA254
End Property
Public Property Get NCTSAccDocHEA601LNG() As String
    NCTSAccDocHEA601LNG = mvarNCTSAccDocHEA601LNG
End Property
Public Property Get NumOfLoaLisHEA304() As String
    NumOfLoaLisHEA304 = mvarNumOfLoaLisHEA304
End Property
Public Property Get TotNumOfIteHEA305() As String
    TotNumOfIteHEA305 = mvarTotNumOfIteHEA305
End Property
Public Property Get TotNumOfPacHEA306() As String
    TotNumOfPacHEA306 = mvarTotNumOfPacHEA306
End Property
Public Property Get TotGroMasHEA307() As String
    TotGroMasHEA307 = mvarTotGroMasHEA307
End Property
Public Property Get DecDatHEA383() As String
    DecDatHEA383 = mvarDecDatHEA383
End Property
Public Property Get DecPlaHEA394() As String
    DecPlaHEA394 = mvarDecPlaHEA394
End Property
Public Property Get DecPlaHEA394LNG() As String
    DecPlaHEA394LNG = mvarDecPlaHEA394LNG
End Property
Public Property Get DecRejDatHEA159() As String
    DecRejDatHEA159 = mvarDecRejDatHEA159
End Property
Public Property Get DecRejReaHEA252() As String
    DecRejReaHEA252 = mvarDecRejReaHEA252
End Property
Public Property Get DecRejReaHEA252LNG() As String
    DecRejReaHEA252LNG = mvarDecRejReaHEA252LNG
End Property
Public Property Get AccDotHEA158() As String
    AccDotHEA158 = mvarAccDotHEA158
End Property
Public Property Get NCTRetCopHEA104() As String
    NCTRetCopHEA104 = mvarNCTRetCopHEA104
End Property
Public Property Get AccDatHEA158() As String
    AccDatHEA158 = mvarAccDatHEA158
End Property
Public Property Get IssDatHEA186() As String
    IssDatHEA186 = mvarIssDatHEA186
End Property
Public Property Get WriOffDat() As String
    WriOffDat = mvarWriOffDat
End Property
Public Property Get DatOfConNotHEA148() As String
    DatOfConNotHEA148 = mvarDatOfConNotHEA148
End Property
Public Property Get OriMesIdeMES22() As String
    OriMesIdeMES22 = mvarOriMesIdeMES22
End Property


'Public Property Set DateTimeOfPreparation(ByVal vdata As CDateTimeElement)
''used when assigning an Object to the property, on the left side of a Set statement.
''Syntax: Set x.DateTimeOfPreparation = Form1
'    Set mvarDateTimeOfPreparation = vdata
'End Property
'
'
'Public Property Get DateTimeOfPreparation() As CDateTimeElement
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.DateTimeOfPreparation
'    Set DateTimeOfPreparation = mvarDateTimeOfPreparation
'End Property
'
'
'
'Public Property Let MessageSender(ByVal vdata As String)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.MessageSender = 5
'    mvarMessageSender = vdata
'End Property
'
'
'Public Property Get MessageSender() As String
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.MessageSender
'    MessageSender = mvarMessageSender
'End Property
'
'
'
'
