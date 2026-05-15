' NOTE: Generated code may require at least .NET Framework 4.5 or .NET Core/Standard 2.0.
'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03"),
 System.Xml.Serialization.XmlRootAttribute([Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03", IsNullable:=False)>
Partial Public Class ChaseFromXML

    Private cstmrCdtTrfInitnField As DocumentCstmrCdtTrfInitn

    '''<remarks/>
    Public Property CstmrCdtTrfInitn() As DocumentCstmrCdtTrfInitn
        Get
            Return Me.cstmrCdtTrfInitnField
        End Get
        Set
            Me.cstmrCdtTrfInitnField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitn

    Private grpHdrField As DocumentCstmrCdtTrfInitnGrpHdr

    Private pmtInfField As DocumentCstmrCdtTrfInitnPmtInf

    '''<remarks/>
    Public Property GrpHdr() As DocumentCstmrCdtTrfInitnGrpHdr
        Get
            Return Me.grpHdrField
        End Get
        Set
            Me.grpHdrField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property PmtInf() As DocumentCstmrCdtTrfInitnPmtInf
        Get
            Return Me.pmtInfField
        End Get
        Set
            Me.pmtInfField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnGrpHdr

    Private msgIdField As String

    Private creDtTmField As Date

    Private nbOfTxsField As Byte

    Private ctrlSumField As Decimal

    Private initgPtyField As DocumentCstmrCdtTrfInitnGrpHdrInitgPty

    '''<remarks/>
    Public Property MsgId() As String
        Get
            Return Me.msgIdField
        End Get
        Set
            Me.msgIdField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property CreDtTm() As Date
        Get
            Return Me.creDtTmField
        End Get
        Set
            Me.creDtTmField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property NbOfTxs() As Byte
        Get
            Return Me.nbOfTxsField
        End Get
        Set
            Me.nbOfTxsField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property CtrlSum() As Decimal
        Get
            Return Me.ctrlSumField
        End Get
        Set
            Me.ctrlSumField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property InitgPty() As DocumentCstmrCdtTrfInitnGrpHdrInitgPty
        Get
            Return Me.initgPtyField
        End Get
        Set
            Me.initgPtyField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnGrpHdrInitgPty

    Private nmField As String

    '''<remarks/>
    Public Property Nm() As String
        Get
            Return Me.nmField
        End Get
        Set
            Me.nmField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInf

    Private pmtInfIdField As String

    Private pmtMtdField As String

    Private nbOfTxsField As Byte

    Private ctrlSumField As Decimal

    Private pmtTpInfField As DocumentCstmrCdtTrfInitnPmtInfPmtTpInf

    Private reqdExctnDtField As Date

    Private dbtrField As DocumentCstmrCdtTrfInitnPmtInfDbtr

    Private dbtrAcctField As DocumentCstmrCdtTrfInitnPmtInfDbtrAcct

    Private dbtrAgtField As DocumentCstmrCdtTrfInitnPmtInfDbtrAgt

    Private cdtTrfTxInfField() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInf

    '''<remarks/>
    Public Property PmtInfId() As String
        Get
            Return Me.pmtInfIdField
        End Get
        Set
            Me.pmtInfIdField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property PmtMtd() As String
        Get
            Return Me.pmtMtdField
        End Get
        Set
            Me.pmtMtdField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property NbOfTxs() As Byte
        Get
            Return Me.nbOfTxsField
        End Get
        Set
            Me.nbOfTxsField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property CtrlSum() As Decimal
        Get
            Return Me.ctrlSumField
        End Get
        Set
            Me.ctrlSumField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property PmtTpInf() As DocumentCstmrCdtTrfInitnPmtInfPmtTpInf
        Get
            Return Me.pmtTpInfField
        End Get
        Set
            Me.pmtTpInfField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlElementAttribute(DataType:="date")>
    Public Property ReqdExctnDt() As Date
        Get
            Return Me.reqdExctnDtField
        End Get
        Set
            Me.reqdExctnDtField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property Dbtr() As DocumentCstmrCdtTrfInitnPmtInfDbtr
        Get
            Return Me.dbtrField
        End Get
        Set
            Me.dbtrField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property DbtrAcct() As DocumentCstmrCdtTrfInitnPmtInfDbtrAcct
        Get
            Return Me.dbtrAcctField
        End Get
        Set
            Me.dbtrAcctField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property DbtrAgt() As DocumentCstmrCdtTrfInitnPmtInfDbtrAgt
        Get
            Return Me.dbtrAgtField
        End Get
        Set
            Me.dbtrAgtField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlElementAttribute("CdtTrfTxInf")>
    Public Property CdtTrfTxInf() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInf()
        Get
            Return Me.cdtTrfTxInfField
        End Get
        Set
            Me.cdtTrfTxInfField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfPmtTpInf

    Private svcLvlField As DocumentCstmrCdtTrfInitnPmtInfPmtTpInfSvcLvl

    Private lclInstrmField As DocumentCstmrCdtTrfInitnPmtInfPmtTpInfLclInstrm

    Private ctgyPurpField As DocumentCstmrCdtTrfInitnPmtInfPmtTpInfCtgyPurp

    '''<remarks/>
    Public Property SvcLvl() As DocumentCstmrCdtTrfInitnPmtInfPmtTpInfSvcLvl
        Get
            Return Me.svcLvlField
        End Get
        Set
            Me.svcLvlField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property LclInstrm() As DocumentCstmrCdtTrfInitnPmtInfPmtTpInfLclInstrm
        Get
            Return Me.lclInstrmField
        End Get
        Set
            Me.lclInstrmField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property CtgyPurp() As DocumentCstmrCdtTrfInitnPmtInfPmtTpInfCtgyPurp
        Get
            Return Me.ctgyPurpField
        End Get
        Set
            Me.ctgyPurpField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfPmtTpInfSvcLvl

    Private cdField As String

    '''<remarks/>
    Public Property Cd() As String
        Get
            Return Me.cdField
        End Get
        Set
            Me.cdField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfPmtTpInfLclInstrm

    Private cdField As String

    '''<remarks/>
    Public Property Cd() As String
        Get
            Return Me.cdField
        End Get
        Set
            Me.cdField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfPmtTpInfCtgyPurp

    Private prtryField As String

    '''<remarks/>
    Public Property Prtry() As String
        Get
            Return Me.prtryField
        End Get
        Set
            Me.prtryField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfDbtr

    Private nmField As String

    Private pstlAdrField As DocumentCstmrCdtTrfInitnPmtInfDbtrPstlAdr

    Private idField As DocumentCstmrCdtTrfInitnPmtInfDbtrID

    '''<remarks/>
    Public Property Nm() As String
        Get
            Return Me.nmField
        End Get
        Set
            Me.nmField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property PstlAdr() As DocumentCstmrCdtTrfInitnPmtInfDbtrPstlAdr
        Get
            Return Me.pstlAdrField
        End Get
        Set
            Me.pstlAdrField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property Id() As DocumentCstmrCdtTrfInitnPmtInfDbtrID
        Get
            Return Me.idField
        End Get
        Set
            Me.idField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfDbtrPstlAdr

    Private ctryField As String

    '''<remarks/>
    Public Property Ctry() As String
        Get
            Return Me.ctryField
        End Get
        Set
            Me.ctryField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfDbtrID

    Private orgIdField As DocumentCstmrCdtTrfInitnPmtInfDbtrIDOrgId

    '''<remarks/>
    Public Property OrgId() As DocumentCstmrCdtTrfInitnPmtInfDbtrIDOrgId
        Get
            Return Me.orgIdField
        End Get
        Set
            Me.orgIdField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfDbtrIDOrgId

    Private othrField As DocumentCstmrCdtTrfInitnPmtInfDbtrIDOrgIdOthr

    '''<remarks/>
    Public Property Othr() As DocumentCstmrCdtTrfInitnPmtInfDbtrIDOrgIdOthr
        Get
            Return Me.othrField
        End Get
        Set
            Me.othrField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfDbtrIDOrgIdOthr

    Private idField As UInteger

    Private schmeNmField As DocumentCstmrCdtTrfInitnPmtInfDbtrIDOrgIdOthrSchmeNm

    '''<remarks/>
    Public Property Id() As UInteger
        Get
            Return Me.idField
        End Get
        Set
            Me.idField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property SchmeNm() As DocumentCstmrCdtTrfInitnPmtInfDbtrIDOrgIdOthrSchmeNm
        Get
            Return Me.schmeNmField
        End Get
        Set
            Me.schmeNmField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfDbtrIDOrgIdOthrSchmeNm

    Private prtryField As String

    '''<remarks/>
    Public Property Prtry() As String
        Get
            Return Me.prtryField
        End Get
        Set
            Me.prtryField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfDbtrAcct

    Private idField As DocumentCstmrCdtTrfInitnPmtInfDbtrAcctID

    Private ccyField As String

    '''<remarks/>
    Public Property Id() As DocumentCstmrCdtTrfInitnPmtInfDbtrAcctID
        Get
            Return Me.idField
        End Get
        Set
            Me.idField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property Ccy() As String
        Get
            Return Me.ccyField
        End Get
        Set
            Me.ccyField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfDbtrAcctID

    Private othrField As DocumentCstmrCdtTrfInitnPmtInfDbtrAcctIDOthr

    '''<remarks/>
    Public Property Othr() As DocumentCstmrCdtTrfInitnPmtInfDbtrAcctIDOthr
        Get
            Return Me.othrField
        End Get
        Set
            Me.othrField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfDbtrAcctIDOthr

    Private idField As UInteger

    '''<remarks/>
    Public Property Id() As UInteger
        Get
            Return Me.idField
        End Get
        Set
            Me.idField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfDbtrAgt

    Private finInstnIdField As DocumentCstmrCdtTrfInitnPmtInfDbtrAgtFinInstnId

    '''<remarks/>
    Public Property FinInstnId() As DocumentCstmrCdtTrfInitnPmtInfDbtrAgtFinInstnId
        Get
            Return Me.finInstnIdField
        End Get
        Set
            Me.finInstnIdField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfDbtrAgtFinInstnId

    Private bICField As String

    Private clrSysMmbIdField As DocumentCstmrCdtTrfInitnPmtInfDbtrAgtFinInstnIdClrSysMmbId

    Private pstlAdrField As DocumentCstmrCdtTrfInitnPmtInfDbtrAgtFinInstnIdPstlAdr

    '''<remarks/>
    Public Property BIC() As String
        Get
            Return Me.bICField
        End Get
        Set
            Me.bICField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property ClrSysMmbId() As DocumentCstmrCdtTrfInitnPmtInfDbtrAgtFinInstnIdClrSysMmbId
        Get
            Return Me.clrSysMmbIdField
        End Get
        Set
            Me.clrSysMmbIdField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property PstlAdr() As DocumentCstmrCdtTrfInitnPmtInfDbtrAgtFinInstnIdPstlAdr
        Get
            Return Me.pstlAdrField
        End Get
        Set
            Me.pstlAdrField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfDbtrAgtFinInstnIdClrSysMmbId

    Private mmbIdField As UInteger

    '''<remarks/>
    Public Property MmbId() As UInteger
        Get
            Return Me.mmbIdField
        End Get
        Set
            Me.mmbIdField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfDbtrAgtFinInstnIdPstlAdr

    Private ctryField As String

    '''<remarks/>
    Public Property Ctry() As String
        Get
            Return Me.ctryField
        End Get
        Set
            Me.ctryField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInf

    Private pmtIdField As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfPmtId

    Private pmtTpInfField As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfPmtTpInf

    Private amtField As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfAmt

    Private cdtrAgtField As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAgt

    Private cdtrField As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtr

    Private cdtrAcctField As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAcct

    Private rmtInfField As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfRmtInf

    '''<remarks/>
    Public Property PmtId() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfPmtId
        Get
            Return Me.pmtIdField
        End Get
        Set
            Me.pmtIdField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property PmtTpInf() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfPmtTpInf
        Get
            Return Me.pmtTpInfField
        End Get
        Set
            Me.pmtTpInfField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property Amt() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfAmt
        Get
            Return Me.amtField
        End Get
        Set
            Me.amtField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property CdtrAgt() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAgt
        Get
            Return Me.cdtrAgtField
        End Get
        Set
            Me.cdtrAgtField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property Cdtr() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtr
        Get
            Return Me.cdtrField
        End Get
        Set
            Me.cdtrField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property CdtrAcct() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAcct
        Get
            Return Me.cdtrAcctField
        End Get
        Set
            Me.cdtrAcctField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property RmtInf() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfRmtInf
        Get
            Return Me.rmtInfField
        End Get
        Set
            Me.rmtInfField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfPmtId

    Private instrIdField As String

    Private endToEndIdField As String

    '''<remarks/>
    Public Property InstrId() As String
        Get
            Return Me.instrIdField
        End Get
        Set
            Me.instrIdField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property EndToEndId() As String
        Get
            Return Me.endToEndIdField
        End Get
        Set
            Me.endToEndIdField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfPmtTpInf

    Private lclInstrmField As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfPmtTpInfLclInstrm

    '''<remarks/>
    Public Property LclInstrm() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfPmtTpInfLclInstrm
        Get
            Return Me.lclInstrmField
        End Get
        Set
            Me.lclInstrmField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfPmtTpInfLclInstrm

    Private cdField As String

    '''<remarks/>
    Public Property Cd() As String
        Get
            Return Me.cdField
        End Get
        Set
            Me.cdField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfAmt

    Private instdAmtField As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfAmtInstdAmt

    '''<remarks/>
    Public Property InstdAmt() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfAmtInstdAmt
        Get
            Return Me.instdAmtField
        End Get
        Set
            Me.instdAmtField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfAmtInstdAmt

    Private ccyField As String

    Private valueField As Decimal

    '''<remarks/>
    <System.Xml.Serialization.XmlAttributeAttribute()>
    Public Property Ccy() As String
        Get
            Return Me.ccyField
        End Get
        Set
            Me.ccyField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlTextAttribute()>
    Public Property Value() As Decimal
        Get
            Return Me.valueField
        End Get
        Set
            Me.valueField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAgt

    Private finInstnIdField As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAgtFinInstnId

    '''<remarks/>
    Public Property FinInstnId() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAgtFinInstnId
        Get
            Return Me.finInstnIdField
        End Get
        Set
            Me.finInstnIdField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAgtFinInstnId

    Private clrSysMmbIdField As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAgtFinInstnIdClrSysMmbId

    Private nmField As String

    Private pstlAdrField As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAgtFinInstnIdPstlAdr

    '''<remarks/>
    Public Property ClrSysMmbId() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAgtFinInstnIdClrSysMmbId
        Get
            Return Me.clrSysMmbIdField
        End Get
        Set
            Me.clrSysMmbIdField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property Nm() As String
        Get
            Return Me.nmField
        End Get
        Set
            Me.nmField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property PstlAdr() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAgtFinInstnIdPstlAdr
        Get
            Return Me.pstlAdrField
        End Get
        Set
            Me.pstlAdrField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAgtFinInstnIdClrSysMmbId

    Private mmbIdField As UInteger

    '''<remarks/>
    Public Property MmbId() As UInteger
        Get
            Return Me.mmbIdField
        End Get
        Set
            Me.mmbIdField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAgtFinInstnIdPstlAdr

    Private twnNmField As String

    Private ctrySubDvsnField As String

    Private ctryField As String

    '''<remarks/>
    Public Property TwnNm() As String
        Get
            Return Me.twnNmField
        End Get
        Set
            Me.twnNmField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property CtrySubDvsn() As String
        Get
            Return Me.ctrySubDvsnField
        End Get
        Set
            Me.ctrySubDvsnField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property Ctry() As String
        Get
            Return Me.ctryField
        End Get
        Set
            Me.ctryField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtr

    Private nmField As String

    Private pstlAdrField As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrPstlAdr

    '''<remarks/>
    Public Property Nm() As String
        Get
            Return Me.nmField
        End Get
        Set
            Me.nmField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property PstlAdr() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrPstlAdr
        Get
            Return Me.pstlAdrField
        End Get
        Set
            Me.pstlAdrField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrPstlAdr

    Private ctryField As String

    Private adrLineField() As String

    '''<remarks/>
    Public Property Ctry() As String
        Get
            Return Me.ctryField
        End Get
        Set
            Me.ctryField = Value
        End Set
    End Property

    '''<remarks/>
    <System.Xml.Serialization.XmlElementAttribute("AdrLine")>
    Public Property AdrLine() As String()
        Get
            Return Me.adrLineField
        End Get
        Set
            Me.adrLineField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAcct

    Private idField As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAcctID

    Private tpField As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAcctTP

    '''<remarks/>
    Public Property Id() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAcctID
        Get
            Return Me.idField
        End Get
        Set
            Me.idField = Value
        End Set
    End Property

    '''<remarks/>
    Public Property Tp() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAcctTP
        Get
            Return Me.tpField
        End Get
        Set
            Me.tpField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAcctID

    Private othrField As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAcctIDOthr

    '''<remarks/>
    Public Property Othr() As DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAcctIDOthr
        Get
            Return Me.othrField
        End Get
        Set
            Me.othrField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAcctIDOthr

    Private idField As UInteger

    '''<remarks/>
    Public Property Id() As UInteger
        Get
            Return Me.idField
        End Get
        Set
            Me.idField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfCdtrAcctTP

    Private cdField As String

    '''<remarks/>
    Public Property Cd() As String
        Get
            Return Me.cdField
        End Get
        Set
            Me.cdField = Value
        End Set
    End Property
End Class

'''<remarks/>
<System.SerializableAttribute(),
 System.ComponentModel.DesignerCategoryAttribute("code"),
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03")>
Partial Public Class DocumentCstmrCdtTrfInitnPmtInfCdtTrfTxInfRmtInf

    Private ustrdField As String

    '''<remarks/>
    Public Property Ustrd() As String
        Get
            Return Me.ustrdField
        End Get
        Set
            Me.ustrdField = Value
        End Set
    End Property
End Class

