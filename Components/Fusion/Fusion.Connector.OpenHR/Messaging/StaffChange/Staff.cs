using System;
using Fusion.Messages.SocialCare;
using System.Xml;
using System.Xml.Serialization;

namespace Fusion.Connector.OpenHR.Messaging
{
    
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
[System.Xml.Serialization.XmlRootAttribute(Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare", IsNullable = false)]
    public partial class staffChange
{

    private staffChangeData _dataField;

    private int _versionField;

    private string _staffRefField;

    public staffChangeData data
    {
        get
        {
            return this._dataField;
        }
        set
        {
            this._dataField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public int version
    {
        get
        {
            return this._versionField;
        }
        set
        {
            this._versionField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string staffRef
    {
        get
        {
            return this._staffRefField;
        }
        set
        {
            this._staffRefField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]    
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]

public partial class staffChangeData
{

    private staffChangeDataStaff staffField;

    private string auditUserNameField;

    private recordStatusRescindable recordStatusField;

    public staffChangeDataStaff staff
    {
        get
        {
            return this.staffField;
        }
        set
        {
            this.staffField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string auditUserName
    {
        get
        {
            return this.auditUserNameField;
        }
        set
        {
            this.auditUserNameField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public recordStatusRescindable recordStatus
    {
        get
        {
            return this.recordStatusField;
        }
        set
        {
            this.recordStatusField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
public partial class staffChangeDataStaff
{

    private string titleField;

    private string forenamesField;

    private string surnameField;

    private string preferredNameField;

    private string payrollNumberField;

    private System.DateTime? dOBField;

    private staffChangeDataStaffEmployeeType employeeTypeField;

    private string workMobileField;

    private string personalMobileField;

    private string workPhoneNumberField;

    private string homePhoneNumberField;

    private string emailField;

    private string personalEmailField;

    private gender? genderField;

    private System.DateTime? startDateField;

    private System.Nullable<System.DateTime> leavingDateField;

    private bool leavingDateFieldSpecified;

    private string leavingReasonField;

    private string companyNameField;

    private string jobTitleField;

    private string managerRefField;

    private staffChangeDataStaffHomeAddress homeAddressField;

    private string nationalInsuranceNumberField;

    [System.Xml.Serialization.XmlElementAttribute(IsNullable = true)]
    public string title {get;set;}
    //{
    //    get
    //    {
    //        return this.titleField;
    //    }
    //    set
    //    {
    //        this.titleField = value;
    //    }
    //}

    /// <remarks/>
    public string forenames
    {
        get
        {
            return this.forenamesField;
        }
        set
        {
            this.forenamesField = value;
        }
    }

    /// <remarks/>
    public string surname
    {
        get
        {
            return this.surnameField;
        }
        set
        {
            this.surnameField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(IsNullable = true)]
    public string preferredName
    {
        get
        {
            return this.preferredNameField;
        }
        set
        {
            this.preferredNameField = value;
        }
    }

    /// <remarks/>
    public string payrollNumber
    {
        get
        {
            return this.payrollNumberField;
        }
        set
        {
            this.payrollNumberField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(DataType = "date")]
//    [System.Xml.Serialization.XmlElementAttribute(IsNullable = false)]
    public System.DateTime? DOB
    {
        get
        {
            return this.dOBField;
        }
        set
        {
            this.dOBField = value;
        }
    }

    /// <remarks/>
    public staffChangeDataStaffEmployeeType employeeType
    {
        get
        {
            return this.employeeTypeField;
        }
        set
        {
            this.employeeTypeField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(IsNullable = true)]
    public string workMobile
    {
        get
        {
            return this.workMobileField;
        }
        set
        {
            this.workMobileField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(IsNullable = true)]
    public string personalMobile
    {
        get
        {
            return this.personalMobileField;
        }
        set
        {
            this.personalMobileField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(IsNullable = true)]
    public string workPhoneNumber
    {
        get
        {
            return this.workPhoneNumberField;
        }
        set
        {
            this.workPhoneNumberField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(IsNullable = true)]
    public string homePhoneNumber
    {
        get
        {
            return this.homePhoneNumberField;
        }
        set
        {
            this.homePhoneNumberField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(IsNullable = true)]
    public string email
    {
        get
        {
            return this.emailField;
        }
        set
        {
            this.emailField = value;
        }
    }

    /// <remarks/>
    /// 

    [System.Xml.Serialization.XmlElementAttribute(IsNullable = true)]
    public string personalEmail
    {
        get
        {
            return this.personalEmailField;
        }
        set
        {
            this.personalEmailField = value;
        }
    }

    /// <remarks/>
    public gender? gender
    {
        get
        {
            return this.genderField;
        }
        set
        {
            this.genderField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(DataType = "date")]
    public System.DateTime? startDate
    {
        get
        {
            return this.startDateField;
        }
        set
        {
            this.startDateField = value;
        }
    }

    [System.Xml.Serialization.XmlElementAttribute(DataType = "date", IsNullable = true)]
    public DateTime? leavingDate
    {
        get
        {
            return this.leavingDateField;
        }
        set
        {
            this.leavingDateField = value;
        }
    }

    [System.Xml.Serialization.XmlIgnoreAttribute()]
    public bool leavingDateSpecified
    {
        get
        {
            return this.leavingDateFieldSpecified;
        }
        set
        {
            this.leavingDateFieldSpecified = value;
        }
    }

    [System.Xml.Serialization.XmlElementAttribute(IsNullable = true)]
    public string leavingReason
    {
        get
        {
            return this.leavingReasonField;
        }
        set
        {
            this.leavingReasonField = value;
        }
    }

    /// <remarks/>
    public string companyName
    {
        get
        {
            return this.companyNameField;
        }
        set
        {
            this.companyNameField = value;
        }
    }

    /// <remarks/>
    public string jobTitle
    {
        get
        {
            return this.jobTitleField;
        }
        set
        {
            this.jobTitleField = value;
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlElementAttribute(IsNullable = true)]
    public string managerRef
    {
        get
        {
            return this.managerRefField;
        }
        set
        {
            this.managerRefField = value;
        }
    }

    /// <remarks/>
    public staffChangeDataStaffHomeAddress homeAddress
    {
        get
        {
            return this.homeAddressField;
        }
        set
        {
            this.homeAddressField = value;
        }
    }

    /// <remarks/>
    public string nationalInsuranceNumber
    {
        get
        {
            return this.nationalInsuranceNumberField;
        }
        set
        {
            this.nationalInsuranceNumberField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
[System.SerializableAttribute()]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
public enum staffChangeDataStaffEmployeeType
{

    /// <remarks/>
    [System.Xml.Serialization.XmlEnumAttribute("Agency Worker")]
    AgencyWorker,

    /// <remarks/>
    Employee,
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
[System.SerializableAttribute()]
[System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://advancedcomputersoftware.com/xml/fusion")]
public enum gender
{

    /// <remarks/>
    Male,

    /// <remarks/>
    Female,

    /// <remarks/>
    Unknown,
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
public partial class staffChangeDataStaffHomeAddress
{

    private string addressLine1Field;

    private string addressLine2Field;

    private string addressLine3Field;

    private string addressLine4Field;

    private string addressLine5Field;

    private string postCodeField;

    [System.Xml.Serialization.XmlElementAttribute(IsNullable = false)]
    public string addressLine1
    {
        get
        {
            return this.addressLine1Field;
        }
        set
        {
            this.addressLine1Field = value;
        }
    }

    [System.Xml.Serialization.XmlElementAttribute(IsNullable = false)]
    public string addressLine2
    {
        get
        {
            return this.addressLine2Field;
        }
        set
        {
            this.addressLine2Field = value;
        }
    }

    [System.Xml.Serialization.XmlElementAttribute(IsNullable = false)]
    public string addressLine3
    {
        get
        {
            return this.addressLine3Field;
        }
        set
        {
            this.addressLine3Field = value;
        }
    }

    [System.Xml.Serialization.XmlElementAttribute(IsNullable = false)]
    public string addressLine4
    {
        get
        {
            return this.addressLine4Field;
        }
        set
        {
            this.addressLine4Field = value;
        }
    }

    [System.Xml.Serialization.XmlElementAttribute(IsNullable = false)]
    public string addressLine5
    {
        get
        {
            return this.addressLine5Field;
        }
        set
        {
            this.addressLine5Field = value;
        }
    }

    [System.Xml.Serialization.XmlElementAttribute(IsNullable = false)]
    public string postCode
    {
        get
        {
            return this.postCodeField;
        }
        set
        {
            this.postCodeField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
[System.SerializableAttribute()]
[System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://advancedcomputersoftware.com/xml/fusion")]
public enum recordStatusRescindable
{

    /// <remarks/>
    Active,

    /// <remarks/>
    Inactive,

    /// <remarks/>
    RecordCreatedInError,
}
}


