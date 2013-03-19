using System;
using Fusion.Connector.OpenHR.MessageComponents.Enums;

namespace Fusion.Connector.OpenHR.MessageComponents.Component
{
    
/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
public partial class Staff
{


    private string forenamesField;

    private string surnameField;

    private string preferredNameField;

    private string payrollNumberField;

    private System.DateTime? dOBField;

    private EmployeeType employeeTypeField;

    private string workMobileField;

    private string personalMobileField;

    private string workPhoneNumberField;

    private string homePhoneNumberField;

    private string emailField;

    private string personalEmailField;

    private Gender? genderField;

    private System.DateTime? startDateField;

    private DateTime? leavingDateField;

    private bool leavingDateFieldSpecified;

    private string leavingReasonField;

    private string companyNameField;

    private string jobTitleField;

    private string managerRefField;

    private Address homeAddressField;

    private string nationalInsuranceNumberField;

    [System.Xml.Serialization.XmlElementAttribute(IsNullable = true)]
    public string title {get;set;}


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
    public EmployeeType employeeType
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
    public Gender? gender
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
            return leavingDateField;
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
    public Address homeAddress
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



}


