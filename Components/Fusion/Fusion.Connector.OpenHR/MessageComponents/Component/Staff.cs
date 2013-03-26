using System;
using System.Xml.Serialization;
using Fusion.Connector.OpenHR.MessageComponents.Enums;

namespace Fusion.Connector.OpenHR.MessageComponents.Component
{
    
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
[SerializableAttribute]
[System.Diagnostics.DebuggerStepThroughAttribute]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[XmlTypeAttribute(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
public class Staff
{
    [XmlElementAttribute(IsNullable = true)]
    public string title {get;set;}

    public string forenames { get; set; }

    public string surname { get; set; }

    [XmlElementAttribute(IsNullable = true)]
    public string preferredName { get; set; }

    public string payrollNumber { get; set; }

    [XmlElementAttribute(DataType = "date", ElementName = "DOB")]
    public DateTime? dob { get; set; }

    public string employeeType {get;set;}

    [XmlElementAttribute(IsNullable = true)]
    public string workMobile { get; set; }

    [XmlElementAttribute(IsNullable = true)]
    public string personalMobile { get; set; }

    [XmlElementAttribute(IsNullable = true)]
    public string workPhoneNumber { get; set; }

    [XmlElementAttribute(IsNullable = true)]
    public string homePhoneNumber { get; set; }

    [XmlElementAttribute(IsNullable = true)]
    public string email { get; set; }

    [XmlElementAttribute(IsNullable = true)]
    public string personalEmail { get; set; }

    public Gender? gender { get; set; }

    [XmlElementAttribute(DataType = "date")]
    public DateTime? startDate { get; set; }

    [XmlElementAttribute(DataType = "date", IsNullable = true)]
    public DateTime? leavingDate { get; set; }

    [XmlIgnoreAttribute]
    public bool leavingDateSpecified { get; set; }

    [XmlElementAttribute(IsNullable = true)]
    public string leavingReason { get; set; }

    public string companyName { get; set; }

    public string jobTitle { get; set; }

    [XmlElementAttribute(IsNullable = true)]
    public string managerRef { get; set; }

    [XmlElementAttribute(IsNullable = true)]
    public Address homeAddress { get; set; }

    public string nationalInsuranceNumber { get; set; }
}



}


