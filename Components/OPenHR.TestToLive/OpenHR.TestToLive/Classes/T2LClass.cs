
using OpenHRTestToLive;
using System;
using System.Collections.Generic;

[Serializable()]
public partial class T2LClass
{

    public List<ASRSysWorkflows> AllWorkflows;
    public List<ASRSysWorkflowLinks> AllLinks;
    public List<ASRSysWorkflowElement> AllElements;
    public List<ASRSysWorkflowElementColumn> AllColumns;
    public List<ASRSysWorkflowElementValidation> AllValidations;
    public List<ASRSysExpression> AllExpressions;
    public List<ASRSysExprComponent> AllComponents;
    public List<ASRSysWorkflowElementItem> AllItems;
    public List<ASRSysWorkflowElementItemValue> AllValues;

    public T2LClass()
    {
        AllWorkflows = new List<ASRSysWorkflows>();
        AllLinks = new List<ASRSysWorkflowLinks>();
        AllElements = new List<ASRSysWorkflowElement>();
        AllColumns = new List<ASRSysWorkflowElementColumn>();
        AllValidations = new List<ASRSysWorkflowElementValidation>();
        AllExpressions = new List<ASRSysExpression>();
        AllComponents = new List<ASRSysExprComponent>();
        AllItems = new List<ASRSysWorkflowElementItem>();
        AllValues = new List<ASRSysWorkflowElementItemValue>();
    }
}