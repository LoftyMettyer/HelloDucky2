Attribute VB_Name = "modScrMgr"
Option Explicit

Public Function LabelledControl(pLngControlType As ControlTypes) As Boolean
  '
  ' Return True if the given control requires labelling.
  '
  LabelledControl = False
  
  Select Case pLngControlType
    Case giCTRL_LABEL
      
    Case giCTRL_TEXTBOX
      LabelledControl = True
      
    Case giCTRL_COMBOBOX
      LabelledControl = True
      
    Case giCTRL_SPINNER
      LabelledControl = True
      
    Case giCTRL_CHECKBOX
    Case giCTRL_OPTIONGROUP
      
    Case giCTRL_OLE
      LabelledControl = True
      
Case giCTRL_PHOTO
  LabelledControl = True
      
    Case giCTRL_FRAME
    Case giCTRL_IMAGE
  End Select

End Function

