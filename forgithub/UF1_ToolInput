Attribute VB_Name = "UF1_ToolInput"
Attribute VB_Base = "0{FB675304-B994-4841-A202-C80CBC64A683}{69E51FB6-098F-4AFB-81B2-D26BE14AB02E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private m_ClosedByOK As Boolean ' New property to track how the form was closed
Private m_AnalysisType As String ' New property to store the selected analysis type

Public Property Get ClosedByOK() As Boolean
    ClosedByOK = m_ClosedByOK
End Property

Public Property Get AnalysisType() As String
    AnalysisType = m_AnalysisType
End Property

Private Sub UserForm_Initialize()
    ' Initialize the Option Buttons to "Same Direction" by default
    OptionButtonSameDirection.Value = True
    m_AnalysisType = "Same Direction"
End Sub

Private Sub OKButton_Click()
    ' Set the ClosedByOK property to True when the OK button is clicked
    m_ClosedByOK = True
    
    ' Get the selected Time Buffer value
    Dim TimeBufferValue As Double
    TimeBufferValue = CDbl(TextBoxTimeBuffer.Value)
    
    ' Determine the selected analysis type
    If OptionButtonSameDirection.Value Then
        m_AnalysisType = "Same Direction"
    ElseIf OptionButtonHeadOn.Value Then
        m_AnalysisType = "Head On"
    ElseIf OptionButtonBoth.Value Then
        m_AnalysisType = "Both"
    End If
    
    ' Call the main code with the selected Time Buffer and Analysis Type
    FlagMatchingRowsDynamicTimeBuffer m_AnalysisType, OptionButtonSameDirection.Value, TimeBufferValue
    
    ' Close the user form
    Unload Me
End Sub

Private Sub CancelButton_Click()
    ' Set the ClosedByOK property to False when the Cancel button is clicked
    m_ClosedByOK = False
    
    ' Close the user form
    Unload Me
End Sub

