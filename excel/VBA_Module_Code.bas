' ===========================================
' VBA Code for CartaTransformer.xlsm
' ===========================================
' 
' SETUP INSTRUCTIONS:
' 1. Open Excel, create new workbook
' 2. Save as "CartaTransformer.xlsm" (macro-enabled)
' 3. Press Alt+F11 to open VBA Editor
' 4. In VBA Editor: Tools > References > Check "xlwings"
' 5. Insert > Module
' 6. Paste ALL code below into the module
' 7. Close VBA Editor
' 8. Add a button: Insert > Shapes > Rectangle
' 9. Right-click button > Assign Macro > Select "TransformCarta"
' 10. Save workbook
'
' ===========================================

Sub TransformCarta()
    ' Main entry point - called by button click
    ' Runs the Python transformation script via xlwings
    
    On Error GoTo ErrorHandler
    
    ' Run the Python script
    RunPython "import carta_to_cap_table; carta_to_cap_table.main()"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error running transformation: " & Err.Description & vbCrLf & vbCrLf & _
           "Make sure:" & vbCrLf & _
           "1. Python is installed" & vbCrLf & _
           "2. xlwings is installed (pip install xlwings)" & vbCrLf & _
           "3. xlwings add-in is installed (xlwings addin install)" & vbCrLf & _
           "4. carta_to_cap_table.py is in the same folder as this workbook", _
           vbCritical, "Carta Transformer Error"
End Sub

Sub TransformCartaStandalone()
    ' Alternative: Run without xlwings using Shell
    ' Use this if xlwings isn't working
    
    Dim pythonPath As String
    Dim scriptPath As String
    Dim cartaFile As String
    Dim templateFile As String
    Dim cmd As String
    
    ' Get the folder where this workbook is saved
    Dim workbookFolder As String
    workbookFolder = ThisWorkbook.Path
    
    ' Prompt for Carta file
    cartaFile = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx), *.xlsx", _
        Title:="Select Carta Export")
    
    If cartaFile = "False" Then
        MsgBox "Cancelled", vbInformation
        Exit Sub
    End If
    
    ' Set paths (adjust pythonPath for your system)
    pythonPath = "python"  ' Or full path like "C:\Python39\python.exe"
    scriptPath = workbookFolder & "\carta_to_cap_table.py"
    templateFile = workbookFolder & "\Cap_Table_Template.xlsx"
    
    ' Check files exist
    If Dir(scriptPath) = "" Then
        MsgBox "Script not found: " & scriptPath, vbCritical
        Exit Sub
    End If
    
    If Dir(templateFile) = "" Then
        MsgBox "Template not found: " & templateFile, vbCritical
        Exit Sub
    End If
    
    ' Build command
    cmd = pythonPath & " """ & scriptPath & """ """ & cartaFile & """ """ & templateFile & """"
    
    ' Run Python script
    Dim result As Double
    result = Shell(cmd, vbNormalFocus)
    
    If result = 0 Then
        MsgBox "Error running Python script", vbCritical
    Else
        MsgBox "Transformation started. Check the folder for output file.", vbInformation
    End If
    
End Sub

Sub CheckSetup()
    ' Diagnostic function to verify setup
    
    Dim msg As String
    msg = "Setup Check:" & vbCrLf & vbCrLf
    
    ' Check workbook location
    msg = msg & "Workbook folder: " & ThisWorkbook.Path & vbCrLf
    
    ' Check for Python script
    If Dir(ThisWorkbook.Path & "\carta_to_cap_table.py") <> "" Then
        msg = msg & "✓ Python script found" & vbCrLf
    Else
        msg = msg & "✗ Python script NOT found" & vbCrLf
    End If
    
    ' Check for template
    If Dir(ThisWorkbook.Path & "\Cap_Table_Template.xlsx") <> "" Then
        msg = msg & "✓ Template found" & vbCrLf
    Else
        msg = msg & "✗ Template NOT found" & vbCrLf
    End If
    
    ' Check xlwings reference
    On Error Resume Next
    Dim xlwingsRef As Boolean
    xlwingsRef = False
    Dim ref As Object
    For Each ref In ThisWorkbook.VBProject.References
        If ref.Name = "xlwings" Then
            xlwingsRef = True
            Exit For
        End If
    Next
    On Error GoTo 0
    
    If xlwingsRef Then
        msg = msg & "✓ xlwings reference found" & vbCrLf
    Else
        msg = msg & "✗ xlwings reference NOT found (add via Tools > References)" & vbCrLf
    End If
    
    MsgBox msg, vbInformation, "Setup Check"
    
End Sub
