"""
Build Fuzzy Lookup Add-In for Excel 2024 (64-bit)
Creates a .xlam file with VBA modules and UserForm
"""
import win32com.client
import os
import sys
import time

ADDIN_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_PATH = os.path.join(ADDIN_DIR, "FuzzyLookup.xlam")

# VBA source files to import
VBA_MODULES = [
    os.path.join(ADDIN_DIR, "FuzzyAlgorithms.bas"),
    os.path.join(ADDIN_DIR, "FuzzyLookupEngine.bas"),
    os.path.join(ADDIN_DIR, "FuzzyLookupUI.bas"),
]

# UserForm code (created programmatically since .frm files are complex)
USERFORM_CODE = r'''
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FuzzyLookupForm
   Caption         =   "Fuzzy Lookup"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   OleObjectBlob   =   "FuzzyLookupForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
'''

USERFORM_VBA = r'''
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Fuzzy Lookup"
    Me.Width = 460
    Me.Height = 540

    BuildUI
    PopulateSheetLists
End Sub

Private Sub BuildUI()
    Dim ctrl As MSForms.Control
    Dim topPos As Single
    topPos = 10

    ' === Left Table Section ===
    AddLabel "lblLeftHeader", "LEFT TABLE (Lookup Values)", 10, topPos, 430, 18, True
    topPos = topPos + 24

    AddLabel "lblLeftSheet", "Sheet:", 10, topPos + 3, 60, 15, False
    AddComboBox "cboLeftSheet", 75, topPos, 180, 20
    AddLabel "lblLeftRange", "Range:", 265, topPos + 3, 45, 15, False
    AddRefEdit "txtLeftRange", 315, topPos, 120, 20
    topPos = topPos + 30

    AddLabel "lblLeftCol", "Match Column:", 10, topPos + 3, 90, 15, False
    AddComboBox "cboLeftCol", 105, topPos, 150, 20
    topPos = topPos + 36

    ' === Right Table Section ===
    AddLabel "lblRightHeader", "RIGHT TABLE (Match Against)", 10, topPos, 430, 18, True
    topPos = topPos + 24

    AddLabel "lblRightSheet", "Sheet:", 10, topPos + 3, 60, 15, False
    AddComboBox "cboRightSheet", 75, topPos, 180, 20
    AddLabel "lblRightRange", "Range:", 265, topPos + 3, 45, 15, False
    AddRefEdit "txtRightRange", 315, topPos, 120, 20
    topPos = topPos + 30

    AddLabel "lblRightCol", "Match Column:", 10, topPos + 3, 90, 15, False
    AddComboBox "cboRightCol", 105, topPos, 150, 20
    topPos = topPos + 36

    ' === Options Section ===
    AddLabel "lblOptionsHeader", "OPTIONS", 10, topPos, 430, 18, True
    topPos = topPos + 24

    AddLabel "lblThreshold", "Similarity Threshold:", 10, topPos + 3, 120, 15, False
    AddTextBox "txtThreshold", 135, topPos, 50, 20, "0.65"
    AddLabel "lblThresholdHint", "(0.0 - 1.0)", 190, topPos + 3, 70, 15, False
    topPos = topPos + 28

    AddLabel "lblMaxMatches", "Max Matches Per Row:", 10, topPos + 3, 120, 15, False
    AddTextBox "txtMaxMatches", 135, topPos, 50, 20, "1"
    topPos = topPos + 28

    AddLabel "lblMethod", "Algorithm:", 10, topPos + 3, 65, 15, False
    AddComboBox "cboMethod", 80, topPos, 175, 20
    Me.Controls("cboMethod").AddItem "Auto (Best Combined)"
    Me.Controls("cboMethod").AddItem "Levenshtein"
    Me.Controls("cboMethod").AddItem "Jaro-Winkler"
    Me.Controls("cboMethod").AddItem "Jaccard (Token)"
    Me.Controls("cboMethod").AddItem "Contains"
    Me.Controls("cboMethod").ListIndex = 0
    topPos = topPos + 28

    AddLabel "lblOutput", "Output Sheet Name:", 10, topPos + 3, 115, 15, False
    AddTextBox "txtOutput", 130, topPos, 125, 20, "Fuzzy_Results"
    topPos = topPos + 40

    ' === Buttons ===
    AddButton "btnGo", "Fuzzy Match!", 10, topPos, 200, 36
    AddButton "btnCancel", "Cancel", 220, topPos, 100, 36
    AddButton "btnHelp", "?", 400, topPos, 30, 36
End Sub

Private Sub AddLabel(nm As String, cap As String, l As Single, t As Single, w As Single, h As Single, bold As Boolean)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", nm)
    lbl.Caption = cap
    lbl.Left = l: lbl.Top = t: lbl.Width = w: lbl.Height = h
    lbl.Font.Bold = bold
    If bold Then lbl.Font.Size = 10
End Sub

Private Sub AddComboBox(nm As String, l As Single, t As Single, w As Single, h As Single)
    Dim cbo As MSForms.ComboBox
    Set cbo = Me.Controls.Add("Forms.ComboBox.1", nm)
    cbo.Left = l: cbo.Top = t: cbo.Width = w: cbo.Height = h
    cbo.Style = fmStyleDropDownList
End Sub

Private Sub AddTextBox(nm As String, l As Single, t As Single, w As Single, h As Single, defaultVal As String)
    Dim txt As MSForms.TextBox
    Set txt = Me.Controls.Add("Forms.TextBox.1", nm)
    txt.Left = l: txt.Top = t: txt.Width = w: txt.Height = h
    txt.Text = defaultVal
End Sub

Private Sub AddRefEdit(nm As String, l As Single, t As Single, w As Single, h As Single)
    ' Use TextBox as RefEdit replacement (RefEdit is buggy in 64-bit)
    Dim txt As MSForms.TextBox
    Set txt = Me.Controls.Add("Forms.TextBox.1", nm)
    txt.Left = l: txt.Top = t: txt.Width = w: txt.Height = h
    txt.Text = ""
End Sub

Private Sub AddButton(nm As String, cap As String, l As Single, t As Single, w As Single, h As Single)
    Dim btn As MSForms.CommandButton
    Set btn = Me.Controls.Add("Forms.CommandButton.1", nm)
    btn.Caption = cap
    btn.Left = l: btn.Top = t: btn.Width = w: btn.Height = h
    btn.Font.Size = 10
End Sub

Private Sub PopulateSheetLists()
    Dim ws As Worksheet
    Dim cboLeft As MSForms.ComboBox
    Dim cboRight As MSForms.ComboBox

    Set cboLeft = Me.Controls("cboLeftSheet")
    Set cboRight = Me.Controls("cboRightSheet")

    cboLeft.Clear
    cboRight.Clear

    For Each ws In ActiveWorkbook.Worksheets
        cboLeft.AddItem ws.Name
        cboRight.AddItem ws.Name
    Next ws

    If cboLeft.ListCount > 0 Then cboLeft.ListIndex = 0
    If cboRight.ListCount > 1 Then
        cboRight.ListIndex = 1
    ElseIf cboRight.ListCount > 0 Then
        cboRight.ListIndex = 0
    End If
End Sub

Private Sub cboLeftSheet_Change()
    PopulateColumns Me.Controls("cboLeftSheet"), Me.Controls("cboLeftCol"), Me.Controls("txtLeftRange")
End Sub

Private Sub cboRightSheet_Change()
    PopulateColumns Me.Controls("cboRightSheet"), Me.Controls("cboRightCol"), Me.Controls("txtRightRange")
End Sub

Private Sub PopulateColumns(cboSheet As MSForms.ComboBox, cboCol As MSForms.ComboBox, txtRange As MSForms.TextBox)
    Dim ws As Worksheet
    Dim usedRng As Range
    Dim col As Long

    cboCol.Clear

    If cboSheet.ListIndex < 0 Then Exit Sub

    Set ws = ActiveWorkbook.Worksheets(cboSheet.Value)
    Set usedRng = ws.UsedRange

    If usedRng Is Nothing Then Exit Sub

    ' Set default range
    txtRange.Text = usedRng.Address(False, False)

    ' Populate column headers
    For col = 1 To usedRng.Columns.Count
        Dim headerVal As String
        headerVal = CStr(usedRng.Cells(1, col).Value)
        If Len(headerVal) = 0 Then headerVal = "Column " & col
        cboCol.AddItem headerVal
    Next col

    If cboCol.ListCount > 0 Then cboCol.ListIndex = 0
End Sub

Private Sub btnGo_Click()
    ' Validate inputs
    If Me.Controls("cboLeftSheet").ListIndex < 0 Then
        MsgBox "Please select a Left Table sheet.", vbExclamation
        Exit Sub
    End If
    If Me.Controls("cboRightSheet").ListIndex < 0 Then
        MsgBox "Please select a Right Table sheet.", vbExclamation
        Exit Sub
    End If
    If Me.Controls("cboLeftCol").ListIndex < 0 Then
        MsgBox "Please select a Left Match Column.", vbExclamation
        Exit Sub
    End If
    If Me.Controls("cboRightCol").ListIndex < 0 Then
        MsgBox "Please select a Right Match Column.", vbExclamation
        Exit Sub
    End If

    Dim threshold As Double
    If Not IsNumeric(Me.Controls("txtThreshold").Text) Then
        MsgBox "Threshold must be a number between 0 and 1.", vbExclamation
        Exit Sub
    End If
    threshold = CDbl(Me.Controls("txtThreshold").Text)
    If threshold < 0 Or threshold > 1 Then
        MsgBox "Threshold must be between 0 and 1.", vbExclamation
        Exit Sub
    End If

    Dim maxMatches As Long
    If Not IsNumeric(Me.Controls("txtMaxMatches").Text) Then
        MsgBox "Max Matches must be a positive number.", vbExclamation
        Exit Sub
    End If
    maxMatches = CLng(Me.Controls("txtMaxMatches").Text)
    If maxMatches < 1 Then maxMatches = 1

    Dim outputName As String
    outputName = Trim$(Me.Controls("txtOutput").Text)
    If Len(outputName) = 0 Then outputName = "Fuzzy_Results"

    ' Parse method
    Dim method As String
    Select Case Me.Controls("cboMethod").ListIndex
        Case 0: method = "Auto"
        Case 1: method = "Levenshtein"
        Case 2: method = "JaroWinkler"
        Case 3: method = "Jaccard"
        Case 4: method = "Contains"
        Case Else: method = "Auto"
    End Select

    ' Build ranges
    Dim wsLeft As Worksheet, wsRight As Worksheet
    Set wsLeft = ActiveWorkbook.Worksheets(Me.Controls("cboLeftSheet").Value)
    Set wsRight = ActiveWorkbook.Worksheets(Me.Controls("cboRightSheet").Value)

    Dim leftRange As Range, rightRange As Range
    Dim leftRangeAddr As String, rightRangeAddr As String
    leftRangeAddr = Me.Controls("txtLeftRange").Text
    rightRangeAddr = Me.Controls("txtRightRange").Text

    On Error Resume Next
    If Len(leftRangeAddr) > 0 Then
        Set leftRange = wsLeft.Range(leftRangeAddr)
    Else
        Set leftRange = wsLeft.UsedRange
    End If
    If Len(rightRangeAddr) > 0 Then
        Set rightRange = wsRight.Range(rightRangeAddr)
    Else
        Set rightRange = wsRight.UsedRange
    End If
    On Error GoTo 0

    If leftRange Is Nothing Then
        MsgBox "Invalid Left Table range.", vbExclamation
        Exit Sub
    End If
    If rightRange Is Nothing Then
        MsgBox "Invalid Right Table range.", vbExclamation
        Exit Sub
    End If

    ' Build config
    Dim cfg As FuzzyConfig
    Set cfg.LeftRange = leftRange
    Set cfg.RightRange = rightRange

    ReDim cfg.LeftMatchCols(0 To 0)
    cfg.LeftMatchCols(0) = Me.Controls("cboLeftCol").ListIndex + 1

    ReDim cfg.RightMatchCols(0 To 0)
    cfg.RightMatchCols(0) = Me.Controls("cboRightCol").ListIndex + 1

    cfg.Threshold = threshold
    cfg.MaxMatches = maxMatches
    cfg.Method = method
    cfg.OutputSheet = outputName

    ' Hide form and run
    Me.Hide

    RunFuzzyLookup cfg

    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnHelp_Click()
    MsgBox "FUZZY LOOKUP ADD-IN" & vbCrLf & vbCrLf & _
           "1. Select the LEFT table (your lookup values)" & vbCrLf & _
           "2. Select the RIGHT table (to match against)" & vbCrLf & _
           "3. Choose which columns to compare" & vbCrLf & _
           "4. Set similarity threshold (0.65 = 65% match)" & vbCrLf & _
           "5. Click 'Fuzzy Match!'" & vbCrLf & vbCrLf & _
           "WORKSHEET FUNCTIONS:" & vbCrLf & _
           "  =FUZZYMATCH(A1, B1)  - Returns similarity 0-1" & vbCrLf & _
           "  =FUZZYMATCH(A1, B1, ""JaroWinkler"")" & vbCrLf & vbCrLf & _
           "  =FUZZYVLOOKUP(A1, B:D, 1, 2, 0.6)" & vbCrLf & _
           "    Like VLOOKUP but fuzzy!" & vbCrLf & _
           "    Args: lookup, range, match_col, return_col, threshold", _
           vbInformation, "Help"
End Sub
'''

# Ribbon XML for custom tab
RIBBON_XML = '''<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="FuzzyLookupTab" label="Fuzzy Lookup">
        <group id="FuzzyGroup" label="Fuzzy Matching">
          <button id="btnFuzzyLookup" label="Fuzzy Lookup"
                  size="large" onAction="ShowFuzzyLookupForm"
                  imageMso="QueryAppend"
                  screentip="Open Fuzzy Lookup"
                  supertip="Match similar text between two tables using fuzzy algorithms"/>
          <button id="btnAbout" label="About"
                  size="normal" onAction="ShowAbout"
                  imageMso="Info"
                  screentip="About Fuzzy Lookup Add-In"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>'''


def build_addin():
    print("=" * 60)
    print("Building Fuzzy Lookup Add-In for Excel 2024 (64-bit)")
    print("=" * 60)

    # Remove existing output
    if os.path.exists(OUTPUT_PATH):
        try:
            os.remove(OUTPUT_PATH)
            print(f"Removed existing: {OUTPUT_PATH}")
        except Exception as e:
            print(f"Warning: Could not remove existing file: {e}")

    excel = None
    wb = None
    try:
        print("\nStarting Excel...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        # Check if VBA access is enabled
        print("Creating workbook...")
        wb = excel.Workbooks.Add()

        # Import VBA modules
        vba_project = wb.VBProject

        for mod_path in VBA_MODULES:
            if os.path.exists(mod_path):
                mod_name = os.path.splitext(os.path.basename(mod_path))[0]
                print(f"  Importing module: {mod_name}")
                vba_project.VBComponents.Import(mod_path)
            else:
                print(f"  WARNING: Module not found: {mod_path}")

        # Create UserForm programmatically
        print("  Creating UserForm...")
        frm = vba_project.VBComponents.Add(3)  # 3 = vbext_ct_MSForm
        frm.Name = "FuzzyLookupForm"
        frm.Properties.Item("Caption").Value = "Fuzzy Lookup"
        frm.Properties.Item("Width").Value = 460
        frm.Properties.Item("Height").Value = 540

        # Add the UserForm code
        frm.CodeModule.DeleteLines(1, frm.CodeModule.CountOfLines)
        # Remove the first line "Option Explicit" that VBA adds automatically
        code_lines = USERFORM_VBA.strip().split('\n')
        code_text = '\n'.join(code_lines)
        frm.CodeModule.AddFromString(code_text)

        # Add ThisWorkbook auto-open code
        print("  Adding auto-open handler...")
        wb_module = vba_project.VBComponents("ThisWorkbook")
        wb_code = (
            "Private Sub Workbook_Open()\n"
            "    ' Add menu item to Ribbon via custom function\n"
            "    AddFuzzyLookupMenu\n"
            "End Sub\n"
        )

        # Add a menu-based fallback (since custom ribbon XML requires .xlam editing post-save)
        menu_module = vba_project.VBComponents.Add(1)  # Standard module
        menu_module.Name = "MenuSetup"
        menu_code = (
            "Option Explicit\n\n"
            "Public Sub AddFuzzyLookupMenu()\n"
            "    Dim cmdBar As Object\n"
            "    Dim ctrl As Object\n"
            "    \n"
            "    On Error Resume Next\n"
            "    ' Remove existing menu item\n"
            "    Application.CommandBars(\"Worksheet Menu Bar\").Controls(\"Fuzzy Lookup\").Delete\n"
            "    On Error GoTo 0\n"
            "    \n"
            "    ' Add to menu bar\n"
            "    Set cmdBar = Application.CommandBars(\"Worksheet Menu Bar\")\n"
            "    Set ctrl = cmdBar.Controls.Add(Type:=10) ' msoControlPopup\n"
            "    ctrl.Caption = \"Fuzzy Lookup\"\n"
            "    \n"
            "    ' Add sub-items\n"
            "    Dim btn As Object\n"
            "    Set btn = ctrl.Controls.Add(Type:=1) ' msoControlButton\n"
            "    btn.Caption = \"Open Fuzzy Lookup...\"\n"
            "    btn.OnAction = \"ShowFuzzyLookupForm\"\n"
            "    btn.FaceId = 1849\n"
            "    \n"
            "    Set btn = ctrl.Controls.Add(Type:=1)\n"
            "    btn.Caption = \"About Fuzzy Lookup\"\n"
            "    btn.OnAction = \"ShowAbout\"\n"
            "    btn.FaceId = 487\n"
            "End Sub\n\n"
            "Public Sub RemoveFuzzyLookupMenu()\n"
            "    On Error Resume Next\n"
            "    Application.CommandBars(\"Worksheet Menu Bar\").Controls(\"Fuzzy Lookup\").Delete\n"
            "    On Error GoTo 0\n"
            "End Sub\n"
        )
        menu_module.CodeModule.AddFromString(menu_code)

        # Update ThisWorkbook with both open and close handlers
        wb_module.CodeModule.DeleteLines(1, wb_module.CodeModule.CountOfLines)
        wb_full_code = (
            "Private Sub Workbook_Open()\n"
            "    AddFuzzyLookupMenu\n"
            "End Sub\n\n"
            "Private Sub Workbook_BeforeClose(Cancel As Boolean)\n"
            "    RemoveFuzzyLookupMenu\n"
            "End Sub\n"
        )
        wb_module.CodeModule.AddFromString(wb_full_code)

        # Save as .xlam (AddIn format = 55)
        print(f"\nSaving as: {OUTPUT_PATH}")
        wb.SaveAs(OUTPUT_PATH, FileFormat=55)  # 55 = xlOpenXMLAddIn (.xlam)
        print("Save successful!")

        wb.Close(False)
        wb = None

        print("\n" + "=" * 60)
        print("BUILD COMPLETE!")
        print("=" * 60)
        print(f"\nAdd-in saved to: {OUTPUT_PATH}")
        print("\nTo install:")
        print("  1. Open Excel 2024")
        print("  2. File > Options > Add-ins")
        print("  3. At bottom: Manage 'Excel Add-ins' > Go...")
        print("  4. Click 'Browse...' and select:")
        print(f"     {OUTPUT_PATH}")
        print("  5. Check the box and click OK")
        print("\nA 'Fuzzy Lookup' menu will appear in the menu bar.")
        print("You can also use these formulas in any cell:")
        print("  =FUZZYMATCH(A1, B1)")
        print("  =FUZZYVLOOKUP(A1, Sheet2!A:C, 1, 2, 0.6)")

    except Exception as e:
        print(f"\nERROR: {e}")
        if "programmatic access" in str(e).lower() or "1004" in str(e):
            print("\n>>> VBA Trust Settings Need Update <<<")
            print("In Excel: File > Options > Trust Center > Trust Center Settings")
            print("> Macro Settings > Check 'Trust access to the VBA project object model'")
            print("\nThen run this script again.")
        raise
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except:
                pass


if __name__ == "__main__":
    build_addin()
