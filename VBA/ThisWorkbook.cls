VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public CategorySheet As Worksheet   'This will be how this workbook references the category worksheet.
Dim ChangedSheets As Dictionary     'This will contain an entry for each worksheet and a boolean as to whether or not the worksheet was changed.
Dim Saving As Boolean               'This will be set to True while we're saving.
Dim Init As Boolean                 'This will be set to True while Workbook_Open() is running.


Private Sub Workbook_Open()
    'Called whever this workbook opens.
    
    Dim Sheet As Worksheet  'This will be used in the For Each loop below.
    
    Set ChangedSheets = New Dictionary
    
    Application.ScreenUpdating = False
    Init = True
    For Each Sheet In ThisWorkbook.Worksheets
    
        Sheet.Activate
        ChangedSheets.Add Sheet.Name, False
        
    Next
    Init = False
    Application.ScreenUpdating = True
    
    ThisWorkbook.Worksheets(1).Activate

End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    'Called whenever the workbook is about to save.
    
    Dim Key As Variant  'This will be used in the For Each loop below.
    
    Saving = True
    
    Call CategorizeAllSheets(False)
    
    For Each Key In ChangedSheets.Keys
    
        ChangedSheets(Key) = False
    
    Next
    
    Saving = False
    
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    'Called right before this workbook attempts to close.
    
    Cancel = Saving
    If Saving Then MsgBox "Please wait until the automatic categorization has finished.", vbExclamation, "Not Ready To Close"
    
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    'Called whenever a worksheet within this workbook is activated.
    
    Dim Sheet As Worksheet  'This will be set to the activated worksheet.
    
    Set Sheet = Sh
    
    If Not Init Then Sheet.Cells(1, 1).Activate
    
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    'When called will mark the Target's worksheet as having been changed, therefore we should categorize it.

    ChangedSheets.Item(Target.Worksheet.Name) = True

End Sub

Private Sub CategorizeAllSheets(ByRef ForceAll As Boolean)
    'When called will loop through each sheet with a transactions table and attempt to categorize it.
    
    Dim Sheet As Worksheet          'This will be used in the For Each loop below.
    Dim Table As ListObject         'This will be used in the nexted For Each loop below.
    Dim DescriptionHeader As Range  'This will be set to the "Description" header within the found table.
    Dim CategoryHeader As Range     'This will be set to the "Description" header within the found table.
    Dim Categorized As Boolean      'This will be set to True if we categorized anything.
    
    Application.ScreenUpdating = False
    
    For Each Sheet In ThisWorkbook.Worksheets
    
        If ChangedSheets(Sheet.Name) Or ForceAll Then
        
            For Each Table In Sheet.ListObjects
            
                Set DescriptionHeader = Table.HeaderRowRange.Find("Description")
                Set CategoryHeader = Table.HeaderRowRange.Find("Category")
            
                If Not DescriptionHeader Is Nothing And Not CategoryHeader Is Nothing Then
                
                    Application.StatusBar = "Categorizing " & Sheet.Name & "..."
            
                    Categorized = True
                    Call Categorize(Table.ListColumns(DescriptionHeader.Column).DataBodyRange, False)
            
                End If
        
            Next
            
            Call RefreshPivotTables(Sheet)
        
        End If
    
    Next
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    If Categorized Then MsgBox "All data has been categorized.", vbInformation, "Pre-Save Categorization Complete"
    
End Sub

Public Sub CategorizeSelection()
Attribute CategorizeSelection.VB_Description = "Will categorize the selected Description cells if they aren't already."
Attribute CategorizeSelection.VB_ProcData.VB_Invoke_Func = "O\n14"
    'When called will call Catgegorize() with the current selection as the argument.
    
    Call Categorize(Selection, True)
    
End Sub

Private Sub Categorize(ByRef DescriptionRange As Range, ByRef DisplaySummary As Boolean)
    'When called will lookup the categorization for the DescriptionRange via the "Categories" worksheet.
    
    Dim Cel As Range                'This will be set to each cell within the For Each loop below.
    Dim Cel2 As Range               'This will be set to each cell within the For Each loop below.
    Dim Count As Integer            'This will track the total number of changes we've made.
    Dim Total As Integer            'This will track the number of selected cells.
    Dim CategoryColumn As Integer   'This will be set to the column index of the "Category" column within the table.

    If DescriptionRange Is Nothing Then Exit Sub
    
    Total = DescriptionRange.Count
    Set Cel = DescriptionRange.Worksheet.ListObjects(1).HeaderRowRange.Find("Category")
    
    If Cel Is Nothing Then
        MsgBox "Unable to locate the Category column within this table. Categorization aborted.", vbCritical, "Missing Category Column"
        Exit Sub
    End If
    
    CategoryColumn = Cel.Column
    
    
    For Each Cel In DescriptionRange
    
        If Intersect(Cel, Cel.Worksheet.UsedRange) Is Nothing Then Exit For
        
        If Cel.Value <> Empty Then
            For Each Cel2 In ThisWorkbook.CategorySheet.UsedRange
            
                If Cel2.Column <> 1 And Cel2.Value <> Empty Then
                
                        If InStr(UCase(Cel.Value), UCase(Cel2.Value)) Then
                        
                            Cel.Worksheet.Cells(Cel.Row, CategoryColumn).Value = Cel2.Worksheet.Cells(Cel2.Row, 1).Value
                            Count = Count + 1
                        
                        End If
                    
                End If
                
                DoEvents
            Next
            
        End If
        
    Next
    
    If DisplaySummary Then MsgBox "Categorized " & Count & "/" & Total & " entries.", vbInformation, "Selection Categorization Complete"
    
End Sub

Private Sub RefreshPivotTables(ByRef Sheet As Worksheet)
    'When called will referesh every table and chart on the given worksheet.
    
    Dim PTable As PivotTable
    Dim Chart As ChartObject
    
    
    For Each PTable In Sheet.PivotTables
    
        PTable.RefreshTable
    
    Next
    
End Sub

Public Sub RefreshSheet()
    'When called will call RefreshPivotTables() using ActiveSheet.

    Call RefreshPivotTables(ActiveSheet)

End Sub

Public Sub ReCategorizeAll()
    'When called will call CategorizeAllSheets() with ForceAll = True.
    
    Dim Key As Variant  'This will be used in the For Each loop below.
    
    
    Call CategorizeAllSheets(True)
    
    For Each Key In ChangedSheets.Keys
    
        ChangedSheets(Key) = False
    
    Next
    
End Sub

Public Sub MakeSelectionNegative()
Attribute MakeSelectionNegative.VB_Description = "Will prefix every selected cell with ""-"" if the cell contains only a number."
Attribute MakeSelectionNegative.VB_ProcData.VB_Invoke_Func = "I\n14"
    'When called will find each cell containing a number within the selected range and prefix it with a "-".
    
    Dim Cel As Range        'This will be set to each cell within the For Each loop below.
    Dim Count As Integer    'This will track the total number of changes we've made.

    
    For Each Cel In Selection
    
        If Intersect(Cel, Application.ActiveSheet.UsedRange) Is Nothing Then Exit For
        
        If Cel.Value <> Empty And Application.WorksheetFunction.IsNumber(Cel.Value) And Not InStr(Cel.Value, "-") Then
        
            Cel.Value = "-" & Cel.Value
            Count = Count + 1
            
        End If
        
        DoEvents
    Next
    
    MsgBox "Inverted " & Count & " values.", vbInformation, "Selection Inversion Complete"
    
End Sub

Public Sub UpperCase()
Attribute UpperCase.VB_Description = "Makes all text within selection upper case."
Attribute UpperCase.VB_ProcData.VB_Invoke_Func = "U\n14"
    'When called will make all text within the selection upper case.
    
    Dim Cel As Range        'This will be used in the For Each loop below.
    Dim Count As Integer    'This will track the total number of changes we've made.
    
    For Each Cel In Selection
    
        If Cel.Value <> Empty And Not Application.WorksheetFunction.IsNumber(Cel.Value) Then
        
            Cel.Value = UCase(Cel.Value)
            Count = Count + 1
            
        End If
        
        DoEvents
    Next
    
    MsgBox "Capitalized " & Count & " values.", vbInformation, "Selection Capitalization Complete"
    
End Sub

Public Sub AutoSizeSelectedChart()
Attribute AutoSizeSelectedChart.VB_Description = "AutoSizes the currently selected Chart."
Attribute AutoSizeSelectedChart.VB_ProcData.VB_Invoke_Func = "R\n14"
    'When called will automatically size the selected chart based on its contents.
    
    With ActiveChart
    
        If .HasTitle Then
    
            .ChartTitle.Top = 0
            .ChartTitle.Left = ((0.66 * .ChartArea.Width) - .ChartTitle.Width) / 2
    
        End If
    

        If .HasTitle Then .PlotArea.Top = 0 + .ChartTitle.Height Else .PlotArea.Top = 0
    
        .PlotArea.Left = 0
        .PlotArea.Width = 0.66 * .ChartArea.Width
        .PlotArea.Height = .ChartArea.Height

    
        If .HasLegend Then
    
            .Legend.Top = 0
            .Legend.Left = .ChartArea.Width - (0.33 * .ChartArea.Width)
            .Legend.Width = 0.33 * .ChartArea.Width
            .Legend.Height = .ChartArea.Height
    
        End If
    
    End With
    
End Sub
