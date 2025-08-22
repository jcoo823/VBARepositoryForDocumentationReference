# Comprehensive VBA Reference Guide for AI

This document serves as a comprehensive reference for Visual Basic for Applications (VBA) programming, designed specifically for AI systems to reference when providing VBA instruction and guidance.

## Table of Contents

1. [VBA Fundamentals](#vba-fundamentals)
2. [Data Types and Variables](#data-types-and-variables)
3. [Operators and Expressions](#operators-and-expressions)
4. [Control Structures](#control-structures)
5. [Procedures and Functions](#procedures-and-functions)
6. [Object-Oriented Programming](#object-oriented-programming)
7. [Error Handling](#error-handling)
8. [File Operations](#file-operations)
9. [Excel VBA Specifics](#excel-vba-specifics)
10. [Word VBA Specifics](#word-vba-specifics)
11. [Access VBA Specifics](#access-vba-specifics)
12. [Best Practices](#best-practices)
13. [Common Patterns and Solutions](#common-patterns-and-solutions)
14. [Debugging Techniques](#debugging-techniques)
15. [Performance Optimization](#performance-optimization)
16. [Advanced Techniques](#advanced-techniques)
17. [Tips and Tricks](#tips-and-tricks)

---

## VBA Fundamentals

### Basic Structure
```vba
' This is a comment
Sub ProcedureName()
    ' Code goes here
End Sub

Function FunctionName() As DataType
    ' Code goes here
    FunctionName = returnValue
End Function
```

### Module Types
- **Standard Modules**: .bas files, contain general procedures
- **Class Modules**: .cls files, define custom objects
- **UserForm Modules**: .frm files, contain form code
- **Document Modules**: ThisWorkbook, Sheet modules in Excel

### VBA Environment Access
- **Alt + F11**: Open VBA Editor
- **F5**: Run current procedure
- **F8**: Step through code
- **F9**: Toggle breakpoint
- **Ctrl + G**: Immediate window
- **Ctrl + R**: Project Explorer

### Option Statements
```vba
Option Explicit          ' Force variable declaration
Option Compare Text      ' Case-insensitive string comparison
Option Compare Binary    ' Case-sensitive string comparison
Option Base 1           ' Arrays start at index 1
```

---

## Data Types and Variables

### Primitive Data Types
```vba
Dim myByte As Byte              ' 0 to 255
Dim myBoolean As Boolean        ' True or False
Dim myInteger As Integer        ' -32,768 to 32,767
Dim myLong As Long             ' -2,147,483,648 to 2,147,483,647
Dim mySingle As Single         ' Single-precision floating-point
Dim myDouble As Double         ' Double-precision floating-point
Dim myCurrency As Currency     ' Currency data type
Dim myDate As Date             ' Date and time
Dim myString As String         ' Text data
Dim myVariant As Variant       ' Any data type
```

### String Data Types
```vba
Dim fixedString As String * 10    ' Fixed-length string
Dim dynamicString As String       ' Variable-length string
```

### Variable Declaration Tips
```vba
' Multiple declarations
Dim x As Integer, y As Integer, z As Integer

' Initialize variables
Dim count As Integer: count = 0

' Constants
Const PI As Double = 3.14159265359
Const MAX_ROWS As Long = 1048576

' Public variables (accessible across modules)
Public globalVar As String

' Private variables (module-level)
Private moduleVar As Integer

' Static variables (retain value between calls)
Static persistentVar As Integer
```

### Arrays
```vba
' Static arrays
Dim arr(1 To 10) As Integer
Dim matrix(1 To 5, 1 To 5) As String

' Dynamic arrays
Dim dynamicArr() As Integer
ReDim dynamicArr(1 To 100)
ReDim Preserve dynamicArr(1 To 150)  ' Preserve existing data

' Array functions
UBound(arr)     ' Upper bound
LBound(arr)     ' Lower bound
```

### User-Defined Types (Structures)
```vba
Type Employee
    Name As String
    ID As Long
    Salary As Currency
    HireDate As Date
End Type

Dim emp As Employee
emp.Name = "John Doe"
emp.ID = 12345
```

### Enumerations
```vba
Enum DaysOfWeek
    Sunday = 1
    Monday = 2
    Tuesday = 3
    Wednesday = 4
    Thursday = 5
    Friday = 6
    Saturday = 7
End Enum

Dim today As DaysOfWeek
today = DaysOfWeek.Monday
```

---

## Operators and Expressions

### Arithmetic Operators
```vba
+ Addition
- Subtraction
* Multiplication
/ Division (floating-point)
\ Integer division
Mod Modulus (remainder)
^ Exponentiation
```

### Comparison Operators
```vba
= Equal to
<> Not equal to
< Less than
> Greater than
<= Less than or equal to
>= Greater than or equal to
Like Pattern matching
Is Object comparison
```

### Logical Operators
```vba
And     ' Logical AND
Or      ' Logical OR
Not     ' Logical NOT
Xor     ' Exclusive OR
Eqv     ' Equivalence
Imp     ' Implication
```

### String Operators
```vba
& Concatenation
+ Concatenation (use & instead)

' Examples
Dim fullName As String
fullName = firstName & " " & lastName
```

### Operator Precedence (highest to lowest)
1. Parentheses ()
2. Exponentiation ^
3. Unary minus -
4. Multiplication *, Division /
5. Integer division \
6. Modulus Mod
7. Addition +, Subtraction -
8. String concatenation &
9. Comparison operators
10. Logical operators

---

## Control Structures

### If...Then...Else Statements
```vba
' Simple If
If condition Then
    ' Code
End If

' If...Else
If condition Then
    ' Code if true
Else
    ' Code if false
End If

' If...ElseIf...Else
If condition1 Then
    ' Code for condition1
ElseIf condition2 Then
    ' Code for condition2
Else
    ' Default code
End If

' Single-line If
If x > 0 Then y = x

' IIf function (immediate if)
result = IIf(x > 0, "Positive", "Not positive")
```

### Select Case Statements
```vba
Select Case variable
    Case value1
        ' Code for value1
    Case value2, value3
        ' Code for value2 or value3
    Case value4 To value6
        ' Code for range value4 to value6
    Case Is > value7
        ' Code for values greater than value7
    Case Else
        ' Default code
End Select

' Example with strings
Select Case UCase(userInput)
    Case "YES", "Y"
        result = True
    Case "NO", "N"
        result = False
    Case Else
        MsgBox "Invalid input"
End Select
```

### Loop Structures

#### For...Next Loops
```vba
' Basic For loop
For i = 1 To 10
    ' Code
Next i

' Step increment
For i = 0 To 100 Step 5
    ' Code
Next i

' Countdown
For i = 10 To 1 Step -1
    ' Code
Next i

' For Each loop
For Each cell In Range("A1:A10")
    ' Code
Next cell

' For Each with collections
For Each ws In ThisWorkbook.Worksheets
    ' Code
Next ws
```

#### Do...Loop Structures
```vba
' Do While (test at beginning)
Do While condition
    ' Code
Loop

' Do Until (test at beginning)
Do Until condition
    ' Code
Loop

' Do...Loop While (test at end)
Do
    ' Code
Loop While condition

' Do...Loop Until (test at end)
Do
    ' Code
Loop Until condition
```

#### While...Wend Loop
```vba
While condition
    ' Code
Wend
```

### Loop Control Statements
```vba
' Exit loop early
For i = 1 To 100
    If someCondition Then Exit For
Next i

' Continue to next iteration
For i = 1 To 100
    If skipCondition Then GoTo NextIteration
    ' Code
NextIteration:
Next i
```

---

## Procedures and Functions

### Subroutines (Procedures)
```vba
' Basic subroutine
Sub ProcedureName()
    ' Code
End Sub

' Subroutine with parameters
Sub ProcedureWithParams(param1 As String, param2 As Integer)
    ' Code
End Sub

' Optional parameters
Sub OptionalParams(required As String, Optional optional1 As Integer = 0)
    ' Code
End Sub

' ParamArray for variable arguments
Sub VariableArgs(ParamArray args() As Variant)
    Dim i As Integer
    For i = 0 To UBound(args)
        Debug.Print args(i)
    Next i
End Sub
```

### Functions
```vba
' Basic function
Function FunctionName() As ReturnType
    ' Code
    FunctionName = returnValue
End Function

' Function with parameters
Function AddNumbers(a As Double, b As Double) As Double
    AddNumbers = a + b
End Function

' Function returning array
Function GetArray() As Variant
    Dim arr(1 To 3) As String
    arr(1) = "One"
    arr(2) = "Two"
    arr(3) = "Three"
    GetArray = arr
End Function
```

### Parameter Passing
```vba
' By Reference (default) - modifies original variable
Sub ByRefExample(ByRef x As Integer)
    x = x + 1
End Sub

' By Value - copies value, doesn't modify original
Sub ByValExample(ByVal x As Integer)
    x = x + 1  ' Original remains unchanged
End Sub
```

### Scope and Lifetime
```vba
' Private (module level)
Private Sub PrivateProcedure()
End Sub

' Public (accessible from other modules)
Public Sub PublicProcedure()
End Sub

' Friend (accessible within same project only)
Friend Sub FriendProcedure()
End Sub
```

### Recursive Functions
```vba
Function Factorial(n As Long) As Long
    If n <= 1 Then
        Factorial = 1
    Else
        Factorial = n * Factorial(n - 1)
    End If
End Function
```

---

## Object-Oriented Programming

### Classes and Objects
```vba
' Class Module: clsPerson
Private m_Name As String
Private m_Age As Integer

' Property procedures
Public Property Let Name(value As String)
    m_Name = value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Age(value As Integer)
    If value >= 0 Then m_Age = value
End Property

Public Property Get Age() As Integer
    Age = m_Age
End Property

' Method
Public Sub Celebrate()
    m_Age = m_Age + 1
    MsgBox m_Name & " is now " & m_Age & " years old!"
End Sub

' Class initialization
Private Sub Class_Initialize()
    m_Name = ""
    m_Age = 0
End Sub

' Class termination
Private Sub Class_Terminate()
    ' Cleanup code
End Sub
```

### Using Classes
```vba
Sub UsePersonClass()
    Dim person As clsPerson
    Set person = New clsPerson
    
    person.Name = "John Doe"
    person.Age = 30
    person.Celebrate()
    
    Set person = Nothing  ' Clean up object reference
End Sub
```

### Collections
```vba
' Built-in Collection object
Dim myCollection As Collection
Set myCollection = New Collection

' Add items
myCollection.Add "Item1", "Key1"
myCollection.Add "Item2", "Key2"

' Access items
Dim item As Variant
item = myCollection("Key1")  ' By key
item = myCollection(1)       ' By index

' Remove items
myCollection.Remove "Key1"
myCollection.Remove 1

' Iterate through collection
For Each item In myCollection
    Debug.Print item
Next item
```

### Events
```vba
' Class with events
Public Event StatusChanged(newStatus As String)

Public Sub ChangeStatus(newStatus As String)
    RaiseEvent StatusChanged(newStatus)
End Sub

' In another module
Dim WithEvents myObject As clsMyClass

Private Sub myObject_StatusChanged(newStatus As String)
    MsgBox "Status changed to: " & newStatus
End Sub
```

---

## Error Handling

### On Error Statements
```vba
' Resume next line after error
On Error Resume Next

' Go to error handler
On Error GoTo ErrorHandler

' Turn off error handling
On Error GoTo 0

' Example with error handler
Sub ErrorHandlingExample()
    On Error GoTo ErrorHandler
    
    ' Code that might cause error
    Dim result As Double
    result = 10 / 0  ' This will cause error
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume Next  ' or Resume, or Exit Sub
End Sub
```

### Err Object Properties
```vba
Err.Number         ' Error number
Err.Description    ' Error description
Err.Source        ' Source of error
Err.HelpFile      ' Help file path
Err.HelpContext   ' Help context ID

' Methods
Err.Clear         ' Clear error
Err.Raise number:=1001, Description:="Custom error"
```

### Custom Error Handling
```vba
Sub CustomErrorHandling()
    On Error GoTo ErrorHandler
    
    ' Validate input
    If value < 0 Then
        Err.Raise Number:=1001, Description:="Value cannot be negative"
    End If
    
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case 1001
            MsgBox "Validation Error: " & Err.Description
        Case 11  ' Division by zero
            MsgBox "Cannot divide by zero"
        Case Else
            MsgBox "Unexpected error: " & Err.Description
    End Select
    Resume Next
End Sub
```

### Error Handling Best Practices
```vba
' Always clean up resources
Sub ProperCleanup()
    Dim obj As Object
    On Error GoTo ErrorHandler
    
    Set obj = CreateObject("Excel.Application")
    ' Work with object
    
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "Error occurred: " & Err.Description
    
CleanUp:
    If Not obj Is Nothing Then
        obj.Quit
        Set obj = Nothing
    End If
End Sub
```

---

## File Operations

### File System Functions
```vba
' File existence and attributes
Dir("C:\path\file.txt")           ' Returns filename if exists
FileExists = (Dir("filename") <> "")
FileLen("filename")               ' File size in bytes
FileDateTime("filename")          ' Last modified date
GetAttr("filename")              ' File attributes

' File attributes constants
vbNormal, vbReadOnly, vbHidden, vbSystem, vbDirectory, vbArchive
```

### File Input/Output

#### Sequential File Access
```vba
' Write to text file
Sub WriteToFile()
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open "C:\temp\output.txt" For Output As #fileNum
    Print #fileNum, "Hello World"
    Print #fileNum, "Line 2"
    Close #fileNum
End Sub

' Read from text file
Sub ReadFromFile()
    Dim fileNum As Integer
    Dim textLine As String
    
    fileNum = FreeFile
    Open "C:\temp\input.txt" For Input As #fileNum
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, textLine
        Debug.Print textLine
    Loop
    
    Close #fileNum
End Sub

' Append to file
Sub AppendToFile()
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open "C:\temp\log.txt" For Append As #fileNum
    Print #fileNum, Now & ": Log entry"
    Close #fileNum
End Sub
```

#### Binary File Access
```vba
Sub BinaryFileExample()
    Dim fileNum As Integer
    Dim data As String
    
    fileNum = FreeFile
    Open "C:\temp\binary.dat" For Binary As #fileNum
    
    ' Write binary data
    data = "Binary Data"
    Put #fileNum, 1, data
    
    ' Read binary data
    data = Space(11)  ' Allocate space
    Get #fileNum, 1, data
    
    Close #fileNum
End Sub
```

### FileSystemObject (Advanced File Operations)
```vba
Sub FileSystemObjectExample()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' File operations
    If fso.FileExists("C:\temp\test.txt") Then
        fso.DeleteFile "C:\temp\test.txt"
    End If
    
    fso.CopyFile "source.txt", "destination.txt"
    fso.MoveFile "old.txt", "new.txt"
    
    ' Folder operations
    If Not fso.FolderExists("C:\temp\newfolder") Then
        fso.CreateFolder "C:\temp\newfolder"
    End If
    
    ' Get file/folder info
    Dim file As Object
    Set file = fso.GetFile("C:\temp\test.txt")
    Debug.Print file.Size
    Debug.Print file.DateLastModified
    
    ' Text file operations
    Dim textFile As Object
    Set textFile = fso.CreateTextFile("C:\temp\output.txt", True)
    textFile.WriteLine "Hello World"
    textFile.Close
    
    Set textFile = fso.OpenTextFile("C:\temp\input.txt", 1)  ' ForReading
    Do While Not textFile.AtEndOfStream
        Debug.Print textFile.ReadLine
    Loop
    textFile.Close
End Sub
```

### File Dialog Boxes
```vba
' File Open Dialog
Function GetOpenFileName() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogOpen)
    
    With fd
        .Title = "Select File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls"
        .Filters.Add "All Files", "*.*"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            GetOpenFileName = .SelectedItems(1)
        End If
    End With
End Function

' File Save Dialog
Function GetSaveFileName() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    
    With fd
        .Title = "Save File"
        .InitialFileName = "MyFile.xlsx"
        
        If .Show = -1 Then
            GetSaveFileName = .SelectedItems(1)
        End If
    End With
End Function
```

---

## Excel VBA Specifics

### Workbook and Worksheet Objects
```vba
' Reference workbooks
ThisWorkbook                    ' The workbook containing the code
ActiveWorkbook                  ' Currently active workbook
Workbooks("WorkbookName.xlsx")  ' Specific workbook by name
Workbooks(1)                   ' First open workbook

' Reference worksheets
ActiveSheet                     ' Currently active worksheet
ThisWorkbook.Sheets("Sheet1")   ' Specific sheet by name
ThisWorkbook.Sheets(1)         ' First sheet
ThisWorkbook.Worksheets("Sheet1") ' Worksheet object specifically
```

### Range Objects and Operations
```vba
' Range references
Range("A1")                    ' Single cell
Range("A1:B10")               ' Range of cells
Range("A1,C1,E1")             ' Non-contiguous range
Cells(1, 1)                   ' Cell by row/column numbers
Rows(1)                       ' Entire row
Columns(1)                    ' Entire column
Range("A:A")                  ' Entire column A
Range("1:1")                  ' Entire row 1

' Range properties
Range("A1").Value             ' Cell value
Range("A1").Formula           ' Cell formula
Range("A1").Text              ' Displayed text
Range("A1").Comment.Text      ' Cell comment
Range("A1").Address           ' Cell address
Range("A1").Row               ' Row number
Range("A1").Column            ' Column number

' Range methods
Range("A1:B10").Select        ' Select range
Range("A1:B10").Copy          ' Copy range
Range("A1:B10").Clear         ' Clear contents
Range("A1:B10").Delete        ' Delete range
Range("A1:B10").Insert        ' Insert cells
```

### Working with Data
```vba
' Read data from range
Sub ReadRangeData()
    Dim dataRange As Range
    Dim cell As Range
    
    Set dataRange = Range("A1:C10")
    
    For Each cell In dataRange
        Debug.Print cell.Value
    Next cell
    
    ' Or read as array for better performance
    Dim dataArray As Variant
    dataArray = dataRange.Value
    
    Dim i As Long, j As Long
    For i = 1 To UBound(dataArray, 1)
        For j = 1 To UBound(dataArray, 2)
            Debug.Print dataArray(i, j)
        Next j
    Next i
End Sub

' Write data to range
Sub WriteRangeData()
    ' Write single value
    Range("A1").Value = "Hello"
    
    ' Write array
    Dim dataArray(1 To 3, 1 To 2) As String
    dataArray(1, 1) = "A1"
    dataArray(1, 2) = "B1"
    dataArray(2, 1) = "A2"
    dataArray(2, 2) = "B2"
    dataArray(3, 1) = "A3"
    dataArray(3, 2) = "B3"
    
    Range("A1:B3").Value = dataArray
End Sub
```

### Find and Replace
```vba
Sub FindAndReplace()
    Dim foundCell As Range
    
    ' Find first occurrence
    Set foundCell = Range("A:A").Find("SearchText")
    If Not foundCell Is Nothing Then
        foundCell.Value = "ReplacementText"
    End If
    
    ' Find all occurrences
    Dim firstAddress As String
    Set foundCell = Range("A:A").Find("SearchText")
    
    If Not foundCell Is Nothing Then
        firstAddress = foundCell.Address
        Do
            foundCell.Value = "ReplacementText"
            Set foundCell = Range("A:A").FindNext(foundCell)
        Loop While Not foundCell Is Nothing And foundCell.Address <> firstAddress
    End If
    
    ' Replace all at once
    Range("A:A").Replace What:="SearchText", Replacement:="ReplacementText"
End Sub
```

### Formatting
```vba
Sub FormatCells()
    With Range("A1:B10")
        ' Font formatting
        .Font.Name = "Arial"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Italic = False
        .Font.Color = RGB(255, 0, 0)  ' Red
        
        ' Border formatting
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        
        ' Interior formatting
        .Interior.Color = RGB(255, 255, 0)  ' Yellow
        .Interior.Pattern = xlSolid
        
        ' Alignment
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        
        ' Number format
        .NumberFormat = "#,##0.00"  ' Number with commas and 2 decimals
        .NumberFormat = "mm/dd/yyyy"  ' Date format
        .NumberFormat = "0.00%"     ' Percentage
    End With
End Sub
```

### Charts
```vba
Sub CreateChart()
    Dim chartRange As Range
    Dim newChart As Chart
    
    Set chartRange = Range("A1:B10")
    
    ' Create chart
    Set newChart = Charts.Add
    With newChart
        .SetSourceData Source:=chartRange
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "My Chart"
        .HasLegend = True
    End With
    
    ' Or create embedded chart
    Dim embeddedChart As ChartObject
    Set embeddedChart = ActiveSheet.ChartObjects.Add(100, 100, 400, 300)
    With embeddedChart.Chart
        .SetSourceData Source:=chartRange
        .ChartType = xlLine
    End With
End Sub
```

### PivotTables
```vba
Sub CreatePivotTable()
    Dim sourceRange As Range
    Dim pivotSheet As Worksheet
    Dim pivotTable As PivotTable
    
    Set sourceRange = Range("A1:D100")
    Set pivotSheet = Worksheets.Add
    
    Set pivotTable = pivotSheet.PivotTables.Add( _
        PivotCache:=ThisWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=sourceRange), _
        TableDestination:=pivotSheet.Range("A1"))
    
    With pivotTable
        .PivotFields("Category").Orientation = xlRowField
        .PivotFields("Product").Orientation = xlColumnField
        .PivotFields("Sales").Orientation = xlDataField
    End With
End Sub
```

### Events in Excel
```vba
' Workbook events (in ThisWorkbook module)
Private Sub Workbook_Open()
    MsgBox "Workbook opened!"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If MsgBox("Save before closing?", vbYesNo) = vbYes Then
        Me.Save
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Code before saving
End Sub

' Worksheet events (in Sheet module)
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$A$1" Then
        MsgBox "Cell A1 was changed!"
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Code when selection changes
End Sub

Private Sub Worksheet_Calculate()
    ' Code when worksheet recalculates
End Sub
```

### Application Object
```vba
Sub ApplicationSettings()
    ' Turn off screen updating for performance
    Application.ScreenUpdating = False
    
    ' Turn off automatic calculation
    Application.Calculation = xlCalculationManual
    
    ' Turn off alerts
    Application.DisplayAlerts = False
    
    ' Your code here
    
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Sub
```

---

## Word VBA Specifics

### Document Object Model
```vba
' Reference documents
ActiveDocument              ' Currently active document
ThisDocument               ' Document containing the code
Documents("DocName.docx")   ' Specific document by name
Documents(1)               ' First open document

' Create new document
Dim newDoc As Document
Set newDoc = Documents.Add
```

### Working with Text
```vba
Sub WorkWithText()
    ' Insert text at cursor
    Selection.TypeText "Hello World"
    
    ' Insert text at end of document
    ActiveDocument.Range.InsertAfter "Text at end"
    
    ' Insert text at beginning
    ActiveDocument.Range.InsertBefore "Text at beginning"
    
    ' Replace text
    With Selection.Find
        .Text = "FindThis"
        .Replacement.Text = "ReplaceWith"
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Select all text
    ActiveDocument.Range.Select
    
    ' Get word count
    Dim wordCount As Long
    wordCount = ActiveDocument.Words.Count
End Sub
```

### Formatting in Word
```vba
Sub FormatText()
    With Selection.Font
        .Name = "Arial"
        .Size = 12
        .Bold = True
        .Italic = False
        .Color = RGB(255, 0, 0)
    End With
    
    With Selection.ParagraphFormat
        .Alignment = wdAlignParagraphCenter
        .LineSpacing = LinesToPoints(2)  ' Double spacing
        .SpaceAfter = 6
    End With
End Sub
```

### Tables in Word
```vba
Sub CreateTable()
    Dim newTable As Table
    
    ' Create table
    Set newTable = ActiveDocument.Tables.Add( _
        Range:=Selection.Range, _
        NumRows:=3, _
        NumColumns:=4)
    
    ' Add content to table
    newTable.Cell(1, 1).Range.Text = "Header 1"
    newTable.Cell(1, 2).Range.Text = "Header 2"
    
    ' Format table
    newTable.AutoFormat Format:=wdTableFormatClassic2
    
    ' Add row
    newTable.Rows.Add
    
    ' Delete row
    newTable.Rows(newTable.Rows.Count).Delete
End Sub
```

### Headers and Footers
```vba
Sub WorkWithHeaders()
    ' Add header
    With ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary)
        .Range.Text = "Document Header"
        .Range.Font.Bold = True
    End With
    
    ' Add footer with page numbers
    With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary)
        .PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter
    End With
End Sub
```

---

## Access VBA Specifics

### Database Operations
```vba
' Current database reference
CurrentDb

' Open external database
Dim db As Database
Set db = OpenDatabase("C:\path\to\database.accdb")

' Execute SQL
CurrentDb.Execute "UPDATE Customers SET City = 'New York' WHERE ID = 1"

' Open recordset
Dim rs As Recordset
Set rs = CurrentDb.OpenRecordset("SELECT * FROM Customers")

' Navigate recordset
rs.MoveFirst
rs.MoveLast
rs.MoveNext
rs.MovePrevious

' Read data
While Not rs.EOF
    Debug.Print rs("CustomerName")
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
```

### Forms and Controls
```vba
' Form events (in form module)
Private Sub Form_Load()
    ' Code when form loads
End Sub

Private Sub Form_Current()
    ' Code when record changes
End Sub

' Control events
Private Sub CommandButton1_Click()
    ' Button click event
End Sub

Private Sub TextBox1_BeforeUpdate(Cancel As Integer)
    ' Validate input before updating
    If Len(Me.TextBox1.Value) < 3 Then
        MsgBox "Text must be at least 3 characters"
        Cancel = True
    End If
End Sub

' Reference form controls
Me.TextBox1.Value = "New Value"
Forms("FormName").Controls("ControlName").Value = "Value"
```

### Reports
```vba
' Open report
DoCmd.OpenReport "ReportName", acViewPreview

' Print report
DoCmd.OpenReport "ReportName", acViewNormal

' Export report
DoCmd.OutputTo acOutputReport, "ReportName", acFormatPDF, "C:\output.pdf"
```

---

## Best Practices

### Code Organization
```vba
' Use descriptive names
Dim customerCount As Integer  ' Good
Dim cc As Integer            ' Bad

' Use consistent naming conventions
Dim strCustomerName As String    ' Hungarian notation
Dim customerName As String       ' Preferred

' Group related constants
Const MIN_AGE As Integer = 18
Const MAX_AGE As Integer = 65
Const DEFAULT_SALARY As Currency = 50000
```

### Performance Tips
```vba
' Turn off screen updating for better performance
Application.ScreenUpdating = False
' Your code here
Application.ScreenUpdating = True

' Use arrays instead of cell-by-cell operations
Dim dataArray As Variant
dataArray = Range("A1:Z1000").Value  ' Read all at once
' Process array
Range("A1:Z1000").Value = dataArray  ' Write all at once

' Avoid using Select and Activate
' Bad
Range("A1").Select
Selection.Value = "Hello"

' Good
Range("A1").Value = "Hello"

' Use With statements for multiple property access
With Range("A1")
    .Value = "Hello"
    .Font.Bold = True
    .Interior.Color = RGB(255, 255, 0)
End With
```

### Error Prevention
```vba
' Always use Option Explicit
Option Explicit

' Validate parameters
Function Divide(numerator As Double, denominator As Double) As Double
    If denominator = 0 Then
        Err.Raise 11, , "Division by zero"
    End If
    Divide = numerator / denominator
End Function

' Check object references
If Not obj Is Nothing Then
    ' Use object
    Set obj = Nothing
End If
```

### Memory Management
```vba
' Set object variables to Nothing when done
Dim ws As Worksheet
Set ws = ActiveSheet
' Use ws
Set ws = Nothing

' Use local variables when possible
' Avoid global variables unless necessary

' Clean up external object references
Dim xlApp As Object
Set xlApp = CreateObject("Excel.Application")
' Use Excel application
xlApp.Quit
Set xlApp = Nothing
```

---

## Common Patterns and Solutions

### Data Validation
```vba
Function IsValidEmail(email As String) As Boolean
    Dim atPos As Integer, dotPos As Integer
    atPos = InStr(email, "@")
    dotPos = InStrRev(email, ".")
    
    IsValidEmail = (atPos > 1) And (dotPos > atPos + 1) And (dotPos < Len(email))
End Function

Function IsNumeric(value As Variant) As Boolean
    IsNumeric = VBA.IsNumeric(value) And Not IsEmpty(value)
End Function

Function IsValidDate(dateString As String) As Boolean
    On Error GoTo ErrorHandler
    Dim testDate As Date
    testDate = CDate(dateString)
    IsValidDate = True
    Exit Function
    
ErrorHandler:
    IsValidDate = False
End Function
```

### String Manipulation
```vba
' Remove extra spaces
Function CleanText(text As String) As String
    CleanText = Trim(Replace(Replace(text, vbCrLf, " "), "  ", " "))
End Function

' Title case conversion
Function ProperCase(text As String) As String
    Dim words As Variant
    Dim i As Integer
    
    words = Split(LCase(text), " ")
    For i = 0 To UBound(words)
        If Len(words(i)) > 0 Then
            words(i) = UCase(Left(words(i), 1)) & Mid(words(i), 2)
        End If
    Next i
    
    ProperCase = Join(words, " ")
End Function

' Extract numbers from text
Function ExtractNumbers(text As String) As String
    Dim i As Integer
    Dim result As String
    
    For i = 1 To Len(text)
        If IsNumeric(Mid(text, i, 1)) Then
            result = result & Mid(text, i, 1)
        End If
    Next i
    
    ExtractNumbers = result
End Function
```

### Array Operations
```vba
' Find value in array
Function FindInArray(arr As Variant, searchValue As Variant) As Integer
    Dim i As Integer
    FindInArray = -1  ' Not found
    
    For i = LBound(arr) To UBound(arr)
        If arr(i) = searchValue Then
            FindInArray = i
            Exit Function
        End If
    Next i
End Function

' Remove duplicates from array
Function RemoveDuplicates(arr As Variant) As Variant
    Dim dict As Object
    Dim i As Integer
    Dim result() As Variant
    Dim counter As Integer
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Add unique values to dictionary
    For i = LBound(arr) To UBound(arr)
        If Not dict.Exists(arr(i)) Then
            dict.Add arr(i), ""
        End If
    Next i
    
    ' Convert back to array
    ReDim result(0 To dict.Count - 1)
    counter = 0
    For Each key In dict.Keys
        result(counter) = key
        counter = counter + 1
    Next
    
    RemoveDuplicates = result
End Function
```

### Working with Dictionaries
```vba
Sub DictionaryExample()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Add items
    dict.Add "Name", "John Doe"
    dict.Add "Age", 30
    dict.Add "City", "New York"
    
    ' Check if key exists
    If dict.Exists("Name") Then
        Debug.Print dict("Name")
    End If
    
    ' Iterate through dictionary
    Dim key As Variant
    For Each key In dict.Keys
        Debug.Print key & ": " & dict(key)
    Next key
    
    ' Remove item
    dict.Remove "Age"
    
    ' Clear dictionary
    dict.RemoveAll
End Sub
```

### Progress Indicators
```vba
Sub ShowProgress()
    Dim i As Long
    Dim totalItems As Long
    totalItems = 1000
    
    Application.StatusBar = "Processing..."
    
    For i = 1 To totalItems
        ' Your processing code here
        
        ' Update progress every 50 items
        If i Mod 50 = 0 Then
            Application.StatusBar = "Processing... " & Format(i / totalItems, "0%")
            DoEvents  ' Allow Windows to update
        End If
    Next i
    
    Application.StatusBar = "Complete!"
    Application.StatusBar = False  ' Reset status bar
End Sub
```

---

## Debugging Techniques

### Debug Object Methods
```vba
' Print to Immediate window
Debug.Print "Variable value: " & variableName

' Print with formatting
Debug.Print Format(Now, "yyyy-mm-dd hh:mm:ss") & ": Started processing"

' Conditional debugging
If debugMode Then Debug.Print "Debug info: " & debugInfo
```

### Breakpoints and Stepping
```vba
' Set breakpoints with F9
' Step through code with F8
' Step over procedures with Shift+F8
' Step out of procedures with Ctrl+Shift+F8

Sub DebuggingExample()
    Dim i As Integer
    
    For i = 1 To 10
        Debug.Print i  ' Set breakpoint here
        ' Step through to watch i change
    Next i
End Sub
```

### Watch Window
```vba
' Add variables to Watch window to monitor values
' Right-click variable and "Add Watch"
' Set break conditions in Watch window
```

### Assertion Testing
```vba
Sub TestFunction()
    Dim result As Integer
    result = MyFunction(5, 10)
    
    Debug.Assert result = 15  ' Will break if false
    
    If result <> 15 Then
        Debug.Print "Test failed: Expected 15, got " & result
    Else
        Debug.Print "Test passed"
    End If
End Sub
```

### Logging
```vba
Sub WriteToLog(message As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open "C:\temp\debug.log" For Append As #fileNum
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:mm:ss") & ": " & message
    Close #fileNum
End Sub

' Usage
Call WriteToLog("Function started")
Call WriteToLog("Variable value: " & variableValue)
```

---

## Performance Optimization

### General Performance Tips
```vba
Sub OptimizedCode()
    ' 1. Turn off screen updating
    Application.ScreenUpdating = False
    
    ' 2. Turn off automatic calculation
    Application.Calculation = xlCalculationManual
    
    ' 3. Turn off events
    Application.EnableEvents = False
    
    ' 4. Use arrays for bulk operations
    Dim dataArray As Variant
    dataArray = Range("A1:Z1000").Value
    
    ' Process array instead of individual cells
    Dim i As Long, j As Long
    For i = 1 To UBound(dataArray, 1)
        For j = 1 To UBound(dataArray, 2)
            ' Process dataArray(i, j)
        Next j
    Next i
    
    ' Write array back to range
    Range("A1:Z1000").Value = dataArray
    
    ' 5. Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
```

### Loop Optimization
```vba
' Avoid using .End(xlUp) in loops - cache the value
Dim lastRow As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

Dim i As Long
For i = 1 To lastRow
    ' Process row i
Next i

' Use For Each for collections
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    ' Process worksheet
Next ws

' Use arrays for large data sets
Dim data As Variant
data = Range("A1:A10000").Value

Dim i As Long
For i = 1 To UBound(data)
    ' Process data(i, 1)
Next i
```

### Memory Optimization
```vba
' Use appropriate data types
Dim counter As Long        ' Not Integer if > 32,767
Dim percentage As Single   ' Not Double if precision not needed

' Release object references
Dim obj As Object
Set obj = CreateObject("Excel.Application")
' Use object
obj.Quit
Set obj = Nothing

' Use local variables instead of global when possible
' Avoid ReDim Preserve if possible - it's slow
```

### String Operations
```vba
' Use Join instead of concatenation in loops
Dim parts(1 To 1000) As String
Dim i As Long

' Fill array
For i = 1 To 1000
    parts(i) = "Part " & i
Next i

' Join all at once (fast)
Dim result As String
result = Join(parts, ", ")

' Instead of this (slow):
' result = ""
' For i = 1 To 1000
'     result = result & "Part " & i & ", "
' Next i
```

---

## Advanced Techniques

### API Declarations
```vba
' 32-bit and 64-bit compatible API declarations
#If VBA7 Then
    Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongPtr
    Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    Declare Function GetTickCount Lib "kernel32" () As Long
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

Sub UseAPI()
    Dim tickCount As LongPtr
    tickCount = GetTickCount()
    Debug.Print "System uptime: " & tickCount & " milliseconds"
End Sub
```

### Regular Expressions
```vba
Function RegExTest(text As String, pattern As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Pattern = pattern
        .IgnoreCase = True
        .Global = True
    End With
    
    RegExTest = regEx.Test(text)
End Function

Function RegExReplace(text As String, pattern As String, replacement As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Pattern = pattern
        .IgnoreCase = True
        .Global = True
    End With
    
    RegExReplace = regEx.Replace(text, replacement)
End Function

' Extract email addresses
Function ExtractEmails(text As String) As Variant
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim emails() As String
    Dim i As Integer
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"
    regEx.Global = True
    
    Set matches = regEx.Execute(text)
    
    If matches.Count > 0 Then
        ReDim emails(0 To matches.Count - 1)
        For Each match In matches
            emails(i) = match.Value
            i = i + 1
        Next match
    End If
    
    ExtractEmails = emails
End Function
```

### Dynamic Programming
```vba
' Memoization example - Fibonacci with caching
Private fibCache As Object

Function FibonacciMemo(n As Long) As Long
    If fibCache Is Nothing Then
        Set fibCache = CreateObject("Scripting.Dictionary")
    End If
    
    If fibCache.Exists(n) Then
        FibonacciMemo = fibCache(n)
        Exit Function
    End If
    
    If n <= 1 Then
        FibonacciMemo = n
    Else
        FibonacciMemo = FibonacciMemo(n - 1) + FibonacciMemo(n - 2)
    End If
    
    fibCache.Add n, FibonacciMemo
End Function
```

### Working with XML
```vba
Sub ParseXML()
    Dim xmlDoc As Object
    Dim nodes As Object
    Dim node As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.async = False
    
    If xmlDoc.Load("C:\path\to\file.xml") Then
        Set nodes = xmlDoc.SelectNodes("//item")
        
        For Each node In nodes
            Debug.Print node.Text
        Next node
    Else
        Debug.Print "Failed to load XML: " & xmlDoc.parseError.reason
    End If
End Sub
```

### Working with JSON
```vba
Function ParseJSON(jsonString As String) As Object
    Dim scriptControl As Object
    Set scriptControl = CreateObject("ScriptControl")
    scriptControl.Language = "JScript"
    
    Set ParseJSON = scriptControl.Eval("(" + jsonString + ")")
End Function

Sub JSONExample()
    Dim json As String
    Dim obj As Object
    
    json = "{""name"":""John"",""age"":30,""city"":""New York""}"
    Set obj = ParseJSON(json)
    
    Debug.Print obj.name
    Debug.Print obj.age
    Debug.Print obj.city
End Sub
```

---

## Tips and Tricks

### Useful Built-in Functions
```vba
' Date/Time functions
Now                         ' Current date and time
Date                        ' Current date
Time                        ' Current time
DateAdd("m", 1, Date)      ' Add 1 month to current date
DateDiff("d", startDate, endDate)  ' Days between dates
Format(Now, "yyyy-mm-dd")   ' Format date
Weekday(Date)              ' Day of week (1=Sunday)

' String functions
Left("Hello", 3)           ' "Hel"
Right("Hello", 3)          ' "llo"
Mid("Hello", 2, 3)         ' "ell"
InStr("Hello", "ll")       ' Position of "ll" (3)
InStrRev("Hello", "l")     ' Last position of "l" (4)
Replace("Hello", "l", "x") ' "Hexxo"
Split("a,b,c", ",")        ' Array: ("a", "b", "c")
Join(array, ",")           ' "a,b,c"
StrReverse("Hello")        ' "olleH"

' Conversion functions
CInt(3.7)                  ' 4 (rounds)
CLng("12345")              ' Convert to Long
CStr(123)                  ' "123"
CBool(1)                   ' True
CDate("1/1/2023")          ' Date value

' Math functions
Abs(-5)                    ' 5
Sqr(16)                    ' 4
Round(3.7, 0)              ' 4
Int(3.7)                   ' 3 (truncates)
Rnd()                      ' Random number 0-1
```

### Keyboard Shortcuts in VBA Editor
```
F5              - Run procedure
F8              - Step into
Shift+F8        - Step over
Ctrl+Shift+F8   - Step out
F9              - Toggle breakpoint
Ctrl+Shift+F9   - Clear all breakpoints
Ctrl+G          - Immediate window
Ctrl+R          - Project Explorer
F7              - View code
Shift+F7        - View object
Ctrl+H          - Find and Replace
F2              - Object Browser
Ctrl+Space      - List Properties/Methods
Ctrl+Shift+I    - Parameter Info
Ctrl+I          - Quick Info
```

### Conditional Compilation
```vba
#Const DEBUG_MODE = True

Sub ConditionalCode()
    #If DEBUG_MODE Then
        Debug.Print "Debug mode is on"
    #Else
        ' Production code
    #End If
    
    #If VBA7 Then
        ' Code for VBA 7 (Office 2010+)
    #Else
        ' Code for earlier versions
    #End If
End Sub
```

### Speed Testing
```vba
Sub SpeedTest()
    Dim startTime As Single
    Dim endTime As Single
    
    startTime = Timer
    
    ' Code to test
    Dim i As Long
    For i = 1 To 1000000
        ' Some operation
    Next i
    
    endTime = Timer
    Debug.Print "Execution time: " & Format(endTime - startTime, "0.00") & " seconds"
End Sub
```

### Environment Information
```vba
Sub SystemInfo()
    Debug.Print "Excel version: " & Application.Version
    Debug.Print "Operating system: " & Application.OperatingSystem
    Debug.Print "User name: " & Application.UserName
    Debug.Print "Computer name: " & Environ("COMPUTERNAME")
    Debug.Print "Current directory: " & CurDir
    Debug.Print "Excel path: " & Application.Path
    
    #If VBA7 Then
        Debug.Print "VBA7 - 64-bit capable"
    #Else
        Debug.Print "VBA6 - 32-bit only"
    #End If
    
    #If Win64 Then
        Debug.Print "Running 64-bit Office"
    #Else
        Debug.Print "Running 32-bit Office"
    #End If
End Sub
```

### Memory Usage
```vba
' Check available memory (requires API declaration)
#If VBA7 Then
    Declare PtrSafe Function GlobalMemoryStatus Lib "kernel32" (lpBuffer As Any) As Long
#Else
    Declare Function GlobalMemoryStatus Lib "kernel32" (lpBuffer As Any) As Long
#End If

Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Sub CheckMemory()
    Dim ms As MEMORYSTATUS
    ms.dwLength = Len(ms)
    GlobalMemoryStatus ms
    
    Debug.Print "Memory load: " & ms.dwMemoryLoad & "%"
    Debug.Print "Available physical memory: " & ms.dwAvailPhys \ 1024 \ 1024 & " MB"
End Sub
```

### Registry Operations
```vba
Function GetRegistryValue(keyPath As String, valueName As String) As String
    On Error GoTo ErrorHandler
    GetRegistryValue = CreateObject("WScript.Shell").RegRead(keyPath & "\" & valueName)
    Exit Function
    
ErrorHandler:
    GetRegistryValue = ""
End Function

Sub SetRegistryValue(keyPath As String, valueName As String, value As String)
    On Error GoTo ErrorHandler
    CreateObject("WScript.Shell").RegWrite keyPath & "\" & valueName, value
    Exit Sub
    
ErrorHandler:
    MsgBox "Could not write to registry"
End Sub
```

### Working with Command Line
```vba
Sub RunCommandLine()
    Dim command As String
    command = "dir C:\ > C:\temp\dirlist.txt"
    
    ' Run command and wait
    Shell command, vbHide
    
    ' Run command without waiting
    Shell command, vbNormalFocus
End Sub

Function GetCommandLineOutput(command As String) As String
    Dim wsh As Object
    Dim exec As Object
    
    Set wsh = CreateObject("WScript.Shell")
    Set exec = wsh.Exec(command)
    
    Do While exec.Status = 0
        DoEvents
    Loop
    
    GetCommandLineOutput = exec.StdOut.ReadAll
End Function
```

### Email Integration
```vba
Sub SendEmail()
    Dim outlookApp As Object
    Dim mail As Object
    
    Set outlookApp = CreateObject("Outlook.Application")
    Set mail = outlookApp.CreateItem(0)  ' olMailItem
    
    With mail
        .To = "recipient@example.com"
        .CC = "cc@example.com"
        .Subject = "Test Email"
        .Body = "This is a test email from VBA"
        .Attachments.Add "C:\path\to\file.xlsx"
        .Send  ' or .Display to show before sending
    End With
    
    Set mail = Nothing
    Set outlookApp = Nothing
End Sub
```

### Progress Bars and User Interaction
```vba
' Simple progress using status bar
Sub ProgressBar()
    Dim i As Long
    For i = 1 To 100
        Application.StatusBar = "Progress: " & i & "%"
        ' Your code here
        DoEvents
    Next i
    Application.StatusBar = False
End Sub

' Input validation
Function GetValidInput(prompt As String, validationType As String) As Variant
    Dim userInput As String
    Dim isValid As Boolean
    
    Do While Not isValid
        userInput = InputBox(prompt)
        
        Select Case LCase(validationType)
            Case "number"
                isValid = IsNumeric(userInput)
            Case "date"
                isValid = IsDate(userInput)
            Case "email"
                isValid = InStr(userInput, "@") > 0 And InStr(userInput, ".") > 0
            Case Else
                isValid = True
        End Select
        
        If Not isValid Then
            MsgBox "Invalid input. Please try again."
        End If
    Loop
    
    GetValidInput = userInput
End Function
```

---

## Common VBA Errors and Solutions

### Runtime Errors
```vba
' Error 9: Subscript out of range
' Solution: Check array bounds and object references
If i >= LBound(arr) And i <= UBound(arr) Then
    value = arr(i)
End If

' Error 13: Type mismatch
' Solution: Validate data types before assignment
If IsNumeric(userInput) Then
    number = CDbl(userInput)
End If

' Error 91: Object variable or With block variable not set
' Solution: Always check object references
If Not obj Is Nothing Then
    obj.Method
End If

' Error 1004: Application-defined or object-defined error
' Solution: Check Excel object references and ranges
If Not Intersect(targetRange, usedRange) Is Nothing Then
    ' Range is valid
End If
```

### Memory Errors
```vba
' Out of memory errors
' Solution: Clean up object references and use arrays efficiently
Set obj = Nothing
Erase largeArray
Application.CutCopyMode = False  ' Clear clipboard
```

### Performance Issues
```vba
' Slow code solutions
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

' Use arrays instead of cell-by-cell operations
' Avoid selecting ranges unnecessarily
' Use With statements for multiple property access
```

---

This comprehensive VBA reference guide contains thousands of tips, tricks, code examples, and best practices that AI systems can reference when providing VBA programming guidance. The content is organized into logical sections covering everything from basic syntax to advanced techniques, making it easy to find relevant information for any VBA programming task.

The guide includes practical examples, common patterns, error handling strategies, performance optimization techniques, and application-specific guidance for Excel, Word, and Access. It serves as a complete reference for VBA development across all Microsoft Office applications.