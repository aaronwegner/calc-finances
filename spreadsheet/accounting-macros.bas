REM  *****  BASIC  *****

' TODO list, dash means done
'
' -  1) total all columns of master sheet
' -  2) accurately label all master card transactions pos and neg
' -  3) find state transactions and refunds and total Gross and Tax
' -  4) match PayPal 1099-K accurately
' -  5) resize and delete various columns in order to present to accountant, test both for totaling and filling
' -  6) use search descriptors combined with trim to match strings
' -  7) write the hide/show/empty (0/optimal/default) routine
' -  8) use globals for all strings repeated
' -  9) test both sheets from scratch and ensure matches
' - 10) include third spreadsheet for expenses not in PayPal history or MasterCard, see notes

Option Explicit

Sub Main

End Sub

' these variables are global to this module only

' the following are applicable primarily to totaling
'
' special name for master totals sheet
Dim total_sheet_str As String
' column title for subcategories
Dim subtotals_col_str As String
' do not add to any total subcategories matching this string
Dim dnt_str As String
' this is the sum of all the subtotals that run down a row of the master sheet
Dim subtotals_totals_str As String
' this is the sum of all entries of a number column 
Dim all_totals_str As String

' the following are applicable primarily to filling
'
' special name for master fillers sheet
Dim filler_sheet_str As String
' column in the fillers sheet containing the type of the data in a row
Dim fill_type_col_str As String
' column in the fillers sheet instructing whether or not to overwrite existing data
Dim overwrite_col_str As String
' this row contains column width data
Dim col_width_str As String
' this row contains the part of the function before the cell string
Dim before_str As String
' this row contains the part of the function after the cell string
Dim after_str As String

' global error message
Dim error_message_str As String

' globals common to many routines
'
'
Dim max_col_index As Long
Dim max_row_index As Long

Sub InitAccountingGlobals
  total_sheet_str = "Totals"
  subtotals_col_str = "Tax Category"
  dnt_str = "DNT *"
  subtotals_totals_str = "Tax Categories Totals"
  all_totals_str = "All Categories Totals"

  filler_sheet_str = "Fillers"
  fill_type_col_str = "Fillers Type"
  overwrite_col_str = "Overwrite"
  col_width_str = "column width"
  before_str = "before cell"
  after_str = "after cell"

  ' max columns 1024
  ' max rows 65536 
  ' we'll probably never have more than 100 columns of data
  'max_col_index = 1024 - 1
  max_col_index = 99
  ' set rows to max regardless, the more sales the better
  max_row_index = 65536 - 1
End Sub

' http://www.openoffice.org/api/docs/common/ref/com/sun/star/table/TableColumn.html 
' this is for column width, colWidth = -1 means optimal width
Type ColumnDisplayType
  colTitle As String
  colWidth As Long
End Type

' this comes in handy for adding members to an array, but doesn't work for Types
Function VarArrayPreserveAdd(a() As Variant, increase as Long) as Long
  Dim l As Long

  l = UBound(a()) + increase
  ReDim Preserve a(l)
  VarArrayPreserveAdd = l
End Function

Sub VarArrayReDim(a() As Variant, ub as Long)
  ReDim a(ub)
End Sub

Sub ColumnDisplayAddToSet(cdt() As ColumnDisplayType, colTitle As String, colWidth As Long)
  Dim l As Long

  l = UBound(cdt()) + 1
  ReDim Preserve cdt(l)

  cdt(l).colTitle = colTitle
  cdt(l).colWidth = colWidth
End Sub

Sub ColumnDisplayCopy(dst_cdt() As ColumnDisplayType, src_cdt() As ColumnDisplayType)
  Dim l As Long

  ReDim dst_cdt(UBound(src_cdt()))
  For l = LBound(src_cdt()) To UBound(src_cdt())
    dst_cdt(l).colTitle = src_cdt(l).colTitle
    dst_cdt(l).colWidth = src_cdt(l).colWidth
  Next l
End Sub

'Sub RowFillerAddToSet(rft() As RowFillerType, oCellRange As Object)
''writeStr As String, colTitle As String, before As String, after As String)
'  Dim b As Boolean
'  Dim c As Long
'  Dim tmp_lng As Long
'  Dim tmp_str0 As String
'  Dim tmp_str1 As String
'  Dim oRangeAddress As Object
'
'  ' Sheet   is the index of the sheet that contains the cell range.  
'  ' StartColumn   is the index of the column of the left edge of the range.  
'  ' StartRow  is the index of the row of the top edge of the range.  
'  ' EndColumn   is the index of the column of the right edge of the range.  
'  ' EndRow  is the index of the row of the bottom edge of the range.  
'  'Print "oRangeAddress sheet " + oRangeAddress.Sheet + _
'  '" start col " + oRangeAddress.StartColumn + _
'  '" start row " + oRangeAddress.StartRow + _
'  '" end col " + oRangeAddress.EndColumn + _
'  '" end row " + oRangeAddress.EndRow
'  oRangeAddress = oCellRange.getRangeAddress
'
'  tmp_lng = oRangeAddress.StartRow
'  If ((tmp_lng + 1) <> oRangeAddress.EndRow) Then
'    ' quick pointless error check
'    Print oCellRange.AbsoluteName + " is not appropriate for row filler rules.  This function expects two adjacent rows."
'  End If
'  ' have we seen our first strings that aren't null?
'  b = False
'  tmp_lng = oRangeAddress.EndColumn - oRangeAddress.StartColumn
'  For c = 0 To tmp_lng
'    tmp_str0 = oCellRange.getCellByPosition(c, 0).getString()
'    tmp_str1 = oCellRange.getCellByPosition(c, 1).getString()
'    ' StrComp(a, b, 0) not case-sensitive
'    ' StrComp(a, b, 1) case-sensitive
'    If ((StrComp(tmp_str0, "", 1) <> 0) Or (StrComp(tmp_str1, "", 1) <> 0)) Then
'      If (b = False) Then
'        b = True
'        RowFillerAddToRule(owFiller(UBound(owFiller())), "Gross", "0.0 bef", "0.0 aft")
'      Else
'        RowFillerAddToSet(owFiller(), "DNT mc", "Type", "0 bef", "0 aft")
'      End If
'    End If
'  Next c
'
'
'End Sub

' this is a fill rule generated from the Filler sheet, may span multiple columns
Type RowFillerType
  isActive As Boolean
  writeStr As String
  colTitle() As String
  colIndex() As Long
  before() As String
  after() As String
End Type

' here we'll add a new filler rule as well as it's first condition
Sub RowFillerAddToSetRF(rft() As RowFillerType, newRF As RowFillerType)
  Dim j As Long
  Dim l As Long
  Dim k As Long

  l = UBound(rft()) + 1
  k = UBound(newRF.colTitle())

  ReDim Preserve rft(l)
  VarArrayPreserveAdd(rft(l).colTitle(), k + 1)
  VarArrayPreserveAdd(rft(l).colIndex(), k + 1)
  VarArrayPreserveAdd(rft(l).before(), k + 1)
  VarArrayPreserveAdd(rft(l).after(), k + 1)

  rft(l).isActive = newRF.isActive
  rft(l).writeStr = newRF.writeStr

  For j = 0 To k
    rft(l).colTitle(j) = newRF.colTitle(j)
    rft(l).colIndex(j) = newRF.colIndex(j)
    rft(l).before(j) = newRF.before(j)
    rft(l).after(j) = newRF.after(j)
  Next j
End Sub

' here we'll add a new filler rule as well as it's first condition
Sub RowFillerAddToSet(rft() As RowFillerType, isActive As Boolean, writeStr As String, _
                      colTitle As String, colIndex As Long, before As String, after As String)
  Dim l As Long

  l = UBound(rft()) + 1
  ReDim Preserve rft(l)
  VarArrayPreserveAdd(rft(l).colTitle(), 1)
  VarArrayPreserveAdd(rft(l).colIndex(), 1)
  VarArrayPreserveAdd(rft(l).before(), 1)
  VarArrayPreserveAdd(rft(l).after(), 1)

  rft(l).isActive = isActive
  rft(l).writeStr = writeStr

  rft(l).colTitle(0) = colTitle
  rft(l).colIndex(0) = colIndex
  rft(l).before(0) = before
  rft(l).after(0) = after
End Sub

' here we'll add another condition to an existing rule
Sub RowFillerAddToRule(rft As RowFillerType, colTitle As String, colIndex As Long, before As String, after As String)
  Dim l as Long

  l = VarArrayPreserveAdd(rft.colTitle(), 1)
  VarArrayPreserveAdd(rft.colIndex(), 1)
  VarArrayPreserveAdd(rft.before(), 1)
  VarArrayPreserveAdd(rft.after(), 1)

  rft.colTitle(l) = colTitle
  rft.colIndex(l) = colIndex
  rft.before(l) = before
  rft.after(l) = after
End Sub

' here we'll add another condition to an existing rule
Sub RowFillerNewRule(rft As RowFillerType, isActive As Boolean, writeStr As String, _
                     colTitle As String, colIndex As Long, before As String, after As String)
  VarArrayReDim(rft.colTitle(), 0)
  VarArrayReDim(rft.colIndex(), 0)
  VarArrayReDim(rft.before(), 0)
  VarArrayReDim(rft.after(), 0)

  rft.isActive = isActive
  rft.writeStr = writeStr
  rft.colTitle(0) = colTitle
  rft.colIndex(0) = colIndex
  rft.before(0) = before
  rft.after(0) = after
End Sub

Sub PrintRowFillerRule(rft As RowFillerType)
  Dim l As Long
  Print "isActive " + rft.isActive + " writeStr " + rft.writeStr

  For l = LBound(rft.colTitle()) To UBound(rft.colTitle())
    Print  "colTitle(" + l + ") " + rft.colTitle(l) + _
          " colIndex(" + l + ") " + rft.colIndex(l) + _
            " before(" + l + ") " + rft.before(l) + _
             " after(" + l + ") " + rft.after(l)
  Next l
'    Print  "Len(colTitle(" + l + ")) " + Len(rft.colTitle(l)) + _
'          " IsNumeric(colIndex(" + l + ")) " + IsNumeric(rft.colIndex(l)) + _
'            " Len(before(" + l + ")) " + Len(rft.before(l)) + _
'             " Len(after(" + l + ")) " + Len(rft.after(l))
End Sub

Sub PrintRowFillerRuleSet(rft() As RowFillerType)
  Dim l As Long

  For l = LBound(rft()) To UBound(rft())
    PrintRowFillerRule(rft(l))
  Next l
End Sub
'
'Sub RetestRowFillerType
'  Dim owFiller() As RowFillerType
'  Dim dwFiller() As RowFillerType
'  Dim tmp_fill As RowFillerType
'  
'  RowFillerNewRule(tmp_fill, "DNT mc", "Type", 13, "0 bef", "0 aft")
'  RowFillerAddToRule(tmp_fill, "Gross", 14, "0.0 bef", "0.0 aft")
'  'PrintRowFillerRule(tmp_fill)
'  RowFillerAddToSetRF(owFiller(), tmp_fill)
'  RowFillerNewRule(tmp_fill, "DNT cm", "Type", 13, "1 bef", "1 aft")
'  'PrintRowFillerRule(tmp_fill)
'  RowFillerAddToSetRF(owFiller(), tmp_fill)
'  RowFillerNewRule(tmp_fill, "DNT cc", "Type", 13, "2 bef", "2 aft")
'  RowFillerAddToRule(tmp_fill, "Gross", 14, "2.2 bef", "2.2 aft")
'  RowFillerAddToRule(tmp_fill, "Gross", 14, "2.3 bef", "2.3 aft")
'  'PrintRowFillerRule(tmp_fill)
'  RowFillerAddToSetRF(dwFiller(), tmp_fill)
'  RowFillerAddToSetRF(owFiller(), tmp_fill)
'  PrintRowFillerRuleSet(owFiller())
'  PrintRowFillerRuleSet(dwFiller())
'
'End Sub
'
'Sub TestRowFillerType
'  Dim owFiller() As RowFillerType
'  Dim dwFiller() As RowFillerType
'  
'  RowFillerAddToSet(owFiller(), "DNT mc", "Type", "0 bef", "0 aft")
'  RowFillerAddToRule(owFiller(UBound(owFiller())), "Gross", "0.0 bef", "0.0 aft")
'  RowFillerAddToSet(owFiller(), "DNT cm", "Type", "1 bef", "1 aft")
'  RowFillerAddToSet(owFiller(), "DNT cc", "Type", "2 bef", "2 aft")
'  RowFillerAddToRule(owFiller(UBound(owFiller())), "Gross", "2.2 bef", "2.2 aft")
'  RowFillerAddToRule(owFiller(UBound(owFiller())), "Gross", "2.3 bef", "2.3 aft")
'
'End Sub

Function StrStarEndComp(searched_str As String, search_for_str As String, strcomp_cmp As Integer) 
  Dim i As Long
  Dim j As Long
  Dim m() As String
  Dim done_by_instr As Boolean
  Dim instr_cmp As Integer
  Dim ret As Integer

  done_by_instr = False
  If strcomp_cmp = 0 Then
    instr_cmp = 1
  Else
    instr_cmp = 0
  End If

  ' returns position of string if found, or zero if not found
  i = InStr(1, search_for_str, "*", instr_cmp)
  j = Len(search_for_str)
  If (i <> 0) And (i = j) Then 
    ' allow for string matching with star at end
    m() = split(search_for_str, "*")
    If InStr(1, searched_str, m(0), instr_cmp) = 1 Then
      ' we only want it if it is a one
      ' here we declare a match
      done_by_instr = True
      ret = 0
    End If
  End If

  If done_by_instr = False Then
    ret = StrComp(search_for_str, searched_str, strcomp_cmp)
  End If

  'MsgBox("StrStarEndComp(" + searched_str + ", " + search_for_str + ", " + strcomp_cmp + ") returns " + ret)

  StrStarEndComp = ret
End Function

Function CCellPositionToAbsoluteName(nCol As Long, nRow As Long, nSheet As Long) As String
  Dim oCellRange As Object

  oCellRange = ThisComponent.Sheets.getCellRangeByPosition(nCol, nRow, nCol, nRow, nSheet)
  CCellPositionToAbsoluteName = oCellRange.AbsoluteName
End Function 

Function CCellToAbsoluteName(oCell As Object) As String
  Dim oCellRange As Object
  Dim c, r, s As Long

  c = oCell.CellAddress.Column
  r = oCell.CellAddress.Row
  s = oCell.CellAddress.Sheet
  oCellRange = ThisComponent.Sheets.getCellRangeByPosition(c, r, c, r, s)
  CCellToAbsoluteName = oCellRange.AbsoluteName
End Function

Function CellGetOnlyValue(oCell As Object) As Double
  Dim d As Double

  d = CellGetOnlyValueErr(oCell)
  If StrComp(error_message_str, "", 0) <> 0 Then
    Print error_message_str
    error_message_str = ""
  End If
  CellGetOnlyValue = d
End Function

Function CellGetOnlyValueErr(oCell As Object) As Double
  On Error Goto ErrorHandler
  Dim d As Double
  Dim b As Boolean
  Dim tmp_str As String

  b = True
  If (oCell.getType() = com.sun.star.table.CellContentType.VALUE) Then
    d = oCell.getValue()
    'Print "Cell " + AbsoluteName + " type VALUE, value is " + d + "."
  ElseIf (oCell.getType() <> com.sun.star.table.CellContentType.EMPTY) Then
    ' allow type EMPTY for a zero value, sometimes the Sales Tax column is empty on certain rows
    d = 0.0
    b = False
    GoTo ErrorHandler
  End If

  CellGetOnlyValueErr = d
  Exit Function

  ErrorHandler:
  If (b = True) Then
    error_message_str = "Error on line " + Erl + ", error number " + Err + ": " + Error(Err) + _
          "The cell """ + CCellToAbsoluteName(oCell) + """ returned no content type.  Please check " + _
          "the the string and ensure it references a valid sheet and cell."
  Else
    Select Case oCell.getType()
    'Case com.sun.star.table.CellContentType.EMPTY
      'Print "Cell " + AbsoluteName + " type EMPTY"
      'tmp_str = "EMPTY"
    Case com.sun.star.table.CellContentType.TEXT
      'Print "Cell " + AbsoluteName + " type TEXT"
      tmp_str = "TEXT"
    Case com.sun.star.table.CellContentType.FORMULA
      'Print "Cell " + AbsoluteName + " type FORMULA"
      tmp_str = "FORMULA"
    Case Else
      'Print "Cell " + AbsoluteName + " type not found"
      tmp_str = "UNKNOWN"
    End Select
    error_message_str = "The cell """ + CCellToAbsoluteName(oCell) + """ has content type " + tmp_str + ".  " + _
          "The macros of this spreadsheet were expecting a number value in this cell. " + _
          "This could be an error that occurred while importing data into the spreadsheet. " + _
          "If the numbers of the data being imported are formatted like $4.00 and ($4.00), " + _
          "then note while importing in the category ""Other options"" that ""Detect special numbers"" " + _
          "must be selected.  If the numbers of the data being imported are formatted like 4.00 and -4.00 " + _
          "then ""Quoted field as text"" should not be selected, and ""Detect special numbers"" need not be selected.  " + _
          "If the cell type is FORMULA and the formula in the cell returns a valid number, then edit the macros to " + _
          "allow FORMULA content type as well."
  End If

  CellGetOnlyValueErr = d
End Function

Function CellAbsoluteNameGetValue(AbsoluteName As String) As Double
  On Error Goto ErrorHandler
  Dim oCellRanges(), oCell As Object
  Dim d As Double
  Dim tmp_str As String

  d = 0.0
  oCellRanges = ThisComponent.Sheets.getCellRangesByName(AbsoluteName)
  If UBound(oCellRanges()) < 0 Then
    GoTo ErrorHandler
  Else
    oCell = oCellRanges(0).getCellByPosition(0, 0)
    ' EMPTY   cell is empty.
    ' VALUE   cell contains a constant value.
    ' TEXT    cell contains text.
    ' FORMULA cell contains a formula.
    Select Case oCell.getType()
    Case com.sun.star.table.CellContentType.EMPTY
      'Print "Cell " + AbsoluteName + " type EMPTY"
      tmp_str = "EMPTY"
      GoTo ErrorHandler
    Case com.sun.star.table.CellContentType.VALUE
      'Print "Cell " + AbsoluteName + " type VALUE"
    Case com.sun.star.table.CellContentType.TEXT
      'Print "Cell " + AbsoluteName + " type TEXT"
      tmp_str = "TEXT"
      GoTo ErrorHandler
    Case com.sun.star.table.CellContentType.FORMULA
      'Print "Cell " + AbsoluteName + " type FORMULA"
      tmp_str = "FORMULA"
      GoTo ErrorHandler
    Case Else
      'Print "Cell " + AbsoluteName + " type not found"
      tmp_str = "UNKNOWN"
      GoTo ErrorHandler
    End Select
    d = oCell.getValue()
    Print "Cell " + AbsoluteName + " type VALUE, value is " + d + "."
  End If

  CellAbsoluteNameGetValue = d
  Exit Function

  ErrorHandler:
  If UBound(oCellRanges()) < 0 Then
    Print "Error on line " + Erl + ", error number " + Err + ": " + Error(Err) + _
          "The cell range """ + AbsoluteName + """ returned no cell ranges.  Please check " + _
          "the the string and ensure it references a valid sheet and cell range."
  Else
    Print "The cell """ + AbsoluteName + """ has content type " + tmp_str + ".  " + _
          "The macros of this spreadsheet were expecting a number value in this cell. " + _
          "This could be an error that occurred while importing data into the spreadsheet. " + _
          "If the numbers of the data being imported are formatted like $4.00 and ($4.00), " + _
          "then note while importing in the category ""Other options"" that ""Detect special numbers"" " + _
          "must be selected.  If the numbers of the data being imported are formatted like 4.00 and -4.00 " + _
          "then ""Quoted field as text"" should not be selected, and ""Detect special numbers"" need not be selected.  " + _
          "If the cell type is FORMULA and the formula in the cell returns a valid number, then edit the macros to " + _
          "allow FORMULA content type as well."
  End If

  CellAbsoluteNameGetValue = d
End Function

Function CAbsoluteNameToCell(AbsoluteName As String) As Object
  Dim oCellRanges(), oCell As Object

  oCellRanges = ThisComponent.Sheets.getCellRangesByName(AbsoluteName)
  If UBound(oCellRanges()) < 0 Then
    Print AbsoluteName + " returned no ranges!"
    oCell = NULL
  Else
    oCell = oCellRanges(0).getCellByPosition(0, 0)
  End If
  CAbsoluteNameToCell = oCell
End Function

' example:
'
' TestCellRangeEmptyByName(0; 0; 0; 5; 7)
'
'Function TestCellRangeEmptyByPosition(sheet, left, top, right, bottom) As Boolean
'  Dim oSheet
'  Dim oCellRange
'
'  oSheet = ThisComponent.getSheets().getByIndex(sheet)
'  oCellRange = oSheet.getCellRangeByPosition(left,top,right,bottom)
'
'  TestCellRangeEmptyByPosition = IsCellRangeEmpty(oCellRange)
'End Function

Function IsCellRangeEmpty(oRange) As Boolean
  Dim oRanges      'Ranges returned after querying for the cells
  Dim oAddrs()     'Array of CellRangeAddress
 
  oRanges = oRange.queryContentCells(_
    com.sun.star.sheet.CellFlags.VALUE OR _
    com.sun.star.sheet.CellFlags.DATETIME OR _
    com.sun.star.sheet.CellFlags.STRING OR _
    com.sun.star.sheet.CellFlags.FORMULA)
  oAddrs() = oRanges.getRangeAddresses()
 
  IsCellRangeEmpty = UBound(oAddrs()) < 0
End Function

Function SheetRowIndexFirstEmpty(oSheet As Object, row As Long, left As Long, right As Long) As Long
  Dim ret As Long
  Dim i As Long

  ret = -1

  ' cycle through all columns if necessary
  For i = left To right
    If (oSheet.getCellByPosition(i, row).getType() = com.sun.star.table.CellContentType.EMPTY) Then
      ret = i
      Exit For
    End If
  Next i

  SheetRowIndexFirstEmpty = ret
End Function

Function CellRangeSearchTrimCol(oCellRange As Object, match As String) As Long
  Dim oCell As Object
  Dim col_index As Long

  col_index = -1

  oCell = CellRangeSearchTrim(oCellRange, match)
  If Not IsNull(oCell) Then
    col_index = oCell.CellAddress.Column
  End If
  CellRangeSearchTrimCol = col_index
End Function

Function CellRangeSearchTrimRow(oCellRange As Object, match As String) As Long
  Dim oCell As Object
  Dim row_index As Long

  row_index = -1

  oCell = CellRangeSearchTrim(oCellRange, match)
  If Not IsNull(oCell) Then
    row_index = oCell.CellAddress.Row
  End If
  CellRangeSearchTrimRow = row_index
End Function

Function CellRangeSearchTrim(oCellRange As Object, match As String) As Object
  Dim oDescriptor As object
  Dim oCell As Object

  oDescriptor = oCellRange.createSearchDescriptor()
  oDescriptor.SearchString = match
  ' If true, the search will match only complete words. One white space character can cause a mismatch.
  'oDescriptor.SearchWords = True
  oDescriptor.SearchWords = False
  oDescriptor.SearchCaseSensitive = True
  ' Instead, set this to false to not have to deal with white space
  'oDescriptor.SearchWords = False
  oCell = oCellRange.findFirst(oDescriptor)
  Do While Not IsNull(oCell)
    If StrComp(Trim(oCell.getString()), Trim(oDescriptor.SearchString)) = 0 Then
      Exit Do
    End If
    oCell = oCellRange.findNext(oCell, oDescriptor)
  Loop
  If IsNull(oCell) Then
    ' didn't find it, so try again when we first trim white space off the search string
    ' we could have done it this way first but then we wouldn't be able to match a string
    ' of white space to another string of white space, sounds goofy but there is a column
    ' that only has one space at the end of the PayPal history stuff and I want this column
    ' adjusted to the default width of 2267, not optimal width, which looks skinny with 1 space
    'Print "searching again this time with trimmed searcher"
    oDescriptor = oCellRange.createSearchDescriptor()
    oDescriptor.SearchString = Trim(match)
    oCell = oCellRange.findFirst(oDescriptor)
    Do While Not IsNull(oCell)
      If StrComp(Trim(oCell.getString()), oDescriptor.SearchString) = 0 Then
        Exit Do
      End If
      oCell = oCellRange.findNext(oCell, oDescriptor)
    Loop
  End If

'  If Not IsNull(oCell) Then
'    ' string found is exact match ignoring leading or trailing white space
'    Print oCell.getString()
'    Print "Column index = " + oCell.CellAddress.Column
'    Print "Row index    = " + oCell.CellAddress.Row
'  End If

  CellRangeSearchTrim = oCell
End Function

Sub AccEnsureTaxColumnViewAll
  Dim sht as Object
  Dim Rng as object
  Dim i as Long
  Dim j as Long

  'Dim oDocument as Object
  Dim oController as Object
  Dim oSheets as Object
  Dim oSheet as Object
  Dim oSheetBack as Object
  Dim oCellRange as Object
  Dim tmp_str as String
  'Dim taxstring as String
  Dim oDispatcher As Object
  Dim stepLast(0) As new com.sun.star.beans.PropertyValue
  
  InitAccountingGlobals()

  ' this stuff moves the active cell from wherever it is to the top left
  stepLast(0).Name = "ToPoint"
  stepLast(0).Value = "$A$1"

  oDispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
  oSheets = ThisComponent.Sheets
  oSheetBack = ThisComponent.CurrentController.getActiveSheet()  

  For i = 0 To oSheets.Count - 1
    oSheet = oSheets.getByIndex(i)
    If StrComp(oSheet.getName(), filler_sheet_str, 1) = 0 Then
      ' freeze top row and leftmost two columns of this sheets
      ThisComponent.CurrentController.setActiveSheet(oSheet)
      oDispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoToCell", "", 0, stepLast())
      'oController = oDocument.getCurrentController
      'oController.freezeAtPosition(0, 0)
      'oController.freezeAtPosition(2, 1)
      'ThisComponent.CurrentController.freezeAtPosition(0, 0)
      'ThisComponent.CurrentController.freezeAtPosition(2, 1)
      ThisComponent.CurrentController.freezeAtPosition(3, 1)
    ElseIf StrComp(oSheet.getName(), total_sheet_str, 1) <> 0 Then
      'Print oSheet.getName()
      tmp_str = oSheet.getCellByPosition(0, 0).getString()
      ' 1 for case-sensitive, 0 for case-insensitive
      If StrComp(tmp_str, subtotals_col_str, 1) <> 0 Then
        'oSheet.Columns.removeByIndex(3, 2)
        oSheet.Columns.insertByIndex(0, 1)
        oSheet.getCellByPosition(0, 0).setString(subtotals_col_str)
      End If
      ' freeze top row and leftmost column of this sheet
      ThisComponent.CurrentController.setActiveSheet(oSheet)
      oDispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoToCell", "", 0, stepLast())
      'ThisComponent.CurrentController.freezeAtPosition(0, 0)
      ThisComponent.CurrentController.freezeAtPosition(1, 1)
      'oController = oDocument.getCurrentController
      'oController.freezeAtPosition(0, 0)
      'oController.freezeAtPosition(1, 1)
    End If
    For j = 0 To max_col_index
      oSheet.getColumns().getByIndex(j).OptimalWidth = True
      If (j mod 4) = 3 Then
        oCellRange = oSheet.getCellRangeByPosition(j, 0, j + 1, 0)
        If IsCellRangeEmpty(oCellRange) Then
          Exit For
        End If
      End If
    Next j
  Next i
  ThisComponent.CurrentController.setActiveSheet(oSheetBack)
End Sub

Sub AccColumnsHideSubset
  ColumnsSubsetOperator(False)
End Sub

Sub AccColumnsDeleteSubset
  ColumnsSubsetOperator(True)
End Sub

Sub ColumnsSubsetOperator(delete As Boolean)
  'Dim pphCols(0 To 40) As ColumnDisplayType
  Dim tmp_str As String
  Dim sheetCount As Integer
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim oDoc As object
  Dim oDescriptor As object
  Dim oCellRange As object
  Dim oSheet As Object
  Dim oCell As Object
  Dim col_index As Long
  Dim row_index As Long
  Dim fill_type_col_idx As Long
  Dim overwrite_col_idx As Long
  Dim col_width_idx As Long
  Dim cdt() As ColumnDisplayType
  'Dim oConv As Object

  InitAccountingGlobals()

  If (delete = True) Then
    Print "This macro searches for a sheet named """ + filler_sheet_str + """, a column titled """ + _ 
          fill_type_col_str + """, and the first cell in that column with a string matching """ + col_width_str + """.  " + _
          "In the row of this cell, for all columns that are not titled """ + fill_type_col_str + """ or """ + overwrite_col_str + _
          """, if that column has False or 0 entered into the cell, then the header of this column will be searched for " + _
          "in all sheets except for the """ + filler_sheet_str + """ and """ + total_sheet_str + """ sheets, and the " + _
          "matching columns in the sheets will be deleted permanently.  If you do not want to delete " + _
          "these columns labeled False or 0, click to Cancel now, otherwise click OK."
  End If


  'oConv = ThisComponent.createInstance("com.sun.star.table.CellAddressConversion") 
  oDoc = ThisComponent
  sheetCount = ThisComponent.Sheets.getCount() 

  ' get data from Fillers sheet
  For i = 0 to sheetCount - 1
    oSheet = ThisComponent.Sheets.getByIndex(i)
    If StrComp(oSheet.getName(), filler_sheet_str, 1) = 0 Then
      col_index = SheetRowIndexFirstEmpty(oSheet, 0, 0, max_col_index)
      If (col_index < 0) Then
        Print "Could not find the end of the column headers in sheet " + oSheet.getName()
        Exit Sub
      End If
      'Print "first empty column index " + col_index + " first empty row " + row_index
      'Exit Sub
      oCellRange = oSheet.getCellRangeByPosition(0, 0, col_index - 1, 0)
      ' this index is not critical
      overwrite_col_idx = CellRangeSearchTrimCol(oCellRange, overwrite_col_str)
      ' this is
      fill_type_col_idx = CellRangeSearchTrimCol(oCellRange, fill_type_col_str)
      If (fill_type_col_idx < 0) Then
        Print "sheet " + oSheet.getName() + " could not find column matching " + _
              fill_type_col_str + " in the first row."
        Exit Sub
      End If
      'tmp_str = " Date"
      'Print "search for """ + tmp_str + """ returns " + CellRangeSearchTrimCol(oCellRange, tmp_str)
      'overwrite_col_idx = CellRangeSearchTrimCol(oCellRange, overwrite_col_str)
      oCellRange = oSheet.getCellRangeByPosition(fill_type_col_idx, 0, fill_type_col_idx, max_row_index)
      col_width_idx = CellRangeSearchTrimRow(oCellRange, col_width_str)
      If (col_width_idx < 0) Then
        Print "sheet " + oSheet.getName() + " could not find row matching " + _
              col_width_str + " in column " + (fill_type_col_idx + 1)
        Exit Sub
      End If
      For j = 0 to col_index - 1
        If (j <> fill_type_col_idx) And (j <> overwrite_col_idx) Then
          ' EMPTY   cell is empty.
          ' VALUE   cell contains a constant value.
          ' TEXT    cell contains text.
          ' FORMULA cell contains a formula.
          'k = oSheet.getCellByPosition(j, col_width_idx).getType()
          'oConv.Address = oSheet.getCellByPosition(j, col_width_idx).getCellAddress
          'Print oConv.UserInterfaceRepresentation + " get type " + k
          tmp_str = oSheet.getCellByPosition(j, col_width_idx).getString()
          'Print "tmp_str " + tmp_str
          ' StrComp(a, b, 0) not case-sensitive
          ' StrComp(a, b, 1) case-sensitive
          If StrComp(tmp_str, "true", 0) = 0 Then
            ' set optimal width
            k = -1
          ElseIf StrComp(tmp_str, "false", 0) = 0 Then
            ' hide column
            k = 0
          ElseIf StrComp(tmp_str, "", 0) = 0 Then
            ' set default width
            k = 2267
          Else
            ' set user defined width
            k = CLng(tmp_str)
          End If
          tmp_str = oSheet.getCellByPosition(j, 0).getString()
          ColumnDisplayAddToSet(cdt(), tmp_str, k)
          'Print tmp_str + " " + k + " UBound(cdt()) " + UBound(cdt())
        End If
      Next j
      Exit For
    End If
  Next i

  For i = 0 to sheetCount - 1
    oSheet = ThisComponent.Sheets.getByIndex(i)
    If (StrComp(oSheet.getName(), filler_sheet_str, 1) <> 0) And _
       (StrComp(oSheet.getName(), total_sheet_str, 1) <> 0) Then
      col_index = SheetRowIndexFirstEmpty(oSheet, 0, 0, max_col_index)
      If (col_index < 0) Then
        Print "Could not find the end of the column headers in sheet " + oSheet.getName()
      Else
      'Print oSheet.getName() + " first empty column index " + col_index + " first empty row index " + row_index
      'Exit Sub
        oCellRange = oSheet.getCellRangeByPosition(0, 0, col_index - 1, 0)
        For j = LBound(cdt()) to UBound(cdt())
          col_index = CellRangeSearchTrimCol(oCellRange, cdt(j).colTitle)
          If (col_index >= 0) Then
            If cdt(j).colWidth = -1 Then
              ' set optimal width
              oSheet.getColumns().getByIndex(col_index).OptimalWidth = True
            ElseIf cdt(j).colWidth = 0 Then
              If (delete = True) Then
                oSheet.Columns.removeByIndex(col_index, 1)
              Else
                oSheet.getColumns().getByIndex(col_index).IsVisible = False
              End If
            Else
              oSheet.getColumns().getByIndex(col_index).Width = cdt(j).colWidth
            End If
          Else
            Print "Could not find column header " + cdt(j).colTitle + " in sheet " + oSheet.getName()
          End If
        Next j
      End If
    End If
  Next i

'Sub ColumnsHideSubset
End Sub

' Iterate over the set of RowFillerTypes and over all rules in each RowFillerType.
' Check that the header in the rule can be found in the cell range, and if not,
' then deactivate the rule for this sheet.  The return value is the first column found
' in the sheet, and therefore a negative return value indicates that no rules are active.
'
Function RowFillerSetFillColumnIndicies(rfSet() as RowFillerType, oCellRange As Object) As Long
  Dim j, k, col_index, ret As Long
  Dim tmp_str As String
  Dim b As Boolean
  Dim oRangeAddress As Object

  ret = -1
  oRangeAddress = oCellRange.getRangeAddress
'  Print "oRangeAddress sheet " + oRangeAddress.Sheet + _
'  " start col " + oRangeAddress.StartColumn + _
'  " start row " + oRangeAddress.StartRow + _
'  " end col " + oRangeAddress.EndColumn + _
'  " end row " + oRangeAddress.EndRow
'  Exit Sub

  For j = LBound(rfSet()) To UBound(rfSet())
    ' activate all fill rules and deactivate if the column cannot be found
    rfSet(j).isActive = True
    For k = LBound(rfSet(j).colTitle()) To UBound(rfSet(j).colTitle())
      b = False
      If ((rfSet(j).colIndex(k) >= oRangeAddress.StartColumn)) And ((rfSet(j).colIndex(k) <= oRangeAddress.EndColumn)) Then
        tmp_str = oCellRange.getCellByPosition(rfSet(j).colIndex(k), 0).getString()
        ' StrComp(a, b, 0) not case-sensitive
        ' StrComp(a, b, 1) case-sensitive
        If (StrComp(Trim(tmp_str), Trim(rfSet(j).colTitle(k)), 1) = 0) Then
          If (ret = -1) Then
            ret = rfSet(j).colIndex(k)
          End If
          b = True
        End If
      End If
      If (b = False) Then
        'Print "Having to search for column title """ + rfSet(j).colTitle(k) + """ in cell range " + _
              'oCellRange.AbsoluteName + " since it was not found at numerical column index " + rfSet(j).colIndex(k)
        col_index = CellRangeSearchTrimCol(oCellRange, rfSet(j).colTitle(k))
        If (col_index < 0) Then
          Print "A column titled """ + rfSet(j).colTitle(k) + """ could not be found in cell range " + oCellRange.AbsoluteName + _
                ".  Please check this range as well as the " + filler_sheet_str + " sheet to verify this fill rule.  " + _
                "This fill rule will be deactivated for this particular sheet."
          rfSet(j).isActive = False
          Exit For
        Else
          ' set a new column index, probably won't happen too much
          rfSet(j).colIndex(k) = col_index
          If (ret = -1) Then
            ret = rfSet(j).colIndex(k)
          End If
        End If
      End If
'    Print "writeStr " + rfSet(j).writeStr
'    'For l = LBound(blFiller(j).colTitle()) To UBound(blFiller(j).colTitle())
'      Print  "colTitle(" + k + ") " + rfSet(j).colTitle(k) + _
'            " colIndex(" + k + ") " + rfSet(j).colIndex(k) + _
'              " before(" + k + ") " + rfSet(j).before(k) + _
'               " after(" + k + ") " + rfSet(j).after(k)
    Next k
  Next j
  RowFillerSetFillColumnIndicies = ret
End Function

'Name  Type  Status  Gross
'TRUE  FALSE FALSE TRUE
'TRUE  TRUE  TRUE  TRUE
'UpThisDownBeginsNumOpp("$January.$E$22"; "PayPal Extras MasterCard<sup>&#174;</sup>"; "Gross") And ($January.$H$22 < 0)


Function RowFillerRunFillRuleSet(rfSet() As RowFillerType, oCellRange As Object, oCellFill As Object, oSuperCell As Object)
  Dim ret As Integer
  Dim i, j, k As Long
  Dim run_formula As String
  Dim oCell As Object
  Dim tmp_str As String
  Dim wrote_cell As Boolean
  Dim clauses_match As Boolean
'  Dim avoid_super_cell As Boolean


  ' I was curious to know the speed difference of executing formulas in the super cell versus executing
  ' them in a macro.  While there is a speed difference it's not huge, and so stick with the super cell.
  ' However, going to need an active sheet routine because doing all sheets is slow.  The results of the
  ' speed test are below.

  ' [aaron@arwlinux ~ Wed Mar 28 10:33:22]
  ' $ cat /proc/cpuinfo | egrep '^model name|^cpu MHz'
  ' model name  : Intel(R) Pentium(R) 4 CPU 3.00GHz
  ' cpu MHz   : 3000.000
  ' model name  : Intel(R) Pentium(R) 4 CPU 3.00GHz
  ' cpu MHz   : 3000.000
  ' [aaron@arwlinux ~ Wed Mar 28 10:33:34]
  ' $ echo 68 seconds compared to 91 seconds
  ' 22 seconds with inline functions compared to 48 seconds with super cell

  ' The speed test must have a Fillers sheet with only two entries for fill rules and it assumes this
  ' structure.  It knows that there is data in the Name and Gross columns and therefore makes the rules
  ' properly, but it bypasses using that data and instead calls the functions based on integer values.
  ' To do the speed check the Fillers sheet for this structure and then uncomment the appropriate lines
  ' in this function.

  ' Fillers Type | Overwrite | Tax Category            | Name                                                     | Gross
  ' before cell  | TRUE      | DNT mastercard negative | UpThisDownBeginsNumOpp("                                 | 
  ' after cell   |           |                         | "; "PayPal Extras MasterCard<sup>&#174;</sup>"; "Gross") | < 0
  ' before cell  | TRUE      | DNT mastercard positive | UpThisDownBeginsNumOpp("                                 | 
  ' after cell   |           |                         | "; "PayPal Extras MasterCard<sup>&#174;</sup>"; "Gross") | > 0
  
'  avoid_super_cell = True


  ' return 0 if we didn't fill, otherwise return 1
  ret = 0
  wrote_cell = False

  For j = LBound(rfSet()) To UBound(rfSet())
    ' activate all fill rules and deactivate if the column cannot be found
    If (rfSet(j).isActive = True) Then
      clauses_match = True
      For k = LBound(rfSet(j).colTitle()) To UBound(rfSet(j).colTitle())
        oCell = oCellRange.getCellByPosition(rfSet(j).colIndex(k), 0)

' speed tester compare to calling internal functions
'        If (avoid_super_cell = True) Then
'          tmp_str = CCellToAbsoluteName(oCell)
'          If (k = 0) Then
'            clauses_match = UpThisDownBeginsNumOpp(tmp_str, "PayPal Extras MasterCard<sup>&#174;</sup>", "Gross")
'            'Print "UpThisDownBeginsNumOpp(""" & tmp_str & """" + "; PayPal Extras MasterCard<sup>&#174;</sup>" + "; ""Gross"") is " + clauses_match
'          ElseIf (j = 0) Then
'            clauses_match = (oCell.getValue() < 0)
'            'Print tmp_str + " < 0 is " + clauses_match
'          Else
'            clauses_match = (oCell.getValue() > 0)
'            'Print tmp_str + " > 0 is " + clauses_match
'          End If
'          If (clauses_match = False) Then
'            Exit For
'          End If
'        Else

        ' normal code here
        run_formula = "=" & rfSet(j).before(k) & CCellToAbsoluteName(oCell) & rfSet(j).after(k)
        oSuperCell.setFormula(run_formula)
        ' TRUE or FALSE, else error
        tmp_str = oSuperCell.getString()
        ' for print statements
        ' clear right away
        oSuperCell.clearContents(23)
        If StrComp(error_message_str, "", 0) <> 0 Then
          Print error_message_str
          error_message_str = ""
        End If
        If StrComp(tmp_str, "true", 0) = 0 Then
          'Print run_formula + " evaluates to " + tmp_str
          ' continue through clauses of this rule
        ElseIf StrComp(tmp_str, "false", 0) = 0 Then
          'Print run_formula + " evaluates to " + tmp_str
          clauses_match = False
          ' can stop processing this rule, clause does not match
          Exit For
        Else
          ' test this warning
          clauses_match = False
          Print "The formula " + run_formula + " generated by the combination of the fill rules and the cell range " + _
                + oCellRange.AbsoluteName + " did not evaluate to either TRUE or FALSE.  This is more than likely an error in the " + _
                before_str + " string or " + after_str + " string entered into the " + filler_sheet_str + " sheet.  Check the sheet and try " + _
                "entering the formula by hand into a cell to test that it returns TRUE or FALSE.  In particular, ensure that the " + _
                "double quotes of the formula are straight double quotes and not being automatically corrected to curly start and end quotes.  " + _
                "You can disable automatic correction of quotes by clicking Tools -> AutoCorrect Options -> Localized Options -> Double Quotes " + _
                "and removing the check from the Replace box.  Also, keep in mind that formulas must either be the macros of this sheet " + _
                "or a function that LibreOffice Calc can evaluate.  LibreOffice Basic functions are only available at run time and must " + _
                "be wrapped by a function in the macros of this sheet to be utilized.  Click Cancel to exit this macro."
          Exit For
        End If

' End If for the speed tester, comment out if not using
'        End If

      Next k
      If (clauses_match = True) Then
        ' automatically fill cell, easy as pi
        oCellFill.setString(rfSet(j).writeStr)
        wrote_cell = True
        ret = 1
      End If
    End If
    If (wrote_cell = True) Then
      Exit For
    End If
  Next j

  'oCellRange = oSheet.getCellRangeByPosition(col_index, UBound(subtotals_rows()) + 2, col_index + UBound(totals_columns()), UBound(subtotals_rows()) + 2)
  'oCellRange.clearContents(23)

  RowFillerRunFillRuleSet = ret
End Function


Sub AccFillActiveSheet()
  SheetFiller(ThisComponent.CurrentController.getActiveSheet().getName())
End Sub

Sub AccFillAllSheets()
  SheetFiller("")
End Sub

' For the PayPal history of 2011 the SheetFiller with all overwrite flags set to True takes about 3:17.
'
' [aaron@arwlinux ~ Wed Apr 04 10:58:46]
' $ 
' [aaron@arwlinux ~ Wed Apr 04 11:02:03]
' $ cat /proc/cpuinfo | egrep '^model name|^cpu MHz'
' model name  : Intel(R) Pentium(R) 4 CPU 3.00GHz
' cpu MHz   : 3000.000
' model name  : Intel(R) Pentium(R) 4 CPU 3.00GHz
' cpu MHz   : 3000.000

Sub SheetFiller(aSheet As String)
  Dim bool0 As Boolean
  Dim bool1 As Boolean
  Dim tmp_str0 As String
  Dim tmp_str1 As String
  Dim tmp_str2 As String
  Dim tmp_str3 As String
  Dim sheetCount As Integer
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim l As Long
  Dim oDoc As object
  Dim oDescriptor As object
  Dim oCellRange As object
  Dim oSheet As Object
  Dim oCell As Object
  Dim oSuperCell As Object
  Dim col_index As Long
  Dim row_index As Long
  Dim fill_type_col_idx As Long
  Dim overwrite_col_idx As Long
  Dim col_width_idx As Long
  Dim subtotals_col_idx As Long
  ' only write data if cell is blank
  Dim blFiller() As RowFillerType
  ' overwrite existing cell data
  Dim owFiller() As RowFillerType
  Dim tmpFiller As RowFillerType

  'Dim cdt() As ColumnDisplayType
  'Dim oConv As Object

  'subtotals_col_str = "Tax Category"

  'filler_sheet_str = "Fillers"
  'fill_type_col_str = "Fillers Type"
  'overwrite_col_str = "Overwrite"
  'col_width_str = "column width"
  'before_str = "before cell"
  'after_str = "after cell"

  InitAccountingGlobals()

  'oConv = ThisComponent.createInstance("com.sun.star.table.CellAddressConversion") 
  oDoc = ThisComponent
  sheetCount = ThisComponent.Sheets.getCount() 

  ' get data from Fillers sheet
  For i = 0 to sheetCount - 1
    oSheet = ThisComponent.Sheets.getByIndex(i)
    If StrComp(oSheet.getName(), filler_sheet_str, 1) = 0 Then
      col_index = SheetRowIndexFirstEmpty(oSheet, 0, 0, max_col_index)
      If (col_index < 0) Then
        Print "Could not find the end of the column headers in sheet " + oSheet.getName()
        Exit Sub
      End If
      oSuperCell = oSheet.getCellByPosition(col_index + 1, 0)
      If (oSuperCell.getType() <> com.sun.star.table.CellContentType.EMPTY) Then
        Print "The cell after the first empty cell in the column headers of sheet " + oSheet.getName() + _
              " will be used to execute all fill rule formulas for all sheets and must be blank." + _
              "  Please delete the contents of cell " + CCellToAbsoluteName(oSuperCell)
        Exit Sub
      End If
      'Print "first empty column index " + col_index + " first empty row " + row_index
      'Exit Sub
      oCellRange = oSheet.getCellRangeByPosition(0, 0, col_index - 1, 0)
      overwrite_col_idx = CellRangeSearchTrimCol(oCellRange, overwrite_col_str)
      If (overwrite_col_idx < 0) Then
        Print "sheet " + oSheet.getName() + " could not find column matching " + _
              overwrite_col_str + " in the first row."
        Exit Sub
      End If
      fill_type_col_idx = CellRangeSearchTrimCol(oCellRange, fill_type_col_str)
      If (fill_type_col_idx < 0) Then
        Print "sheet " + oSheet.getName() + " could not find column matching " + _
              fill_type_col_str + " in the first row."
        Exit Sub
      End If
      subtotals_col_idx = CellRangeSearchTrimCol(oCellRange, subtotals_col_str)
      If (subtotals_col_idx < 0) Then
        Print "sheet " + oSheet.getName() + " could not find column matching " + _
              subtotals_col_str + " in the first row."
        Exit Sub
      End If

      For j = 1 to max_row_index
        'tmpFiller = NULL
        If (oSheet.getCellByPosition(fill_type_col_idx, j).getType() = com.sun.star.table.CellContentType.EMPTY) Then
          'Print CCellPositionToAbsoluteName(totals_columns_index(0), j, i) + " is EMPTY"
          oCellRange = oSheet.getCellRangeByPosition(0, j, col_index - 1, j)
          If IsCellRangeEmpty(oCellRange) Then
            'Print oCellRange.AbsoluteName + " is EMPTY"
            Exit For
          End If
        ElseIf ((StrComp(oSheet.getCellByPosition(fill_type_col_idx, j).getString(), after_str) = 0) And _
                (StrComp(oSheet.getCellByPosition(fill_type_col_idx, j - 1).getString(), before_str) <> 0)) Then
          ' error checking after without before 
          Print "Fill rules entered in sheet " + oSheet.getName() + " in column " + fill_type_col_str + _
                " must have both before and after parts on adjacent rows.  The before string must be """ + before_str + _
                """ and the after string must be """ + after_str + """, however, cell " + CCellPositionToAbsoluteName(fill_type_col_str, j, i) + _
                " does not have a corresponding before string on the previous row.  Please fix this to make a complete rule."
        ElseIf ((StrComp(oSheet.getCellByPosition(fill_type_col_idx, j).getString(), before_str) = 0) And _
                (StrComp(oSheet.getCellByPosition(fill_type_col_idx, j + 1).getString(), after_str) <> 0)) Then
          ' error checking before without after
          Print "Fill rules entered in sheet " + oSheet.getName() + " in column " + fill_type_col_str + _
                " must have both before and after parts on adjacent rows.  The before string must be """ + before_str + _
                """ and the after string must be """ + after_str + """, however, cell " + CCellPositionToAbsoluteName(fill_type_col_str, j, i) + _
                " does not have a corresponding after string on the next row.  Please fix this to make a complete rule."
        ElseIf ((StrComp(oSheet.getCellByPosition(fill_type_col_idx, j).getString(), before_str) = 0) And _
                (StrComp(oSheet.getCellByPosition(fill_type_col_idx, j + 1).getString(), after_str) = 0)) Then
          ' before and after, now check for overwrite boolean and string
          tmp_str0 = oSheet.getCellByPosition(overwrite_col_idx, j).getString()
          tmp_str1 = oSheet.getCellByPosition(overwrite_col_idx, j + 1).getString()
          tmp_str2 = oSheet.getCellByPosition(subtotals_col_idx, j).getString()
          tmp_str3 = oSheet.getCellByPosition(subtotals_col_idx, j + 1).getString()

          ' StrComp(a, b, 0) not case-sensitive
          ' StrComp(a, b, 1) case-sensitive
          If StrComp(tmp_str0, "true", 0) = 0 Then
            ' overwrite data
            bool0 = True
            tmp_str0 = ""
          ElseIf StrComp(tmp_str0, "false", 0) = 0 Then
            ' write blank cell
            bool0 = False
            tmp_str0 = ""
          Else
            ' warning before row overwrite column not TRUE or FALSE
            tmp_str0 = CCellPositionToAbsoluteName(overwrite_col_idx, j, i) + " "
          End If
          If (StrComp(tmp_str1, "", 1) <> 0) Then
            ' warning after row overwrite column not blank
            tmp_str0 = tmp_str0 + CCellPositionToAbsoluteName(overwrite_col_idx, j + 1, i) + " "
          End If
          If (StrComp(tmp_str2, "", 1) = 0) Then
            ' warning before row subtotals column blank
            tmp_str0 = tmp_str0 + CCellPositionToAbsoluteName(subtotals_col_idx, j, i) + " "
          End If
          If (StrComp(tmp_str3, "", 1) <> 0) Then
            ' warning after row subtotals column not blank
            tmp_str0 = tmp_str0 + CCellPositionToAbsoluteName(subtotals_col_idx, j + 1, i) + " "
          End If

          If (StrComp(tmp_str0, "", 1) <> 0) Then
            ' in before row, need boolean in overwrite col and string in subtotals col
            ' in after row need empty values
            ' if tmp_str0 is not empty then at least one cell is in violation
            Print "Fill rules entered in sheet " + oSheet.getName() + " require the """ + before_str + """ row to have a TRUE or FALSE " + _
                  "value entered in the " + overwrite_col_str + " column, and a string value in the " + subtotals_col_str + " column " + _
                  "that is used to automatically fill cells of the spreadsheet.  The """ + after_str + """ row must be blank by convention " + _
                  "in these columns.  Please check and correct the following cells [" + Trim(tmp_str0) + "]."
          Else
            ' looking good, let's process these two rows to create a fill rule
            bool1 = False
            For k = 0 To col_index - 1
              If (k <> fill_type_col_idx) And (k <> overwrite_col_idx) And (k <> subtotals_col_idx) Then
                ' avoid processing the two columns that have our control info and the column that we want to fill
                tmp_str0 = oSheet.getCellByPosition(k, j).getString()
                tmp_str1 = oSheet.getCellByPosition(k, j + 1).getString()
                ' StrComp(a, b, 0) not case-sensitive
                ' StrComp(a, b, 1) case-sensitive
                If ((StrComp(tmp_str0, "", 1) <> 0) Or (StrComp(tmp_str1, "", 1) <> 0)) Then
                  tmp_str3 = oSheet.getCellByPosition(k, 0).getString()
                  If (bool1 = False) Then
                    bool1 = True
                    ' a little sloppy but tmp_str2 already has our write value,
                    ' and k - 2 is a good guess of the column index for the other sheets
                    ' 2 because of the fill type and overwrite columns unique to this sheet
                    RowFillerNewRule(tmpFiller, False, tmp_str2, tmp_str3, k - 2, tmp_str0, tmp_str1)
                  Else
                    RowFillerAddToRule(tmpFiller, tmp_str3, k - 2, tmp_str0, tmp_str1)
                  End If
                End If
              End If
            Next k
            If (bool1 = True) Then
              If (bool0 = False) Then
                ' only write data if cell is blank
                RowFillerAddToSetRF(blFiller(), tmpFiller)
              Else
                ' overwrite existing cell data
                RowFillerAddToSetRF(owFiller(), tmpFiller)
              End If
            Else
              Print "The fill rule entered in sheet " + oSheet.getName() + " rows " + (j + 1) + " and " + (j + 2) + _
                    " contains no before or after strings in any column, and therefore has no effect.  Delete these " + _
                    " rows or insert at least one before or after string in the rows."
            End If
          End If
        End If

      Next j
      'Print "now printing blank filler rules"
      'PrintRowFillerRuleSet(blFiller())
      'Print "now printing write filler rules"
      'PrintRowFillerRuleSet(owFiller())

      ' getting close, all rules formed correctly, just need to apply

    ' here we're done with the fill sheet and exit the loop of sheets
    Exit For
    End If
  Next i

  For i = 0 to sheetCount - 1
    If (StrComp(aSheet, "", 1) = 0) Then
      oSheet = ThisComponent.Sheets.getByIndex(i)
    Else
      oSheet = ThisComponent.Sheets.getByName(aSheet)
    End If
    If (StrComp(oSheet.getName(), filler_sheet_str, 1) <> 0) And _
       (StrComp(oSheet.getName(), total_sheet_str, 1) <> 0) Then

      bool0 = True
      col_index = SheetRowIndexFirstEmpty(oSheet, 0, 0, max_col_index)
      If (col_index < 0) Then
        Print "Could not find the end of the column headers in sheet " + oSheet.getName()
        bool0 = False
      Else
        oCellRange = oSheet.getCellRangeByPosition(0, 0, col_index - 1, 0)
        subtotals_col_idx = CellRangeSearchTrimCol(oCellRange, subtotals_col_str)
        If (subtotals_col_idx < 0) Then
          Print "sheet " + oSheet.getName() + " could not find column matching " + _
                subtotals_col_str + " in the first row."
          bool0 = False
        End If
      End If

      If (bool0 = True) Then
        ' we have an end to columns as well as a subcategory column

        ' TODO: test with empty set
        k = -1
        j = RowFillerSetFillColumnIndicies(blFiller(), oCellRange)
        If (j < 0) Then
          bool0 = False
        Else
          ' k gets at least one valid column
          k = j
          bool0 = True
        End If
        j = RowFillerSetFillColumnIndicies(owFiller(), oCellRange)
        If (j < 0) Then
          bool1 = False
        Else
          ' k gets at least one valid column
          k = j
          bool1 = True
        End If

        If (k >= 0) Then
          For j = 1 to max_row_index
            oCellRange = oSheet.getCellRangeByPosition(0, j, col_index - 1, j)

            If (oSheet.getCellByPosition(k, j).getType() = com.sun.star.table.CellContentType.EMPTY) Then
              'Print CCellPositionToAbsoluteName(totals_columns_index(0), j, i) + " is EMPTY"
              If IsCellRangeEmpty(oCellRange) Then
                ' the current row contains no data, presumably the end of data for this sheet
                'Print oCellRange.AbsoluteName + " is EMPTY"
                Exit For
              End If
            End If

            ' for fill rules it's the first one matching then done
            ' don't cycle through remaining rules and fill a cell twice
            oCell = oSheet.getCellByPosition(subtotals_col_idx, j)
            tmp_str0 = oCell.getString()
            k = 0
            If ((bool0 = True) And (StrComp(tmp_str0, "", 1) = 0)) Then
              ' if the cell is blank fill it with the first matching fill rule
              k = RowFillerRunFillRuleSet(blFiller(), oCellRange, oCell, oSuperCell)
            End If
            If ((bool1 = True) And (k = 0)) Then
              ' if we didn't already fill it then fill it with the first matching fill rule
              RowFillerRunFillRuleSet(owFiller(), oCellRange, oCell, oSuperCell)
            End If
          Next j

        End If


'      'Print oSheet.getName() + " first empty column index " + col_index + " first empty row index " + row_index
'      'Exit Sub
'        oCellRange = oSheet.getCellRangeByPosition(0, 0, col_index - 1, 0)
'        For j = LBound(cdt()) to UBound(cdt())
'          col_index = CellRangeSearchTrimCol(oCellRange, cdt(j).colTitle)
'          If (col_index >= 0) Then
'            If cdt(j).colWidth = -1 Then
'              ' set optimal width
'              oSheet.getColumns().getByIndex(col_index).OptimalWidth = True
'            ElseIf cdt(j).colWidth = 0 Then
'              oSheet.getColumns().getByIndex(col_index).IsVisible = False
'            Else
'              oSheet.getColumns().getByIndex(col_index).Width = cdt(j).colWidth
'            End If
'          Else
'            Print "Could not find column header " + cdt(j).colTitle + " in sheet " + oSheet.getName()
'          End If
'        Next j
      End If
    End If
    If (StrComp(aSheet, "", 1) <> 0) Then
      ' just do this one sheet
      Exit For
    End If
  Next i
        'Print "early exit"
        'Exit Sub

  MsgBox("The fill is complete.")
'Sub ColumnsHideSubset
End Sub


'' SheetBcolMcolMatchFill("20111225", "a", "e", "USPS", "shipping")
'Sub PayPalHistoryAutoFill
'  Dim overwrite As Boolean
'  Dim dontoverwrite As Boolean
'  Dim matchbeginning As Boolean
'  Dim matchanywhere As Boolean
'
'  overwrite = True
'  dontoverwrite = False
'  matchbeginning = False
'  matchanywhere = True
'
'  ' GoDaddy.com, Inc.
'  ' 
'  ' strings beginning with this text that we should only write the tax category of if it is blank
'  'SheetWcolMcolMatchFill("Tax Category", " Name", "GoDaddy.com, Inc.", "C advertising", dontoverwrite, matchbeginning)
'
'  ' strings beginning with this text that we should overwrite the tax category of
'  SheetWcolMcolMatchFill("Tax Category", " Name", "GoDaddy.com, Inc.", "C advertising", overwrite, matchbeginning)
'  SheetWcolMcolMatchFill("Tax Category", " Name", "Zazzle Inc.", "C product expenses", overwrite, matchbeginning)
'  SheetWcolMcolMatchFill("Tax Category", " Name", "eBay Inc.", "C advertising", overwrite, matchbeginning)
'  SheetWcolMcolMatchFill("Tax Category", " Name", "Skype Communications Sarl", "C office expenses", overwrite, matchbeginning)
'
'  'SheetWcolMcolMatchFill("Tax Category", " Name", "Bank Account", "do not total", overwrite, matchbeginning)
'  SheetWcolMcolMatchFill("Tax Category", " Type", "Withdraw Funds to a Bank Account", "DNT withdrawal to bank", overwrite, matchbeginning)
'  SheetWcolMcolMatchFill("Tax Category", " Type", "PayPal Extras MasterCard<sup>&#174;</sup> Payment", "DNT mastercard payment", overwrite, matchbeginning)
'  SheetWcolMcolMatchFill("Tax Category", " Type", "Debit Card Purchase", "DNT debit card purchase", overwrite, matchbeginning)
'  'SheetWcolMcolMatchFill("Tax Category", " Type", "Cancelled Fee", "do not total", overwrite, matchbeginning)
'  SheetWcolMcolMatchFill("Tax Category", " Type", "Cancelled Fee", "DNT paypal canceled fee", overwrite, matchbeginning)
'
'  SheetWcolMcolNcolMatchFill("Tax Category", " Name", " Gross",  0, "PayPal Extras MasterCard<sup>&#174;</sup>", "DNT mastercard", overwrite, matchbeginning)
'  ' a type refund may be a positive value as well, if it's a refund to the MasterCard from a merchant
'  ' in this case we only want corporate tax category customer refund for deductions
'  SheetWcolMcolNcolMatchFill("Tax Category", " Type", " Gross", -1, "Refund", "C customer refund", overwrite, matchbeginning)
'
'  ' =SumSheetNumberTextStringMatch("January", " Gross", " Type", "Refund", 1)
'  ' =SUMSHEETRANGENUMBERTEXTSTRING("February","2:65536","j","f","cancelled fee,Withdraw Funds to a Bank Account", 0)
'End Sub
'
'' SheetBcolMcolMatchFill("20111225", "a", "e", "USPS", "shipping")
'Sub PayPalExtrasMasterCardAutoFill
'  Dim overwrite As Boolean
'  Dim dontoverwrite As Boolean
'  Dim matchbeginning As Boolean
'  Dim matchanywhere As Boolean
'
'  overwrite = True
'  dontoverwrite = False
'  matchbeginning = False
'  matchanywhere = True
'
'  ' strings beginning with this text that we should only write the tax category of if it is blank
'  SheetWcolMcolMatchFill("Tax Category", "Description of Transaction or Credit", "WEST ELM", "C advertising", dontoverwrite, matchbeginning)
'
'  ' strings containing this text that we should only write the tax category of if it is blank
'  'SheetWcolMcolMatchFill("Tax Category", "Description of Transaction or Credit", "xxxx", "yyyy", dontoverwrite, matchanywhere)
'
'  ' strings beginning with this text that we should overwrite the tax category of
'  SheetWcolMcolMatchFill("Tax Category", "Description of Transaction or Credit", "PAYMENT - THANK YOU", "DNT mastercard payment", overwrite, matchbeginning)
'  SheetWcolMcolMatchFill("Tax Category", "Description of Transaction or Credit", "USPS", "C shipping", overwrite, matchbeginning)
'  SheetWcolMcolMatchFill("Tax Category", "Description of Transaction or Credit", "IKEA", "C office expenses", overwrite, matchbeginning)
'
'  ' strings containing this text that we should overwrite the tax category of
'  SheetWcolMcolMatchFill("Tax Category", "Description of Transaction or Credit", "ETSY, INC", "C advertising", overwrite, matchanywhere)
'  SheetWcolMcolMatchFill("Tax Category", "Description of Transaction or Credit", "SKYPE", "C office expenses", overwrite, matchanywhere)
'
'End Sub

Sub AccTotalActiveSheet()
  SheetTotaler(ThisComponent.CurrentController.getActiveSheet().getName())
End Sub

Sub AccTotalAllSheets()
  SheetTotaler("")
End Sub

Sub SheetTotaler(aSheet As String)
  Dim totals_first_header As String
  Dim totals_columns() As String
  Dim totals_columns_index() As Long
  Dim totals_columns_sub_val() As Double
  Dim totals_columns_all_val() As Double
  Dim subtotals_rows() As String
  Dim totals_row_col_vals() As Double
  Dim sheet_tmp_idx() As Long
  Dim sheet_tmp_dbl() As Double
  Dim sheet_row_col_vals() As Double
  Dim col_index As Long
  Dim row_index As Long
  Dim numSheets As Integer
  Dim oSheet As Object
  Dim oSheetBack As Object
  Dim oSheetTotals As Object
  Dim oSheets As Object
  Dim oDocument As Object
  Dim oCellRangeTotals As Object
  Dim b As Boolean
  Dim match As Boolean
  Dim tmp_str As String
  Dim tmp_dbl As Double
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim m As Long
  Dim total_sheet_index As Integer
  Dim subtotals_col_index As Long
  Dim sheet_subtotals_col_index As Long
  'Dim no_tax_category_str As String
  Dim oCellRange As Object
  Dim oDescriptor As Object
  Dim oCell As Object
  'Dim sheet_tmp_idx(3) As Long
  'Dim sheet_tmp_dbl(3) As Double
  'Dim adbl() As Double 
  'Dim aidx() As Long
  'Dim adbl() As Double 
  'Dim aidx() As Long
'
'  aidx(0) = 45
'  aidx(1) = 46
'  aidx(2) = 47
'  aidx(3) = 48
'  adbl(0) = 0.0
'  adbl(1) = 0.0
'  adbl(2) = 0.0
'  adbl(3) = 0.0

  InitAccountingGlobals()
  
  'num_totals_columns = 0
  'num_subtotals_rows = 0

  oDocument = ThisComponent
  ' good for subroutine, not function, fuction will generate error when opening doc
  oSheetBack = oDocument.CurrentController.getActiveSheet()
  oSheets = oDocument.Sheets
  numSheets = oSheets.Count

  For i = 0 To numSheets - 1
    oSheet = oSheets.getByIndex(i)
    If StrComp(oSheet.getName(), total_sheet_str, 1) = 0 Then
      total_sheet_index = i

      ' get rows of subtotals column from Totals sheet
      'subtotals_col_index = SheetFindColumn(oSheet.getName(), subtotals_col_str, 0, 0, max_col_index, 0)
      oCellRange = oSheet.getCellRangeByPosition(0, 0, max_col_index, 0)
      subtotals_col_index = CellRangeSearchTrimCol(oCellRange, subtotals_col_str)
      If (subtotals_col_index < 0) Then
        Print "This subroutine requires a subtotals column with a string in the first row matching """ + subtotals_col_str + """"
        Exit Sub
      End If
      col_index = SheetRowIndexFirstEmpty(oSheet, 0, 0, max_col_index)
      If (col_index < 2) Then
        Print oSheet.getName() + " must have a subtotals column with categories and at least one numbers column to total."
        Exit Sub
      End If
      ' the row is empty from zero all the way to before the first empty column
      ' this allows for the empty string "" in the subtotals column
      ' but also have to be careful, either the next two rows are also blank,
      ' or the next two rows contain the subtotals_totals_str and the all_totals_str strings
      row_index = 0
      For j = 1 to max_row_index
        If (oSheet.getCellByPosition(subtotals_col_index, j).getType() = com.sun.star.table.CellContentType.EMPTY) Then
          If (oSheet.getCellByPosition(subtotals_col_index, j + 1).getType() = com.sun.star.table.CellContentType.EMPTY) Then
            ' two empty cells in a row
            row_index = j
            Exit For
          ElseIf (StrComp(oSheet.getCellByPosition(subtotals_col_index, j + 1).getString(), subtotals_totals_str) = 0) And _
                 (StrComp(oSheet.getCellByPosition(subtotals_col_index, j + 2).getString(), all_totals_str) = 0) Then
            ' blank cell then our final two totals rows
            row_index = j
            Exit For
          End If
        End If
      Next j
      If (row_index < 2) Then
        Print oSheet.getName() + " may only contain " + max_row_index + _
              " subtotals, but also must have at least one category to total."
        Exit Sub
      End If

      ' subtract one because rows start at index one, not zero
      row_index = row_index - 1
      ' subtract one because one of the columns is the subtotals column, not a number column
      ' col_index = 5
      col_index = col_index - 1
      ' col_index = 4

      'VarArrayPreserveAdd(sheet_tmp_idx(), col_index)
      'VarArrayPreserveAdd(sheet_tmp_dbl(), col_index)

      ' this cell range is used over and over to match strings in the sheets to the Totals sheet
      ' this creates the subtotals
      oCellRangeTotals = oSheet.getCellRangeByPosition(subtotals_col_index, 1, subtotals_col_index, row_index)

      ' arrays for the values in the Totals sheet
      ' array indexing starts at 0
      ' subtract one again from both to become array sizes
      row_index = row_index - 1
      col_index = col_index - 1

      ' col_index = 3
      totals_columns = DimArray(col_index)
      ' NOTE: big time trouble here...DimArray works but is not passable as a parameter to a subroutine
      ' need to use ReDim and need to specify the type
      'totals_columns_index = DimArray(col_index)
      ReDim totals_columns_index(col_index) As Long
      totals_row_col_vals = DimArray(col_index, row_index)
      totals_columns_sub_val = DimArray(col_index)
      totals_columns_all_val = DimArray(col_index)
      subtotals_rows = DimArray(row_index)

      ' arrays for the values in a sheet
      sheet_row_col_vals = DimArray(col_index, row_index)
      'aidx = DimArray(col_index)
      'adbl = DimArray(col_index)
      ' NOTE: big time trouble here...DimArray works but is not passable as a parameter to a subroutine
      ' need to use ReDim and need to specify the type
      ReDim sheet_tmp_idx(col_index) As Long
      ReDim sheet_tmp_dbl(col_index) As Double
'      Print "li ui ld ud " + LBound(sheet_tmp_idx()) + " " + _
'      UBound(sheet_tmp_idx()) + " " + _
'      LBound(sheet_tmp_dbl()) + " " + _
'      UBound(sheet_tmp_dbl()) 

      'Print "the number of subtotals rows is " + num_subtotals_rows
      'Print "the number of totals columns is " + num_totals_columns
      'Exit Sub

      ' initialize all arrays, arrays containing our strings and zeros for all number arrays
      For j = LBound(subtotals_rows()) To UBound(subtotals_rows())
        subtotals_rows(j) = oSheet.getCellByPosition(subtotals_col_index, j + 1).getString()
        'Print "ARW subtotals_rows(" + j + ") is " + subtotals_rows(j)
      Next j

      ' add one for the subtotals column, but don't index the array based on this
      k = 0
      For j = LBound(totals_columns()) To UBound(totals_columns()) + 1
        If (j <> subtotals_col_index) Then
          totals_columns(k) = oSheet.getCellByPosition(j, 0).getString()
          k = k + 1
        End If
      Next j
      ' first header used to compare a previously written location just to the right of
      ' the normal headers on each sheet
      totals_first_header = oSheet.getCellByPosition(0, 0).getString()

      For j = LBound(totals_columns()) To UBound(totals_columns())
        For k = LBound(subtotals_rows()) To UBound(subtotals_rows())
          totals_row_col_vals(j, k) = 0.0
        Next k
        totals_columns_sub_val(j) = 0.0
        totals_columns_all_val(j) = 0.0
      Next j

      ' we have data from master sheet, exit for
      Exit For
    End If

  Next i

  For i = 0 To numSheets - 1
  'For i = 1 To 1
  'For i = 2 To 2
    If (StrComp(aSheet, "", 1) = 0) Then
      oSheet = ThisComponent.Sheets.getByIndex(i)
    Else
      oSheet = ThisComponent.Sheets.getByName(aSheet)
    End If

    If ((StrComp(oSheet.getName(), total_sheet_str, 1) <> 0) And _
        (StrComp(oSheet.getName(), filler_sheet_str, 1) <> 0)) Then

      ' find the first column with empty first row cell
      b = False
      col_index = SheetRowIndexFirstEmpty(oSheet, 0, 0, max_col_index)
      If col_index < 0 Then
        Print oSheet.getName() + " does not have a column with an empty first row cell"
      Else
        ' determine if either the cell range is blank or we are presumably overwriting an area which we previously wrote
        col_index = col_index + 1
        'b = TestCellRangeEmptyByPosition(oSheet.RangeAddress.sheet, col_index, 0, col_index + num_totals_columns - 1, num_subtotals_rows)
        ' row adds one because they start at one
        ' col adds one because we took the subtotals column out of the array and need to account for this
        oCellRange = oSheet.getCellRangeByPosition(col_index, 0, col_index + UBound(totals_columns()) + 1, UBound(subtotals_rows()) + 1)
        b = IsCellRangeEmpty(oCellRange)
        If b = False Then
          If StrComp(totals_first_header, oSheet.getCellByPosition(col_index, 0).getString(), 1) = 0 Then
            b = True
          Else
            ' TODO: test this printout
            ' ReturnAbsoluteAddressAsString(x, y As Long)
            Print oSheet.getName() + " cell " + CCellPositionToAbsoluteName(col_index, 0, i) + " is not empty, and its content " _
            + oSheet.getCellByPosition(col_index, 0).getString() + " is not equal to " + totals_first_header
            'Print "This subroutine requires a subtotals column with a string in the first row matching """ + subtotals_col_str + """"
            'b = False
          End If
        End If
      End If

      ' make sure we can find all the columns that are in the Totals sheet
      If b = True Then
        oCellRange = oSheet.getCellRangeByPosition(0, 0, col_index - 2, 0)
        For j = LBound(totals_columns()) To UBound(totals_columns())
          'totals_columns_index(j) = SheetFindColumn(oSheet.getName(), totals_columns(j), 0, 0, max_col_index, 0)
          totals_columns_index(j) = CellRangeSearchTrimCol(oCellRange, totals_columns(j))
          If (totals_columns_index(j) < 0) Then
            b = False
            j = oCellRange.getRangeAddress.StartColumn
            k = oCellRange.getRangeAddress.EndColumn
            Print "Could not find a cell matching """ + totals_columns(j) + """ between cell " + _
                  CCellPositionToAbsoluteName(j, 0, i) + " and cell " + CCellPositionToAbsoluteName(k, 0, i)
            Print "Please ensure that all sheets contain the headers of the " + total_sheet_str + " sheet."
            Exit For
          End If
        Next j
        sheet_subtotals_col_index = CellRangeSearchTrimCol(oCellRange, subtotals_col_str)
        If (sheet_subtotals_col_index < 0) Then
          b = False
          j = oCellRange.getRangeAddress.StartColumn
          k = oCellRange.getRangeAddress.EndColumn
          Print "Could not find a cell matching """ + subtotals_col_str + """ between cell " + _
                CCellPositionToAbsoluteName(j, 0, i) + " and cell " + CCellPositionToAbsoluteName(k, 0, i)
          Print "Please ensure that all sheets contain the headers of the " + total_sheet_str + " sheet."
          Exit Sub
        End If
      End If

      If b = True Then

        ' zero out the arrays that hold the sheets subtotals totals, all categories totals,
        ' and individual subcategories totals for each numbers column
        For j = LBound(totals_columns()) To UBound(totals_columns())
          For k = LBound(subtotals_rows()) To UBound(subtotals_rows())
            sheet_row_col_vals(j, k) = 0.0
          Next k
          sheet_tmp_dbl(j) = 0.0
        Next j

        ' add one for the subtotals column, but don't index the array based on this
        k = 0
        For j = LBound(totals_columns()) To UBound(totals_columns()) + 1
          If (j <> subtotals_col_index) Then
            sheet_tmp_idx(k) = col_index + j
            k = k + 1
          End If
        Next j

        ' now just calculate all the values and populate the rest
        For j = 1 To max_row_index
          If (oSheet.getCellByPosition(totals_columns_index(0), j).getType() = com.sun.star.table.CellContentType.EMPTY) Then
            'Print CCellPositionToAbsoluteName(totals_columns_index(0), j, i) + " is EMPTY"
            oCellRange = oSheet.getCellRangeByPosition(0, j, col_index - 2, j)
            If IsCellRangeEmpty(oCellRange) Then
              'Print oCellRange.AbsoluteName + " is EMPTY"
              Exit For
            End If
          End If
          tmp_str = oSheet.getCellByPosition(sheet_subtotals_col_index, j).getString()
          'oCellRange = oSheet.getCellRangeByPosition(subtotals_col_index + col_index, 1, _
          '                                           subtotals_col_index + col_index, num_subtotals_rows)
          oDescriptor = oCellRangeTotals.createSearchDescriptor()
          oDescriptor.SearchString = tmp_str
          ' If true, the search will match only complete words. One white space character can cause a mismatch.
          oDescriptor.SearchWords = True
          ' Instead, set this to false to not have to deal with white space
          'oDescriptor.SearchWords = False
          oCell = oCellRangeTotals.findFirst(oDescriptor)
          If Not IsNull(oCell) Then
            'Print "oCell.CellAddress.Row " + oCell.CellAddress.Row
            'Print "oCell.getString() " + oCell.getString()
            'Print "tmp_str " + tmp_str + " row_index " + row_index + " sheet " + i
            'col_index = oCell.CellAddress.Column
            For k = LBound(totals_columns()) To UBound(totals_columns())
              'Print "k " + k + " .Row " + oCellAddress.Row

              ' NOTE: when inserting a sheet from a file and under other options checking quoted field as text all of our numbers are
              ' inserted in the sheet as text values.  So, even if the text appears to be a negative number like -5.95 or ($8.00), it
              ' is still evaluated as text, so something like "= C3 < 0" will return False.  In order to convert the text to a value,
              ' use the Value function, so something like "Value(C3) < 0" now returns true, since it converted the string to a number.
              ' For these functions that means calling the getString function for a cell then converting that to a double.  Speed
              ' tests with the large data sets give comparable results: around 58 seconds for the getValue version and around 66
              ' seconds for the Value(getString) version.  Since there is not a great difference, I'll just take the more flexible and
              ' slower way.
              sheet_row_col_vals(k, oCell.CellAddress.Row - 1) = sheet_row_col_vals(k, oCell.CellAddress.Row - 1) + _
                                                                 CellGetOnlyValue(oSheet.getCellByPosition(totals_columns_index(k), j))
'                                                                 oSheet.getCellByPosition(totals_columns_index(k), j).getValue()
            Next k
          Else
            Print "In sheet " + oSheet.getName() + ", row " + (j + 1) + ", in the " + subtotals_col_str + " column, " + _
                  "there is a string """ + tmp_str + """ which is not a subtotal found in the " + total_sheet_str + " sheet.  " + _
                  "Add this string to a row of the " + total_sheet_str + " sheet or it will not be totaled."
          End If
        Next j

        ' do the first row which are the column headers of the Totals sheet
        oSheet.getCellByPosition(col_index + subtotals_col_index, 0).setString(subtotals_col_str)
        For j = LBound(totals_columns()) To UBound(totals_columns())
          oSheet.getCellByPosition(sheet_tmp_idx(j), 0).setString(totals_columns(j)) 
        Next j
        ' fill out the number values of the final two rows, 
        ' one for total of subtotals, and one for independent total of all categories
        ' this helps us check that the sum of all subtotals equals the overall total
        ' first clear that row if we've added or subtracted subcategories so the function can work
        ' need to add one to the right column index because we took the subtotals column out of the array
        oCellRange = oSheet.getCellRangeByPosition(col_index, UBound(subtotals_rows()) + 2, _
                                                   col_index + UBound(totals_columns()) + 1, UBound(subtotals_rows()) + 2)
        oCellRange.clearContents(23)

        ' do the subtotals column which are the number of rows in the Totals sheet
        For j = LBound(subtotals_rows()) To UBound(subtotals_rows())
          oSheet.getCellByPosition(col_index + subtotals_col_index, j + 1).setString(subtotals_rows(j))
        Next j
        ' in the subtotals column, after all subcategories and one blank row, write two strings
        oSheet.getCellByPosition(col_index + subtotals_col_index, UBound(subtotals_rows()) + 3).setString(subtotals_totals_str) 
        oSheet.getCellByPosition(col_index + subtotals_col_index, UBound(subtotals_rows()) + 4).setString(all_totals_str) 
        ' fill in the remaining rows under all these columns with strings and numeric values
        m = 0
        For j = LBound(totals_columns()) To UBound(totals_columns()) + 1
          If (j <> subtotals_col_index) Then
            ' in all other columns, write totals of all subcategories, as well as the sum of these totals
            For k = LBound(subtotals_rows()) To UBound(subtotals_rows())
              oSheet.getCellByPosition(col_index + j, k + 1).setValue(sheet_row_col_vals(m, k))
              totals_row_col_vals(m, k) = totals_row_col_vals(m, k) + sheet_row_col_vals(m, k)
            Next k
            m = m + 1
          End If
        Next j
          ' need to put values in these first
          ' not a subtotal string, must calculate a total based on the string
          ' potential difference between subcategories total and all categories total
          ' is if there is a label in the subcategory column that is not listed in the Totals sheet

          ' TODO: change these to arrays of Long and Double values
          ' then loop through and set values
'          tmp_dbl = SumSheetNumberiTextiStringMatch( _
'            oSheet.getName(), col_index + j, col_index + subtotals_col_index, dnt_str, 0)
'          oSheet.getCellByPosition(col_index + j, num_subtotals_rows + 2).setValue(tmp_dbl)
'
'          tmp_dbl = SumSheetNumberiTextiStringMatch( _
'            oSheet.getName(), totals_columns_index(j), sheet_subtotals_col_index, dnt_str, 0)
'          oSheet.getCellByPosition(col_index + j, num_subtotals_rows + 3).setValue(tmp_dbl)
'
'          ' Here are the lines to instead enter them as formulas, perhaps helping debugging
'          '
'          'oSheet.getCellByPosition(col_index + j, num_subtotals_rows + 2).setFormula( _
'          '  "=SumSheetNumberiTextiStringMatch(""" & oSheet.getName() & """; " & col_index + j & "; " _
'          '    & col_index + subtotals_col_index & "; """ & dnt_str & """; 0)" )
'          'oSheet.getCellByPosition(col_index + j, num_subtotals_rows + 3).setFormula( _
'          '  "=SumSheetNumberiTextiStringMatch(""" & oSheet.getName() & """; """ & totals_columns_index(j) & """; """ _
'          '  & totals_columns_index(subtotals_col_index) & """; """ & dnt_str & """; 0)" )
'
'          'Print "col_index + j "+ col_index + j + " num_subtotals_rows + 3 " + num_subtotals_rows + 3
'          'Print "oSheet.getName() " + oSheet.getName() + " totals_columns_index(k) " + totals_columns_index(k) 
'          'Print "totals_columns_index(subtotals_col_index) " + totals_columns_index(subtotals_col_index) + " dnt_str " + dnt_str
'
'          totals_columns_sub_val(j) = totals_columns_sub_val(j) + _
'                                      oSheet.getCellByPosition(col_index + j, num_subtotals_rows + 2).getValue()
'          totals_columns_all_val(j) = totals_columns_all_val(j) + _
'                                      oSheet.getCellByPosition(col_index + j, num_subtotals_rows + 3).getValue()


        ' for each numbers column, total the subtotals column that we've written to the right of all the data
        SumSheetNumberiaTextiStringMatch(sheet_tmp_dbl(), oSheet.getName(), sheet_tmp_idx(), (col_index + subtotals_col_index), dnt_str, 0)
        For j = LBound(totals_columns()) To UBound(totals_columns())
          'Print "Column header " + totals_columns(j) + " index " + totals_columns_index(j) + " double value " + sheet_tmp_dbl(j) + "."
          oSheet.getCellByPosition(sheet_tmp_idx(j), UBound(subtotals_rows()) + 3).setValue(sheet_tmp_dbl(j))
          totals_columns_sub_val(j) = totals_columns_sub_val(j) + _
                                      oSheet.getCellByPosition(sheet_tmp_idx(j), UBound(subtotals_rows()) + 3).getValue()
        Next j

        ' for each numbers column, total the subtotals column of this sheet
        '
        ' This can be different than the totals produced above if we encountered a string in the subtotals
        ' column that is not present in the totals sheet.  In this case it would not be totaled in any 
        ' subcategory column and therefore not included in the values we write to the right of all the data
        SumSheetNumberiaTextiStringMatch(sheet_tmp_dbl(), oSheet.getName(), totals_columns_index(), sheet_subtotals_col_index, dnt_str, 0)
        For j = LBound(totals_columns()) To UBound(totals_columns())
          'Print "Column header " + totals_columns(j) + " index " + totals_columns_index(j) + " double value " + sheet_tmp_dbl(j) + "."
          oSheet.getCellByPosition(sheet_tmp_idx(j), UBound(subtotals_rows()) + 4).setValue(sheet_tmp_dbl(j))
          totals_columns_all_val(j) = totals_columns_all_val(j) + _
                                      oSheet.getCellByPosition(sheet_tmp_idx(j), UBound(subtotals_rows()) + 4).getValue()
        Next j

      ' end if we're going to write data on the end of the sheet
      End If

    ' end if not Totals sheet
    End If

    If (StrComp(aSheet, "", 1) <> 0) Then
      ' just do this one sheet
      'Exit For
      Exit Sub
    End If

  ' cycle through remaining sheets in document
  Next i

  oSheet = oSheets.getByIndex(total_sheet_index)

  ' in the subtotals column, after all subcategories and one blank row, write two strings
  oSheet.getCellByPosition(subtotals_col_index, UBound(subtotals_rows()) + 3).setString(subtotals_totals_str) 
  oSheet.getCellByPosition(subtotals_col_index, UBound(subtotals_rows()) + 4).setString(all_totals_str) 
  ' get columns of row 0 from Totals sheet
  k = 0
  For i = LBound(totals_columns()) To UBound(totals_columns()) + 1
    If (i <> subtotals_col_index) Then
      ' in all other columns, write totals of all subcategories, as well as the sum of these totals
      oSheet.getCellByPosition(i, UBound(subtotals_rows()) + 3).setValue(totals_columns_sub_val(k))
      oSheet.getCellByPosition(i, UBound(subtotals_rows()) + 4).setValue(totals_columns_all_val(k))
      For j = LBound(subtotals_rows()) To UBound(subtotals_rows())
        oSheet.getCellByPosition(i, j + 1).setValue(totals_row_col_vals(k, j))
      Next j
      ' k is used so we don't over-index our array, i.e. skip the subtotals_col_index
      k = k + 1
    End If
  Next i

  ThisComponent.CurrentController.setActiveSheet(oSheetBack)

  MsgBox("The total is complete.")
'Sub AccTotalAllSheets()
End Sub

Function SumSheetNumberTextStringMatch(sheet As String, num_col As String, str_col As String, match As String, addmatching As Boolean) As Double
  Dim num_index As Long
  Dim str_index As Long
  Dim oSheet As Object
  Dim oCellRange As Object

  'MsgBox("sheet is " + "20110125")
  'MsgBox("sheet is " + sheet)

  'oSheet = ThisComponent.Sheets.getByName("20110125")
  oSheet = ThisComponent.Sheets.getByName(sheet)
  oCellRange = oSheet.getCellRangeByPosition(0, 0, max_col_index, 0)
  'tax_column_index = SheetFindColumn(oSheet.getName(), "Tax Category", 0, 0, 99, 0)
  str_index = CellRangeSearchTrimCol(oCellRange, str_col)
  'str_index = SheetFindColumn(oSheet.getName(), str_col, 0, 0, 99, 0)
  'num_index = SheetFindColumn(oSheet.getName(), num_col, 0, 0, 99, 0)
  num_index = CellRangeSearchTrimCol(oCellRange, num_col)

  SumSheetNumberTextStringMatch = SumSheetNumberiTextiStringMatch(sheet, num_index, str_index, match, addmatching)
End Function

Function SumSheetNumberiTextiStringMatch(sheet As String, num_index As Long, str_index As Long, match As String, addmatching As Boolean) As Double
  Dim l(0) As Long
  Dim d(0) As Double

  l(0) = num_index
  SumSheetNumberiaTextiStringMatch(d(), sheet, l(), str_index, match, addmatching)
  SumSheetNumberiTextiStringMatch = d(0)
End Function

Sub SumSheetNumberiaTextiStringMatch(ret_dbl() As Double, sheet As String, num_index() As Long, _
                                     str_index As Long, match As String, addmatching As Boolean)
  'Dim TheSum As Double
  Dim b As Boolean
  Dim oFunctionAccess As Variant
  Dim col_index As Long
  Dim min_index As Long
  Dim max_index As Long
  Dim row_index As Long
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim oSheet
  Dim tval As String
  Dim m As Variant
  Dim word As String
  Dim oDescriptor As Object
  Dim oCellRange As Object

  oSheet = ThisComponent.Sheets.getByName(sheet)

  'Print "UBound(num_index()) is " + UBound(num_index())
  
  If StrComp (match, "", 0) = 0 Then
    m = split("hack", ",")
    m(0) = ""
  Else
    m = split(match, ",")
  End If

  ' initialize all doubles and ensure non-negative column indices
  b = True
  min_index = str_index
  If (LBound(num_index()) = LBound(ret_dbl())) And _
     (UBound(num_index()) = UBound(ret_dbl())) Then
    For i = LBound(num_index()) To UBound(num_index())
      If (num_index(i) < 0) Then
        b = False
        Print oSheet.getName() + " number columns array index is negative"
        Exit For
      ElseIf (num_index(i) < min_index) Then
        min_index = num_index(i)
      End If
      ret_dbl(i) = 0.0
    Next i
  Else
    Print oSheet.getName() + " number columns array size does not match doubles array size"
    b = False
  End If

  If (b = True) Then
    ' find an empty column in the first row starting from the lesser of the
    ' string column index and the number column index
    col_index = SheetRowIndexFirstEmpty(oSheet, 0, min_index, max_col_index)
    If col_index < 0 Then
      Print oSheet.getName() + " does not have a column with an empty first row cell starting from " + min_index
    End If
  End If

  If ((col_index >= 0) And (str_index >= 0) And (b = True)) Then
    For i = 1 To max_row_index
      ' the numbers columns are usually always full, as opposed to the string index column
      ' testing for EMPTY in any numbers column therefore is better than in the string column
      ' because if we're EMPTY there then we're most likely empty in the whole row
      If (oSheet.getCellByPosition(num_index(0), i).getType() = com.sun.star.table.CellContentType.EMPTY) Then
        'Print "test empty i is " + i
        'b = TestCellRangeEmptyByPosition(oSheet.RangeAddress.sheet, min_index, i, col_index - 1, i)
        oCellRange = oSheet.getCellRangeByPosition(min_index, i, col_index - 1, i)
        b = IsCellRangeEmpty(oCellRange)
        If b = True Then
          'Print "index = " + i
          Exit For
        End If
      End If
      ' iterate over range of rows
      If addmatching Then
        ' add matching
        ' TODO: in all functions check StrComp
        For Each word in m
          ' iterate over the strings in to match
          tval = oSheet.getCellByPosition(str_index, i).getString()
          ' 1 for case-sensitive, 0 for case-insensitive
          'If StrComp (tval, word, 0) = 0 Then
          ' let's try out the bash stype wildcard string compare
          If StrStarEndComp (tval, word, 0) = 0 Then
            For j = LBound(num_index()) To UBound(num_index())
              ret_dbl(j) = ret_dbl(j) + CellGetOnlyValue(oSheet.getCellByPosition(num_index(j), i))
'              ret_dbl(j) = ret_dbl(j) + oSheet.getCellByPosition(num_index(j), i).getValue()
            Next j
            Exit For
          End If
        Next word
      Else
        ' add non-matching
        j = 0
        k = 0
        For Each word in m
          ' iterate over the strings in to match
          tval = oSheet.getCellByPosition(str_index, i).getString()
          ' k is the total loop count
          k = k + 1
          If StrStarEndComp (tval, word, 0) <> 0 Then
            ' j is incremented if there is no match
            j = j + 1
          End If
        Next word
        If j = k Then
          For j = LBound(num_index()) To UBound(num_index())
            ret_dbl(j) = ret_dbl(j) + CellGetOnlyValue(oSheet.getCellByPosition(num_index(j), i))
'            ret_dbl(j) = ret_dbl(j) + oSheet.getCellByPosition(num_index(j), i).getString()
          Next j
        End If
      End If
    Next i
  End If
End Sub

' The whole point of this routine was to match the gross profit of the spreadsheet to the gross profit
' reported on the 2011 1099-K form from PayPal.  Sheet September row 88 has a payment received of $5.45
' and no PayPal fee associated with it.  In then end I was off by 5.45 from the 1099-K form so
' apparently PayPal didn't count this toward gross profit and therefore I will categorize it as a
' personal refund from a personal purchase.  Don't really care what it was, just want the spreadsheet
' to match the 1099-K.  Sheet October row 58 also has such a payment, but apparently PayPal included
' this in the gross profit total for 2011, even though they did not extract a fee from the gross
' payment similar to September row 88.
' 
' Going to change the name of this subroutine to print any information that might help with cells that
' still need filling or that might not be filled correctly.  In particular, going to add a check for
' any blank cell that has a negative gross, that being an expense.  Pretty sure I won't find anything
' since the totals match the gross total on the 2011 1099-K form, but might be good for the future.
' Also, going to print more verbose messages to the screen.
' 
Sub AccPayPalHistoryCheckFills()
  Dim b As Boolean
  Dim use_ncol As Boolean
  Dim match_two As Boolean
  'Dim lower_row As Long
  'Dim upper_row As Long
  Dim col_index As Long
  Dim row_index As Long
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim l As Integer
  Dim numSheets As Integer
  Dim oSheet
  Dim oSheetBack
  Dim oSheets
  Dim oDocument
  Dim state_index As Long
  Dim add_line_one_index As Long
  Dim tval As String
  Dim tmp_str0 As String
  Dim tmp_str1 As String
  Dim tmp_str2 As String
  Dim m() As String
  Dim n() As String
  Dim o() As String
  Dim word As String
  Dim tmpdbl0 As Double
  Dim tmpdbl1 As Double
  Dim oDescriptor As Object
  Dim oCellRange As Object

  'Dim tax_category_str As String
  Dim status_str As String
  Dim type_str As String
  Dim categorize_str As String

  Dim status_col_str As String
  Dim gross_col_str As String
  Dim fee_col_str As String
  Dim type_col_str As String

  Dim status_col_idx As Long
  Dim gross_col_idx As Long
  Dim fee_col_idx As Long
  Dim type_col_idx As Long

  Dim subtotals_col_idx As Long

  InitAccountingGlobals()

  'tax_category_str = "DNT mastercard positive,DNT paypal canceled fee"
  status_str = "Removed,Placed"
  type_str = "Currency Conversion"

  categorize_str = "DNT personal refund"


  status_col_str = "Status"
  gross_col_str = "Gross"
  fee_col_str = "Fee"
  type_col_str = "Type"

  oDocument = ThisComponent
  oSheetBack = oDocument.CurrentController.getActiveSheet()
  oSheets = oDocument.Sheets
  'numSheets = oSheets.Count

  'm() = split(tax_category_str, ",")
  'For Each word in m
    'tmp_str0 = tmp_str0 + word + " "
  'Next word
  'tmp_str0 = Trim(tmp_str0)

  n() = split(status_str, ",")
  For Each word in n
    ' status column
    tmp_str1 = tmp_str1 + """" + word + """ "
  Next word
  tmp_str1 = Trim(tmp_str1)

  o() = split(type_str, ",")
  For Each word in o
    ' status column
    tmp_str2 = tmp_str2 + """" + word + """ "
  Next word
  tmp_str2 = Trim(tmp_str2)

  ' do all rows until we hit no more data
  'lower_row = 2 - 1
  'upper_row = 65536 - 1

  numSheets = oSheets.Count
  For i = 0 To numSheets - 1
  'For i = 0 To 0
    oSheet = oSheets.getByIndex(i)

  'total_sheet_str = "Totals"
  'subtotals_col_str = "Tax Category"
    If (StrComp(oSheet.getName(), total_sheet_str, 1) <> 0) And _
       (StrComp(oSheet.getName(), filler_sheet_str, 1) <> 0) Then

      col_index = SheetRowIndexFirstEmpty(oSheet, 0, 0, max_col_index)
      If (col_index < 0) Then
        Print "Could not find the end of the column headers in sheet " + oSheet.getName()
        Exit Sub
      Else
        oCellRange = oSheet.getCellRangeByPosition(0, 0, col_index - 1, 0)

        subtotals_col_idx = CellRangeSearchTrimCol(oCellRange, subtotals_col_str)
        status_col_idx = CellRangeSearchTrimCol(oCellRange, status_col_str)
        gross_col_idx = CellRangeSearchTrimCol(oCellRange, gross_col_str)
        fee_col_idx = CellRangeSearchTrimCol(oCellRange, fee_col_str)
        type_col_idx = CellRangeSearchTrimCol(oCellRange, type_col_str)
      End If


'      ' remove these later
'      ' look for "480 NE 30th"
'      add_line_one_index = CellRangeSearchTrimCol(oCellRange, "Address Line 1")
'      ' look for "FL"
'      state_index = CellRangeSearchTrimCol(oCellRange, "State/Province/Region/County/Territory/Prefecture/Republic")

      If ((subtotals_col_idx < 0) Or (status_col_idx < 0) Or (gross_col_idx < 0) Or (fee_col_idx < 0) Or (type_col_idx < 0)) Then
          Print "Could not find one or more of column headers " + subtotals_col_str + ", " + status_col_str + _
                ", " + gross_col_str + ", " + type_col_str + ", and " + fee_col_str + " in sheet " + oSheet.getName() + "."
          Exit Sub
      Else
        ' iterate over range of rows
        ' the row is empty from zero all the way to before the first empty column
'        row_index = SheetFindEmptyRow(oSheet, 0, 1, col_index - 1, max_row_index)
'        If (row_index < 2) Then
'          Print oSheet.getName() + " may only contain " + max_row_index + _
'                " rows of fill data, but also must have at least one row of data."
'          Exit Sub
'        End If

        For j = 1 To max_row_index
          If (oSheet.getCellByPosition(gross_col_idx, j).getType() = com.sun.star.table.CellContentType.EMPTY) Then
            oCellRange = oSheet.getCellRangeByPosition(0, j, col_index - 1, j)
            If IsCellRangeEmpty(oCellRange) Then
              Exit For
            End If
          End If

          ' even if the array is empty we'll still pass these tests since
          ' UBound(m()) = UBount(n()) = -1, so k = UBound(n()) + 1

          'k = 0

          'If (oSheet.getCellByPosition(subtotals_col_idx, j).getType() = com.sun.star.table.CellContentType.EMPTY) Then
          '  k = 0
          'Else
          '  k = -1
          'End If

          ' rows we can ignore if the status column matches strings we define
          ' works even if there are no such strings we care about
          tval = oSheet.getCellByPosition(status_col_idx, j).getString()
          k = 0
          For Each word in n
            ' status column
            If StrComp(tval, word, 1) <> 0 Then
              k = k + 1
            End If
          Next word

          tval = oSheet.getCellByPosition(type_col_idx, j).getString()
          l = 0
          For Each word in o
            ' status column
            If StrComp(tval, word, 1) <> 0 Then
              l = l + 1
            End If
          Next word


          tval = oSheet.getCellByPosition(subtotals_col_idx, j).getString()
          'For Each word in m
            'If StrComp(tval, word, 1) <> 0 Then
              ' tax column
              'k = k + 1
            'End If
          'Next word
          If (k = (UBound(n()) + 1)) And (l = (UBound(o()) + 1)) Then
            If (StrComp(tval, "", 1) = 0) Then
              ' tax category empty, check for negative gross values and (positive gross and zero fee) values
              If (CellGetOnlyValue(oSheet.getCellByPosition(gross_col_idx, j)) < 0) Then
'              If (oSheet.getCellByPosition(gross_col_idx, j).getValue() < 0) Then
                Print "In " + oSheet.getName() + ", row " + (j + 1) + ", there is no string in the " + subtotals_col_str + " column, " + _
                      "the text in the " + status_col_str + " column does not match any of the strings [ " + tmp_str1 + " ], " + _
                      "and there is a negative value in the " + gross_col_str + " column.  This might indicate a " + _
                      "corporate expense that still needs to be categorized and totaled so that it can become a tax deduction."
              ElseIf (CellGetOnlyValue(oSheet.getCellByPosition(fee_col_idx, j)) = 0) Then
'              ElseIf (oSheet.getCellByPosition(fee_col_idx, j).getValue() = 0) Then
                  'Print oSheet.getName() + " row " + (j + 1) + ", the text in the " + subtotals_col_str + " column does not match " + _
                        '"any of the strings [ " + tmp_str0 + " ], the text in the " + status_col_str + " column does not match " + _
                        '"any of the strings [ " + tmp_str1 + " ], there is a positive value in the " + gross_col_str + _
                  'Print "In " + oSheet.getName() + ", row " + (j + 1) + ", the text in the " + status_col_str + " column does not match " + _
                  '      "any of the strings [ " + tmp_str1 + " ], there is a positive value in the " + gross_col_str + _
                  Print "In " + oSheet.getName() + ", row " + (j + 1) + ", there is no string in the " + subtotals_col_str + " column, " + _
                        "the text in the " + status_col_str + " column does not match any of the strings [ " + tmp_str1 + " ], " + _
                        "the text in the " + type_col_str + " column does not match any of the strings [ " + tmp_str2 + " ], " + _
                        "there is a positive value in the " + gross_col_str + " column, " + _
                        "and there is a zero value in the " + fee_col_str + " column.  In the past, this has led to the totals of " + _
                        "the spreadsheet not matching the totals on the tax form 1099-K sent from PayPal.  " + _
                        "If there is such a mismatch, then simply categorize the payment as something like " + _
                        categorize_str + " so that the spreadsheet totals can match the 1099-K totals exactly."
              End If
            End If
          End If

        ' go to the next row
        Next j

      ' end if we have all of our columns in this sheet, probably could have used an exit instead
      End If

    ' end if we're not the totals or fillers sheets
    End If

  ' go to the next sheet
  Next i
  ThisComponent.CurrentController.setActiveSheet(oSheetBack)
End Sub

' the row up, this row, or the row down begins with a string and the numbers values are opposites
Function UpThisDownBeginsNumOpp(cell_abs_name As String, begin_str As String, num_col_str As String) As Boolean
  On Error Goto ErrorHandler
  Dim oConv, oSheet, oCell, oCellRange As Object
  Dim str() As String
  Dim c, r, j As Long
  Dim tmp0_str As String
  Dim tmp1_str As String
  Dim tmp0_dbl As Double
  Dim tmp1_dbl As Double
  'Dim i As Integer
  Dim bool0 As Boolean
  Dim bool1 As Boolean
  Dim err_is_cell As Boolean
  Dim verbose As Boolean
  Dim empty_col_idx As Long
  Dim num_col_idx As Long
  Dim cdt() As ColumnDisplayType
  Static save_empty_col_idx As Long
  Static save_sheet_name As String
  'Static save_num_col_idx As Long
  'Static save_num_col_str As Long
  Static save_cdt() As ColumnDisplayType

  err_is_cell = True
  verbose = False

  InitAccountingGlobals()

  bool0 = False


  ' for instance
  '
  ' UpThisDownBeginsNumOpp("$September.$E$60"; "PayPal Extras MasterCard<sup>&#174;</sup>"; "Gross")
  ' UpThisDownBeginsNumOpp("$September.$E$61"; "PayPal Extras MasterCard<sup>&#174;</sup>"; "Gross")
  ' UpThisDownBeginsNumOpp("$September.$E$62"; "PayPal Extras MasterCard<sup>&#174;</sup>"; "Gross")
  ' UpThisDownBeginsNumOpp("$September.$E$63"; "PayPal Extras MasterCard<sup>&#174;</sup>"; "Gross")
  '
  ' Now thanks to the array implementation more than one number column can be stored and remembered.  Note that for
  ' future it is still better to keep track of all sheets all columns that are accessed in any subroutine or function as
  ' a global variable.  This is left as a future exercise for now.
  '
  ' UpThisDownBeginsNumOpp("$September.$E$63"; "PayPal Extras MasterCard<sup>&#174;</sup>"; "Net")


  'oSheet = ThisComponent.Sheets.getByName(str(LBound(str())))

  bool1 = False
  oCell = CAbsoluteNameToCell(cell_abs_name)
  If IsNull(oCell) Then
    Print "cound not locate cell whose absolute name is " + cell_abs_name
    GoTo ErrorHandler
  End If
  err_is_cell = False

  oSheet = ThisComponent.Sheets.getByIndex(oCell.CellAddress.Sheet)

  str() = split(cell_abs_name, ".")
'  If ((StrComp(str(0), save_sheet_name, 1) <> 0) Or (StrComp(save_num_col_str, num_col_str, 1) <> 0)) Then
'    'Print "old sheet " + save_sheet_name + " new sheet " + str(0)
'    ' a new sheet so neet to look up stuff again
'    save_sheet_name = str(0)
'    col_index = SheetRowIndexFirstEmpty(oSheet, 0, 0, max_col_index)
'    If (col_index < 0) Then
'      Print "Could not find the end of the column headers in sheet " + oSheet.getName()
'    Else
'      'save_num_col_idx = SheetFindColumn(oSheet.getName(), num_col_str, 0, 0, 99, 0)
'      oCellRange = oSheet.getCellRangeByPosition(0, 0, col_index - 1, 0)
'      save_num_col_idx = CellRangeSearchTrimCol(oCellRange, num_col_str)
'      If (save_num_col_idx >= 0) Then
'        save_num_col_str = num_col_str
'        bool1 = True
'      End If
'    End If
'  Else
'    'Print "old sheet " + save_sheet_name
'    If (save_num_col_idx >= 0) Then
'      bool1 = True
'    End If
'  End If

  If (StrComp(str(0), save_sheet_name, 1) = 0) Then
    'Print "Sheet matches " + save_sheet_name
    ' Sheet names match.  At least we can avoid looking up the empty col index since we have looked it up at least once previously.
    empty_col_idx = save_empty_col_idx
    For j = LBound(save_cdt()) To UBound(save_cdt())
      If (StrComp(save_cdt(j).colTitle, num_col_str, 1) = 0) Then
        ' we found our number column index previously, but it still could be -1
        ' so set bool1 to True to bypass a lookup
        num_col_idx = save_cdt(j).colWidth
        bool1 = True
        If (verbose = True) Then
          Print "Function call (""" + cell_abs_name + """, """ + begin_str + """, """ + num_col_str + """) the index " + _
                num_col_idx + " of the """ + num_col_str + """ was recovered from saved data obtained from a prior " + _
                "call involving this same sheet."
          'Print "From saved data, the sheet " + oSheet.getName() + " col index " + save_cdt(j).colWidth + _
          '      " has header " + save_cdt(j).colTitle
        End If
      End If
    Next j
  Else
    'Print "New sheet " + oSheet.getName()
    ' save the name of the new sheet such as "$January"
    save_sheet_name = str(0)

    ' find the first empty column and save it
    j = 0
    ' on initialization, save_empty_col_idx will be zero
    If (save_empty_col_idx >= 1) Then
      If (oSheet.getCellByPosition(save_empty_col_idx, 0).getType() = com.sun.star.table.CellContentType.EMPTY) Then
        If (oSheet.getCellByPosition(save_empty_col_idx - 1, 0).getType() <> com.sun.star.table.CellContentType.EMPTY) Then
          ' quick stab at eliminating a search for the first empty cell in the top row
          empty_col_idx = save_empty_col_idx
          j = 1
          If (verbose = True) Then
            Print "Function call (""" + cell_abs_name + """, """ + begin_str + """, """ + num_col_str + """) the index " + _
                  empty_col_idx + " of the first empty column in the first row was recovered from saved data obtained from a prior " + _
                  "call involving a previous sheet."
          End If
        End If
      End If
    End If
    If (j = 0) Then
      ' this is sort of a heavy duty operation so try to avoid it
      save_empty_col_idx = SheetRowIndexFirstEmpty(oSheet, 0, 0, max_col_index)
      empty_col_idx = save_empty_col_idx
      If (verbose = True) Then
        Print "Function call (""" + cell_abs_name + """, """ + begin_str + """, """ + num_col_str + """) the index " + _
              empty_col_idx + " of the first empty cell in the first row was obtained from a search."
      End If
    End If

    ' Sheet names don't match, and this is the current sheet.  Either this is a new sheet going forward, or
    ' this is the first time we've called this function.  If it's a new sheet going forward, try to recycle the old data.
    ReDim cdt()
    For j = LBound(save_cdt()) To UBound(save_cdt())
      tmp0_str = oSheet.getCellByPosition(save_cdt(j).colWidth, 0).getString()
      ' colTitle trimmed befoe saving
      If (StrComp(Trim(tmp0_str), save_cdt(j).colTitle, 1) = 0) Then
        ' index from last sheet is correct for this sheet too
        ColumnDisplayAddToSet(cdt(), save_cdt(j).colTitle, save_cdt(j).colWidth)
        If ((bool1 = False) And (StrComp(Trim(tmp0_str), Trim(num_col_str), 1) = 0)) Then
          ' correct index and it's the number column that we're looking for
          num_col_idx = save_cdt(j).colWidth
          bool1 = True
          If (verbose = True) Then
            Print "Function call (""" + cell_abs_name + """, """ + begin_str + """, """ + num_col_str + """) the index " + _
                  num_col_idx + " of the """ + num_col_str + """ was recovered from saved data obtained from a prior " + _
                  "call involving a previous sheet."
          End If
        End If
        'Print "Going forwards, the sheet " + oSheet.getName() + " col index " + save_cdt(j).colWidth + _
        '      " has header " + save_cdt(j).colTitle
      End If
    Next j
    'Print "UBound(cdt()) " + UBound(cdt()) + " UBound(save_cdt()) " + UBound(save_cdt())
    If (UBound(cdt()) <> UBound(save_cdt())) Then
      ' one or more of the columns that we have saved is not in the same position in the new sheet
      ColumnDisplayCopy(save_cdt(), cdt())
      If (verbose = True) Then
      End If
        ' TODO: move this in the if statement
        Print "Only " + (UBound(save_cdt()) + 1) + " ind{ex was,ices were} found again from the previous sheet."
    End If
  ' if this is the first sheet or a new sheet
  End If

  If (bool1 = False) Then
    If (empty_col_idx < 0) Then
      Print "Could not find the end of the column headers in sheet " + oSheet.getName()
    Else
      ' there are various ways that we haven't found the number column yet, so do a search now
      oCellRange = oSheet.getCellRangeByPosition(0, 0, empty_col_idx - 1, 0)
      num_col_idx = CellRangeSearchTrimCol(oCellRange, num_col_str)
      If (verbose = True) Then
        Print "Function call (""" + cell_abs_name + """, """ + begin_str + """, """ + num_col_str + """) the index " + _
              num_col_idx + " of the """ + num_col_str + """ column was obtained from a search."
      End If
      ColumnDisplayAddToSet(save_cdt(), Trim(num_col_str), num_col_idx)
    End If
  End If


  If (num_col_idx < 0) Then
    Print "In sheet " + oSheet.getName() + ", could not find index of column titled " + num_col_str + "."
  Else
    ' http://www.openoffice.org/api/docs/common/ref/com/sun/star/table/CellAddress.html
    'c = oSheet.getCellRangeByName(str(UBound(str()))).CellAddress.Column
    'r = oSheet.getCellRangeByName(str(UBound(str()))).CellAddress.Row
    c = oCell.CellAddress.Column
    r = oCell.CellAddress.Row

    'oCellRange = ThisComponent.Sheets.getCellRangesByName(Text1)

    ' if one of these two matches match, and the corresponding number values are opposites, true
    tmp0_str = oSheet.getCellByPosition(c, r).getString()
    ' first, try the one below it
    tmp1_str = oSheet.getCellByPosition(c, r + 1).getString()
    ' case-sensitive, we find begin_str at the beginning of the text in the cell
    If ((InStr(1, tmp0_str, begin_str, 0) = 1) Or (InStr(1, tmp1_str, begin_str, 0) = 1)) Then
'      tmp0_dbl = oSheet.getCellByPosition(num_col_idx, r).getValue()
      tmp0_dbl = CellGetOnlyValue(oSheet.getCellByPosition(num_col_idx, r))
'      tmp1_dbl = oSheet.getCellByPosition(num_col_idx, r + 1).getValue()
      tmp1_dbl = CellGetOnlyValue(oSheet.getCellByPosition(num_col_idx, r + 1))
      ' TODO: test for blank as well?
      If (tmp0_dbl = (tmp1_dbl * -1)) Then
        bool0 = True
      End If
    End If

    If bool0 = False Then
      ' otherwise try the one above it
      tmp1_str = oSheet.getCellByPosition(c, r - 1).getString()
      ' case-sensitive, we find begin_str at the beginning of the text in the cell
      If ((InStr(1, tmp0_str, begin_str, 0) = 1) Or (InStr(1, tmp1_str, begin_str, 0) = 1)) Then
'        tmp0_dbl = oSheet.getCellByPosition(num_col_idx, r).getValue()
        tmp0_dbl = CellGetOnlyValue(oSheet.getCellByPosition(num_col_idx, r))
'        tmp1_dbl = oSheet.getCellByPosition(num_col_idx, r - 1).getValue()
        tmp1_dbl = CellGetOnlyValue(oSheet.getCellByPosition(num_col_idx, r - 1))
        ' TODO: test for blank as well?
        If (tmp0_dbl = (tmp1_dbl * -1)) Then
          bool0 = True
        End If
      End If
    End If

  ' we have a number column index
  End If

  UpThisDownBeginsNumOpp = bool0
  Exit Function

  ErrorHandler:
    'MsgBox("There has been an error in " & sMacroName & ". " & chr(10) & "Error text is: " & Error(Err) & chr(10) & _
           '"Line number: " & Erl, 48, "Error Message")
  If (err_is_cell = True) Then
    error_message_str = "This function must receive as a parameter the absolute name of a cell surrounded " + _
          "by quotes, such as ""$September.$E$60"" which indicates column E, row 60, in sheet September.  If instead it receives the " + _
          "absolute name without quotes, then the parameter will evaluate to the contents of that cell, which is incorrect.  The " + _
          "function received the following parameter """ + cell_abs_name + """.  This is likely an error in the " + filler_sheet_str + _
          " sheet in the " + before_str + " string or the " + after_str + " string.  For most entries the " + filler_sheet_str + _
          " sheet the last character of the " + before_str + " string should not be a double quote "", and the first character " + _
          "of the " + after_str + " string also should not be a double quote "".  However, for this function quotes are required " + _
          "in those locations so that the absolute name can be passed as an address and not the contents of the cell at that address.  " + _
          "Click Cancel and edit the " + filler_sheet_str + " in order to use this function correctly."
  'Else
  '  error_message_str = "Error on line " + Erl + ": " + Error(Err)
  End If
  UpThisDownBeginsNumOpp = False
'Function UpThisDownBeginsNumOpp(cell_abs_name As String, begin_str As String, num_col_str As String) As Boolean
End Function

Function BeginsNotMatchDownCol(cell_abs_name As String, begin_str As String, match_col_str As String)  As Boolean
  BeginsNotMatchDownCol = InStrAndMatchDownCol(cell_abs_name, begin_str, match_col_str, True)
End Function

Function    BeginsMatchDownCol(cell_abs_name As String, begin_str As String, match_col_str As String)  As Boolean
  BeginsMatchDownCol    = InStrAndMatchDownCol(cell_abs_name, begin_str, match_col_str, False)
End Function

' So even faster than the approach of functions saving information with static variables is all functions in a module
' sharing global variables and storing information of sheets in those.  For now the static variables inside the functions
' approach is fast enough but later improvements can be made.
'
' The PayPal history is a series of transactions as if you piled your receipts on a table.  The oldest
' receipt is on the bottom of the pile.  Similarly, the oldest transaction on any given sheet is on the
' bottom, the most recent is on the top.  If we search down on the sheet where the row number is
' increasing then we are going towards older dates, so we are looking for something that happened
' earlier in time.  We want the strings in the match column to match, and when we find this we also
' want to compare the strings in the cell_abs_name to the string begin_str using the basic function
' InStr and case-sensitive.  If invert is false then we return true if InStr is 1.  If invert is true
' then we return true of InStr is not 1.  Otherwise return false for anything not matching these
' requirements.
'
' If this function will be used on a regular basis with more than one match column string, then edit
' this function to use and array to save the column index data for all strings.  This will save a lot
' of processing time since the indices don't have to be looked up again for each function call.  This
' array structure can be implemented for the UpThisDownBeginsNumOpp function too.
'
Function InStrAndMatchDownCol(cell_abs_name As String, begin_str As String, match_col_str As String, invert As Boolean) As Boolean
  On Error Goto ErrorHandler
  Dim oConv, oSheets, oSheet, oCell, oCellRange As Object
  Dim str() As String
  Dim c, r As Long
  Dim i, j, k, start_row As Long
  Dim curr_sheet As Integer
  Dim begin_col_str As String
  Dim match_str As String
  Dim tmp0_str As String
  Dim tmp1_str As String
  'Dim tmp0_dbl As Double
  'Dim tmp1_dbl As Double
  'Dim i As Integer
  Dim bool0 As Boolean
  Dim bool1 As Boolean
  Dim err_is_cell As Boolean
  Dim verbose As Boolean

  Dim empty_col_idx As Long

  ' BeginsNotMatchDownCol("$April.$AM$73"; "FL"; "Name")
  ' BeginsMatchDownCol("$April.$AM$69"; "FL"; "From Email Address")


  ' The index of the column that we're looking for matching strings in is the important part of this function.
  ' There is lots of logic here to preserve the results of our searches to that we don't have to search every
  ' time we enter the function for the same sheet and the same column.  If we see a column once then it should
  ' be good for the rest of the calls for this sheet.
  Dim match_col_idx As Long

  ' in this column we'll try to compare the string contents of a cell with begin_str
  Dim begin_col_idx As Long

  ' sheet name will be taken from the cell absolute name, something like "$January"
  Static save_sheet_name As String
  ' can save the empty column index as well, this doesn't change with the sheet
  Static save_empty_col_idx As Long
  ' here we save the match column index based on the match column string
  Static save_cdt() As ColumnDisplayType

  ' this is what we use to construct a new match index/string array based on the old one
  Dim cdt() As ColumnDisplayType
  ' TODO: change ColumnDisplayType.{colTitle,colWidth} to StringLongType.{s,l}

  'Static save_match_col_idx As Long
  'Static save_match_col_str As Long

  err_is_cell = True
  verbose = False

  oSheets = ThisComponent.Sheets

  InitAccountingGlobals()

  'l = UBound(a()) + increase
  'ColumnDisplayAddToSet(cdt(), tmp_str, k)

  'Print "LBound(cdt()) + " + LBound(cdt()) + " UBound(cdt()) " + UBound(cdt())
  'ReDim Preserve cdt(0)
  'Print "LBound(cdt()) + " + LBound(cdt()) + " UBound(cdt()) " + UBound(cdt())
  'ReDim cdt()
  'Print "LBound(cdt()) + " + LBound(cdt()) + " UBound(cdt()) " + UBound(cdt())
  'InStrAndMatchDownCol = False
  'Exit Function

  ' BeginsNotMatchDownCol("$January.$AL$43"; "FL"; "Name")

  'Print "here"

  ' before cell     UpThisDownBeginsNumOpp("
  ' after cell      "; "PayPal Extras MasterCard<sup>&#174;</sup>"; "Gross")

  ' before cell     BeginsNotMatchDownCol("
  ' after cell      "; "FL"; "Name")

  ' before cell     BeginsMatchDownCol("
  ' after cell      "; "FL"; "Name")



  ' Name      Type     [Status]            [Gross]                    State/Province/Region/County/Territory/Prefecture/Republic
  ' --------------------------------------------------------------------------------------------------
  ' John Doe  Refund   Completed           (ref < 0)&&((ref*-1)
  ' John Doe  Payment  Partially Refunded  (sale>0)&&((ref*-1)<sale)  FL or not FL
  ' John Doe  Payment  Refunded            (sale>0)&&((ref*-1)<sale)  FL or not FL

  ' 1 1/11/2011
  ' 2 1/11/2011
  ' 3 1/9/2011
  ' 4 1/8/2011
  ' 5 1/8/2011
  ' 6 
  ' 7 

  ' for instance
  '
  ' UpThisDownBeginsNumOpp("$September.$E$60"; "PayPal Extras MasterCard<sup>&#174;</sup>"; "Gross")
  ' UpThisDownBeginsNumOpp("$September.$E$61"; "PayPal Extras MasterCard<sup>&#174;</sup>"; "Gross")
  ' UpThisDownBeginsNumOpp("$September.$E$62"; "PayPal Extras MasterCard<sup>&#174;</sup>"; "Gross")
  ' UpThisDownBeginsNumOpp("$September.$E$63"; "PayPal Extras MasterCard<sup>&#174;</sup>"; "Gross")

  'oSheet = ThisComponent.Sheets.getByName(str(LBound(str())))


  ' bool0 represents success at finding a matching row, and applying the string comparison
  ' it is used to exit the loop over the sheets counting down
  bool0 = False

  str() = split(cell_abs_name, ".")
  oCell = CAbsoluteNameToCell(cell_abs_name)
  If IsNull(oCell) Then
    Print "Cound not locate cell whose absolute name is: " + cell_abs_name
    GoTo ErrorHandler
  End If
  err_is_cell = False

  ' The header string 
  ' This cell we're comparing the contents of to begin_str, and from it we also get the coordinates of what row and
  ' sheet we're starting at in order to try to match the string in the match column of this row to an earlier
  ' transaction in a higher number row or prior sheet.

  ' First identify the column index of the column title matching match_col_str.  If this is found, then get the column
  ' title for the column index of this cell, and go to the match column index in the row index of this cell and get the
  ' string that we're trying to match to a higher numbered row possibly on a lower numbered sheet.  If we move back a
  ' sheet, ensure that we have indices for the column title matching match_col_str as well as the title matching the
  ' column title of the column index of this cell.  If we find our row match, then compare the string in the column
  ' title matching the title of the column index of this cell to begin_str, and depending on that match, return true or
  ' false.
 
  ' This cell is the starting point for our search.  Obtain the column title from this sheet and column, row zero.  
  ' we hopefully find a match 
  curr_sheet = oCell.CellAddress.Sheet
  c = oCell.CellAddress.Column
  r = oCell.CellAddress.Row

  ' column title of column with the index of the cell identified by cell_abs_name we only use this if we're moving
  ' back sheets to adapt to changes in the index of this column title
  begin_col_str = oSheets.getCellByPosition(c, 0, curr_sheet).getString()

  For i = 0 To curr_sheet
    k = curr_sheet - i
    ' every sheet we need to reset these
    '
    ' bool1 represents a success in finding the index of the match column
    ' it helps to distinguish between this success and an index < 0
    bool1 = False
    match_col_idx = -1
    empty_col_idx = -1
    'begin_col_idx = -1

    ' go backwards through the seets looking for our match
    oSheet = oSheets.getByIndex(k)
    
    If (StrComp(oSheet.getName(), total_sheet_str, 1) <> 0) And _
       (StrComp(oSheet.getName(), filler_sheet_str, 1) <> 0) Then

      ' We enter into a sheet either from a function call or through the
      ' sheets loop.  If we came from a function call see if we can reuse our
      ' saved data.  If we came from the loop see if any of our data works for that
      ' sheet as well.  Otherwise, do the expensive lookup.
      If (k <> curr_sheet) Then
        'Print "Sheet is not current sheet, has index " + k
        ' going backwards in time through the sheets, don't save data going backwards,
        ' just take a best guess at the match col index before a search
        For j = LBound(save_cdt()) To UBound(save_cdt())
          If (StrComp(save_cdt(j).colTitle, match_col_str, 1) = 0) Then
            tmp0_str = oSheet.getCellByPosition(save_cdt(j).colWidth, 0).getString()
            If (StrComp(Trim(tmp0_str), Trim(match_col_str), 1) = 0) Then
              ' our current data matches an old sheet as well
              match_col_idx = save_cdt(j).colWidth
              bool1 = True
              If (verbose = True) Then
                Print "Function call (""" + cell_abs_name + """, """ + begin_str + """, """ + match_col_str + """) the index " + _
                      match_col_idx + " of the """ + match_col_str + """ was recovered from saved data obtained from a prior " + _
                      "call involving a sheet ordered after the sheet currently being searched."
                'Print "Going backwards through loop, the sheet " + oSheet.getName() + " col index " + save_cdt(j).colWidth + _
                '      " has header " + save_cdt(j).colTitle
              End If
            End If
          End If
        Next j
        If (bool1 = False) Then
          ' find the first empty column but don't save it, we'll need to do a search
          ' only do if bool1 if False however since we might have already found it
          ' this is sort of a heavy duty operation so try to avoid it
          empty_col_idx = SheetRowIndexFirstEmpty(oSheet, 0, 0, max_col_index)
          If (verbose = True) Then
            Print "Function call (""" + cell_abs_name + """, """ + begin_str + """, """ + match_col_str + """) the index " + _
                  empty_col_idx + " of the first empty cell in the first row was obtained from a search."
          End If
        End If
      ElseIf (StrComp(str(0), save_sheet_name, 1) = 0) Then
        'Print "Sheet matches " + save_sheet_name
        ' Sheet names match.  At least we can avoid looking up the empty col index since we have looked it up at least once previously.
        empty_col_idx = save_empty_col_idx
        For j = LBound(save_cdt()) To UBound(save_cdt())
          If (StrComp(save_cdt(j).colTitle, match_col_str, 1) = 0) Then
            ' we found our match column index previously, but it still could be -1
            ' so set bool1 to True to bypass a lookup
            match_col_idx = save_cdt(j).colWidth
            bool1 = True
            If (verbose = True) Then
              Print "Function call (""" + cell_abs_name + """, """ + begin_str + """, """ + match_col_str + """) the index " + _
                    match_col_idx + " of the """ + match_col_str + """ was recovered from saved data obtained from a prior " + _
                    "call involving this same sheet."
              'Print "From saved data, the sheet " + oSheet.getName() + " col index " + save_cdt(j).colWidth + _
              '      " has header " + save_cdt(j).colTitle
            End If
          End If
        Next j
      Else
        'Print "New sheet " + oSheet.getName()
        ' save the name of the new sheet such as "$January"
        save_sheet_name = str(0)

        ' find the first empty column and save it
        j = 0
        If (save_empty_col_idx >= 1) Then
          If (oSheet.getCellByPosition(save_empty_col_idx, 0).getType() = com.sun.star.table.CellContentType.EMPTY) Then
            If (oSheet.getCellByPosition(save_empty_col_idx - 1, 0).getType() <> com.sun.star.table.CellContentType.EMPTY) Then
              ' quick stab at eliminating a search for the first empty cell in the top row
              empty_col_idx = save_empty_col_idx
              j = 1
              If (verbose = True) Then
                Print "Function call (""" + cell_abs_name + """, """ + begin_str + """, """ + match_col_str + """) the index " + _
                      empty_col_idx + " of the first empty column in the first row was recovered from saved data obtained from a prior " + _
                      "call involving a previous sheet."
              End If
            End If
          End If
        End If
        If (j = 0) Then
          ' this is sort of a heavy duty operation so try to avoid it
          save_empty_col_idx = SheetRowIndexFirstEmpty(oSheet, 0, 0, max_col_index)
          empty_col_idx = save_empty_col_idx
          If (verbose = True) Then
            Print "Function call (""" + cell_abs_name + """, """ + begin_str + """, """ + match_col_str + """) the index " + _
                  empty_col_idx + " of the first empty cell in the first row was obtained from a search."
          End If
        End If

        ' Sheet names don't match, and this is the current sheet.  Either this is a new sheet going forward, or
        ' this is the first time we've called this function.  If it's a new sheet going forward, try to recycle the old data.
        ReDim cdt()
        For j = LBound(save_cdt()) To UBound(save_cdt())
          tmp0_str = oSheet.getCellByPosition(save_cdt(j).colWidth, 0).getString()
          ' colTitle trimmed befoe saving
          If (StrComp(Trim(tmp0_str), save_cdt(j).colTitle, 1) = 0) Then
            ' index from last sheet is correct for this sheet too
            ColumnDisplayAddToSet(cdt(), save_cdt(j).colTitle, save_cdt(j).colWidth)
            If ((bool1 = False) And (StrComp(Trim(tmp0_str), Trim(match_col_str), 1) = 0)) Then
              ' correct index and it's the match column that we're looking for
              match_col_idx = save_cdt(j).colWidth
              bool1 = True
              If (verbose = True) Then
                Print "Function call (""" + cell_abs_name + """, """ + begin_str + """, """ + match_col_str + """) the index " + _
                      match_col_idx + " of the """ + match_col_str + """ was recovered from saved data obtained from a prior " + _
                      "call involving a previous sheet."
              End If
            End If
            'Print "Going forwards, the sheet " + oSheet.getName() + " col index " + save_cdt(j).colWidth + _
            '      " has header " + save_cdt(j).colTitle
          End If
        Next j
        'Print "UBound(cdt()) " + UBound(cdt()) + " UBound(save_cdt()) " + UBound(save_cdt())
        If (UBound(cdt()) <> UBound(save_cdt())) Then
          ' one or more of the columns that we have saved is not in the same position in the new sheet
          ColumnDisplayCopy(save_cdt(), cdt())
          If (verbose = True) Then
          End If
            ' TODO: move this in the brackets
            Print "Only " + (UBound(save_cdt()) + 1) + " ind{ex was,icies were} found again from the previous sheet."
        End If

      End If

      'If ((StrComp(str(0), save_sheet_name, 1) <> 0) Or (StrComp(save_match_col_str, match_col_str, 1) <> 0)) Then
      If (bool1 = False) Then
        If (empty_col_idx < 0) Then
          Print "Could not find the end of the column headers in sheet " + oSheet.getName()
        Else
          ' there are various ways that we haven't found the match column yet, so do a search now
          oCellRange = oSheet.getCellRangeByPosition(0, 0, empty_col_idx - 1, 0)
          match_col_idx = CellRangeSearchTrimCol(oCellRange, match_col_str)
          If (verbose = True) Then
            Print "Function call (""" + cell_abs_name + """, """ + begin_str + """, """ + match_col_str + """) the index " + _
                  match_col_idx + " of the """ + match_col_str + """ column was obtained from a search."
          End If
          If (k = curr_sheet) Then
            ' only save data going forward, since we could have multiple match indices
            ' going backwards we're only looking for a single index
            ' trim match column string before saving
            ColumnDisplayAddToSet(save_cdt(), Trim(match_col_str), match_col_idx)
          End If
        End If
      End If

      ' if match column index is less than zero, force bail of next if comparison through the begin column index
      begin_col_idx = -1
      If (match_col_idx >= 0) Then
        If (k = curr_sheet) Then
          begin_col_idx = c
        Else
          tmp0_str = oSheet.getCellByPosition(c, 0).getString()
          If (StrComp(Trim(tmp0_str), Trim(begin_col_str), 1) = 0) Then
            begin_col_idx = c
          Else
            If (empty_col_idx < 0) Then
              Print "Could not find the end of the column headers in sheet " + oSheet.getName()
            Else
              ' forced to search for the begin column string
              oCellRange = oSheet.getCellRangeByPosition(0, 0, empty_col_idx - 1, 0)
              begin_col_idx = CellRangeSearchTrimCol(oCellRange, begin_col_str)
              If (verbose = True) Then
                Print "Function call (""" + cell_abs_name + """, """ + begin_str + """, """ + begin_col_str + """) the index " + _
                      begin_col_idx + " of the """ + begin_col_str + """ column was obtained from a search."
              End If
            End If
          End If
        End If
      End If

      'Print "Sheet " + oSheet.getName() + " mci " + match_col_idx + " bci " + begin_col_idx

      If (begin_col_idx < 0) Then
        Print "In sheet " + oSheet.getName() + ", could not find column indices for one or more of columns titled " + _
              match_col_str + ", " + begin_col_str + "."
      Else
        'Print "save_match_col_idx " + save_match_col_idx
        ' we found the match column in this sheet
        If (k = curr_sheet) Then
          match_str = oSheet.getCellByPosition(match_col_idx, r).getString()
          ' in the current sheet, only search from r + 1 and greater
          start_row = (r + 1)
        Else
          ' in prior sheets, search starting at 1
          start_row = 1
        End If

        ' very complicated setup, but we have both our column indices of interest, now just go row to row searching
        ' reuse bool1, new meaning is found matching row and exit all loops
        bool1 = False
        For j = start_row To max_row_index
          tmp0_str = oSheet.getCellByPosition(match_col_idx, j).getString()
          If (StrComp(match_str, tmp0_str, 1) = 0) Then
            ' we're going to exit this loop but we also want to exit the sheets loop
            bool1 = True
            ' OK so we matched the strings in the match column, now apply the "begins" logic, and the inversion
            tmp0_str = oSheet.getCellByPosition(begin_col_idx, j).getString()
            If (InStr(1, tmp0_str, begin_str, 0) = 1) Then
              bool0 = True
            End If
            If (invert = True) Then
              bool0 = Not bool0
            End If
            'If (bool0 = True) Then
            '  Print "Sheet " + oSheet.getName() + ", row idx " + j + ", match str is " + match_str + _
            '        ", state is " + tmp0_str + ", invert is " + invert + ", returns True."
            'End If
            Exit For
          ElseIf (oSheet.getCellByPosition(match_col_idx, j).getType() = com.sun.star.table.CellContentType.EMPTY) Then
            oCellRange = oSheet.getCellRangeByPosition(0, j, empty_col_idx - 1, j)
            If IsCellRangeEmpty(oCellRange) Then
              If (verbose = True) Then
                Print "Sheet " + oSheet.getName() + ", row " + j + " is empty."
              End If
              Exit For
            End If
          End If
        Next j
        If (bool1 = True) Then
          ' we have our match so get out
          Exit For
        End If

      ' we have a match column index
      End If
    
    ' if we're not Filler or Totals sheets
    End If

  ' go to a prior sheet to look for the data
  Next i

  ' We can't match a refund to a purchase that was placed in a previous year.
  ' This data will need to be looked up and the tax category copy and pasted by the user.

  If (bool1 = True) Then
'    If (bool0 = True) Then
'      Print "Function call (""" + cell_abs_name + """, """ + begin_str + """, """ + match_col_str + """, " + invert + ") returns " + bool0 + ".  " + _
'            "Starting from sheet " + save_sheet_name + ", row " + (r + 1) + ", and searching for an exact match in the """ + match_col_str + _
'            """ column in higher number row indices and lower number sheet indices (both should be backwards in time), a match was found.  " + _
'            "The string from the """ + begin_col_str + """ column in that matching row is """ + tmp0_str + """."
'    End If
  Else
    Print "Function call (""" + cell_abs_name + """, """ + begin_str + """, """ + match_col_str + """, " + invert + ") returns " + bool0 + ".  " + _
          "Starting from sheet " + save_sheet_name + ", row " + (r + 1) + ", and searching for an exact match to the string """ + _
          match_str + """ in the """ + match_col_str + _
          """ column in higher number row indices and lower number sheet indices (both should be backwards in time), a match could not be " + _
          "found.  Perform a manual search, perhaps in a spreadsheet of a prior year, to determine if the refund was from a purchase in which " + _
          "sales tax applies.  Such sales should have separate tax category labels since the sales tax is also refunded, and therefore the " + _
          "tax obligation of the company is less."
  End If

  InStrAndMatchDownCol = bool0
  Exit Function

  ErrorHandler:
  If (err_is_cell = True) Then
    error_message_str = "This function must receive as a parameter the absolute name of a cell surrounded " + _
          "by quotes, such as ""$September.$E$60"" which indicates column E, row 60, in sheet September.  If instead it receives the " + _
          "absolute name without quotes, then the parameter will evaluate to the contents of that cell, which is incorrect.  The " + _
          "function received the following parameter """ + cell_abs_name + """.  This is likely an error in the " + filler_sheet_str + _
          " sheet in the " + before_str + " string or the " + after_str + " string.  For most entries the " + filler_sheet_str + _
          " sheet the last character of the " + before_str + " string should not be a double quote "", and the first character " + _
          "of the " + after_str + " string also should not be a double quote "".  However, for this function quotes are required " + _
          "in those locations so that the absolute name can be passed as an address and not the contents of the cell at that address.  " + _
          "Click Cancel and edit the " + filler_sheet_str + " in order to use this function correctly."
  'Else
  '  error_message_str = "Error on line " + Erl + ": " + Error(Err)
  End If
  InStrAndMatchDownCol = False
End Function

Sub TestBeginsMatch()
  '=BeginsNotMatchDownCol("$January.$AM$20"; "FL"; "Name")
  '=BeginsNotMatchDownCol("$January.$AM$20"; "FL"; "From Email Address")
  '=BeginsMatchDownCol("$January.$AM$20"; "FL"; "Name")
  BeginsNotMatchDownCol("$January.$AM$20", "FL", "Name")
  BeginsNotMatchDownCol("$January.$AM$20", "FL", "From Email Address")
  BeginsMatchDownCol("$January.$AM$20", "FL", "Name")
End Sub

Function BeginsNot(Text1 As String, Text2 As String) As Boolean
  ' case-sensitive, we find Text2 at the beginning of Text1
  BeginsNot = (InStr(1, Text1, Text2, 0) <> 1)
'Function Begins(Text1 As String, Text2 As String) As Boolean
End Function

Function ContainsNot(Text1 As String, Text2 As String) As Boolean
  ' case-sensitive, we find Text2 somewhere in Text1
  ContainsNot = (InStr(1, Text1, Text2, 0) = 0)
'Function Contains(Text1 As String, Text2 As String) As Boolean
End Function

Function Begins(Text1 As String, Text2 As String) As Boolean
  ' case-sensitive, we find Text2 at the beginning of Text1
  Begins = (InStr(1, Text1, Text2, 0) = 1)
'Function Begins(Text1 As String, Text2 As String) As Boolean
End Function

Function Contains(Text1 As String, Text2 As String) As Boolean
  ' case-sensitive, we find Text2 somewhere in Text1
  Contains = (InStr(1, Text1, Text2, 0) <> 0)
'Function Contains(Text1 As String, Text2 As String) As Boolean
End Function

