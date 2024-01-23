Option Explicit

' Copy
Public copyMeals As Range
Public copyMealsWithLineItemAmount As Range

' Paste
Public pasteMeals As Range
Public pasteMealsWithoutName As Range

'Sort
Public sortMeals As Range
Public sortMealsKey As Range
Public sortCurrentMealItem As Range
Public sortCurrentCellForCount As Range
Public sortEmptyCellsToZero As Range
Public sortOnItemCount As Range
Public sortOnItemCountKey As Range

'Clear
Public clearMeals As Range
Public clearSortedMeals As Range
Public clearMealsWithZeroItems As Range
Public clearAll As Range
Public ClearHighlighting As Range

' Format
Public formatCurrency
Public formatSelectedCell

Const topRow = 4
Const bottomRow = 24

Private Sub Initialise()
    Call InitialiseForCopying
    Call InitialiseForPasting
    Call InitialiseForClearing
    Call InitialiseForSorting
    Call InitialiseForFormatting
End Sub

Private Sub InitialiseForCopying()
    Set copyMeals = Range("A" & topRow & ":" & "C" & bottomRow)
    Set copyMealsWithLineItemAmount = Range("E" & topRow & ":" & "G" & bottomRow)
End Sub

Private Sub InitialiseForPasting()
    Set pasteMeals = Range("F" & topRow)
    Set pasteMealsWithoutName = Range("J" & topRow)
End Sub

Private Sub InitialiseForClearing()
    Set clearMeals = Range("E" & topRow & ":" & "H" & bottomRow)
    Set clearSortedMeals = Range("J" & topRow & ":" & "L" & bottomRow)
    Set clearMealsWithZeroItems = Range("J" & topRow & ":" & "L" & bottomRow)
    Set clearAll = Range("E" & topRow & ":" & "L" & bottomRow)
    Set ClearHighlighting = Range("A" & topRow & ":C" & bottomRow & "," & "F" & topRow & ":H" & bottomRow)
End Sub

Private Sub InitialiseForFormatting()
    Set formatCurrency = Range("G" & topRow & ":G" & bottomRow & "," & "L" & topRow & ":L" & bottomRow)
    Set formatSelectedCell = Range("A" & topRow & ":A" & bottomRow & "," & "F" & topRow & ":F" & bottomRow)
End Sub

Private Sub InitialiseForSorting()

    Set sortMeals = Range("F" & topRow & ":" & "H" & bottomRow)
    Set sortMealsKey = Range("F" & topRow & ":" & "F" & bottomRow)

    Set sortCurrentCellForCount = Range("E" & topRow)
    Set sortCurrentMealItem = Range("F" & topRow)
    Set sortEmptyCellsToZero = Range("J" & topRow & ":" & "J" & bottomRow)

    Set sortOnItemCount = Range("J" & topRow & ":" & "L" & bottomRow)
    Set sortOnItemCountKey = Range("J" & topRow & ":" & "J" & bottomRow)

End Sub

Private Sub btnOrder_Click()
    If IsEmpty(Range("A4").Value) Then
        Exit Sub    ' A4 is the first cell in the 'order' range. Exit if data is not present.
    End If
    Application.ScreenUpdating = False
    Call Initialise
    Call ResetHighlighting
    Call CopyOrderedMeals
    Call SortCopiedMealsByFoodItem
    Call CountItemByGroup
    Call CopyAndPasteSortedOrder
    Call PopulateEmptyWithZero
    Call sortMealOnItemCount
    Call RemoveItemsWithAmountZero
    formatCurrency.Style = "Currency"
    Range("A1").Select ' Park the cursor in the corner of screen
    Application.ScreenUpdating = True
End Sub

Private Sub CopyOrderedMeals()
    clearMeals.Clear
    copyMeals.Copy
    pasteMeals.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub

Private Sub SortCopiedMealsByFoodItem()
    sortMeals.Select
    sortMealsKey.Select
    sortMeals.Sort Key1:=sortMealsKey, Order1:=xlAscending, Header:=xlNo
End Sub

Private Sub CountItemByGroup()
    Dim rngMeals
    Set rngMeals = Range("F" & topRow & ":" & "F" & bottomRow)
    Dim rngMarkerForAmount: Set rngMarkerForAmount = sortCurrentCellForCount
    Dim currentMeal As String
    Dim activeCell As Range
    For Each activeCell In rngMeals
        If IsEmpty(activeCell.Value) Then
            Exit Sub
        End If
        If currentMeal = Empty Or currentMeal = activeCell.Value Then
            rngMarkerForAmount.Value = rngMarkerForAmount.Value + 1
        Else
            Set rngMarkerForAmount = activeCell.Offset(0, -1) ' Move marker down to next cell in column
            rngMarkerForAmount.Value = rngMarkerForAmount.Value + 1
        End If
        currentMeal = activeCell.Value
    Next
End Sub

Private Sub CopyAndPasteSortedOrder()
    clearSortedMeals.Clear
    copyMealsWithLineItemAmount.Copy
    pasteMealsWithoutName.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub

Private Sub PopulateEmptyWithZero()
    Dim activeCell As Range
    For Each activeCell In sortEmptyCellsToZero
        If IsEmpty(activeCell.Value) Then
            activeCell.Value = 0
        End If
    Next
End Sub

Private Sub sortMealOnItemCount()
    sortOnItemCount.Sort Key1:=sortOnItemCountKey, Order1:=xlDescending, Header:=xlNo
End Sub

Private Sub RemoveItemsWithAmountZero()
    Dim activeCell As Range
    For Each activeCell In clearMealsWithZeroItems
        If activeCell.Column = 10 And activeCell.Value = 0 Then
            activeCell.Value = ""
            activeCell.Offset(0, 1).Value = ""
            activeCell.Offset(0, 2).Value = ""
        End If
    Next
End Sub

Private Sub ResetHighlighting()
    With ClearHighlighting.Interior
       .Pattern = xlNone
        .TintAndShade = 0
    End With
End Sub

Private Sub btnClear_Click()
    Call Initialise
    Call ResetHighlighting
    clearAll.Clear
End Sub

Private Sub btnClear_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Range("A4", Range("A4").End(xlDown)).ClearContents
    Range("C4", Range("C4").End(xlDown)).ClearContents
End Sub


Private Sub cmdHighlight_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call Initialise
    Call ResetHighlighting
End Sub


Private Sub cmdHighlight_Click()
    Call Initialise
    Dim strAddressOfMatchingItems As String
    Dim rngCurrent As Range

    ' Clear previous highlighting
    Call ResetHighlighting

    If IsEmpty(activeCell) Then
        Exit Sub
    End If

   ' The first time a match has been found, i.e. when there is a match AND strAddressOfMatchingItems = Empty, add the
     ' current' cell to start the creation of the string of matching cells.

     ' This part of the loop will only be hit one, when strAddressOfMatchingItems = Empty AND when there is a match.

     ' The second time a match is found, the ElseIf part of the loop will be hit. In this we need to concatenate, and add
     ' a comma before a range is added.

    For Each rngCurrent In formatSelectedCell
        If strAddressOfMatchingItems = Empty And rngCurrent = activeCell Then
            With rngCurrent
                strAddressOfMatchingItems = .Address(0, 0) & ","
                strAddressOfMatchingItems = strAddressOfMatchingItems & .Offset(0, 1).Address(0, 0) & ","
                strAddressOfMatchingItems = strAddressOfMatchingItems & .Offset(0, 2).Address(0, 0)
            End With
        ElseIf rngCurrent = activeCell Then
            With rngCurrent
                strAddressOfMatchingItems = strAddressOfMatchingItems & "," & .Address(0, 0) & ","
                strAddressOfMatchingItems = strAddressOfMatchingItems & .Offset(0, 1).Address(0, 0) & ","
                strAddressOfMatchingItems = strAddressOfMatchingItems & .Offset(0, 2).Address(0, 0)
            End With
        End If
    Next rngCurrent

    If strAddressOfMatchingItems <> Empty Then
        With Range(strAddressOfMatchingItems).Interior
                .Pattern = xlSolid
                .ThemeColor = xlThemeColorAccent5
                .TintAndShade = 0.399975585192419
        End With
    End If
End Sub


Private Sub SEP_26_2022()

 Range("A4").Value = "Special Chow Mein"
 Range("A5").Value = "Special Curry"
 Range("A6").Value = "Chips"
 Range("A7").Value = "Curry Sauce"
 Range("A8").Value = "Egg Fried Rice"
 Range("A9").Value = "Chicken Curry"
 Range("A10").Value = "Egg Fried Rice"
 Range("A11").Value = "Chips"
 Range("A12").Value = "Curry Sauce"
 Range("A13").Value = "Special Chow Mein"


 Range("C4").Value = "Angela & Mike"
 Range("C5").Value = "Angela & Mike"
 Range("C6").Value = "Angela & Mike"
 Range("C7").Value = "Angela & Mike"
 Range("C8").Value = "Angela & Mike"
 Range("C9").Value = "Charles"
 Range("C10").Value = "Celia & Charles"
 Range("C11").Value = "Celia & Charles"
 Range("C12").Value = "Celia & Charles"
 Range("C13").Value = "Celia"

End Sub


