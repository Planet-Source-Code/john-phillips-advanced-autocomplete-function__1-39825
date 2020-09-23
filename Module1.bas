Attribute VB_Name = "Module1"
' uses a refrence to thye MS DAO 3.6 Object Library
' if you want to use this with an older version of the DAO
' select which one you awant to use, you may have some problems
' with some of the functions for DAO 3.6 change them as you need

' Advanced AutoComplete function, by John Phillips
' I was inspired by code found on PSC, I liked what was there
' but need more options with the autocomplete from a database
' and I wanted to wrap it all into one function instead of several
' functions for each operation I need to do.

Private DB As Database
Private RS As Recordset
Public bExit As Boolean
Private tb As TableDef
Private fl As Field
Private ix As Index


Public Function AutoComplete(sText As String, sDB As String _
                , sTable As String, sField As String, lType As Long, bColumn As Boolean _
                , Optional oListBox As ListBox, Optional oListView As ListView _
                , Optional oTextBox As TextBox) As Boolean

On Error GoTo errNoComplete ' basic error handling
Dim lLen As Long ' for the length of the search string
Dim X As Long ' place holder
Dim c As Long ' place holder

If bExit = True Or sText = "" Then Exit Function ' if this value is true then backspace or del was pressed

' open the database
Set DB = OpenDatabase(sDB)

' set the function to false for now
AutoComplete = False

' get the length of the search string
lLen = Len(sText)

' open the recordset
Set RS = DB.OpenRecordset("SELECT * FROM " & sTable & " WHERE " & sField & " LIKE '" & sText & "*'")

' if the EOF and BOF are both true then
' that meens no records were found
If RS.EOF = True And RS.BOF = True Then
    Select Case lType
        Case 0 ' textbox
            ' no records were found so just exit the function
            Exit Function
        Case 1 ' listbox
            ' no records were found, so give an indicator
            ' in the listbox and exit the function
            oListBox.Clear
            oListBox.AddItem "No Matching Records"
            Exit Function
        Case 2 ' listview
            ' no records were found, so give an indicator
            ' in the listview and exit the function
            
            ' check to see if the RS defines the column headers
            ' if so the clear them also
            If bColumn = True Then
                oListView.ColumnHeaders.Clear
            End If
            oListView.ListItems.Clear
            oListView.ListItems.Add 1, , "No Matching Records"
            Exit Function
    End Select

End If

Select Case lType
    Case 0 ' textbox
        'set the textbox to the found value
        oTextBox.Text = RS(sField)
        
        ' select the part of the text that hasnt been typed yet
        If oTextBox.SelText = "" Then
            oTextBox.SelStart = lLen
        Else
            oTextBox.SelStart = InStr(oTextBox.Text, oTextBox.SelText)
        End If
        
        oTextBox.SelLength = Len(oTextBox.Text)
        
        ' set the function to true and ext the function
        AutoComplete = True
                             
        Exit Function
    Case 1 ' listbox
        ' make sure we are at the first record found
        RS.MoveFirst
        ' clear the listbox
        oListBox.Clear
        
        ' fill the listbox with all the values found
        ' by looping through the enitre recordset
        Do While RS.EOF = False
        ' add the current ield to the listbox
        oListBox.AddItem RS(sField)
        ' move to the next record
        RS.MoveNext
        Loop
        
        ' set the function to true
        AutoComplete = True
        ' finally exit the function
        Exit Function
    Case 2 ' listview
        ' clear the listview
        oListView.ListItems.Clear
            If bColumn = True Then ' add the columns also
                ' clear the listview's columns
                oListView.ColumnHeaders.Clear
                
                ' set the tabledef to the currewnt table
                Set tb = DB.TableDefs(sTable)
                
                ' now loop through all the field names
                ' fl is declared as a field
                ' tb is declared as the table
                For Each fl In tb.Fields
                    ' we do this just to make sure it isnt
                    ' a hidden propertie in the table
                    If Left(fl.Name, 4) <> "MSys" Then
                        ' now add the field name to the column
                        ' header of the listview
                        oListView.ColumnHeaders.Add , , _
                            fl.Name, (Len(fl.Name) * 260) + 520
                        ' the last calculation
                        ' ie. (Len(fl.Name) * 260) + 520
                        ' is used to calculate the cloumn width
                        ' 260 is apporx. the size of 1 cahracter
                    
                    ' get a coun of the fields for use later
                    X = X + 1
                    End If
                Next
                ' set the listview view type to report
                oListView.View = lvwReport
            End If
        
        ' set c to 1
        c = 1
        
        ' move to the first record in the recordset
        RS.MoveFirst
        
        ' loop through the entire recordset and add all the
        ' found values
        Do While RS.EOF = False
            ' add the first field to the first field in the listview
            oListView.ListItems.Add c, , RS.Fields(0).Value & ""
            
            ' now loop through the rest of the fields in
            ' the current record and add them to the listview also
            For r = 1 To X - 1
                ' c = the current listview item
                ' r = the current sub item of the current line
                ' r also represents the current field index
                ' to add
                oListView.ListItems(c).SubItems(r) = RS.Fields(r).Value & ""
            Next r
        
        ' increment c to move to the next line in the listview
        c = 1 + 1
        ' move to the next record in the recordset
        RS.MoveNext
        Loop
        
        ' set the function to true and exit
        AutoComplete = True
        
        Exit Function
    Case Else ' just incase of an error
        Exit Function
End Select

AutoComplete = True
Exit Function
errNoComplete:
    If Err.Number = 3075 Then
        ' this happens when a table name has a space in it
        ' why it happens I am not sure
        ' but it does and i havent put the time in to figure
        ' out how to fix this as of yet
        Exit Function
    End If
    
    ' display the error to the user and try to resume the
    ' next line
    MsgBox "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description
    Resume Next
End Function
                
Public Function ChkChar(lChar As Integer)
' this function does a check to see if the
' backspace key (Character #8) or the Delete Key (charatcer #46)
' was pressed

If lChar = 8 Then
    bExit = True
    Exit Function
ElseIf lChar = 46 Then
    bExit = True
    Exit Function
Else
    bExit = False
    Exit Function
End If

End Function
