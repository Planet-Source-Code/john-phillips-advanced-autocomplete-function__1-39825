VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Auto Complete"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option3 
      Caption         =   "Listview Auto Complete"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Listbox Auto Complete"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Textbox Auto Complete"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Value           =   -1  'True
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5318
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7455
   End
   Begin VB.Label Label3 
      Caption         =   "Type Text In This Box"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "ListView Auto Complete"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Listbox Auto Complete"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub Text1_Change()
Dim DB As Database

' here is how to call the autocomplete function
' in the keydown event of the textbox put
' ChkChar KeyCode, this will let the function know
' the user is trying to delete a character and not to
' execute the autocomplete function

' in the change event of the textbox put
' AutoComplete TheTextBox, TheDatabase, TheSearchTable, _
'   TheSearchField, TheAutoCompleteType, AddColumnHeaders, _
'   TheListBox, TheListView, TheTextBox
'
' TheListBox, TheListView and TheTextBox are all optional
' depending on the autocomplete type you select
'
' TheTextBox        -   The Textbox the text is being entered from
' TheDatabase       -   The Database being serached
' TheSearchTable    -   The Table being searched
' TheSearchField    -   The Field being searched
' TheAutoCompleteType
'       0 = textbox autocomplete
'       1 = listbox Autocomplete
'       2 = listview autocomplete
' AddCoulmnHeaders  -   Either True or False (Used on with the listview)
' TheListBox        -   Name of the listbox to display the values found
' TheListView       -   Name of the listview being used to display the values found
' TheTextBox        -   Name of the textbox to display the values found


If Option1.Value = True Then ' do the textbox auto complete
    AutoComplete Text1.Text, App.Path & "\nwind.mdb", "customers", "city", 0, False, , , Text1
ElseIf Option2.Value = True Then
    AutoComplete Text1.Text, App.Path & "\nwind.mdb", "customers", "city", 1, False, List1
ElseIf Option3.Value = True Then
    AutoComplete Text1.Text, App.Path & "\nwind.mdb", "customers", "city", 2, True, , ListView1
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
ChkChar KeyCode
End Sub
