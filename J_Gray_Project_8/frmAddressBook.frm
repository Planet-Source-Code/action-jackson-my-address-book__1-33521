VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmAddressBook 
   Caption         =   "Address Book"
   ClientHeight    =   5235
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9045
   Icon            =   "frmAddressBook.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdgAddress 
      Left            =   4800
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   960
      TabIndex        =   20
      Top             =   3840
      Width           =   3495
      Begin VB.CommandButton cmdButtonArray 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1800
         TabIndex        =   24
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdButtonArray 
         Caption         =   "Edit Record"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1800
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdButtonArray 
         Caption         =   "Delete Record"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdButtonArray 
         Caption         =   "Add Record"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fraInput 
      Caption         =   "New Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5415
      Begin VB.TextBox txtInputBoxes 
         Height          =   375
         Index           =   8
         Left            =   2880
         TabIndex        =   10
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox txtInputBoxes 
         Height          =   375
         Index           =   7
         Left            =   1080
         TabIndex        =   9
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox txtInputBoxes 
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   8
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtInputBoxes 
         Height          =   375
         Index           =   5
         Left            =   3720
         TabIndex        =   7
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtInputBoxes 
         Height          =   375
         Index           =   4
         Left            =   3120
         TabIndex        =   6
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtInputBoxes 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtInputBoxes 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   4935
      End
      Begin VB.TextBox txtInputBoxes 
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   3
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtInputBoxes 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblAC 
         Caption         =   "AC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblZip 
         Caption         =   "Zip"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   16
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblState 
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblCity 
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblStreet 
         Caption         =   "Street Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblLast 
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblFirst 
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.ListBox lstNames 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      Left            =   5760
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Sav&e As..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmAddressBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Filename As String       'Hold Filename string for Common Dialog

'Type definition to declare Type variables (members)

Private Type AddressBk
     ID As Integer
     FName As String
     LName As String
     Street As String
     City As String
     State As String
     Zip As String
     AC As String
     Phone As String
     Email As String
End Type

'Declare this  variable array as the AddressBk Type
Private m_AddressBook() As AddressBk

'txtInputBoxes control array constants

Private Const txtFName = 0
Private Const txtLName = 1
Private Const txtStreet = 2
Private Const txtCity = 3
Private Const txtState = 4
Private Const txtZip = 5
Private Const txtAC = 6
Private Const txtPhone = 7
Private Const txtEmail = 8

'ButtonArray control array constants

Private Const cmdAdd = 0
Private Const cmdDelete = 1
Private Const cmdEdit = 2
Private Const cmdExit = 3

Private Sub cmdButtonArray_Click(Index As Integer)

'    If the Add Record button is selected _
          Call the AddAddressBk routine _
     If the Delete Record button is selected _
          Call the DeleteAddressRec routine _
     If the Edit Record button is selected _
          Call the RecordEdit routine _
     If the Exit button is selected _
          MsgBox _
          Save the file by calling SaveAddressBk routine _
          Unload the program
      
     Select Case Index
     Case cmdAdd
          AddRecord
     Case cmdDelete
          DeleteAddressRec
     Case cmdEdit
          EditRecord
     Case cmdExit
           If MsgBox("Are you sure you want to Exit & Save this file?", vbQuestion + vbYesNo, _
                    "Address Book") = vbYes Then
               SaveAddressBk m_Filename
               Unload Me
          End If
     End Select
End Sub

Private Sub LoadAddressBook(ByVal p_sFileName As String)

Dim l_intAdd As Integer          'Input file number
Dim l_intCounter As Integer   'Counts array size

'    Title bar caption consisting of "Address Book -"  and the filename using the _
               ExtractFileName function of the selected file _
     Assign the next available free file number to the input file _
     Open the addressbook database text file
     
     Me.Caption = "Address Book - " & ExtractFileName(p_sFileName)
     l_intAdd = FreeFile()
     Open p_sFileName For Input As #l_intAdd
     
'    Initialize the array size counter to 0 _
     Do While not at the End of File _
          Increment the counter by 1 _
          Resize the array based on the counter _
               Read the values in the data file into the array _
               Add item to show LastName, FirstName to the list box _
               Assign ID number _
     Close the File
    
     l_intCounter = 0
     lstNames.Clear
     Do While Not EOF(l_intAdd)
          l_intCounter = l_intCounter + 1
          ReDim Preserve m_AddressBook(1 To l_intCounter) As AddressBk
          With m_AddressBook(l_intCounter)
               Input #l_intAdd, .FName, .LName, .Street, .City, .State, .Zip, .AC, .Phone, .Email
               lstNames.AddItem .LName & ", " & .FName
               lstNames.ItemData(lstNames.NewIndex) = l_intCounter
               .ID = l_intCounter
          End With
     Loop
     Close #l_intAdd
    
'    If at least one record is read, show the first record _
     Enable the delete button if the count is > 0
    
     If l_intCounter > 0 Then lstNames.ListIndex = 0
     cmdButtonArray(cmdDelete).Enabled = l_intCounter > 0
    'LockupTextBoxes
End Sub

Private Sub ShowAddressBk(p_intIndex As Integer)

'    Assign the data from the array into the appropriate text boxes _
     Change the frame caption to reflect the First and last name
     
     With m_AddressBook(p_intIndex)
          txtInputBoxes(txtFName) = .FName
          txtInputBoxes(txtLName) = .LName
          txtInputBoxes(txtStreet) = .Street
          txtInputBoxes(txtCity) = .City
          txtInputBoxes(txtState) = .State
          txtInputBoxes(txtZip) = .Zip
          txtInputBoxes(txtAC) = .AC
          txtInputBoxes(txtPhone) = .Phone
          txtInputBoxes(txtEmail) = .Email
          fraInput.Caption = .FName & "" & .LName
     End With
     LockupTextBoxes               'Lock the text boxes
End Sub

Private Sub ClearInputBoxes()
Dim l_intIndex As Integer          'Index number for control array

'    For each text box in the array, set the value to empty

     For l_intIndex = 0 To txtInputBoxes.UBound
          txtInputBoxes(l_intIndex) = ""
     Next l_intIndex
     
End Sub

Private Sub Form_Load()

'    Open the prog with the LoadAddressBook routine and the Address_Book._db file
'    Assign the filename to the filename string variable

     LoadAddressBook App.Path & "\Address_Book._db"
     m_Filename = App.Path & "\Address_Book._db"

End Sub

Private Sub lstNames_Click()

'     Change frame caption to Firstname Lastname _
     Lock the text boxes
     
     ShowAddressBk lstNames.ItemData(lstNames.ListIndex)
     fraInput.Caption = txtInputBoxes(txtFName) & " " & txtInputBoxes(txtLName)
     LockupTextBoxes
     
End Sub

Private Sub DeleteAddressRec()
Dim l_intIndex As Integer               'Index of items to be deleted

'    If Yes is selected from the MsgBox Then _
          Call ClearInputBoxes _
          Assign the number of items to be deleted to the index variable _
          Check to see if something is selected for deletion _
               Delete the data from the array _
               Remove the items from the listbox _

     If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, _
                    "Address Book") = vbYes Then
          ClearInputBoxes
          l_intIndex = lstNames.ListIndex
          If l_intIndex >= 0 Then
               ClearArrayEntry lstNames.ItemData(l_intIndex)
               lstNames.RemoveItem (l_intIndex)
               
'              Check how many records are left in the addressbook so that we know which one to _
               display next after we delete the currently selected one
               
               If l_intIndex = 0 Then
               
'                   We deleted the 1st entry so check to see if any more exists. If there are no more, _
                    don't select any by assigning the ListIndex property of -1. If there is more, make _
                    the next one the current one.
                    
                    lstNames.ListIndex = IIf(lstNames.ListCount > 0, l_intIndex, -1)
               Else
               
'                   We deleted one other than the first one so show the next one, unless we deleted _
                    the last one which in this case, show the previous one.
                    
                    lstNames.ListIndex = IIf(l_intIndex > lstNames.ListCount - 1, l_intIndex - 1, l_intIndex)
               End If
          End If
          cmdButtonArray(cmdDelete).Enabled = lstNames.ListCount > 0  'Enable the delete _
                                                                                                         button if the ListCount is > 0.
     End If
     LockupTextBoxes          'Lock the text boxes
End Sub

Private Sub ClearArrayEntry(p_intIndex As Integer)
Dim l_intCounter As Integer

'This sub deletes an address entry from the array by moving up all the entries below it up one _
 more, then deleting the last entry that is now blank by ReDimming the array. The entries are _
 assigned new ID's as they are moved up one in the array. If there are no more entries in the _
 addressbook, simply ReDim the array to zero (0).
 
     If p_intIndex < UBound(m_AddressBook) Then
          For l_intCounter = p_intIndex To UBound(m_AddressBook) - 1
               m_AddressBook(l_intCounter + 1).ID = m_AddressBook(l_intCounter).ID
               m_AddressBook(l_intCounter) = m_AddressBook(l_intCounter + 1)
          Next l_intCounter
     End If

     If lstNames.ListCount = 1 Then
          ReDim m_AddressBook(0) As AddressBk
          fraInput.Caption = ""
     Else
          ReDim Preserve m_AddressBook(1 To UBound(m_AddressBook) - 1) As AddressBk
          For l_intCounter = 0 To UBound(m_AddressBook)
               If lstNames.ItemData(l_intCounter) >= p_intIndex Then
                    lstNames.ItemData(l_intCounter) = lstNames.ItemData(l_intCounter) - 1
               End If
          Next l_intCounter
     End If
     
End Sub

Private Sub SaveAddressBk(ByVal p_FileName As String)
Dim l_intCounter As Integer
Dim l_intFile As Integer      'File handle variable of the next available file handle

'    If there is nothing in the AddressBook then Exit the subroutine _
     Assign the next free file number to the file variable _
     OPEN the path to the Addressbook database for Output to the file variable
     
     If UBound(m_AddressBook) = 0 Then Exit Sub
     l_intFile = FreeFile()
     Open p_FileName For Output As #l_intFile

'    For each item in the AddressBook variable _
          Write each item to the database _
     Lock up the text boxes
     
     For l_intCounter = 1 To UBound(m_AddressBook)
          With m_AddressBook(l_intCounter)
               Write #l_intFile, .FName, .LName, .Street, .City, .State, .Zip, .AC, .Phone, .Email
          End With
     Next l_intCounter
     LockupTextBoxes
End Sub

Private Sub AddRecord()
Dim l_intNewIndex As Integer

'    In order to add an entry into the address book, the following steps need to occurr: _
          1.   Add a new entry to the array to hold the data in _
          2.   Add a new entry into the names listbox _
          3.   Assign the new entry the key which is the same as the new number of records in the _
               array _
          4.   Unlock the textboxes for input and set the focus to the first name field
     
     l_intNewIndex = UBound(m_AddressBook) + 1
     ReDim Preserve m_AddressBook(1 To l_intNewIndex)
     lstNames.AddItem "", l_intNewIndex - 1
     lstNames.ItemData(lstNames.NewIndex) = l_intNewIndex
     lstNames.ListIndex = l_intNewIndex - 1
     txtInputBoxes(txtFName).SetFocus
     UnlockTextBoxes

End Sub

Private Sub mnuFileExit_Click()

     cmdButtonArray_Click cmdExit       '    Call the Button array, cmd Exit

End Sub

Private Sub mnuFileOpen_Click()

'    Using the Common Dialog _
          specify the initial directory _
          Filter out all extensions except *._db and *.txt _
          set *.db as the default extension _
          Show the Open Dialog _
               If a Filename has been selected _
                   Call LoadAddressBook _
                   Assign the filename to the Filename string variable
                   
     With cdgAddress
          .InitDir = App.Path
          .Filter = "AddressBook Database (*._db)|*._db| Text Files (*.txt)|*.txt"
          .FilterIndex = 1
          .ShowOpen
          If .FileName <> "" Then
               LoadAddressBook .FileName
               m_Filename = .FileName
          End If
     End With
     
End Sub

Private Sub mnuFileSave_Click()

     SaveAddressBk m_Filename          'Call the SaveAddressBk routine
     
End Sub

Private Sub mnuFileSaveAs_Click()
    
'   Using the Common Dialog _
          specify the initial directory _
          Filter out all extensions except *._db and *.txt _
          set *.db as the default extension _
          Show the Save Dialog _
               If a Filename has been selected _
                   Call SaveAddressBook _
                   Assign the filename to the Filename string variable
     
     With cdgAddress
          .InitDir = App.Path
          .Filter = "AddressBook Database (*._db)|*._db| Text Files (*.txt)|*.txt"
          .FilterIndex = 1
          .ShowSave
               If .FileName <> "" Then
                    SaveAddressBk .FileName
                    m_Filename = .FileName
               End If
     End With
     
End Sub

Private Sub txtInputBoxes_Change(Index As Integer)

'    As soon as the user changes any of the text in the input boxes, reflect the change in the _
     array as well. Also, if the changes are occurring in either the First or Last Name boxes, _
     change the name in the list as well as in the frame caption.

     With m_AddressBook(lstNames.ItemData(lstNames.ListIndex))
          Select Case Index
          Case txtFName
               .FName = txtInputBoxes(Index)
          Case txtLName
               .LName = txtInputBoxes(Index)
          Case txtStreet
               .Street = txtInputBoxes(Index)
          Case txtCity
               .City = txtInputBoxes(Index)
          Case txtState
               .State = txtInputBoxes(Index)
          Case txtZip
               .Zip = txtInputBoxes(Index)
          Case txtAC
               .AC = txtInputBoxes(Index)
          Case txtPhone
               .Phone = txtInputBoxes(Index)
          Case txtEmail
               .Email = txtInputBoxes(Index)
          End Select
     End With
     
'    Change the frame caption to Firstname Lastname _
     If there is only a comma left in the name(meaning both the first and last name fields are _
     empty), don't show anything at all in the frame caption.
     
     fraInput.Caption = txtInputBoxes(txtFName) & " " & txtInputBoxes(txtLName)
     If fraInput.Caption = " " Then fraInput.Caption = ""
     lstNames.List(lstNames.ListIndex) = txtInputBoxes(txtLName) & ", " & _
                                                                      txtInputBoxes(txtFName)
     If lstNames.List(lstNames.ListIndex) = ", " Then _
                                                                      lstNames.List(lstNames.ListIndex) = ""
End Sub

Private Sub EditRecord()
Dim l_blnEdit As Boolean
Dim l_intEditRecNum As Integer            'Store record # being edited
Dim l_intTotalRecs As Integer                'Store total # of records
          
'    check the caption of cmdButtonArry(cmdEdit) to see whether or not it is Edit Record or _
      Replace Record
'     If the Edit button is selected _
          change button caption to "Replace..." _
          Unlock the text boxes for editing
          
     l_blnEdit = cmdButtonArray(cmdEdit).Caption = "Edit Record"
     If l_blnEdit = True Then
          cmdButtonArray(cmdEdit).Caption = "Replace Record"
          UnlockTextBoxes
      Else
      
          'initialize the l_intEditRecNum variable to the correct one selected in the list box using _
          the value stored in the ItemData property so that we know exactly which one to work _
          with in the Array
          
          l_intEditRecNum = lstNames.ItemData(lstNames.ListIndex)
          
'         Assign the changes in the textboxes to the member of the array m_AddressBook

          With m_AddressBook(l_intEditRecNum)
               .FName = txtInputBoxes(txtFName).Text
               .LName = txtInputBoxes(txtLName).Text
               .Street = txtInputBoxes(txtStreet).Text
               .City = txtInputBoxes(txtCity).Text
               .State = txtInputBoxes(txtState).Text
               .Zip = txtInputBoxes(txtZip).Text
               .AC = txtInputBoxes(txtAC).Text
               .Phone = txtInputBoxes(txtPhone).Text
               .Email = txtInputBoxes(txtEmail).Text
               cmdButtonArray(cmdEdit).Caption = "Edit Record"   'Change button caption to "Edit..."
               l_blnEdit = False
          End With
               LockupTextBoxes               'lock the text boxes
     End If
End Sub

Private Sub LockupTextBoxes()
Dim l_intCounter As Integer        'Counter to index the input box control array

'    User defined routine to lock up all text boxes which _
     forces user to select the Edit button to change record content
     
     For l_intCounter = 0 To txtInputBoxes.UBound
               txtInputBoxes(l_intCounter).Locked = True
     Next l_intCounter
         
End Sub

Private Sub UnlockTextBoxes()
Dim l_intCounter As Integer        'Counter to index the input box control array

'    Routine to unlock the text boxes for editing an existing record or to open a new one

     For l_intCounter = 0 To txtInputBoxes.UBound
               txtInputBoxes(l_intCounter).Locked = False
     Next l_intCounter
     
     txtInputBoxes(txtFName).SetFocus        'Set focus to to the First name field
     
End Sub
Private Function ExtractFileName(ByVal p_FileName As String) As String
Dim l_aParse() As String      'dynamic array

'    Split function splits or separates the pathnames separated by the "\" _
     Separate off the parsed string at the last "\", which is the "Filename" as assign it to _
               ExtractFileName function

     l_aParse = Split(p_FileName, "\")
     ExtractFileName = l_aParse(UBound(l_aParse))

End Function
