Attribute VB_Name = "modProg_9_Documents"
Option Explicit
'Title:              Program #8 - AddressBook
'Name:               Jack Gray
'Date:               3/24/2002
'Course:             CSC 265
'Section:            001
'Description:        This program demonstrates the use of the SaveAs and Open Common _
                              Dialog boxes. Arrays were used for text entry boxes and for command _
                              boxes. A module is not used because there are no global variables. _
                              User defined Type is declared in the AddressBook form declaration section. _
                              Named constants are used to hold the text box and command button _
                              arrays. The Address Book also demonstrates the use of sequential data _
                              manipulation. Data entered into the various text boxes is placed in a text(.txt) _
                              file in a sequential manner. The data in the default text file is automatically _
                              loaded when the program is started. The functionality will allow the user to _
                              create new data files using the SaveAs Dialog allowing creation of various _
                              files such as Friends, Family or business data files.
'
'Data Requirements:  Text (.txt) data files are used
'
'Formulas:           None
'
'Initial Algorithm:           B.   frmAddressBook
'                                            1. Menu
'                                                 a. File
'                                                      1) Open...
'                                                      2) Save
'                                                      3) Save As...
'                                                      4) Exit
'                                            2.   Various Text boxes to hold the following data
'                                                 - Data held in a control array
'                                                 a.   First Name
'                                                 b.   Last Name
'                                                 c.   Street Address
'                                                 d.   City
'                                                 e.   State
'                                                 f.   Zip code
'                                                 i.   Area code
'                                                 j.   Phone number
'                                                 k.   Email
'                                            3.   Command buttons
'                                                 - All buttons in a control array
'                                                 a.   Add Record
'                                                      - Used to enter new records to the data base
'                                                 b.   Delete Record
'                                                      - Used to delete a record from the data base
'                                                 c.   Edit Record
'                                                      -Used to make changes to exiting records
'                                                      -User defined LockupTextBoxes routine forces user to select
'                                                           "Edit Record" button which unlocks the text boxes for editing
'                                                      -When selected, the button caption changes to "Replace
'                                                           Record" which makes the needed changes in the database
'                                                           and locks the text boxes again
'                                                 d.   Exit
'                                                      -Exits the program AND saves the file
'                                                      -Uses a MsgBox to confirm
'                                            4.   Listbox
'                                                 -Used to store the name of the person in the data file in the format:
'                                                      Lastname, Firstname
'                                            5.   Common Dialog
'                                                 a.   Used for "Open" on the File menu
'                                                 b.   Used for "Save As..." on the file menu
'                                            6.   Labels appropriate to text boxes
'                                            7.   Frames to hold the Textbox control array and another for the
'                                                      Command buttons
'                                                 -The Textbox frame caption will hold the name of the selected
'                                                      record or the name of the New record being entered using the
'                                                      "Add Record" button. Format is Firstname Lastname.
'

                                                  
