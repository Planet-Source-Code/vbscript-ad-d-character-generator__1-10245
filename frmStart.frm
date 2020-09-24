VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AD&D Character Database"
   ClientHeight    =   5385
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7170
   ControlBox      =   0   'False
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStart.frx":030A
   ScaleHeight     =   5385
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh Count"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtNumChar 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmStart.frx":7EC4C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   5895
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
      NegotiatePosition=   3  'Right
   End
   Begin VB.Menu mnuCharacter 
      Caption         =   "&Character"
      Begin VB.Menu mnuViewChar 
         Caption         =   "&View Character"
      End
      Begin VB.Menu mnuEditChar 
         Caption         =   "&Edit Character"
      End
      Begin VB.Menu mnuAddChar 
         Caption         =   "&Add Character"
      End
      Begin VB.Menu mnuDelChar 
         Caption         =   "&Delete Character"
      End
   End
   Begin VB.Menu mnuEditTables 
      Caption         =   "&Edit Tables"
      Begin VB.Menu mnuClasses 
         Caption         =   "&Classes"
      End
      Begin VB.Menu mnuRaces 
         Caption         =   "&Races"
      End
      Begin VB.Menu mnuAlignments 
         Caption         =   "&Alignments"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+------------------------------------------------------------+
'|
'| Title:   frmStart
'| Purpose: This is the initial form.  All other forms are
'|          accessed through this on via a menu item
'| Author:  Bradley Buskey
'|
'+------------------------------------------------------------+

Option Explicit
    Dim db As Database          'declares the database variable
    Dim rsChar As Recordset     'declares the recordset variable

Private Sub Form_Load()
    
    Set db = OpenDatabase(App.Path + "\cha.mdb")    'opens the database
    Set rsChar = db.OpenRecordset("tblCha")         'populates recordset
        
    'The following lines get the number of records and displays in the textbox.
    'It tests for number of records and whether there are no records.
    If rsChar.EOF Then
        txtNumChar.Text = "There are no records in the database."
    Else
        If rsChar.RecordCount < 2 Then
            txtNumChar.Text = "There is " & rsChar.RecordCount & " character in the database."
        Else
            txtNumChar.Text = "There are " & rsChar.RecordCount & " characters in the database."
        End If
    End If
End Sub

Private Sub mnuExit_Click()     'closes recordset, database and form
    rsChar.Close
    db.Close
    End
End Sub

Private Sub mnuViewChar_Click()     'Shows/Loads the View Character form
    frmViewChar.Show
End Sub

Private Sub mnuEditChar_Click()     'Shows/Loads the Edit Character form
    frmEditChar.Show
End Sub

Private Sub mnuAddChar_Click()      'Shows/Loads the Add Character form
    frmAddChar.Show
End Sub

Private Sub mnuDelChar_Click()      'Shows/Loads the Delete Character form
    frmDelChar.Show
End Sub

Private Sub mnuClasses_Click()      'Shows/Loads the Classes form
    frmEditClasses.Show
End Sub

Private Sub mnuRaces_Click()        'Shows/Loads the Races form
    frmEditRaces.Show
End Sub

Private Sub mnuAlignments_Click()   'Shows/Loads the Alignment form
    frmEditAlign.Show
End Sub

Private Sub mnuAbout_Click()        'Shows/Loads the About form
    frmAbout.Show
End Sub

Private Sub cmdRefresh_Click()
    Set db = OpenDatabase(App.Path + "\cha.mdb")    'opens the database
    Set rsChar = db.OpenRecordset("tblCha")         'populates recordset
    'The following lines get the number of records and displays in the textbox.
    'It tests for number of records and whether there are no records.
    If rsChar.EOF Then
        txtNumChar.Text = "There are no records in the database."
    Else
        If rsChar.RecordCount < 2 Then
            txtNumChar.Text = "There is " & rsChar.RecordCount & " character in the database."
        Else
            txtNumChar.Text = "There are " & rsChar.RecordCount & " characters in the database."
        End If
    End If
End Sub
