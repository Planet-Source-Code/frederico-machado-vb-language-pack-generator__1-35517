VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Language pack test form"
   ClientHeight    =   3840
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5430
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Open frmTest"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enumerate Language Packs"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin VB.ListBox lstLangPacks 
      Height          =   840
      ItemData        =   "frmMain.frx":08D2
      Left            =   240
      List            =   "frmMain.frx":08D4
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Language Pack"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   $"frmMain.frx":08D6
      Height          =   675
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "You can make labels or buttons with tooltiptext."
      Height          =   195
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Test TooltipText..."
      Top             =   1800
      Width           =   3330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This is a test."
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   930
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuTest 
         Caption         =   "You can translate menus.."
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // Example: Use of the Language Pack Class Module
' // Developed by Frederico Machado (indiofu@bol.com.br)
' // Vote for me if you like it please!
' //
' // I don't know if it have bugs, cause I haven't tested
' // it deeply. If you find any bug in the Packer or in
' // the Class Module, please, let me know about it.
' // Thank you!
' // P.S.: Please, don't forget to give me some credit
' // if you use this code in your own VB softwares.
' /////////////////////////////////////////////////////////

Private Sub Command1_Click()
  ' Just to be sure if there is Language Packs loaded or if the user selected one
  If lstLangPacks.ListCount = 0 Or lstLangPacks.ListIndex = -1 Then Exit Sub
  
  ' Lets load the entire language pack. It doesn't apply the language pack in the form.
  cLanguage.LoadLanguagePack App.Path & "\packs\" & lstLangPacks.List(lstLangPacks.ListIndex)
  ' Now it applies the language pack in the form
  cLanguage.SetLanguageInForm Me
End Sub

Private Sub Command2_Click()
  ' Clear the listbox. If we clicked in Command2 more than one time, the packs don't repeat in it.
  lstLangPacks.Clear
  
  Dim sTmp As String, sTmpArray() As String, i As Integer
  
  ' Set the temp variable with the function that returns the packs found separated by |
  sTmp = cLanguage.EnumLanguagePacks(App.Path & "\packs", "*.lpk")
  ' Lets split the temp variable into the temp array
  sTmpArray = Split(sTmp, "|")
  ' Lets put the file into the listbox
  For i = 0 To UBound(sTmpArray)
    ' Just to be sure that it's not empty
    If sTmpArray(i) <> "" Then lstLangPacks.AddItem sTmpArray(i)
  Next
End Sub

Private Sub Command3_Click()
  ' Tcharan!
  frmTest.Show
End Sub

Private Sub Command4_Click()
  ' Don't leave, please! :)
  End
End Sub

Private Sub mnuExit_Click()
  ' Don't leave, please! :)
  End
End Sub
