VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DistInputs_RowOption 
   Caption         =   "UserForm1"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4680
   OleObjectBlob   =   "DistInputs_RowOption.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DistInputs_RowOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOk_Click()
    Hide
End Sub
Private Sub cmdCancel_Click()
    MsgBox "Macro has been halted."
    
    StateToggle.UpdateScreen "On"
    Unload Me
    End
End Sub
Private Sub UserForm_Initialize()
    TransIn = ""
    TransOut = ""
    Distributions = ""
    DistRow = ""
    
    Dim WidthCenter As Variant
    Dim HeightCenter As Variant
    
    WidthCenter = Application.Width / 2
    HeightCenter = Application.Height / 2
    
    Me.StartUpPosition = 0
    Me.Top = Application.Top + HeightCenter - Me.Height / 2
    Me.Left = Application.Left + WidthCenter - Me.Width / 2
End Sub
