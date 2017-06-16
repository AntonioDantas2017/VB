VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   14
      Mask            =   "999.999.999-99"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox Format(MaskEdBox1.Text, "dddd, dd mmmm yyyy")

End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    temp = Replace(Replace(MaskEdBox1.Text, "/", ""), "_", "")
    'MsgBox temp
    'vTemp2 = MaskEdBox1.FormattedText
    'MsgBox vTemp2
    'vTemp = MaskEdBox1.Mask
    'MsgBox vTemp
    'MaskEdBox1.Mask = ""
    If Len(temp) = 8 Then
        Command1.Enabled = True
    End If
    'MaskEdBox1.Mask = vTemp
    'MaskEdBox1.Text = vTemp2
End Sub

Private Sub MaskEdBox1_LostFocus()
    Call ValidaData(MaskEdBox1)
End Sub

Private Sub MaskEdBox2_LostFocus()
    Call ValidaCpf(MaskEdBox2)
End Sub
