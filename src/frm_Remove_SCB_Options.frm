VERSION 5.00
Begin VB.Form frm_Remove_SCB_Options 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Fill/Overwrite Checkblocks with"
   ClientHeight    =   2505
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Opt_Fill 
      Caption         =   "Fill with 'long' asmcodes with don't take that much lines in the disassembling than the nop's"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   2775
   End
   Begin VB.OptionButton Opt_NopOut_normal 
      Caption         =   "Add short Jmp at the beginning to jump over the Nop's"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   840
      UseMaskColor    =   -1  'True
      Value           =   -1  'True
      Width           =   2655
   End
   Begin VB.OptionButton Opt_NopOut_complete 
      Caption         =   "NopOut completely"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   2655
   End
End
Attribute VB_Name = "frm_Remove_SCB_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private bFormLoaded As Boolean

Private Sub Form_Load()
 Select Case frmMain.SCBFillOption
   Case FILL_WITH_LONG_ASM
      Opt_Fill.value = True
   
   Case NOP_OUT_COMPLETE
      Opt_NopOut_complete.value = True
      
   Case NOP_OUT_NORMAL
      Opt_NopOut_normal.value = True
 End Select
 
 bFormLoaded = True
 
End Sub

Private Sub Opt_Fill_Click()
   frmMain.SCBFillOption = FILL_WITH_LONG_ASM
   If bFormLoaded Then Me.Hide
End Sub

Private Sub Opt_NopOut_complete_Click()
   frmMain.SCBFillOption = NOP_OUT_COMPLETE
   If bFormLoaded Then Me.Hide
End Sub

Private Sub Opt_NopOut_normal_Click()
   frmMain.SCBFillOption = NOP_OUT_NORMAL
   If bFormLoaded Then Me.Hide
End Sub
