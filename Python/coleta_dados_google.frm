VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Captar dados"
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim objShell As Object
Dim PythonExePath, PythonScriptPath As String

    Set objShell = VBA.CreateObject("Wscript.Shell")
    
    PythonExePath = """C:\Python310\python.exe"""
    PythonScriptPath = "C:\Users\Lucas Moreira\Desktop\Projetos\Python\Pesquisa Google.py"
    
    objShell.Run PythonExePath & PythonScriptPath
    

End Sub
