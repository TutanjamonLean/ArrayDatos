VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21615
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   21615
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   9600
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Dia*Mes*Año"
      Height          =   615
      Left            =   9360
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As Integer
Dim arraydato() As String
Private Sub Command1_Click()
    
        datos dividir(Label1.Caption, "*")
    
End Sub
Private Function datos(Datosave() As String) As String

    For A = LBound(Datosave) To UBound(Datosave)

        Print Datosave(A)

    Next A
End Function
Private Function dividir(text As String, caract As String) As String()

    dividir = Split(text, caract)

End Function
