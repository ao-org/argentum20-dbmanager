VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock MainSocket 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DataBuffer As cStringBuilder

Private Sub Form_Load()
    Dim Puerto As Integer
    Puerto = Val(GetVar(App.Path & "\..\re20-server\Server.ini", "DBMANAGER", "PUERTO"))

    Call MainSocket.Connect("localhost", Puerto)
    
    Set DataBuffer = New cStringBuilder
End Sub

Private Sub MainSocket_Close()
    End
End Sub

Private Sub MainSocket_DataArrival(ByVal bytesTotal As Long)

    Dim Data As String
    Call MainSocket.GetData(Data, vbString, bytesTotal)

    Call DataBuffer.Append(Data)
    
    Dim Separator As Long, Packet As String
    
    Do
        Separator = DataBuffer.Find(Chr(0))
        
        If Separator > 0 Then
            Packet = DataBuffer.SubStr(0, Separator)

            Debug.Print Packet

            Call DataBuffer.Remove(0, Separator)
        End If

    Loop While Separator > 0
    
End Sub
