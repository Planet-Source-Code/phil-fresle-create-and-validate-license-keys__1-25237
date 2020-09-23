VERSION 5.00
Begin VB.Form FTest 
   Caption         =   "Test Registration (Version 2)"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Registered Owner Licence"
      Height          =   2835
      Left            =   0
      TabIndex        =   15
      Top             =   2520
      Width           =   4155
      Begin VB.TextBox txtOwnerKey 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   3
         Left            =   2700
         TabIndex        =   29
         Top             =   1680
         Width           =   795
      End
      Begin VB.TextBox txtOwnerKey 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   1860
         TabIndex        =   28
         Top             =   1680
         Width           =   795
      End
      Begin VB.TextBox txtOwnerKey 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   1020
         TabIndex        =   27
         Top             =   1680
         Width           =   795
      End
      Begin VB.TextBox txtOwner 
         Height          =   315
         Left            =   1560
         TabIndex        =   21
         Text            =   "John Smith"
         Top             =   600
         Width           =   2115
      End
      Begin VB.CommandButton cmdTestOwner 
         Caption         =   "Test License Key"
         Height          =   495
         Left            =   180
         TabIndex        =   19
         Top             =   2100
         Width           =   2115
      End
      Begin VB.TextBox txtOwnerKey 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   1680
         Width           =   795
      End
      Begin VB.TextBox txtOwnerApp 
         Height          =   315
         Left            =   2520
         TabIndex        =   17
         Text            =   "JJ201X"
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdGenerateOwner 
         Caption         =   "Generate License Key"
         Height          =   495
         Left            =   180
         TabIndex        =   16
         Top             =   1080
         Width           =   2115
      End
      Begin VB.Label Label4 
         Caption         =   "Registered Owner"
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Application Specific Characters"
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   300
         Width           =   2355
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Alpha-Numeric Generic Licence"
      Height          =   2415
      Left            =   4200
      TabIndex        =   9
      Top             =   0
      Width           =   4215
      Begin VB.TextBox txtAlphaNumericKey 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   4
         Left            =   3300
         TabIndex        =   26
         Top             =   1320
         Width           =   795
      End
      Begin VB.TextBox txtAlphaNumericKey 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   3
         Left            =   2520
         TabIndex        =   25
         Top             =   1320
         Width           =   795
      End
      Begin VB.TextBox txtAlphaNumericKey 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   1740
         TabIndex        =   24
         Top             =   1320
         Width           =   795
      End
      Begin VB.TextBox txtAlphaNumericKey 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   23
         Top             =   1320
         Width           =   795
      End
      Begin VB.CommandButton cmdGenerateAlphaNumeric 
         Caption         =   "Generate License Key"
         Height          =   495
         Left            =   180
         TabIndex        =   13
         Top             =   720
         Width           =   2115
      End
      Begin VB.TextBox txtAlphaNumericApp 
         Height          =   315
         Left            =   2520
         TabIndex        =   12
         Text            =   "AB101"
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtAlphaNumericKey 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   1320
         Width           =   795
      End
      Begin VB.CommandButton cmdTestAlphaNumeric 
         Caption         =   "Test License Key"
         Height          =   495
         Left            =   180
         TabIndex        =   10
         Top             =   1740
         Width           =   2115
      End
      Begin VB.Label Label2 
         Caption         =   "Application Specific Characters"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   300
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Original Numeric Licence"
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4155
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate License Key"
         Height          =   495
         Left            =   180
         TabIndex        =   7
         Top             =   720
         Width           =   2115
      End
      Begin VB.TextBox txtUserSpecifiedPart1 
         Height          =   315
         Left            =   3060
         TabIndex        =   6
         Text            =   "00101"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtPart1 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtPart2 
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtPart3 
         Height          =   315
         Left            =   1740
         TabIndex        =   3
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtPart4 
         Height          =   315
         Left            =   2460
         TabIndex        =   2
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test License Key"
         Height          =   495
         Left            =   180
         TabIndex        =   1
         Top             =   1740
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "Part 1 of Key (5 Numeric Characters):"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Width           =   2775
      End
   End
End
Attribute VB_Name = "FTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       FTest
' FILENAME:     C:\My Code\vb\Registration\FTest.frm
' AUTHOR:       Phil Fresle
' CREATED:      06-Sep-2000
' COPYRIGHT:    Copyright 2000 Frez Systems Limited.
'
' DESCRIPTION:
' Used to test class for generating and testing license keys.
'
' This is 'free' software with the following restrictions:
'
' You may not redistribute this code as a 'sample' or 'demo'. However, you are free
' to use the source code in your own code, but you may not claim that you created
' the sample code. It is expressly forbidden to sell or profit from this source code
' other than by the knowledge gained or the enhanced value added by your own code.
'
' Use of this software is also done so at your own risk. The code is supplied as
' is without warranty or guarantee of any kind.
'
' Should you wish to commission some derivative work based on this code provided
' here, or any consultancy work, please do not hesitate to contact us.
'
' Web Site:  http://www.frez.co.uk
' E-mail:    sales@frez.co.uk
'
' MODIFICATION HISTORY:
' 1.0       06-Sep-2000
'           Phil Fresle
'           Initial Version
'*******************************************************************************
Option Explicit

'*******************************************************************************
' cmdGenerate_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Test key generation
'*******************************************************************************
Private Sub cmdGenerate_Click()
    Dim oRegistration   As CRegistration
    Dim sKey            As String
    
    Set oRegistration = New CRegistration
    
    sKey = oRegistration.GenerateKey(txtUserSpecifiedPart1.Text)
    
    txtPart1.Text = Left(sKey, 5)
    txtPart2.Text = Mid(sKey, 6, 4)
    txtPart3.Text = Mid(sKey, 10, 4)
    txtPart4.Text = Mid(sKey, 14, 4)

    Set oRegistration = Nothing
End Sub

'*******************************************************************************
' cmdGenerateAlphaNumeric_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Test key generation for generic alpha numeric keys
'*******************************************************************************
Private Sub cmdGenerateAlphaNumeric_Click()
    Dim oReg As CGenericRegistration
    Dim sKey As String
    
    Set oReg = New CGenericRegistration
    
    sKey = oReg.GenerateKey(txtAlphaNumericApp.Text)
    
    txtAlphaNumericKey(0).Text = Left(sKey, 5)
    txtAlphaNumericKey(1).Text = Mid(sKey, 6, 5)
    txtAlphaNumericKey(2).Text = Mid(sKey, 11, 5)
    txtAlphaNumericKey(3).Text = Mid(sKey, 16, 5)
    txtAlphaNumericKey(4).Text = Mid(sKey, 21, 5)
    
    Set oReg = Nothing
End Sub

'*******************************************************************************
' cmdGenerateOwner_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Test key generation for owner keys
'*******************************************************************************
Private Sub cmdGenerateOwner_Click()
    Dim oReg As COwnerRegistration
    Dim sKey As String
    
    Set oReg = New COwnerRegistration
    
    sKey = oReg.GenerateKey(txtOwner.Text, txtOwnerApp.Text)
    
    txtOwnerKey(0).Text = Left(sKey, 4)
    txtOwnerKey(1).Text = Mid(sKey, 5, 4)
    txtOwnerKey(2).Text = Mid(sKey, 9, 4)
    txtOwnerKey(3).Text = Mid(sKey, 13, 4)
    
    Set oReg = Nothing
End Sub

'*******************************************************************************
' cmdTest_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Test key validation
'*******************************************************************************
Private Sub cmdTest_Click()
    Dim oRegistration   As CRegistration
    Dim sKey            As String
    
    sKey = Trim(txtPart1.Text) & Trim(txtPart2.Text) _
        & Trim(txtPart3.Text) & Trim(txtPart4.Text)
        
    Set oRegistration = New CRegistration
    
    If oRegistration.IsKeyOK(sKey) Then
        MsgBox "Key is valid"
    Else
        MsgBox "Key is NOT valid"
    End If
    
    Set oRegistration = Nothing
End Sub

'*******************************************************************************
' cmdTestAlphaNumeric_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Test key validation
'*******************************************************************************
Private Sub cmdTestAlphaNumeric_Click()
    Dim sKey As String
    Dim oReg As CGenericRegistration
    
    sKey = txtAlphaNumericKey(0).Text & _
        txtAlphaNumericKey(1).Text & _
        txtAlphaNumericKey(2).Text & _
        txtAlphaNumericKey(3).Text & _
        txtAlphaNumericKey(4).Text
        
    Set oReg = New CGenericRegistration
    
    If oReg.IsKeyOK(sKey, txtAlphaNumericApp.Text) Then
        MsgBox "Key is OK"
    Else
        MsgBox "Key is BAD"
    End If
    
    Set oReg = Nothing
End Sub

'*******************************************************************************
' cmdTestOwner_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Test key validation
'*******************************************************************************
Private Sub cmdTestOwner_Click()
    Dim sKey As String
    Dim oReg As COwnerRegistration
    
    sKey = txtOwnerKey(0).Text & _
        txtOwnerKey(1).Text & _
        txtOwnerKey(2).Text & _
        txtOwnerKey(3).Text
        
    Set oReg = New COwnerRegistration
    
    If oReg.IsKeyOK(sKey, txtOwner.Text, txtOwnerApp.Text) Then
        MsgBox "Key is OK"
    Else
        MsgBox "Key is BAD"
    End If
    
    Set oReg = Nothing
End Sub
