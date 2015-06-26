VERSION 5.00
Begin VB.Form GiveAdvice 
   Caption         =   "GiveAdvice"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   5670
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   3735
   End
   Begin VB.CommandButton Send 
      Caption         =   "发送"
      Height          =   615
      Left            =   4080
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "GiveAdvice.frx":0000
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "如果您不介意的话可以留下您的联系方式:"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   3330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "您要提的建议将会发送到开发者,再次感谢您的支持!"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4140
   End
End
Attribute VB_Name = "GiveAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendEmail Lib "SMTPDLL.dll" (ByVal SMTPServer As String, ByVal EMailName As String, ByVal Password As String, ByVal Dest As String, ByVal Sorc As String, ByVal Subject As String, ByVal BodyCaption As String) As Boolean

Private Sub Send_Click()
If SendEmail("smtp.sina.cn", "produceadvice", "A123456", "lcatro@sina.cn", "produceadvice@sina.cn", "图片加密小程序建议", Text1.Text & IIf(Text2.Text = "", "", vbCrLf & vbCrLf & "该网友的联系方式:" & Text2.Text)) Then
    MsgBox "发送成功!", vbOKOnly
    Unload Me
Else
    MsgBox "发送失败!", vbCritical
End If
End Sub
