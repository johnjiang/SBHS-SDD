VERSION 5.00
Begin VB.Form matrixcouncil 
   Caption         =   "The Matrix Council"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   13515
   LinkTopic       =   "Form1"
   Picture         =   "matrixcouncil.frx":0000
   ScaleHeight     =   8115
   ScaleWidth      =   13515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Exit The Council"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton back 
      Caption         =   "Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Nextpage 
      Caption         =   "Next"
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "matrixcouncil.frx":1BEA3
      Left            =   5640
      List            =   "matrixcouncil.frx":1BEB3
      TabIndex        =   0
      Text            =   "Select Your Area of Query"
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   4080
      TabIndex        =   1
      Top             =   840
      Width           =   5415
   End
End
Attribute VB_Name = "matrixcouncil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer

Private Sub Combo1_Click()
    reset
    Select Case Combo1.ListIndex
        Case 0
            a = 1
        Case 1
            b = 1
        Case 2
            c = 1
        Case 3
            d = 1
    End Select
    check
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Form_QueryUnLoad(cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure you want to leave the council?", vbQuestion + vbYesNo, "Leave Council") = vbNo Then
        cancel = True
    End If
End Sub
Private Sub form_unload(cancel As Integer)
    matrixmain.Enabled = True
End Sub

Private Sub Nextpage_Click()
    Select Case Combo1.ListIndex
        Case 0
            a = a + 1
        Case 1
            b = b + 1
        Case 2
            c = c + 1
        Case 3
            d = d + 1
    End Select
    check
    Nextpage.Enabled = False
    back.Enabled = True
End Sub

Private Sub back_Click()
    Select Case Combo1.ListIndex
        Case 0
            a = a - 1
        Case 1
            b = b - 1
        Case 2
            c = c - 1
        Case 3
            d = d - 1
    End Select
    back.Enabled = False
    Nextpage.Enabled = True
    check
End Sub

Public Sub check()
    back.Enabled = False
    Nextpage.Enabled = True
    Select Case a
        Case 1
            Label1.Caption = "Storyline - The Logos is in your hands now Captain. One by one the sentinels are engaging in your path to Zion. You have the only ship left with an active EMP. One shot and *boom*, down come falling 1 million squidies. Yes! A light at the end of the tunnel, will you make it or will the sentinels take down your ship just like they did with the 10 others. Don't let the magnetic lag weigh you down, you must focus and reach the end of the tunnel!"
        Case 2
            Label1.Caption = "Controls and Instructions - Use the arrow keys 'left', 'right', 'up', 'down' to navigate the Logos across the screen. Guide the Logos across the screen to the other side. Once you reach the other side you'll start off again at the beginning however there will be one more sentinel on the screen. Dodge the sentinels, if you touch them, you are as good as dead."
    End Select
    Select Case b
        Case 1
            Label1.Caption = "Storyline - The agents are here and they are hear to kill you. Are you the one? Have you got the skills in you to beat the agents? Remember, no one has taken on agents and survived. The best idea is to run. Are you going to run? If you win...then you are the one...if you lose...you're just another person stupid enough to take on three agents. If you do manage to beat the agents though, there's another surprise at the end of the road for you."
        Case 2
            Label1.Caption = "Controls and Instructions - Use your mouse to shoot the Agents by clicking on them, that's falling down the screen. You have to be quick or else the system exploit will cause severe damage to your health. However, you with your hacking abilities is able to hack the matrix and slow down the Agents. Remember, only the One can succesfully deafeat the agents. Are you the one? Shoot down the Agents as quickly as possible."
    End Select
    Select Case c
        Case 1
            Label1.Caption = "Storyline - The Machines are digging...each minute they are getting closer to Zion. The only way to stop them is to dock all the ships and set up paremeters for defence. You don't have long before the machines reach you. You must station all the ships before you completely lose the dock. You have the lives of the whole of Zion in your hands. If you fail...there could only be one hope left...Neo. Remember light as a feather and don't over squeeze the mouse"
        Case 2
            Label1.Caption = "Controls and Instructions - Use your mouse to drag the hovercrafts onto the designated stations. Do it before the time reaches 0 and help save Zion from destrcution."
    End Select
    Select Case d
        Case 1
            Label1.Caption = "Storyline - The Captain of the hovercraft 'Osiris' has ran out of energy and requires a recharge, unfortunately this process takes up 30 minutes...30 minutes which the crew do not have. There is an incoming wave of sentinels which have only one purporse...'search and destroy'. You must help Thadeus and his crew recharge the ship and deliver Thadeus's message to the rest of Zion warning them of the attack unleashed by the machines. An EMP is up and ready however requires 50% of the ship's total energy, use it sparingly and wisely."
        Case 2
            Label1.Caption = "Controls and Instructions - Click on Roll dice to see your outcome. Select on delay ship or fuel ship to end your turn. The sentinels are arriving at a rate of 2 minutes each turn. You can use your EMP upon reaching 50% however, it will also deplete 50% of your energy whilst setting the sentinel stunned for 30 minutes. Use it wisely."
    End Select
End Sub

Public Sub reset()
    a = 0
    b = 0
    c = 0
    d = 0
End Sub

