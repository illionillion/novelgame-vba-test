VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} addTextForm 
   Caption         =   "�e�L�X�g�o�^"
   ClientHeight    =   7128
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6768
   OleObjectBlob   =   "addTextForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "addTextForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim char, charpic, background, text As String
Dim lastrow As Long

    If ComboBox1.ListIndex = -1 Then
        char = ComboBox1.text
    Else
        char = ComboBox1.List(ComboBox1.ListIndex)
    End If

    If ComboBox2.ListIndex = -1 Then
        charpic = ComboBox2.text
    Else
        charpic = ComboBox2.List(ComboBox2.ListIndex)
    End If

    If ComboBox3.ListIndex = -1 Then
        background = ComboBox3.text
    Else
        background = ComboBox3.List(ComboBox3.ListIndex)
    End If
    
    text = TextBox1.text

    lastrow = Sheets("�V�i���I�V�[�g").Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    Sheets("�V�i���I�V�[�g").Cells(lastrow, 1).Value = lastrow - 1
    Sheets("�V�i���I�V�[�g").Cells(lastrow, 2).Value = char
    Sheets("�V�i���I�V�[�g").Cells(lastrow, 3).Value = charpic
    Sheets("�V�i���I�V�[�g").Cells(lastrow, 4).Value = background
    Sheets("�V�i���I�V�[�g").Cells(lastrow, 5).Value = text
    

End Sub

Private Sub UserForm_Initialize()

Dim charalastrow, charaPiclastrow, backGroundlastrow As Long

charalastrow = Sheets("���X�g�V�[�g").Cells(Rows.Count, 1).End(xlUp).Row
ComboBox1.RowSource = "���X�g�V�[�g!" & Range("A2", "A" & charalastrow).Address

charaPiclastrow = Sheets("���X�g�V�[�g").Cells(Rows.Count, 2).End(xlUp).Row
ComboBox2.RowSource = "���X�g�V�[�g!" & Range("B2", "B" & charaPiclastrow).Address

backGroundlastrow = Sheets("���X�g�V�[�g").Cells(Rows.Count, 3).End(xlUp).Row
ComboBox3.RowSource = "���X�g�V�[�g!" & Range("C2", "C" & backGroundlastrow).Address


End Sub

