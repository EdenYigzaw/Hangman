'*************************************************************
' PROGRAMME	:	Hangman
'  
' OUTLINE		:	This programme plays a game of hangman 
'                   With the user And allows them To choose 
'                   the difficulty level. 
' 
' PROGRAMMER	:	Eden Yigzaw
'
'  DATE		:	June 16, 2017
' ***********************************************************

Imports System.IO
Public Class frmHangman
    Dim letter As String
    Dim level, ans As String
    Dim val As String
    Dim myChar() As String
    Dim count As Integer
    Dim nums(0 To 19) As Integer

    Dim count2 As Integer
    Public btn() As Button = {btnA, btnB, btnC, btnD, btnE, btnF, btnG, btnH, btnI, btnJ, btnK, btnL, btnM,
                              btnN, btnO, btnP, btnQ, btnR, btnS, btnT, btnU, btnV, btnW, btnX, btnY, btnZ}
    Public numGuess As Integer
    Private Sub btnA_Click(sender As Object, e As EventArgs) Handles btnA.Click
        letter = "A"
        btnA.Enabled = False
        PlaceLetter()
    End Sub
    Private Sub btnB_Click(sender As Object, e As EventArgs) Handles btnB.Click
        letter = "B"
        btnB.Enabled = False
        PlaceLetter()
    End Sub
    Private Sub btnC_Click(sender As Object, e As EventArgs) Handles btnC.Click
        letter = "C"
        btnC.Enabled = False
        PlaceLetter()
    End Sub
    Private Sub btnD_Click(sender As Object, e As EventArgs) Handles btnD.Click
        letter = "D"
        btnD.Enabled = False
        PlaceLetter()
    End Sub
    Private Sub btnE_Click(sender As Object, e As EventArgs) Handles btnE.Click
        letter = "E"
        btnE.Enabled = False
        PlaceLetter()
    End Sub
    Private Sub btnF_Click(sender As Object, e As EventArgs) Handles btnF.Click
        letter = "F"
        btnF.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnG_Click(sender As Object, e As EventArgs) Handles btnG.Click
        letter = "G"
        btnG.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnH_Click(sender As Object, e As EventArgs) Handles btnH.Click
        letter = "H"
        btnH.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnI_Click(sender As Object, e As EventArgs) Handles btnI.Click
        letter = "I"
        btnI.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnJ_Click(sender As Object, e As EventArgs) Handles btnJ.Click
        letter = "J"
        btnJ.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnK_Click(sender As Object, e As EventArgs) Handles btnK.Click
        letter = "K"
        btnK.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnL_Click(sender As Object, e As EventArgs) Handles btnL.Click
        letter = "L"
        btnL.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnM_Click(sender As Object, e As EventArgs) Handles btnM.Click
        letter = "M"
        btnM.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnN_Click(sender As Object, e As EventArgs) Handles btnN.Click
        letter = "N"
        btnN.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnO_Click(sender As Object, e As EventArgs) Handles btnO.Click
        letter = "O"
        btnO.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnP_Click(sender As Object, e As EventArgs) Handles btnP.Click
        letter = "P"
        btnP.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnQ_Click(sender As Object, e As EventArgs) Handles btnQ.Click
        letter = "Q"
        btnQ.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnR_Click(sender As Object, e As EventArgs) Handles btnR.Click
        letter = "R"
        btnR.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnS_Click(sender As Object, e As EventArgs) Handles btnS.Click
        letter = "S"
        btnS.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnT_Click(sender As Object, e As EventArgs) Handles btnT.Click
        letter = "T"
        btnT.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnU_Click(sender As Object, e As EventArgs) Handles btnU.Click
        letter = "U"
        btnU.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnV_Click(sender As Object, e As EventArgs) Handles btnV.Click
        letter = "V"
        btnV.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnW_Click(sender As Object, e As EventArgs) Handles btnW.Click
        letter = "W"
        btnW.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnX_Click(sender As Object, e As EventArgs) Handles btnX.Click
        letter = "X"
        btnX.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnY_Click(sender As Object, e As EventArgs) Handles btnY.Click
        letter = "Y"
        btnY.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub btnZ_Click(sender As Object, e As EventArgs) Handles btnZ.Click
        letter = "Z"
        btnZ.Enabled = False
        PlaceLetter()
    End Sub

    Private Sub cboLevel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLevel.SelectedIndexChanged
        btnPlay.Enabled = True

    End Sub

    Private Sub frmHangman_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        btnDisable()
        btnPlay.Enabled = False
        btnGuess.Enabled = False
        btnChange.Enabled = False

    End Sub
    Private Sub PlaceLetter()
        Dim c As Integer

        count += 1
        lblGuessNum.Text = numGuess - count

        lblPicked.Text += letter
        For i As Integer = 0 To val.Length - 1
            If val.Chars(i) = letter Then
                c += 1
                If c = 1 Then
                    lblWord.Text = Nothing
                End If
                myChar(i) = " " & val.Chars(i) & " "
                lblWord.Text = Nothing
                For u As Integer = 0 To val.Length - 1
                    lblWord.Text &= myChar(u)
                Next
            End If
        Next

        If lblWord.Text = val Then
            MessageBox.Show("YOU WON!!!")
            btnGuess.Enabled = False

            Exit Sub
        End If

        If count = numGuess Then
            ans = InputBox("The answer is....?").ToUpper
            If ans <> val Then
                MessageBox.Show("YOU LOSE!!! YOU RAN OUT OF GUESSES!!! The word was " & val)
                btnDisable()
                lblWord.Text = val
                btnGuess.Enabled = False
                btnPlay.Enabled = True

                Exit Sub
            Else
                MessageBox.Show("YOU WON!!!")
                lblWord.Text = val
                btnPlay.Enabled = True
                btnGuess.Enabled = False
                Exit Sub
            End If


        End If

    End Sub
    Private Sub read()
        Dim word As StreamReader
        Dim count1 As Integer
        Dim rand As New Random
        Dim num As Integer


        Dim outer, inner As Integer
        Dim count As Integer
        Dim strResult As String = ""

        If count2 = 0 Then
            Randomize()

            For outer = 0 To 19
                nums(outer) = Int(Rnd() * 20) + 1
                For inner = 0 To outer
                    If inner < outer Then
                        If nums(outer) = nums(inner) Then
                            outer = outer - 1
                            Exit For
                        End If
                    End If
                Next inner
            Next outer
        End If

        word = File.OpenText(level)
        Do
            val = word.ReadLine()
            If count1 = nums(count2) Then
                count2 += 1
                Exit Sub
            End If
            count1 += 1
        Loop

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub

    Private Sub btnGuess_Click(sender As Object, e As EventArgs) Handles btnGuess.Click
        ans = InputBox("The answer is....?").ToUpper
        If ans <> val Then
            MessageBox.Show("INCORRECT. The word was " & val)
            cboLevel.Enabled = False
            btnGuess.Enabled = False
            btnPlay.Enabled = True
            btnDisable()
            lblWord.Text = val

            Exit Sub
        Else
            MessageBox.Show("YOU WON!!!")
            btnGuess.Enabled = False
            btnPlay.Enabled = True
            lblWord.Text = val
            Exit Sub
        End If
    End Sub

    Private Sub btnPlay_Click(sender As Object, e As EventArgs) Handles btnPlay.Click
        btnGuess.Enabled = True
        cboLevel.Enabled = True
        btnChange.Enabled = True
        btnPlay.Enabled = False
        lblPicked.Text = Nothing
        count = 0
        If cboLevel.SelectedItem = "Easy" Then
            level = "Easy.txt"
            numGuess = 8
        ElseIf cboLevel.SelectedItem = "Medium" Then
            level = "Medium.txt"
            numGuess = 10
        ElseIf cboLevel.SelectedItem = "Hard" Then
            level = "Hard.txt"
            numGuess = 12
        End If
        cboLevel.Enabled = False
        lblWord.Text = Nothing
        read()
        ReDim myChar(val.Length - 1)
        For a As Integer = 0 To val.Length - 1
            myChar(a) = " _ "
            lblWord.Text &= myChar(a)
        Next

        btnEnable()

        lblGuessNum.Text = numGuess - count

    End Sub
    Private Sub btnDisable()
        Dim btn() As Button = {btnA, btnB, btnC, btnD, btnE, btnF, btnG, btnH, btnI, btnJ, btnK, btnL, btnM,
                              btnN, btnO, btnP, btnQ, btnR, btnS, btnT, btnU, btnV, btnW, btnX, btnY, btnZ}
        For i As Integer = 0 To btn.Length - 1
            btn(i).Enabled = False
        Next

    End Sub
    Private Sub btnEnable()
        Dim btn() As Button = {btnA, btnB, btnC, btnD, btnE, btnF, btnG, btnH, btnI, btnJ, btnK, btnL, btnM,
                              btnN, btnO, btnP, btnQ, btnR, btnS, btnT, btnU, btnV, btnW, btnX, btnY, btnZ}
        For i As Integer = 0 To btn.Length - 1
            btn(i).Enabled = True
        Next

    End Sub

    Private Sub btnChange_Click(sender As Object, e As EventArgs) Handles btnChange.Click
        Application.Restart()
    End Sub
End Class
