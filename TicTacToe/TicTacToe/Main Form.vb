' Programmer: Eric Patrick

Option Strict On
Option Explicit On
Option Infer Off

Public Class frmMain

    ' File > Exit
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    ' File > Start New Game
    Private Sub NewGameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewGameToolStripMenuItem.Click
        StartNewGame()
    End Sub

    'Help > About
    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        MessageBox.Show("Program by Eric Patrick." & Environment.NewLine & "Source code available at https://github.com/epatr/cop-2837-final", "About", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' Clear all buttons and start over.
    Private Sub StartNewGame()
        btn1.Text = String.Empty
        btn2.Text = String.Empty
        btn3.Text = String.Empty
        btn4.Text = String.Empty
        btn5.Text = String.Empty
        btn6.Text = String.Empty
        btn7.Text = String.Empty
        btn8.Text = String.Empty
        btn9.Text = String.Empty
    End Sub

    ' Boolean check to see if the move wins the game
    Private Function CheckForWin(ByVal player As String) As Boolean
        ' Check all possible wins
        If btn1.Text = player AndAlso btn2.Text = player AndAlso btn3.Text = player Then
            Return True
        ElseIf btn4.Text = player AndAlso btn5.Text = player AndAlso btn6.Text = player Then
            Return True
        ElseIf btn7.Text = player AndAlso btn8.Text = player AndAlso btn9.Text = player Then
            Return True
        ElseIf btn1.Text = player AndAlso btn4.Text = player AndAlso btn7.Text = player Then
            Return True
        ElseIf btn2.Text = player AndAlso btn5.Text = player AndAlso btn8.Text = player Then
            Return True
        ElseIf btn3.Text = player AndAlso btn6.Text = player AndAlso btn9.Text = player Then
            Return True
        ElseIf btn1.Text = player AndAlso btn5.Text = player AndAlso btn9.Text = player Then
            Return True
        ElseIf btn3.Text = player AndAlso btn5.Text = player AndAlso btn7.Text = player Then
            Return True
        Else
            Return False
        End If
    End Function

    ' Ask the winning player if they want to play again
    Private Sub PlayerWins()
        Dim result As Integer = MessageBox.Show("You've won! Play again?", "Victory!", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Me.Close()
        ElseIf result = DialogResult.Yes Then
            StartNewGame()
        End If
    End Sub

    ' Ask the losing player if they want to play again
    Private Sub ComputerWins()
        Dim result As Integer = MessageBox.Show("You've lost! Play again?", "Defeat!", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Me.Close()
        ElseIf result = DialogResult.Yes Then
            StartNewGame()
        End If
    End Sub

    Private Sub ComputerTurn()
        'Check for corners, the center, or sides to block player
        If btn1.Text = String.Empty Then
            btn1.Text = "X"
        ElseIf btn3.Text = String.Empty Then
            btn3.Text = "X"
        ElseIf btn7.Text = String.Empty Then
            btn7.Text = "X"
        ElseIf btn9.Text = String.Empty Then
            btn9.Text = "X"
        ElseIf btn5.Text = String.Empty Then
            btn5.Text = "X"
        ElseIf btn2.Text = String.Empty Then
            btn2.Text = "X"
        ElseIf btn4.Text = String.Empty Then
            btn4.Text = "X"
        ElseIf btn6.Text = String.Empty Then
            btn6.Text = "X"
        ElseIf btn8.Text = String.Empty Then
            btn8.Text = "X"
        End If

        ' Check to see if the computer won
        If CheckForWin("X") Then
            ComputerWins()
        End If
    End Sub

    ' Respond when btn1 is clicked clicked
    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        If btn1.Text = String.Empty Then
            btn1.Text = "O"
        End If

        If CheckForWin("O") Then
            PlayerWins()
        Else
            ComputerTurn()
        End If
    End Sub

    ' Respond when btn2 is clicked clicked
    Private Sub btn2_Click(sender As Object, e As EventArgs) Handles btn2.Click
        If btn2.Text = String.Empty Then
            btn2.Text = "O"
        End If

        If CheckForWin("O") Then
            PlayerWins()
        Else
            ComputerTurn()
        End If
    End Sub

    ' Respond when btn3 is clicked clicked
    Private Sub btn3_Click(sender As Object, e As EventArgs) Handles btn3.Click
        If btn3.Text = String.Empty Then
            btn3.Text = "O"
        End If

        If CheckForWin("O") Then
            PlayerWins()
        Else
            ComputerTurn()
        End If
    End Sub

    ' Respond when btn4 is clicked clicked
    Private Sub btn4_Click(sender As Object, e As EventArgs) Handles btn4.Click
        If btn4.Text = String.Empty Then
            btn4.Text = "O"
        End If

        If CheckForWin("O") Then
            PlayerWins()
        Else
            ComputerTurn()
        End If
    End Sub

    ' Respond when btn5 is clicked clicked
    Private Sub btn5_Click(sender As Object, e As EventArgs) Handles btn5.Click
        If btn5.Text = String.Empty Then
            btn5.Text = "O"
        End If

        If CheckForWin("O") Then
            PlayerWins()
        Else
            ComputerTurn()
        End If
    End Sub

    ' Respond when btn6 is clicked clicked
    Private Sub btn6_Click(sender As Object, e As EventArgs) Handles btn6.Click
        If btn6.Text = String.Empty Then
            btn6.Text = "O"
        End If

        If CheckForWin("O") Then
            PlayerWins()
        Else
            ComputerTurn()
        End If
    End Sub

    ' Respond when btn7 is clicked clicked
    Private Sub btn7_Click(sender As Object, e As EventArgs) Handles btn7.Click
        If btn7.Text = String.Empty Then
            btn7.Text = "O"
        End If

        If CheckForWin("O") Then
            PlayerWins()
        Else
            ComputerTurn()
        End If
    End Sub

    ' Respond when btn8 is clicked clicked
    Private Sub btn8_Click(sender As Object, e As EventArgs) Handles btn8.Click
        If btn8.Text = String.Empty Then
            btn8.Text = "O"
        End If

        If CheckForWin("O") Then
            PlayerWins()
        Else
            ComputerTurn()
        End If
    End Sub

    ' Respond when btn9 is clicked clicked
    Private Sub btn9_Click(sender As Object, e As EventArgs) Handles btn9.Click
        If btn9.Text = String.Empty Then
            btn9.Text = "O"
        End If

        If CheckForWin("O") Then
            PlayerWins()
        Else
            ComputerTurn()
        End If
    End Sub
End Class
