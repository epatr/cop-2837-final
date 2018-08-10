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
    End Sub

    ' Boolean check to see if the move wins the game
    Private Function CheckForWin() As Boolean
        ' Check bottom row
        If btn1.Text = "O" Then
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

    ' Respond when btn1 is clicked clicked
    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        If btn1.Text = String.Empty Then
            btn1.Text = "O"
        End If

        If CheckForWin() Then
            PlayerWins()
        End If
    End Sub
End Class
