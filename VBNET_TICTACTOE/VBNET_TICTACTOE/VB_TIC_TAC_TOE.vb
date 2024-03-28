Public Class VB_TIC_TAC_TOE

    Private Sub VB_TIC_TAC_TOE_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' add event to all buttons inside the panel2
        For Each c As Control In Panel2.Controls

            If c.GetType() = GetType(Button) Then

                AddHandler c.Click, AddressOf btn_Click

            End If

        Next

    End Sub

    Dim XorO As Integer = 0
    ' create button event
    Private Sub btn_Click(sender As Object, e As EventArgs)

        Dim btn As Button = sender

        ' we will clear buttons text later
        If btn.Text.Equals("") Then

            If XorO Mod 2 = 0 Then

                btn.Text = "X"
                btn.ForeColor = Color.Red
                Label1.Text = "[O] Turn"
                getTheWinner()

            Else
                btn.Text = "O"
                btn.ForeColor = Color.Blue
                Label1.Text = "[X] Turn"
                getTheWinner()
            End If

            XorO += 1

        End If

    End Sub

    Dim win As Boolean = False

    ' create a function to get the winner
    Private Sub getTheWinner()

        If Not Button1.Text.Equals("") AndAlso Button1.Text.Equals(Button2.Text) AndAlso Button1.Text.Equals(Button3.Text) Then
            win = True
            winEffect(Button1, Button2, Button3)
        End If

        If Not Button4.Text.Equals("") AndAlso Button4.Text.Equals(Button5.Text) AndAlso Button4.Text.Equals(Button6.Text) Then
            win = True
            winEffect(Button4, Button5, Button6)
        End If

        If Not Button7.Text.Equals("") AndAlso Button7.Text.Equals(Button8.Text) AndAlso Button7.Text.Equals(Button9.Text) Then
            win = True
            winEffect(Button7, Button8, Button9)
        End If

        If Not Button1.Text.Equals("") AndAlso Button1.Text.Equals(Button4.Text) AndAlso Button1.Text.Equals(Button7.Text) Then
            win = True
            winEffect(Button1, Button4, Button7)
        End If

        If Not Button2.Text.Equals("") AndAlso Button2.Text.Equals(Button5.Text) AndAlso Button2.Text.Equals(Button8.Text) Then
            win = True
            winEffect(Button2, Button5, Button8)
        End If

        If Not Button3.Text.Equals("") AndAlso Button3.Text.Equals(Button6.Text) AndAlso Button3.Text.Equals(Button9.Text) Then
            win = True
            winEffect(Button3, Button6, Button9)
        End If

        If Not Button1.Text.Equals("") AndAlso Button1.Text.Equals(Button5.Text) AndAlso Button1.Text.Equals(Button9.Text) Then
            win = True
            winEffect(Button1, Button5, Button9)
        End If

        If Not Button3.Text.Equals("") AndAlso Button3.Text.Equals(Button5.Text) AndAlso Button3.Text.Equals(Button7.Text) Then
            win = True
            winEffect(Button3, Button5, Button7)
        End If

        ' if no one win later
        ' 9 buttons with X or O mean 9 char = no button is empty
        If allbuttonsTextLength() = 9 AndAlso win = False Then

            Label1.Text = "NO Winner"

        End If

    End Sub

    Function allbuttonsTextLength() As Integer

        Dim btnsTxtLength As Integer = 0

        For Each c As Control In Panel2.Controls

            If c.GetType() = GetType(Button) Then

                btnsTxtLength += c.Text.Length

            End If

        Next

        Return btnsTxtLength

    End Function

    Private Sub winEffect(ByVal b1 As Button, ByVal b2 As Button, ByVal b3 As Button)

        b1.BackColor = Color.Red
        b2.BackColor = Color.Red
        b3.BackColor = Color.Red

        b1.ForeColor = Color.White
        b2.ForeColor = Color.White
        b3.ForeColor = Color.White

        Label1.Text = b1.Text + " Win"

    End Sub

    Private Sub ButtonNewPartie_Click(sender As Object, e As EventArgs) Handles ButtonNewPartie.Click

        XorO = 0
        win = False
        Label1.Text = "Play"
        For Each c As Control In Panel2.Controls

            If c.GetType() = GetType(Button) Then

                c.BackColor = Color.White
                c.Text = ""

            End If

        Next

    End Sub
End Class