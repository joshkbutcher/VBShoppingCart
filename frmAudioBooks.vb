Public Class frmAudioBooks
    Private Sub frmAudioBooks_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Add the book titles to the list box
        lstAudioBooks.Items.Add(bookName(4))
        lstAudioBooks.Items.Add(bookName(5))
        lstAudioBooks.Items.Add(bookName(6))
        lstAudioBooks.Items.Add(bookName(7))
    End Sub

    Private Sub btnAddAudio_Click(sender As Object, e As EventArgs) Handles btnAddAudio.Click
        'Call AddBookToCart to add book to global array
        AddBookToCart(lstAudioBooks.SelectedIndex + 4)
    End Sub

    'Add book to global array so it can be found in the shopping cart form
    Private Sub AddBookToCart(bookInt As Integer)
        If bookInt >= 0 Then
            selectedBooks(bookCount) = bookInt
            bookCount += 1
        End If

    End Sub

    'Close form
    Private Sub btnCloseAudio_Click(sender As Object, e As EventArgs) Handles btnCloseAudio.Click
        Close()
    End Sub
End Class