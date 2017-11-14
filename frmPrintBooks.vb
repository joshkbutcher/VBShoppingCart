Public Class frmPrintBooks
    Private Sub frmPrintBooks_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Add book titles to list box
        lstPrintBooks.Items.Add(bookName(0))
        lstPrintBooks.Items.Add(bookName(1))
        lstPrintBooks.Items.Add(bookName(2))
        lstPrintBooks.Items.Add(bookName(3))
    End Sub

    Private Sub btnAddPrint_Click(sender As Object, e As EventArgs) Handles btnAddPrint.Click
        'Call AddBookToCart to add book to global array
        AddBookToCart(lstPrintBooks.SelectedIndex)
    End Sub

    'Add book to global array so it can be found in the shopping cart form
    Private Sub AddBookToCart(bookInt As Integer)
        If bookInt >= 0 Then
            selectedBooks(bookCount) = bookInt
            bookCount += 1
        End If

    End Sub

    'Close form
    Private Sub btnClosePrint_Click(sender As Object, e As EventArgs) Handles btnClosePrint.Click
        Close()
    End Sub
End Class