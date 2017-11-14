Public Class frmShoppingCart
    Private Sub frmShoppingCart_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        'When the Shopping Cart is activated, update the list box and prices
        UpdateListBox()
        UpdatePrices()

    End Sub

    'Open the Print Books form
    Private Sub OpenPrintBooksForm(sender As Object, e As EventArgs) Handles PrintBooksToolStripMenuItem.Click, ctxtPrintBooks.Click
        Dim box = New frmPrintBooks()
        box.Show()
    End Sub


    'Open the Audio Books form
    Private Sub OpenAudioBooksForm(sender As Object, e As EventArgs) Handles AudioBooksToolStripMenuItem.Click, ctxtAudioBooks.Click
        Dim box = New frmAudioBooks()
        box.Show()
    End Sub

    'Open the About window
    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Dim box = New frmAbout()
        box.Show()
    End Sub

    'Remove an item from the list box
    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        Dim bookIndex As Integer = lstProducts.SelectedIndex()

        If bookIndex = -1 Then
            MsgBox("You need to select a book title to remove it.")
        Else
            'Remove item and update the prices
            lstProducts.Items.RemoveAt(bookIndex)
            UpdatePrices()

        End If
    End Sub

    'Update list box from other forms
    Private Sub UpdateListBox()
        Dim intCount As Integer = 0

        'If there are books to load, find the book title and add it to the list box
        Do While intCount < selectedBooks.Length()
            If selectedBooks(intCount) <> -1 Then
                lstProducts.Items.Add(bookName(selectedBooks(intCount)))
            End If
            intCount += 1
        Loop

        'Clear array after it is has populated list box
        For i = 0 To selectedBooks.Length - 1
            selectedBooks(i) = -1
        Next i
    End Sub

    'Update prices from books in shopping list
    Private Sub UpdatePrices()
        Dim intCount As Integer = 0
        subTotal = 0
        ClearTextBoxes()

        Do While intCount < lstProducts.Items.Count
            CalcTotal(CalcSubTotal(lstProducts.Items(intCount).ToString), CalcTax(), CalcShipping())
            intCount += 1
        Loop
    End Sub

    'Calculate subtotal
    Private Function CalcSubTotal(bookInt As String) As Decimal

        'Get price of book based on book title
        Select Case bookInt
            Case bookName(0)
                subTotal += printOnePrice
            Case bookName(1)
                subTotal += printTwoPrice
            Case bookName(2)
                subTotal += printThrPrice
            Case bookName(3)
                subTotal += printFouPrice
            Case bookName(4)
                subTotal += audioOnePrice
            Case bookName(5)
                subTotal += audioTwoPrice
            Case bookName(6)
                subTotal += audioThrPrice
            Case bookName(7)
                subTotal += audioFouPrice
        End Select

        txtSubTotal.Text = subTotal.ToString("C")
        Return subTotal
    End Function

    'Calculate tax
    Private Function CalcTax() As Decimal
        Dim taxCost As Decimal = subTotal * salesTax
        txtTax.Text = taxCost.ToString("C")
        Return taxCost
    End Function

    'Calculate shipping
    Private Function CalcShipping() As Decimal
        Dim shippingCost As Decimal = lstProducts.Items.Count * shippingCharge
        txtShipping.Text = shippingCost.ToString("C")
        Return shippingCost
    End Function

    'Calculate total
    Private Sub CalcTotal(subtotal As Decimal, tax As Decimal, shipping As Decimal)
        Dim total As Decimal = subtotal + tax + shipping
        txtTotal.Text = total.ToString("C")
    End Sub

    'Clear text boxes
    Private Sub ClearTextBoxes()
        txtSubTotal.Text = ""
        txtTax.Text = ""
        txtShipping.Text = ""
        txtTotal.Text = ""
    End Sub

    'Reset menu item clicked
    Private Sub ResetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResetToolStripMenuItem.Click
        Dim frm = New frmShoppingCart
        frm.ShowDialog()
        Me.Close()
    End Sub

    'Exit menu item clicked
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub


End Class