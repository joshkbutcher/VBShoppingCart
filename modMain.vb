Module modMain

    'Declare constants
    Public Const salesTax As Decimal = 0.06D
    Public Const shippingCharge As Decimal = 2D
    Public Const printOnePrice As Decimal = 11.95D
    Public Const printTwoPrice As Decimal = 14.5D
    Public Const printThrPrice As Decimal = 29.95D
    Public Const printFouPrice As Decimal = 18.5D
    Public Const audioOnePrice As Decimal = 29.95D
    Public Const audioTwoPrice As Decimal = 14.5D
    Public Const audioThrPrice As Decimal = 12.95D
    Public Const audioFouPrice As Decimal = 11.5D

    'Declare book titles
    Public bookName() As String = {"I Did It your Way (Print)", "The History of Scotland (Print)", "Learn Calculus in One Day (Print)",
        "Feel the Stress (Print)", "Learn Calculus In One Day (Audio)", "The History Of Scotland (Audio)",
        "The Science Of Body Language (Audio)", "Relaxation Techniques (Audio)"}

    'Array for books selected from the print form and the audio form
    Public selectedBooks(99) As Integer

    'Keep count of total books selected
    Public bookCount As Integer

    'Keep up with subtotal for all books
    Public subTotal As Decimal

    'Declare shopping cart form
    Dim frm As New frmShoppingCart()

    Sub Main()

        'Assign '-1' to entire array so it can be changed in print and audio forms to keep track of selected books
        For i = 0 To selectedBooks.Length - 1
            selectedBooks(i) = -1
        Next i

        'Show shopping cart form
        frm.ShowDialog()

    End Sub

End Module
