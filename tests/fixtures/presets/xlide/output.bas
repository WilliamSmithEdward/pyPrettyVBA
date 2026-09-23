Attribute VB_Name = "Invoices"
Option Explicit
Private Const TAX_RATE As Double = .20
Private mLastTotal    As Currency

Public Function InvoiceTotal(ByVal subtotal As Currency, Optional ByVal discount As Double = 0.0) As Currency
    Dim total As Currency
    If discount < 0 Or discount > 1 Then Err.Raise 5, "InvoiceTotal", "bad discount"
    total = subtotal * (1 - discount)
    total = total + total * TAX_RATE   ' tax on the discounted price
    mLastTotal = total
    InvoiceTotal = total
End Function
Sub PrintInvoice()
    Dim t As Currency: t = InvoiceTotal(100@, .1)
    Let t = t
    Debug.Print "Total: "; Format$(t, "0.00")
    ActiveSheet.range("B2").value = t
End Sub
