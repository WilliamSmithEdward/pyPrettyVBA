Attribute VB_Name = "Orders"
Option Explicit
Public Sub Process()
    Dim basket As shoppingbasket
    Set basket = New SHOPPINGBASKET
    If basket.Count > MAXITEMS Then logmessage "too many"
End Sub
