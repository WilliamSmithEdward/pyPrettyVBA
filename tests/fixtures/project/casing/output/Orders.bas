Attribute VB_Name = "Orders"
Option Explicit

Public Sub Process()
    Dim basket As ShoppingBasket
    Set basket = New ShoppingBasket
    If basket.Count > MaxItems Then LogMessage "too many"
End Sub
