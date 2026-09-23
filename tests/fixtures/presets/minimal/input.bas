Attribute VB_Name = "Invoices"
option explicit
private const TAX_RATE as double = .20
private mLastTotal    as currency

public function InvoiceTotal(byval subtotal as currency, optional byval discount as double = 0.0) as currency
dim total as currency
if discount<0 or discount>1 then err.raise 5,"InvoiceTotal","bad discount"
total=subtotal*(1-discount)
total=total+total*tax_rate   ' tax on the discounted price
mlasttotal=total
invoicetotal=total
end function
sub PrintInvoice()
dim t as currency: t=invoicetotal(100@,.1)
let t = t
debug.print "Total: ";format$(t,"0.00")
activesheet.range("B2").value=t
end sub
