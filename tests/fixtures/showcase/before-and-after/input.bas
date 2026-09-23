Attribute VB_Name = "Inventory"
option explicit
'  Stock levels for the warehouse sheet.
private const LOW_STOCK=10
private mItems as collection
public function RestockList(ws as worksheet) as collection
dim r as long,qty as long,result as new collection
on error goto fail
for r=2 to ws.cells(ws.rows.count,1).end(xlup).row
qty=ws.cells(r,3).value
if qty<low_stock then
result.add ws.cells(r,1).value
elseif qty=0 then result.add ws.cells(r,1).value & " (out)"
end if
next r
set restocklist=result
exit function
fail:
    msgbox "Could not read row "&r &": "&err.description,vbexclamation
end function
