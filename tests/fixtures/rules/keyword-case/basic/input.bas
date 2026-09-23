option explicit
option compare text
deflng a-c, x-z

private declare ptrsafe function tickcount lib "kernel32" alias "GetTickCount" () as long

public sub demo(byval count as long)
    dim i as long
    for i = 1 to count step 2
        if i mod 3 = 0 then
            debug.print "fizz"
        elseif i > 10 then
            exit for
        endif
    next i
    on error goto handler
    workbook.close savechanges:=false
    exit sub
handler:
    rem a remark
    resume next
end sub
