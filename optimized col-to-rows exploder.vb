sub optimized_exploder ()
dim r as range
dim i as long
dim ar
application.screenupdating=false
set r = activesheet.range("p999999").end(xlup)
do while r.row > 1
    if r.row mod 100 = 0 then debug.print(r.row)
    if instr(r.value,",")<1 then
    ' do nothing
    else
        ar =split(r.value),",")
        if ubound(ar)>=0 then r.value = ar(0)
        for i = ubound(ar) to 1 step -1
            r.entirerow.copy
            r.offset(1).entirerow.insert
            r.offset(1).value = trim(ar(i))
        next
    end if
    set r= r.offset(-1)
loop
application.cutcopymode=false
end sub