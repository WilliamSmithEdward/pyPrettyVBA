VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
option explicit
private mValue as long
public property get value() as long
Attribute value.VB_UserMemId = 0
value=mvalue
end property
public sub increment()
mvalue=mvalue+1
end sub
