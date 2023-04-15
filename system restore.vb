' Script Name 		: system_restore.vb
' Description 		: Create a system restore point, it is dated at today's date

localdate=date()
set SRP = getobject("winmgmts:\\.\root\default:Systemrestore")
CSRP = SRP.createrestorepoint ("System Start-up - " & localdate, 0, 100)
