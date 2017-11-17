<HTML>
<HEAD>
<title> Procesa carga de PAQ</title>

</HEAD>

<body bgcolor="#FF00FF" onload="maximizar()">

<%

if request.form("OEP") <> "" then

Const adOpenDynamic = 2
Const adLockOptimistic = 3
set conectarOEP = Server.CreateObject("ADODB.Connection")
conectarOEP.Open "DRIVER={Microsoft Access Driver (*.mdb)};Database Locking Mode=0;DBQ=C:\Inetpub\wwwroot\balanceSFE\CIimposicion.mdb"

set rsOEP = server.CreateObject ("ADODB.recordset")
rsOEP.CursorType= AdOpenDynamic
rsOEP.LockType= adLockOptimistic

set rsCONTROL = server.CreateObject ("ADODB.recordset")
rsCONTROL.CursorType= AdOpenDynamic
rsCONTROL.LockType= adLockOptimistic

control= cstr(request.form("OEP"))
sqlVACIO= "select * from lectPAQ where PAQ_oep ='"& control &"'"

rsCONTROL.open sqlVACIO,conectarOEP

if rsCONTROL.EOF and rsCONTROL.BOF then

sqlOEP= "select * from lectPAQ"

rsOEP.open sqlOEP,conectarOEP

rsOEP.addnew
rsOEP("PAQ_oep")= request.form("OEP")
rsOEP.update
rsOEP.close

end if

rsCONTROL.close
conectarOEP.close


end if

response.redirect("cargaPAQ.asp?target=_self")

%>

<SCRIPT Language="javascript" type="text/javascript">

function maximizar() {

window.moveTo(0,0);

window.resizeTo(screen.width,screen.height);
}
</SCRIPT>

</body>
</HTML>


