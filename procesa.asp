<HTML>
<HEAD>
<title> Procesa carga de CI</title>
<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">

</HEAD>

<body onload="maximizar()">

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
sqlVACIO= "select * from imposicionCI where CI_oep ='"& control &"'"

rsCONTROL.open sqlVACIO,conectarOEP

if rsCONTROL.EOF and rsCONTROL.BOF then

sqlOEP= "select * from imposicionCI"

rsOEP.open sqlOEP,conectarOEP

rsOEP.addnew
rsOEP("CI_oep")= request.form("OEP")
rsOEP.update
rsOEP.close


end if

rsCONTROL.close
conectarOEP.close


end if

response.redirect("cargaCI.asp?target=_self")

%>

<SCRIPT Language="javascript" type="text/javascript">

function maximizar() {

window.moveTo(0,0);

window.resizeTo(screen.width,screen.height);
}
</SCRIPT>

</body>
</HTML>


