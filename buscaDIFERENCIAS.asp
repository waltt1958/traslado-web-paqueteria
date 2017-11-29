<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
<title> VER DIFERENCIAS</title>
</HEAD>
<body bgcolor="WHITE" onload="maximizar()">
<H3 align= "right">Hoy es: <%=date%></H3>

<br>
<FORM >
<%
Const adOpenDynamic = 2
Const adLockOptimistic = 3
set conectarOEP = Server.CreateObject("ADODB.Connection")
conectarOEP.Open "DRIVER={Microsoft Access Driver (*.mdb)};Database Locking Mode=0;DBQ=C:\Inetpub\wwwroot\balanceSFE\CIimposicion.mdb"

set rsNOleyoCI = server.CreateObject ("ADODB.recordset")
rsNOleyoCI.CursorType= AdOpenDynamic
rsNOleyoCI.LockType= adLockOptimistic

set rsNOleyoPAQ = server.CreateObject ("ADODB.recordset")
rsNOleyoPAQ.CursorType= AdOpenDynamic
rsNOleyoPAQ.LockType= adLockOptimistic

sqlPAQ= "select * from NOleyoPAQ"
sqlCI= "select * from NOleyoCI"

rsNOleyoPAQ.open sqlPAQ,conectarOEP

rsNOleyoCI.open sqlCI,conectarOEP


%>
<BR>
<TABLE WIDTH= "100%" border="3" bordercolor="black">

<TR><CENTER><P ALIGN="CENTER"><FONT SIZE="6"><b><U> PAQUETES LEIDOS EN PAQUETERIA Y NO LEIDOS EN EL CI</U></FONT SIZE></b></P></CENTER></TR>
<TR>


<%
if rsNOleyoCI.bof and rsNOleyoCI.eof then

response.write ("<TR><TD ALIGN=CENTER WIDTH=100><B><FONT SIZE= 4>" & ("PAQUETERIA NO LEYO NINGUN PAQUETE QUE NO FUERA LEIDO POR EL CI") & "</FONT SIZE></B></TD></TR>")

else
%>
<TD WIDTH="50%"><CENTER><B>NRO OEP</B></CENTER></TD>
<TD WIDTH="50%"><CENTER><B>FECHA CARGA</B></CENTER></TD>
</TR>
<BR>

<% 


rsNOleyoCI.movefirst


do while not (rsNOleyoCI.eof)

response.write ("<TR><TD ALIGN=CENTER>" & rsNOleyoCI("PAQ_oep") & "</TD><TD ALIGN=CENTER>" & rsNOleyoCI("FECHA_CARGA_oep") & "</TD></TR>")

rsNOleyoCI.movenext
loop


end if

rsNOleyoCI.close




%>

</TABLE>

<BR>
<BR>

<TABLE WIDTH= "100%" border="3" bordercolor="black">
<TR><CENTER><P ALIGN="CENTER"><FONT SIZE="6"><B><U> PAQUETES LEIDOS EN EL CI Y NO LEIDOS EN PAQUETERIA</U></B></FONT SIZE></P></CENTER></TR>
<TR>

<%
if rsNOleyoPAQ.bof and rsNOleyoPAQ.eof then

response.write ("<TR><TD ALIGN=CENTER WIDTH=100><B><FONT SIZE= 4>" & ("EL CI LEYO TODOS LOS PAQUETES QUE VINIERON A PAQUETERIA Y FUERON LEIDOS POR PAQUETERIA") & "</FONT SIZE></B></TD></TR>")

else
%>
<TD WIDTH="50%"><CENTER><B>NRO OEP</B></CENTER></TD>
<TD WIDTH="50%"><CENTER><B>FECHA CARGA</B></CENTER></TD>
</TR>
<BR>

<%



rsNOleyoPAQ.movefirst

do while not (rsNOleyoPAQ.eof)

response.write ("<TR><TD ALIGN=CENTER>" & rsNOleyoPAQ("CI_oep") & "</TD><TD ALIGN=CENTER>" & rsNOleyoPAQ("FECHA_CARGA_ci") & "</TD></TR>")

rsNOleyoPAQ.movenext
loop

end if

conectarOEP.execute ("UPDATE imposicionCI INNER JOIN lectPAQ ON imposicionCI.CI_oep = lectPAQ.PAQ_oep SET imposicionCI.CERRADO_CI = 'SI', imposicionCI.FECHA_CIERRE_CI = Now(), lectPAQ.CERRADO_OEP = 'SI',lectPAQ.FECHA_CIERRE_OEP = Now()")

rsNOleyoPAQ.close
conectarOEP.close
%>

</TABLE>
</FORM>

</body>
<SCRIPT Language="javascript" type="text/javascript">
function validar() {
if (window.document.inicia.OEP.value=="") {
alert("Tiene que colocar el turno OEP previo a llamar");
window.document.inicia.OEP.focus();
return;
}
if (isNaN(window.document.inicia.OEP.value)) {
alert("Tiene que colocar solo numeros en turno OEP");
window.document.inicia.OEP.focus();
return;
}
window.document.inicia.submit();
}
</script>

<SCRIPT Language="javascript" type="text/javascript">
function maximizar() {
window.document.inicia.OEP.focus();
window.moveTo(0,0);
window.resizeTo(screen.width,screen.height);
}
</SCRIPT>
</HTML>