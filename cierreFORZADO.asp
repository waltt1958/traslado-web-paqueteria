<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
<title>CIERRE FORZADO</title>
</HEAD>
<body onload="maximizar()">
<H3 align= "right">Hoy es: <%=date%></H3>

<br>
<FORM >
<%
Const adOpenDynamic = 2
Const adLockOptimistic = 3
set conectarOEP = Server.CreateObject("ADODB.Connection")
conectarOEP.Open "DRIVER={Microsoft Access Driver (*.mdb)};Database Locking Mode=0;DBQ=C:\Inetpub\wwwroot\balanceSFE\CIimposicion.mdb"

set rsCI = server.CreateObject ("ADODB.recordset")
rsCI.CursorType= AdOpenDynamic
rsCI.LockType= adLockOptimistic

set rsPAQ = server.CreateObject ("ADODB.recordset")
rsPAQ.CursorType= AdOpenDynamic
rsPAQ.LockType= adLockOptimistic

sqlPAQ= "select * from lectPAQ where CERRADO_oep='FORZADO'"
sqlCI= "select * from imposicionCI Where CERRADO_ci='FORZADO'"

rsPAQ.open sqlPAQ,conectarOEP

rsCI.open sqlCI,conectarOEP


%>
<br>
<TABLE WIDTH= "100%" border="3" bordercolor="black">

<TR><CENTER><P ALIGN="CENTER"><FONT SIZE="6"><b><U> PAQUETES CON CIERRE FORZADO POR NO SER LEIDOS EN PAQUETERIA Y SI FUERON LEIDOS EN EL CI HASTA EL DIA DE AYER</U></FONT SIZE></b></P></CENTER></TR>
<TR>


<%
if rsCI.bof and rsCI.eof then

response.write ("<TR><TD ALIGN=CENTER WIDTH=100><B><FONT SIZE= 4>" & ("NO SE FORZO EL CIERRE DE NINGUN PAQUETE LEIDO EN EL CI Y NO LEIDO EN PAQUETERIA") & "</FONT SIZE></B></TD></TR>")

else
%>
<TD WIDTH="33%"><CENTER><B>NRO OEP</B></CENTER></TD>
<TD WIDTH="33%"><CENTER><B>FECHA CARGA</B></CENTER></TD>
<TD WIDTH="33%"><CENTER><B>FECHA CIERRE FORZADO</B></CENTER></TD> 
</TR>
<BR>

<% 


rsCI.movefirst


do while not (rsCI.eof)

response.write ("<TR><TD ALIGN=CENTER>" & rsCI("CI_oep") & "</TD><TD ALIGN=CENTER>" & rsCI("FECHA_CARGA_ci") & "</TD><TD ALIGN=CENTER>" & rsCI("FECHA_CIERRE_ci") & "</TD></TR>")

rsCI.movenext
loop


end if

rsCI.close




%>

</TABLE>

<BR>
<BR>

<TABLE WIDTH= "100%" border="3" bordercolor="black">
<TR><CENTER><P ALIGN="CENTER"><FONT SIZE="6"><B><U> PAQUETES CON CIERRE FORZADO POR NO SER LEIDOS EN EL CI Y SI FUERON LEIDOS EN PAQUETERIA HASTA EL DIA DE AYER</U></B></FONT SIZE></P></CENTER></TR>
<TR>

<%
if rsPAQ.bof and rsPAQ.eof then

response.write ("<TR><TD ALIGN=CENTER WIDTH=100><B><FONT SIZE= 4>" & ("NO SE FORZO EL CIERRE DE NINGUN PAQUETE LEIDO EN PAQUETERIA Y NO LEIDO EN EL CI") & "</FONT SIZE></B></TD></TR>")

else
%>
<TD WIDTH="33%"><CENTER><B>NRO OEP</B></CENTER></TD>
<TD WIDTH="33%"><CENTER><B>FECHA CARGA</B></CENTER></TD>
<TD WIDTH="33%"><CENTER><B>FECHA CIERRE FORZADO</B></CENTER></TD> 
</TR>
<BR>

<%



rsPAQ.movefirst

do while not (rsPAQ.eof)

response.write ("<TR><TD ALIGN=CENTER>" & rsPAQ("PAQ_oep") & "</TD><TD ALIGN=CENTER>" & rsPAQ("FECHA_CARGA_oep") & "</TD><TD ALIGN=CENTER>" & rsPAQ("FECHA_CIERRE_oep") & "</TD></TR>")

rsPAQ.movenext
loop

end if

rsPAQ.close
conectarOEP.close
%>

</TABLE>
</FORM>

</body>

<SCRIPT Language="javascript" type="text/javascript">
function maximizar() {
window.document.inicia.OEP.focus();
window.moveTo(0,0);
window.resizeTo(screen.width,screen.height);
}
</SCRIPT>
</HTML>