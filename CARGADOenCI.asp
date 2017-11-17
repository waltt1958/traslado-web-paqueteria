<HTML>

<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title> PAQUETES A PASAR A PAQUETERIA</title>
</HEAD>

<body bgcolor="WHITE" onload="maximizar()">
<H3 align= "right">Hoy es: <%=date%></H3>
<center><h1><p align="center"><u><b><font size="12"> PAQUETES CARGADOS EN EL CI ROSARIO</font size></b></u></p> </h1></center>

<br>
<hr size= 6 color="black"></hr>
<br>
<br>
 
<TABLE WIDTH=100%" BORDER="0" BORDERCOLOR="BLACK">

<TR ALIGN="CENTER">
<TD><B>NRO PAQUETE</B></TD>
<TD><B> FECHA DE CARGA</B></TD>
</TR>

<%

Const adOpenDynamic = 2
Const adLockOptimistic = 3
set conectarOEP = Server.CreateObject("ADODB.Connection")
conectarOEP.Open "DRIVER={Microsoft Access Driver (*.mdb)};Database Locking Mode=0;DBQ=C:\Inetpub\wwwroot\balanceSFE\CIimposicion.mdb"

set rsOEP = server.CreateObject ("ADODB.recordset")
'rsOEP.CursorType= AdOpenDynamic
'rsOEP.LockType= adLockOptimistic

sqlCI="select * from imposicionCI where CERRADO_CI ='NO' order by FECHA_CARGA_CI asc"

rsOEP.open sqlCI, conectarOEP, adOpenDynamic, adLockOptimistic

if not (rsOEP.eof and rsOEP.bof) then
 
rsOEP.movefirst

end if

do while not rsOEP.EOF

response.write ("<TR><TD ALIGN=CENTER>" & rsOEP("CI_oep") & "</TD><TD ALIGN=CENTER>" & rsOEP("FECHA_CARGA_CI") & "</TD></TR>")


rsOEP.movenext
loop

rsOEP.close
conectarOEP.close


%>
</TABLE>
</body>



<SCRIPT Language="javascript" type="text/javascript">
function maximizar() {
window.moveTo(0,0);
window.resizeTo(screen.width,screen.height);
}
</SCRIPT>
</HTML>