<HTML>

<HEAD>
<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title> PAQUETES RECIBIDOS EN PAQUETERIA</title>
</HEAD>

<body >
<H3 align= "right">Hoy es: <%=date%></H3>
<br>
<center><h1><p align="center"><u><b><font size="12"> PAQUETES CARGADOS EN EL SECTOR PAQUETERIA</font size></b></u></p> </h1></center>

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
rsOEP.CursorType= AdOpenDynamic
rsOEP.LockType= adLockOptimistic

sqlPAQ="select * from lectPAQ WHERE CERRADO_OEP ='NO' order by FECHA_CARGA_oep asc"

rsOEP.open sqlPAQ,conectarOEP

if not (rsOEP.eof and rsOEP.bof) then

rsOEP.movefirst

end if

do while not rsOEP.EOF

response.write ("<TR><TD ALIGN=CENTER>" & rsOEP("PAQ_oep") & "</TD><TD ALIGN=CENTER>" & rsOEP("FECHA_CARGA_oep") & "</TD></TR>")


rsOEP.movenext
loop

rsOEP.close
conectarOEP.close


%>
</TABLE>
<BR>
<BR>
<table width="100%">
<td align="center">
<a href="index.asp" target="_self"><input type="button" name="CARGADOenPAQ" value="INICIO" style="FONT-SIZE: 20pt; border: 5px solid; [b]FONT-FAMILY: Verdana, boldt[/b];
BACKGROUND-COLOR: #C0C0C0"></a>
</td>
</table>
</body>


</HTML>