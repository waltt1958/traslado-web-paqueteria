<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
<title>CIERRE HASTA AYER</title>
</HEAD>
<BR>
<BR>
<H3 align= "right">Hoy es: <%=date%></H3>
<body>

<%

Const adOpenDynamic = 2
Const adLockOptimistic = 3
set conectarOEP = Server.CreateObject("ADODB.Connection")
conectarOEP.Open "DRIVER={Microsoft Access Driver (*.mdb)};Database Locking Mode=0;DBQ=C:\Inetpub\wwwroot\balanceSFE\CIimposicion.mdb"

conectarOEP.execute ("UPDATE imposicionCI SET imposicionCI.CERRADO_CI = 'FORZADO', imposicionCI.FECHA_CIERRE_CI = Now() WHERE imposicionCI.CERRADO_CI= 'NO' AND imposicionCI.FECHA_CARGA_CI < DATE()")

conectarOEP.execute ("UPDATE lectPAQ SET lectPAQ.FECHA_CIERRE_OEP = Now(), lectPAQ.CERRADO_OEP = 'FORZADO' WHERE lectPAQ.CERRADO_OEP='NO' AND lectPAQ.FECHA_CARGA_oep < date()")

conectarOEP.close


%>


<BR>
<BR>
<BR>
<BR>
<BR>
<BR>


<center><h1><p align="center"><b><font size="12"> TERMINO EL PROCESO DE CIERRE FORZADO DE LOS PAQUETES QUE NO FUERON LEIDOS EN EL CI O EN PAQUETERIA</font size></b></p> </h1></center>
<BR>
<BR>
<table width="100%">
<td align="center">
<a href="index.asp" target="_self"><input type="button" name="cierreFINAL" value="INICIO" style="FONT-SIZE: 20pt; border: 5px solid; [b]FONT-FAMILY: Verdana, boldt[/b];
BACKGROUND-COLOR: #C0C0C0"></a>
</td>
</table>


</body>

</SCRIPT>
</HTML>