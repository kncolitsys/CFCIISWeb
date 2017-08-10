<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
<cfset iisObj = createObject("component", "com.iisweb").init() />
<cfdump var="#variables.iisObj.start('temp1.knivis.net')#">

</body>
</html>