<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>Insert title here</title>
</head>
<body>
		varfp = new FlexPaperViewer(     

        'FlexPaperViewer',  

        'viewerPlaceHolder', { config : {  

		SwfFile : "{Paper[*,0].swf,28}",  
		
		Scale : 0.6,  
		
		ZoomTime : 0.5,  
		
		ZoomInterval : 0.1,  
		
		FitPageOnLoad : false,  
		
		FitWidthOnLoad : false,  
		
		PrintEnabled : false,  
		
		MinZoomSize : 0.2,  
		
		MaxZoomSize : 5,  
		
		localeChain : "en_US"
</body>
</html>