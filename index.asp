<!DOCTYPE html>
<!--#include virtual="common/DBHelper.asp"-->
<!--#include virtual="common/json_2.0.4.asp"-->
<html lang="ko">
	<meta charset= "utf-8">
	<head>
		<style>
		</style>
		<link href="css/pagination.css" rel="stylesheet" type="text/css">
	</head>
	<body>
				 <div id="data-container"></div>
				
				<div id="pagination-container" style="position: absolute; left: 0px; top: 250px;"></div>
	</body>
	<script src="http://code.jquery.com/jquery-1.8.2.min.js"></script>
	<script type="text/javascript" src="js/pagination.min.js" ></script>
	<script type="text/javascript">
		
		$(document).ready(function(){				
	
			$.ajax({ 
				url: 'test.asp' , 
				dataType : 'json' , 
				success: function(data) {
					 	$('#pagination-container').pagination({
		    			dataSource: data
		    			,callback: function(data, pagination) {
			        	// template method of yourself
			       		var html = simpleTemplating(data);
			        	$('#data-container').html(html);
		    			}
						})
				} 
			});

				
		});
		
		function simpleTemplating(data) {
			
	    var html;
	    $.each(data, function(index, item){
					html +='<tr>';
	        html += '<th>'+ item.L_LAWS +'</th>';
	        html += '</tr>';
	    });
    	return html;
		}
	</script>
</html>
