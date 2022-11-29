// JavaScript Document
  function launchTimelineModal(id) {
		var dialog = "#" + id;
		//alert(dialog);
		
		$(function() {
			$( dialog ).dialog({resizable: false, draggable: false, width: 420, modal: true, position: 'center'  });
			$(".ui-dialog-titlebar").hide();
		});
	
  }
  
  
  
  function CloseTimelineModal(id) {
	 	var dialog = "#" + id;
		
		$(function() {
				   
				   	$( dialog ).dialog("close");

		});
		
  }
  