(function(){tinymce.create('tinymce.plugins.ASPImagePlugin',{init:function(ed,url){ed.addCommand('mceASPImage',function(){var e=ed.selection.getNode();if(ed.dom.getAttrib(e,'class').indexOf('mceItem')!=-1)
return;ed.windowManager.open({file:url+'/image.asp',width:888+ed.getLang('aspimage.delta_width',0),height:400+ed.getLang('aspimage.delta_height',0),inline:1},{plugin_url:url});});ed.addButton('image',{title:'aspimage.image_desc',cmd:'mceASPImage'});},getInfo:function(){return{longname:'ASP image',author:'Erik Oosterwaal',authorurl:'http://www.precompiled.com',infourl:'http://www.precompiled.com',version:tinymce.majorVersion+"."+tinymce.minorVersion};}});tinymce.PluginManager.add('aspimage',tinymce.plugins.ASPImagePlugin);})();