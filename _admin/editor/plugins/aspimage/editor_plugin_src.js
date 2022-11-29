/**
 * $Id: editor_plugin_src.js 324 2007-11-01 12:58:49Z spocke $
 *
 * @author Moxiecode
 * @copyright Copyright © 2004-2007, Moxiecode Systems AB, All rights reserved.
 */

(function() {
	tinymce.create('tinymce.plugins.ASPImagePlugin', {
		init : function(ed, url) {
			// Register commands
			ed.addCommand('mceASPImage', function() {
				var e = ed.selection.getNode();

				// Internal image object like a flash placeholder
				if (ed.dom.getAttrib(e, 'class').indexOf('mceItem') != -1)
					return;

				ed.windowManager.open({
					file : url + '/image.asp',
					width : 888 + ed.getLang('aspimage.delta_width', 0),
					height : 400 + ed.getLang('aspimage.delta_height', 0),
					inline : 1
				}, {
					plugin_url : url
				});
			});

			// Register buttons
			ed.addButton('image', { 
				title : 'aspimage.image_desc',
				cmd : 'mceASPImage'
			});
		},

		getInfo : function() {
			return {
				longname : 'ASP image',
				author : 'Erik Oosterwaal',
				authorurl : 'http://www.precompiled.com',
				infourl : 'http://www.precompiled.com',
				version : tinymce.majorVersion + "." + tinymce.minorVersion
			};
		}
	});

	// Register plugin
	tinymce.PluginManager.add('aspimage', tinymce.plugins.ASPImagePlugin);
})();