// JavaScript Document

function savelist(destlist, savefield) {
	
	var dest = document.forms[0].elements[destlist];
	var sfield = document.forms[0].elements[savefield];
	
    var h = '*!';
    for (var i=0; i < dest.options.length; i++) {
        h += dest.options[i].value + '*!';
    }
	
	sfield.value = h;
	
}

function valueadd_item(availablelist, destlist) {
    var source = document.forms[0].elements[availablelist];
    var dest = document.forms[0].elements[destlist];

    if (false) {
        var selectedIndex = document.forms[0].property.options.selectedIndex;
        var propText = document.forms[0].property.options[selectedIndex].text;
        var propVal = document.forms[0].property.options[selectedIndex].value;
    }

    var h = ',';
    for (var i=0; i < dest.options.length; i++) {
        h += dest.options[i].value + ',';
    }

    var flag = 0;
    for (var i=0; i < source.options.length; i++) {
        if (source.options[i].selected && source.options[i].value != 'ADD') {
            var optionLookup = false ? ','+source.options[i].value+':' : ','+source.options[i].value+',';
            optionLookup += (false && !false) ? propVal+',' : '';
            if (h.indexOf(optionLookup) == -1) {
                var optionText = false ? source.options[i].text + ' [' + propText + ']' : source.options[i].text;
                var optionVal = false ? source.options[i].value + ':' + propVal : source.options[i].value;
                dest.options[dest.length] = new Option(optionText,optionVal);
            }
        }

        if (flag == 1 || source.options[i].selected) flag = 1;
    }

    
    if (flag == 0) alert('Please select an item to add');
}

function valueremove_item(availablelist, destlist) {
    var dest = document.forms[0].elements[destlist];
    var flag = 0;
    for (var i=dest.options.length-1; i >=0;  i--) {
        if (dest.options[i].selected) {
            dest.options[i] = null;
            flag = 1;
        }
    }
    
    if (flag == 0) alert('Please select an item to remove');
}

function valueremove_all(availablelist, destlist) {
	document.forms[0].elements[destlist].options.length = 0;
}

function valueadd_all(availablelist, destlist) {

    if (0==1) {
       valueremove_all(availablelist, destlist);
       return;
    }

    var source = document.forms[0].elements[availablelist];
    var dest = document.forms[0].elements[destlist];
    dest.options.length=0;

    if (false) {
        var selectedIndex = document.forms[0].property.options.selectedIndex;
        var propText = document.forms[0].property.options[selectedIndex].text;
        var propVal = document.forms[0].property.options[selectedIndex].value;
    }

    for (var i=0; i < source.options.length; i++) {
        var optionText = false ? source.options[i].text + ' [' + propText + ']' : source.options[i].text;
        var optionVal = false ? source.options[i].value + ':' + propVal : source.options[i].value;
        dest.options[i] = new Option(optionText,optionVal);
    }
}


function valuemoveSelectedUp(availablelist, destlist) {
    var flag = 0;
    var sel = document.forms[0].elements[destlist];
    for (var i=1; i < sel.options.length; i++) {
        if (sel.options[i].selected == true) {
            var t_value = sel.options[i].value;
            var t_text = sel.options[i].text;
            var t_selected = sel.options[i].selected;
            sel.options[i].value = sel.options[i-1].value;
            sel.options[i].text = sel.options[i-1].text;
            sel.options[i].selected = sel.options[i-1].selected;
            sel.options[i-1].value = t_value;
            sel.options[i-1].text = t_text;
            sel.options[i-1].selected = t_selected;
            flag = 1;
        }
    }
    if (flag == 0) alert('Please highlight a selected item to move it up');
}

function valuemoveSelectedDown(availablelist, destlist) {
    var flag = 0;
    var sel = document.forms[0].elements[destlist];
    for (var i=sel.options.length-2; i>=0; i--) {
        if (sel.options[i].selected == true) {
            var t_value = sel.options[i].value;
            var t_text = sel.options[i].text;
            var t_selected = sel.options[i].selected;
            sel.options[i].value = sel.options[i+1].value;
            sel.options[i].text = sel.options[i+1].text;
            sel.options[i].selected = sel.options[i+1].selected;
            sel.options[i+1].value = t_value;
            sel.options[i+1].text = t_text;
            sel.options[i+1].selected = t_selected;
            flag = 1;
        }
    }
    if (flag == 0) alert('Please highlight a selected item to move it down');
}