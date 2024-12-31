// Run this script from After Effects (File > Scripts > Run Script File) while a keyframed mask is selected.
// The mask will be converted to JSON data, which can then be used by the Python script to create a Premiere mask.

// The next line is a shrunken version of a public domain JSON library.
// https://github.com/douglascrockford/JSON-js/blob/master/json2.js
"object"!=typeof JSON&&(JSON={}),function(){"use strict";var gap,indent,meta,rep,rx_one=/^[\],:{}\s]*$/,rx_two=/\\(?:["\\\/bfnrt]|u[0-9a-fA-F]{4})/g,rx_three=/"[^"\\\n\r]*"|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g,rx_four=/(?:^|:|,)(?:\s*\[)+/g,rx_escapable=/[\\"\u0000-\u001f\u007f-\u009f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,rx_dangerous=/[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g;function f(t){return t<10?"0"+t:t}function this_value(){return this.valueOf()}function quote(t){return rx_escapable.lastIndex=0,rx_escapable.test(t)?'"'+t.replace(rx_escapable,function(t){var e=meta[t];return"string"==typeof e?e:"\\u"+("0000"+t.charCodeAt(0).toString(16)).slice(-4)})+'"':'"'+t+'"'}function str(t,e){var n,o,f,u,r,$=gap,i=e[t];switch(i&&"object"==typeof i&&"function"==typeof i.toJSON&&(i=i.toJSON(t)),"function"==typeof rep&&(i=rep.call(e,t,i)),typeof i){case"string":return quote(i);case"number":return isFinite(i)?String(i):"null";case"boolean":case"null":return String(i);case"object":if(!i)return"null";if(gap+=indent,r=[],"[object Array]"===Object.prototype.toString.apply(i)){for(n=0,u=i.length;n<u;n+=1)r[n]=str(n,i)||"null";return f=0===r.length?"[]":gap?"[\n"+gap+r.join(",\n"+gap)+"\n"+$+"]":"["+r.join(",")+"]",gap=$,f}if(rep&&"object"==typeof rep)for(n=0,u=rep.length;n<u;n+=1)"string"==typeof rep[n]&&(f=str(o=rep[n],i))&&r.push(quote(o)+(gap?": ":":")+f);else for(o in i)Object.prototype.hasOwnProperty.call(i,o)&&(f=str(o,i))&&r.push(quote(o)+(gap?": ":":")+f);return f=0===r.length?"{}":gap?"{\n"+gap+r.join(",\n"+gap)+"\n"+$+"}":"{"+r.join(",")+"}",gap=$,f}}"function"!=typeof Date.prototype.toJSON&&(Date.prototype.toJSON=function(){return isFinite(this.valueOf())?this.getUTCFullYear()+"-"+f(this.getUTCMonth()+1)+"-"+f(this.getUTCDate())+"T"+f(this.getUTCHours())+":"+f(this.getUTCMinutes())+":"+f(this.getUTCSeconds())+"Z":null},Boolean.prototype.toJSON=this_value,Number.prototype.toJSON=this_value,String.prototype.toJSON=this_value),"function"!=typeof JSON.stringify&&(meta={"\b":"\\b","	":"\\t","\n":"\\n","\f":"\\f","\r":"\\r",'"':'\\"',"\\":"\\\\"},JSON.stringify=function(t,e,n){var o;if(gap="",indent="","number"==typeof n)for(o=0;o<n;o+=1)indent+=" ";else"string"==typeof n&&(indent=n);if(rep=e,e&&"function"!=typeof e&&("object"!=typeof e||"number"!=typeof e.length))throw Error("JSON.stringify");return str("",{"":t})}),"function"!=typeof JSON.parse&&(JSON.parse=function(text,reviver){var j;function walk(t,e){var n,o,f=t[e];if(f&&"object"==typeof f)for(n in f)Object.prototype.hasOwnProperty.call(f,n)&&(void 0!==(o=walk(f,n))?f[n]=o:delete f[n]);return reviver.call(t,e,f)}if(text=String(text),rx_dangerous.lastIndex=0,rx_dangerous.test(text)&&(text=text.replace(rx_dangerous,function(t){return"\\u"+("0000"+t.charCodeAt(0).toString(16)).slice(-4)})),rx_one.test(text.replace(rx_two,"@").replace(rx_three,"]").replace(rx_four,"")))return j=eval("("+text+")"),"function"==typeof reviver?walk({"":j},""):j;throw SyntaxError("JSON.parse")})}();

function copyToClipboard(string) {
	// Credit to Tomas Sinkunas on Adobe Community for this function.
	// https://community.adobe.com/t5/after-effects-discussions/how-can-i-copy-string-in-a-variable-to-clipboard-in-extendscript/td-p/10930590
	var cmd, isWindows;

	string = (typeof string === 'string') ? string : string.toString();
	isWindows = $.os.indexOf('Windows') !== -1;
	
	cmd = 'echo "' + string + '" | pbcopy';
	if (isWindows) {
		cmd = 'cmd.exe /c cmd.exe /c "echo ' + string + ' | clip"';
	}
	system.callSystem(cmd);	
}

function buildUI(thisObj, textToDisplay) {
	var win = (thisObj instanceof Panel)? thisObj : new Window('dialog', "Results", undefined, {resizable:true});
		win.spacing = 0;
		win.margins = 4;
		var mainGroup = win.add ("group");
			mainGroup.spacing = 4;
			mainGroup.margins = 0;
			mainGroup.orientation = "column";
			var tbox = mainGroup.add( "edittext", undefined, textToDisplay, { multiline: true} );
			tbox.maximumSize.height = 300;
			tbox.maximumSize.width = 500;
			tbox.minimumSize.height = 300;
			tbox.minimumSize.width = 500;
			
			var buttonGroup = win.add("group");
			buttonGroup.orientation = "row";
			var copyButton = buttonGroup.add("button", undefined, "Copy To Clipboard");
			copyButton.onClick = function () {copyToClipboard(textToDisplay);}
			var closeButton = buttonGroup.add("button", undefined, "Close", {name:"cancel"});		

	win.layout.layout(true);
	return win
}

function AEMasksToPP(){
	var scriptName = "Mask Convert";
	var activeComp = app.project.activeItem;
	if ((activeComp == null) || !(activeComp instanceof CompItem)) {
		alert("Please open a composition first.", scriptName);
		return null;
	}
	var selectedLayers = activeComp.selectedLayers;
	if (activeComp.selectedLayers.length != 1) {
		alert("Please select a mask.", scriptName);
		return null;
		}
	
	var selectedLayer = activeComp.selectedLayers[0];
	var selectedMask;
	for (var i = 0; i < selectedLayer.selectedProperties.length; i++) {
		if (selectedLayer.selectedProperties[i].isMask){
			if (selectedMask == undefined){
			selectedMask = selectedLayer.selectedProperties[i];
			}
			else {
				alert("Only one mask may be selected at a time.");
				return null;
			}
		}
	}
	if (selectedMask == undefined){
		alert("Please select a mask.");
		return null;
	}
	
	var width = selectedLayer.width;
	var height = selectedLayer.height;

	// Loop through all masks
	var maskShape = selectedMask.property("maskShape");
	var numKeys = maskShape.numKeys;
	
	// Loop through all keyframes of mask
	var keyFrameArray = []
	if (numKeys == 0){
		alert("Please select a mask with keyframes.");
		return null;
	}	
	for (var k = 1; k < numKeys + 1; k++){
		var newKeyFrame = {};
		newKeyFrame.time = maskShape.keyTime(k);
		newKeyFrame.closed = maskShape.value.closed;
		shapeThisFrame = maskShape.keyValue(k)
	
		// Loop through all points of the keyframe
		var pointArray = [];	
		for (var p = 0; p < shapeThisFrame.vertices.length; p++){			
			var vertex = shapeThisFrame.vertices[p];
			var normalizedVertex = [vertex[0] / width, vertex[1] / height];
			var outTangent = shapeThisFrame.outTangents[p];
			outTangent = [vertex[0] + outTangent[0], vertex[1] + outTangent[1]];	// Handle positions need to be converted from relative to absolute.
			outTangent = [outTangent[0] / width, outTangent[1] / height]
			var inTangent = shapeThisFrame.inTangents[p];			
			inTangent = [vertex[0] + inTangent[0], vertex[1] + inTangent[1]]		// Ditto
			inTangent = [inTangent[0] / width, inTangent[1] / height]
			pointArray.push({"vertex":normalizedVertex.toString(), "out":outTangent.toString(), "in":inTangent.toString()});
		}
		newKeyFrame.points = pointArray;			
		keyFrameArray.push(newKeyFrame);
	}
	
	if (keyFrameArray.length > 0){
		return JSON.stringify(keyFrameArray);
	}
	else {
		alert("An error occurred. Mask was not processed.");
		return null;
	}
}


// Show the Panel
var text = AEMasksToPP();
if (text != null){
	var w = buildUI(this, text);
	if (w.toString() == "[object Panel]") {
		w;
	} else {
		w.show();
	}
}


