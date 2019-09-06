// This script exports photoshop layers as individual PNGs. It also
// writes a csd file that can be imported into Cocos studio where the images
// will be displayed in the same positions and draw order.

// Setting defaults. 初始变量
var writePngs = true;
var writeTemplate = false;
var writeCsd = true;
var ignoreHiddenLayers = true;
var pngScale = 1;
var groupsAsimgs = false;
var useRulerOrigin = false;
var imagesDir = "./images/";
var projectDir = "";
var padding = 1;
var typeText = "Scene";

// IDs for saving settings.
const settingsID = stringIDToTypeID("settings");
const writePngsID = stringIDToTypeID("writePngs");
const writeTemplateID = stringIDToTypeID("writeTemplate");
const writeCsdID = stringIDToTypeID("writeCsd");
const ignoreHiddenLayersID = stringIDToTypeID("ignoreHiddenLayers");
const groupsAsimgsID = stringIDToTypeID("groupsAsimgs");
const useRulerOriginID = stringIDToTypeID("useRulerOrigin");
const pngScaleID = stringIDToTypeID("pngScale");
const imagesDirID = stringIDToTypeID("imagesDir");
const projectDirID = stringIDToTypeID("projectDir");
const paddingID = stringIDToTypeID("padding");

var originalDoc;
try {
	originalDoc = app.activeDocument;
} catch (ignored) {}
var settings, progress;
loadSettings();
showDialog();

// 获得actionTag的十位数
function getgetTenNum()
{
	return '1' + Math.floor((Math.random() * 9 + 1) * 0x10000000)
      .toString(10)
      .substring(1);
}

// 生成guid独立标识
function guid() {
  function getFourNum() {
    return Math.floor((1 + Math.random()) * 0x10000)  //floor返回不大于的最大整数 
      .toString(16)
      .substring(1);
  }
  return getFourNum() + getFourNum() + '-' + getFourNum() + '-' + getFourNum() + '-' +
    getFourNum() + '-' + getFourNum() + getFourNum() + getFourNum();
}

// 运行模块
function run () {
	// Output dirs.输出文件夹
	var absProjectDir = absolutePath(projectDir);
	new Folder(absProjectDir).create();
	var absImagesDir = absolutePath(imagesDir);
	var imagesFolder = new Folder(absImagesDir);
	imagesFolder.create();
	var relImagesDir = imagesFolder.getRelativeURI(absProjectDir);
	relImagesDir = relImagesDir == "." ? "" : (relImagesDir + "/");

	// Get ruler origin. 获得素材位置信息
	var xOffSet = 0, yOffSet = 0;
	if (useRulerOrigin) {
		var ref = new ActionReference();
		ref.putEnumerated(charIDToTypeID("Dcmn"), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
		var desc = executeActionGet(ref);
		xOffSet = desc.getInteger(stringIDToTypeID("rulerOriginH")) >> 16;
		yOffSet = desc.getInteger(stringIDToTypeID("rulerOriginV")) >> 16;
	}

	activeDocument.duplicate();

	// Output template image. 输出效果图
	if (writeTemplate) {
		if (pngScale != 1) {
			scaleImage();
			storeHistory();
		}

		var file = new File(absImagesDir + "template");
		if (file.exists) file.remove();

		activeDocument.saveAs(file, new PNGSaveOptions(), true, Extension.LOWERCASE);

		if (pngScale != 1) restoreHistory();
	}

	if (!writeCsd && !writePngs) {
		activeDocument.close(SaveOptions.DONOTSAVECHANGES);
		return;
	}

	// Rasterize all layers.栅格化全部图层
	try {
		executeAction(stringIDToTypeID( "rasterizeAll" ), undefined, DialogModes.NO);
	} catch (ignored) {}

	// Collect and hide layers.逐一枚举并隐藏图层
	var layers = [];
	collectLayers(activeDocument, layers);
	var layersCount = layers.length;

	storeHistory();

	// Store the slot names and layers for each img.
	var slots = {}, imgs = { "default": [] };
	for (var i = layersCount - 1; i >= 0; i--) {
		var layer = layers[i];

		// Use groups as img names.
		var potentialimgName = trim(layer.parent.name);
		var layerGroupimg = potentialimgName.indexOf("-NOimg") == -1;
		var imgName = (groupsAsimgs && layer.parent.typename == "LayerSet" && layerGroupimg) ? potentialimgName : "default";

		var imgLayers = imgs[imgName];
		if (!imgLayers) imgs[imgName] = imgLayers = [];
		imgLayers[imgLayers.length] = layer;

		slots[layerName(layer)] = true;
	}

	//
	var docName = decodeURI(originalDoc.name);
		docName = docName.substring(0, docName.indexOf("."));
	var tagIndex = 74;
	var docWidth = activeDocument.width.as("px") * pngScale;
	var	docHeight = activeDocument.height.as("px") * pngScale;

	// Output csd 写入csd
	var csd = '<GameFile>\n';
	csd += '\t<PropertyGroup Name="' + docName + '" Type="' + typeText + '" ID="' + guid() + '" Version="3.10.0.0" />\n';
	csd += '\t<Content ctype="GameProjectContent">\n';
	csd += '\t\t<Content>\n';
	csd += '\t\t<Animation Duration="0" Speed="1.0000" />\n';
	if (typeText == 'Layer'){
	    csd += '\t\t<ObjectData Name="' + typeText + '" Tag="' + tagIndex + '" ctype="GameLayerObjectData">\n';
    } else {
    	csd += '\t\t<ObjectData Name="' + typeText + '" Tag="' + tagIndex + '" ctype="GameNodeObjectData">\n'
    }
	tagIndex += 1;
	if (typeText != 'Node'){
	    csd += '\t\t\t<Size X="' + docWidth + '" Y="' + docHeight + '" />\n';
    }
	csd += '\t\t\t<Children>\n';

	// Output imgs.
	var imgsCount = countAssocArray(imgs);
	var imgIndex = 0;
	for (var imgName in imgs) {
		if (!imgs.hasOwnProperty(imgName)) continue;
		// csd += '\t"' + imgName + '":{\n';

		var imgLayers = imgs[imgName];
		var imgLayersCount = imgLayers.length;
		var imgLayerIndex = 0;
		// for (var i = imgLayersCount - 1; i >= 0; i--) {
		for (var i = 0; i < imgLayersCount; i++) {
			var layer = imgLayers[i];
			var slotName = layerName(layer);
			var placeholderName, attachmentName;
			if (imgName == "default") {
				placeholderName = slotName;
				attachmentName = placeholderName;
			} else {
				placeholderName = slotName;
				attachmentName = imgName + "/" + slotName;
			}

			var x = activeDocument.width.as("px") * pngScale;
			var y = activeDocument.height.as("px") * pngScale;
              
              // 裁切素材
			layer.visible = true;
			if (!layer.isBackgroundLayer) activeDocument.trim(TrimType.TRANSPARENT, false, true, true, false);
			x -= activeDocument.width.as("px") * pngScale;
			y -= activeDocument.height.as("px") * pngScale;
			if (!layer.isBackgroundLayer) activeDocument.trim(TrimType.TRANSPARENT, true, false, false, true);
			var width = activeDocument.width.as("px") * pngScale + padding * 2;
			var height = activeDocument.height.as("px") * pngScale + padding * 2;

			// Save image.
			if (writePngs) {
				if (pngScale != 1) scaleImage();
				if (padding > 0) activeDocument.resizeCanvas(width, height, AnchorPosition.MIDDLECENTER);

				if (imgName != "default") new Folder(absImagesDir + imgName).create();
				activeDocument.saveAs(new File(absImagesDir + attachmentName), new PNGSaveOptions(), true, Extension.LOWERCASE);
			}

			restoreHistory();
			layer.visible = false;

			x += Math.round(width) / 2;
			y += Math.round(height) / 2;

			// Make relative to the Photoshop document ruler origin. 获取素材位置
			if (useRulerOrigin) {
				x -= xOffSet * pngScale;
				y -= activeDocument.height.as("px") * pngScale - yOffSet * pngScale; // Invert y.
			}

			csd += '\t\t\t\t<AbstractNodeData Name="' + slotName + '" ActionTag="' + getTenNum() + '" Tag="' + tagIndex + '" IconVisible="False" ctype="SpriteObjectData">\n';
			tagIndex += 1;
			// csd += '\t\t\t\t\t<Position X="' + (x - Math.round(width) * 0.5000) + '" Y="' + (y - Math.round(height) * 0.5000) + '" />\n';
            csd += '\t\t\t\t\t<Position X="' + (x) + '" Y="' + (y) + '" />\n';
            csd += '\t\t\t\t\t<Scale ScaleX="1.0000" ScaleY="1.0000" />\n';
            csd += '\t\t\t\t\t<AnchorPoint ScaleX="0.5000" ScaleY="0.5000" />\n';
            csd += '\t\t\t\t\t<CColor A="255" R="255" G="255" B="255" />\n';
            csd += '\t\t\t\t\t<Size X="' + Math.round(width) + '" Y="' + Math.round(height) + '" />\n';
            csd += '\t\t\t\t\t<PrePosition X="0.0000" Y="0.0000" />\n';
            csd += '\t\t\t\t\t<PreSize X="0.0000" Y="0.0000" />\n';
            csd += '\t\t\t\t\t<FileData Type="Normal" Path="' + relImagesDir + placeholderName + '.png" />\n';
            csd += '\t\t\t\t</NodeObjectData>\n';

			imgLayerIndex++;
		}

		imgIndex++;
	}
    
	csd += '\t\t\t</Children>\n';
    csd += '\t\t</ObjectData>\n';
    csd += '\t\t</Content>\n';
	csd += '\t</Content>\n';
	csd += '</GameProjectFile>\n';

	activeDocument.close(SaveOptions.DONOTSAVECHANGES);

	// Output csd file.
	if (writeCsd) {
		var name = decodeURI(originalDoc.name);
		name = name.substring(0, name.indexOf("."));
		var file = new File(absProjectDir + name + ".csd");
		file.remove();
		file.open("w", "TEXT");
		file.lineFeed = "\n";
		file.write(csd);
		file.close();
	}
}

// Dialog and settings:

function showDialog () {
	if (!originalDoc) {
		alert("Please open a document before running the LayersToPNG script.");
		return;
	}
	if (!hasFilePath()) {
		alert("Please save the document before running the LayersToPNG script.");
		return;
	}

	var dialog = new Window("dialog", "Layers To Cocos Studio 2 csd");
	dialog.alignChildren = "fill";

	var checkboxGroup = dialog.add("group");
		var group = checkboxGroup.add("group");
			group.orientation = "column";
			group.alignChildren = "left";
			var writePngsCheckbox = group.add("checkbox", undefined, " Write layers as PNGs");
			writePngsCheckbox.value = writePngs;
			var writeTemplateCheckbox = group.add("checkbox", undefined, " Write a template PNG");
			writeTemplateCheckbox.value = writeTemplate;
			var writeCsdCheckbox = group.add("checkbox", undefined, " Write Cocos csd");
			writeCsdCheckbox.value = writeCsd;
		group = checkboxGroup.add("group");
			group.orientation = "column";
			group.alignChildren = "left";
			var ignoreHiddenLayersCheckbox = group.add("checkbox", undefined, " Ignore hidden layers");
			ignoreHiddenLayersCheckbox.value = ignoreHiddenLayers;
			var groupsAsimgsCheckbox = group.add("checkbox", undefined, " Use groups as imgs");
			groupsAsimgsCheckbox.value = groupsAsimgs;
			var useRulerOriginCheckbox = group.add("checkbox", undefined, " Use ruler origin as 0,0");
			useRulerOriginCheckbox.value = useRulerOrigin;

	var slidersGroup = dialog.add("group");
		group = slidersGroup.add("group");
			group.orientation = "column";
			group.alignChildren = "right";
			group.add("statictext", undefined, "PNG scale:");
			group.add("statictext", undefined, "Padding:");
		group = slidersGroup.add("group");
			group.orientation = "column";
			var scaleText = group.add("edittext", undefined, pngScale * 100);
			scaleText.characters = 4;
			var paddingText = group.add("edittext", undefined, padding);
			paddingText.characters = 4;
		group = slidersGroup.add("group");
			group.orientation = "column";
			group.add("statictext", undefined, "%");
			group.add("statictext", undefined, "px");
		group = slidersGroup.add("group");
			group.alignment = ["fill", ""];
			group.orientation = "column";
			group.alignChildren = ["fill", ""];
			var scaleSlider = group.add("slider", undefined, pngScale * 100, 1, 100);
			var paddingSlider = group.add("slider", undefined, padding, 0, 4);
	scaleText.onChanging = function () { scaleSlider.value = scaleText.text; };
	scaleSlider.onChanging = function () { scaleText.text = Math.round(scaleSlider.value); };
	paddingText.onChanging = function () { paddingSlider.value = paddingText.text; };
	paddingSlider.onChanging = function () { paddingText.text = Math.round(paddingSlider.value); };

	var stringOptions = [];
	stringOptions[0] = "Scene";
	stringOptions[1] = "Layer";
	stringOptions[2] = "Node";

	var selGroup = dialog.add("group");
			group = selGroup.add("group");
			group.orientation = "column";
			group.alignChildren = "right";
			group.add("statictext", undefined, "Type:");

			group = selGroup.add("group");
			group.alignment = ["fill", ""];
			group.orientation = "column";
			group.alignChildren = ["fill", ""];
			var typeSel = group.add('dropdownlist', undefined, 'Test');

			var item
			for (var i=0,len=stringOptions.length;i<len;i++){
				item = typeSel.add ('item', "" + stringOptions[i]);     
			};
			typeSel.onChange = function() {typeText = stringOptions[parseInt(this.selection)];};
			typeSel.selection = typeSel.items[0];

	var outputGroup = dialog.add("panel", undefined, "Output directories");
		outputGroup.alignChildren = "fill";
		outputGroup.margins = [10,15,10,10];
		var textGroup = outputGroup.add("group");
			group = textGroup.add("group");
				group.orientation = "column";
				group.alignChildren = "right";
				group.add("statictext", undefined, "Images:");
				group.add("statictext", undefined, ".csd:");
			group = textGroup.add("group");
				group.orientation = "column";
				group.alignChildren = "fill";
				group.alignment = ["fill", ""];
				var imagesDirText = group.add("edittext", undefined, imagesDir);
				var projectDirText = group.add("edittext", undefined, projectDir);
		outputGroup.add("statictext", undefined, "Begin paths with \"./\" to be relative to the PSD file.").alignment = "center";

	var group = dialog.add("group");
		group.alignment = "center";
		var runButton = group.add("button", undefined, "OK");
		var cancelButton = group.add("button", undefined, "Cancel");
		cancelButton.onClick = function () {
			dialog.close(0);
			return;
		};

	function updateSettings () {
		writePngs = writePngsCheckbox.value;
		writeTemplate = writeTemplateCheckbox.value;
		writeCsd = writeCsdCheckbox.value;
		ignoreHiddenLayers = ignoreHiddenLayersCheckbox.value;
		var scaleValue = parseFloat(scaleText.text);
		if (scaleValue > 0 && scaleValue <= 100) pngScale = scaleValue / 100;
		groupsAsimgs = groupsAsimgsCheckbox.value;
		useRulerOrigin = useRulerOriginCheckbox.value;
		imagesDir = imagesDirText.text;
		projectDir = projectDirText.text;
		var paddingValue = parseInt(paddingText.text);
		if (paddingValue >= 0) padding = paddingValue;
	}

	dialog.onClose = function() {
		updateSettings();
		saveSettings();
	};
	
	runButton.onClick = function () {
		if (scaleText.text <= 0 || scaleText.text > 100) {
			alert("PNG scale must be between > 0 and <= 100.");
			return;
		}
		if (paddingText.text < 0) {
			alert("Padding must be >= 0.");
			return;
		}
		dialog.close(0);

		var rulerUnits = app.preferences.rulerUnits;
		app.preferences.rulerUnits = Units.PIXELS;
		try {
			run();
		} catch (e) {
			alert("An unexpected error has occurred.\n\nTo debug, run the LayersToPNG script using Adobe ExtendScript "
				+ "with \"Debug > Do not break on guarded exceptions\" unchecked.");
			debugger;
		} finally {
			if (activeDocument != originalDoc) activeDocument.close(SaveOptions.DONOTSAVECHANGES);
			app.preferences.rulerUnits = rulerUnits;
		}
	};

	dialog.center();
	dialog.show();
}

function loadSettings () {
	try {
		settings = app.getCustomOptions(settingsID);
	} catch (e) {
		saveSettings();
	}
	if (typeof settings == "undefined") saveSettings();
	settings = app.getCustomOptions(settingsID);
	if (settings.hasKey(writePngsID)) writePngs = settings.getBoolean(writePngsID);
	if (settings.hasKey(writeTemplateID)) writeTemplate = settings.getBoolean(writeTemplateID);
	if (settings.hasKey(writeCsdID)) writeCsd = settings.getBoolean(writeCsdID);
	if (settings.hasKey(ignoreHiddenLayersID)) ignoreHiddenLayers = settings.getBoolean(ignoreHiddenLayersID);
	if (settings.hasKey(pngScaleID)) pngScale = settings.getDouble(pngScaleID);
	if (settings.hasKey(groupsAsimgsID)) groupsAsimgs = settings.getBoolean(groupsAsimgsID);
	if (settings.hasKey(useRulerOriginID)) useRulerOrigin = settings.getBoolean(useRulerOriginID);
	if (settings.hasKey(imagesDirID)) imagesDir = settings.getString(imagesDirID);
	if (settings.hasKey(projectDirID)) projectDir = settings.getString(projectDirID);
	if (settings.hasKey(paddingID)) padding = settings.getDouble(paddingID);
}

function saveSettings () {
	var settings = new ActionDescriptor();
	settings.putBoolean(writePngsID, writePngs);
	settings.putBoolean(writeTemplateID, writeTemplate);
	settings.putBoolean(writeCsdID, writeCsd);
	settings.putBoolean(ignoreHiddenLayersID, ignoreHiddenLayers);
	settings.putDouble(pngScaleID, pngScale);
	settings.putBoolean(groupsAsimgsID, groupsAsimgs);
	settings.putBoolean(useRulerOriginID, useRulerOrigin);
	settings.putString(imagesDirID, imagesDir);
	settings.putString(projectDirID, projectDir);
	settings.putDouble(paddingID, padding);
	app.putCustomOptions(settingsID, settings, true);
}

// Photoshop utility:

function scaleImage () {
	var imageSize = activeDocument.width.as("px");
	activeDocument.resizeImage(UnitValue(imageSize * pngScale, "px"), null, null, ResampleMethod.BICUBICSHARPER);
}

var historyIndex;
function storeHistory () {
	historyIndex = activeDocument.historyStates.length - 1;
}
function restoreHistory () {
	activeDocument.activeHistoryState = activeDocument.historyStates[historyIndex];
}

function collectLayers (layer, collect) {
	for (var i = 0, n = layer.layers.length; i < n; i++) {
		var child = layer.layers[i];
		if (ignoreHiddenLayers && !child.visible) continue;
		if (child.bounds[2] == 0 && child.bounds[3] == 0) continue;
		if (child.layers && child.layers.length > 0)
			collectLayers(child, collect);
		else if (child.kind == LayerKind.NORMAL) {
			collect.push(child);
			child.visible = false;
		}
	}
}

function hasFilePath () {
	var ref = new ActionReference();
	ref.putEnumerated(charIDToTypeID("Dcmn"), charIDToTypeID("Ordn"), charIDToTypeID("Trgt"));
	return executeActionGet(ref).hasKey(stringIDToTypeID("fileReference"));
}

function absolutePath (path) {
	path = trim(path);
	if (path.length == 0)
		path = activeDocument.path.toString();
	else if (imagesDir.indexOf("./") == 0)
		path = activeDocument.path + path.substring(1);
	path = path.replace(/\\/g, "/");
	if (path.substring(path.length - 1) != "/") path += "/";
	return path;
}

// JavaScript utility:

function countAssocArray (obj) {
	var count = 0;
	for (var key in obj)
		if (obj.hasOwnProperty(key)) count++;
	return count;
}

function trim (value) {
	return value.replace(/^\s+|\s+$/g, "");
}

function endsWith (str, suffix) {
	return str.indexOf(suffix, str.length - suffix.length) !== -1;
}

function stripSuffix (str, suffix) {
	if (endsWith(str.toLowerCase(), suffix.toLowerCase())) str = str.substring(0, str.length - suffix.length);
	return str;
}

function layerName (layer) {
	return stripSuffix(trim(layer.name), ".png").replace(/[:\/\\*\?\"\<\>\|]/g, "");
}
