/**
 *
 * (c) Copyright Ascensio System SIA 2020
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */
(function(window, undefined){
	Editor = {};
	Editor.callMethod = async function(name, args)
	{
		return new Promise(resolve => (function(){
			Asc.plugin.executeMethod(name, args || [], function(returnValue){
				resolve(returnValue);
			});
		})());
	};

	Editor.callCommand = async function(func)
	{
		return new Promise(resolve => (function(){
			Asc.plugin.callCommand(func, false, true, function(returnValue){
				resolve(returnValue);
			});
		})());
	};

	let codeEditor	= null;
	let xmlData		= [];
	let lastSelectedXmlId = null;
	let originalXmlText = '';

	async function getXmls() {
		Asc.scope.editorType = window.Asc.plugin.info.editorType;
		xmlData = await Editor.callCommand(() => {
			let doc;
			switch(Asc.scope.editorType) {
				case 'word':
					doc = Api.GetDocument();
					break;
				case 'slide':
					doc = Api.GetPresentation();
					break;
				case 'cell':
					doc = Api.GetActiveSheet();
					break;
			}
			let xmlManager = doc.GetCustomXmlParts();
			let xmls = xmlManager.GetAll();
			return xmls.map(xml => ({text: xml.GetXml(), id: xml.GetId()}));
		});

		if (xmlData && xmlData.length)
			createListOfXmls(xmlData);
	}

	async function createXml()
	{
		Asc.scope.editorType = window.Asc.plugin.info.editorType;
		let id = await Editor.callCommand(() => {
			let doc;
			switch(Asc.scope.editorType) {
				case 'word':
					doc = Api.GetDocument();
					break;
				case 'slide':
					doc = Api.GetPresentation();
					break;
				case 'cell':
					doc = Api.GetActiveSheet();
					break;
			}
			let xmlManager = doc.GetCustomXmlParts();
			let xmlDefaultText = `<?xml version="1.0" encoding="UTF-8"?><defalut>text</defalut>`;
			let xml = xmlManager.Add(xmlDefaultText);
			return xml.GetId();
		});

		let listWrapper		= document.getElementById("xmlList");
		let option			= document.createElement("option");
		option.innerHTML	= id;

		listWrapper.appendChild(option);
		document.getElementById('xmlList').value = id;
		loadXmlTextAndStructure();
	}

	function getCurrentXmlId()
	{
		let select			= document.getElementById("xmlList");
		let index			= select.selectedIndex;
		if (index === -1)
			return "";

		let selectedItem	= select.options[index];
		return selectedItem.innerText
	}

	async function deleteXml()
	{
		let id = getCurrentXmlId();
		if (!id)
			return;

		Asc.scope.id = id;
		Asc.scope.editorType = window.Asc.plugin.info.editorType;

		await Editor.callCommand(() => {
			let doc;
			switch(Asc.scope.editorType) {
				case 'word':
					doc = Api.GetDocument();
					break;
				case 'slide':
					doc = Api.GetPresentation();
					break;
				case 'cell':
					doc = Api.GetActiveSheet();
					break;
			}
			let xmlManager = doc.GetCustomXmlParts();
			let xml		   = xmlManager.GetById(Asc.scope.id);
			xml.Delete();
		});

		let select = document.getElementById('xmlList');
		let selectedIndex = select.selectedIndex;
		if (selectedIndex !== -1) {
			select.remove(selectedIndex);
		}
	}

	async function updateXmlText(id, str)
	{
		Asc.scope.id			= id;
		Asc.scope.str			= str;
		Asc.scope.editorType	= window.Asc.plugin.info.editorType;

		await Editor.callCommand(() => {
			function getFirstTagName(xmlText) {
				const parser		= new DOMParser();
				const xmlDoc		= parser.parseFromString(xmlText, "application/xml");
				const parseError	= xmlDoc.getElementsByTagName("parsererror");

				if (parseError.length > 0)
					return false;

				return xmlDoc.documentElement.nodeName;
			}

			let doc;
			switch(Asc.scope.editorType) {
				case 'word':
					doc = Api.GetDocument();
					break;
				case 'slide':
					doc = Api.GetPresentation();
					break;
				case 'cell':
					doc = Api.GetActiveSheet();
					break;
			}
			let xmlManager	= doc.GetCustomXmlParts();
			let xml			= xmlManager.GetById(Asc.scope.id);
			if (!xml)
				return;
			let xmlText		= xml.GetXml();

			let rootNodeName= getFirstTagName(xmlText);
			let rootNodes	= xml.GetNodes('/' + rootNodeName);
			if (rootNodes.length)
			{
				let rootNode = rootNodes[0];
				rootNode.SetXml(Asc.scope.str);
			}
		});

		await updateAllContentControlsFromBinding();
	}

	async function getTextOfXml(id)
	{
		Asc.scope.id = id;
		Asc.scope.editorType = window.Asc.plugin.info.editorType;
		return await Editor.callCommand(() => {
			let doc;
			switch(Asc.scope.editorType) {
				case 'word':
					doc = Api.GetDocument();
					break;
				case 'slide':
					doc = Api.GetPresentation();
					break;
				case 'cell':
					doc = Api.GetActiveSheet();
					break;
			}
			let xmlManager	= doc.GetCustomXmlParts();
			let xml			= xmlManager.GetById(Asc.scope.id);
			return xml.GetXml();
		});
	}

	function createListOfXmls(xmlData)
	{
		let listWrapper			= document.getElementById("xmlList");
		listWrapper.innerHTML	= '';

		xmlData.forEach(xml => {
			let oButton			= document.createElement("option");
			oButton.innerHTML	= xml.id;

			listWrapper.appendChild(oButton);
		});

		loadXmlTextAndStructure();
	}

	async function loadXmlTextAndStructure()
	{
		let id = getCurrentXmlId();
		if (!id) {
			originalXmlText = '';
			updateSaveButtonState();
			return;
		}

		lastSelectedXmlId = id;
		let xmlText = await getTextOfXml(id);
		let prettyXmlText = prettifyXml(xmlText);
		
		originalXmlText = prettyXmlText;
		codeEditor.setValue(prettyXmlText);
		updateSaveButtonState();
		await createStrucOfXml(id);
		updateContentControlButtonStates();
	}

	async function createStrucOfXml(id)
	{
		let oStructure = document.getElementById("structureOfXML");
		oStructure.innerHTML = '';

		Asc.scope.id = id;
		Asc.scope.editorType = window.Asc.plugin.info.editorType;
		let data = await Editor.callCommand(function() {
			function GenerateDataFromNode (node, data)
			{
				let nodeName = node.GetNodeName();
				if (nodeName)
				{
					let nodeData = {
						name: nodeName,
						attributes: [],
						child: [],
						xPath: node.GetXPath()
					}

					let attributes = node.GetAttributes();
					attributes.forEach(attribute => {
						attribute.xPath = nodeData.xPath + "/@" + attribute.name;
						nodeData.attributes.push(attribute);
					})


					let childnodes = node.GetNodes("/*");
					if (childnodes)	
						childnodes.forEach((cnode => {GenerateDataFromNode(cnode, nodeData.child)}))

					data.push(nodeData)
				}
				else
				{
					let childnodes = node.GetNodes("/*");
					if (childnodes)	
						childnodes.forEach((cnode => {GenerateDataFromNode(cnode, data)}))
				}
			}

			function GetFirstTagName(xmlText) {
				const parser = new DOMParser();
				const xmlDoc = parser.parseFromString(xmlText, "application/xml");
				const parseError = xmlDoc.getElementsByTagName("parsererror");
		
				if (parseError.length > 0) {
				  return false;
				}
			  
				return xmlDoc.documentElement.nodeName;
			}

			let doc;
			switch(Asc.scope.editorType) {
				case 'word':
					doc = Api.GetDocument();
					break;
				case 'slide':
					doc = Api.GetPresentation();
					break;
				case 'cell':
					doc = Api.GetActiveSheet();
					break;
			}
			let xmlManager	= doc.GetCustomXmlParts();
			let xml			= xmlManager.GetById(Asc.scope.id);
			let xmlText		= xml.GetXml();
			let rootNodeName= GetFirstTagName(xmlText);
			let rootNodes	= xml.GetNodes('/' + rootNodeName);

			if (rootNodes.length)
			{
				let node	= rootNodes[0];
				let data	= [];
				GenerateDataFromNode(node, data);
				return JSON.stringify(data);
			}
		});

		if (!data)
			return;

		data = JSON.parse(data);

		function selectLi(el)
		{
			el.stopPropagation();
			Array.prototype.slice.call(document.querySelectorAll('li')).forEach(function(element){
				element.classList.remove('selected');
			});
			el.currentTarget.classList.add('selected');
			updateContentControlButtonStates();
		}

		function isContainAttributes(node)
		{
			let arr = node.attributes.filter(att => !(att.name.startsWith('xmlns:')) && att.name !== 'xmlns');
			return arr.length > 0;
		}

		function proceedData(data, oStructure)
		{
			for (let i = 0; i < data.length; i++)
			{
				let oCurrentData = data[i];

				if (oCurrentData.child.length || isContainAttributes(oCurrentData))
				{
					let li		= document.createElement("li");
					let details = document.createElement("details");
					let summury = document.createElement("summary");
					let ul		= document.createElement("ul");
					
					summury.innerText = oCurrentData.name;
	
					oStructure.appendChild(li);
					li.appendChild(details);
					details.appendChild(summury);
					details.appendChild(ul);

					li.onclick	= selectLi;
					li.xPath	= oCurrentData.xPath;

					proceedData(oCurrentData.child, ul);
					
					oCurrentData.attributes.forEach(attribute => {
						if (attribute.name === 'xmlns' || attribute.name.startsWith('xmlns:'))
							return;

						let li			= document.createElement("li");
						let summury		= document.createElement("summary");
		
						ul.appendChild(li);
						li.appendChild(summury);
						
						li.onclick		= selectLi;
						li.xPath		= attribute.xPath;
		
						summury.innerText = "@" + attribute.name;
					})
				}
				else
				{
					let summury			= document.createElement("summary");
					let li				= document.createElement("li");
					
					summury.innerText	= oCurrentData.name;

					li.appendChild(summury);

					li.onclick			= selectLi;
					li.xPath			= oCurrentData.xPath;

					oStructure.appendChild(li);
				}
			}
		}

		proceedData(data, oStructure);
		
		return data;
	}

	function prettifyXml(sourceXml) {
		// Save daclaration
		let xmlDeclaration = '';
		if (sourceXml.startsWith('<?xml')) {
			const match = sourceXml.match(/^<\?xml.*?\?>/);
			if (match) {
				xmlDeclaration = match[0] + '\n';
			}
		}
	
		// Parse xml
		const xmlDoc = new DOMParser().parseFromString(sourceXml, 'application/xml');
		const parsererror = xmlDoc.getElementsByTagName('parsererror');
		if (parsererror.length > 0) {
			console.log('XML parsing error: ' + parsererror[0].textContent);
		}
	
		// Minimal XSLT for formating
		const xsltString = `
			<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
				<xsl:output method="xml" indent="yes"/>
				<xsl:strip-space elements="*"/>
				<xsl:template match="node()|@*">
					<xsl:copy>
						<xsl:apply-templates select="node()|@*"/>
					</xsl:copy>
				</xsl:template>
			</xsl:stylesheet>
		`;
		const xsltDoc = new DOMParser().parseFromString(xsltString, 'application/xml');
	
		// Apply  XSLT
		const xsltProcessor = new XSLTProcessor();
		xsltProcessor.importStylesheet(xsltDoc);
		const resultDoc = xsltProcessor.transformToDocument(xmlDoc);
	
		// Serialize and format
		const rawXml = new XMLSerializer().serializeToString(resultDoc);
		return xmlDeclaration + formatXml(rawXml);
	}

	// Simple tab formatting
	function formatXml(xml) {
		const PADDING = '  ';
		let formatted = '';
		let pad = 0;

		xml.replace(/(>)(<)(\/*)/g, '$1\n$2$3').split('\n').forEach((node) => {
			if (node.match(/^<\/\w/)) pad--;
			formatted += PADDING.repeat(pad) + node + '\n';
			if (node.match(/^<[^!?\/][^>]*[^\/]>$/)) pad++;
		});

		return formatted.trim();
	}

	function getSelectedItemXPath()
	{
		let li = document.querySelectorAll('li.selected');
		
		if (li.length && li[0].xPath)
			return li[0].xPath;
	}

	function hasXmlChanged() {
		if (!codeEditor) return false;
		return codeEditor.getValue().trim() !== originalXmlText.trim();
	}

	function updateSaveButtonState() {
		const saveButton = document.getElementById('updateContentOfXml');
		if (saveButton) {
			saveButton.disabled = !hasXmlChanged();
		}
	}

	function updateContentControlButtonStates() {
		const hasSelection = !!getSelectedItemXPath();
		const bindButton = document.getElementById('match_with_selected_сс');
		const insertButton = document.getElementById('insert_cc');
		
		if (bindButton) {
			bindButton.disabled = !hasSelection;
		}
		if (insertButton) {
			insertButton.disabled = !hasSelection;
		}
	}

	async function updateAllContentControlsFromBinding()
	{
		Asc.scope.editorType	= window.Asc.plugin.info.editorType;
		await Editor.callCommand(() => {
			let doc;
			switch(Asc.scope.editorType) {
				case 'word':
					doc = Api.GetDocument();
					break;
				case 'slide':
					doc = Api.GetPresentation();
					break;
				case 'cell':
					doc = Api.GetActiveSheet();
					break;
			}

			let controls = doc.GetAllContentControls();

			for (let i = 0; i < controls.length; i++)
			{
				let cc = controls[i];
				if (cc)
					cc.UpdateFromXmlMapping();
			}
		});
	}

	async function updateXml(){
		let id = getCurrentXmlId()
		if (!id)
			return;

		let text	= codeEditor.getValue();
		await updateXmlText(id, text);
		originalXmlText = text;
		updateSaveButtonState();
		await createStrucOfXml(id);
	}

    window.Asc.plugin.init = async function()
    {
		// Hide content control elements for non-Word editors
		if (window.Asc.plugin.info.editorType !== 'word') {
			document.getElementById("structureLabel").style.display = 'none';
			document.getElementById("structureDiv").style.display = 'none';
			document.getElementById("contentControlsDiv").style.display = 'none';
		}

		if (!codeEditor) {
			codeEditor = CodeMirror(document.getElementById("main"), {
				mode: "text/xml",
				value: "",
				lineNumbers: true,
				lineWrapping: false,
				matchTags: {bothTags: true},
				autoCloseTags: {whenClosing: true},
				extraKeys: {
					"Ctrl-J": "toMatchingTag", 
					"Ctrl-Q": function(cm){cm.foldCode(cm.getCursor())},
					"Ctrl-S": async function(){await updateXml()}
				},
				autoCloseBrackets: true,
				foldGutter: true,
				gutters: ["CodeMirror-linenumbers", "CodeMirror-foldgutter"]
			});

			codeEditor.on('change', function() {
				updateSaveButtonState();
			});
		}

		document.getElementById("xmlList").addEventListener("input", loadXmlTextAndStructure);

		updateSaveButtonState();
		updateContentControlButtonStates();
		getXmls();
		
		document.getElementById("reloadContentOfXml").onclick = getXmls;

		document.getElementById("updateContentOfXml").onclick = async function(e) {
			await updateXml();
		};

		document.getElementById("createContentOfXml").onclick = createXml;

		document.getElementById("deleteContentOfXml").onclick = async function(e) {
			await deleteXml();
			codeEditor.setValue("");
			document.getElementById("structureOfXML").innerHTML = '';
			await getXmls();
		};

		document.getElementById("match_with_selected_сс").onclick = async function(e) {
			let xmlId	= document.getElementById('xmlList').value;

			Asc.scope.xmlId	= xmlId;
			Asc.scope.xPath = getSelectedItemXPath();

			await Editor.callCommand(() => {
				let Doc			= Api.GetDocument();
				let cc			= Doc.GetCurrentContentControl();
				if (!cc)
					return;
				cc.SetDataBinding({prefixMapping: "", storeItemID: Asc.scope.xmlId, xpath: Asc.scope.xPath});
			});
		}

		document.getElementById("insert_cc").onclick = async function(e) {

			let select			= document.getElementById("ccType");
			let index			= select.selectedIndex;
			let selectedItem	= select.options[index];
			let value			= selectedItem.value

			let id				= getCurrentXmlId();
			if (!id)
				return;

			Asc.scope.xmlId = id;
			Asc.scope.xPath = getSelectedItemXPath();

			if (!Asc.scope.xPath)
				return;

			if (value === 'block')
			{
				await Editor.callCommand(() => {
					let doc = Api.GetDocument();
					let sdt	= Api.CreateBlockLvlSdt();
					doc.InsertContent([sdt]);
					sdt.SetDataBinding({prefixMapping: "", storeItemID: Asc.scope.xmlId, xpath: Asc.scope.xPath});
				});
			}
			else if (value === 'inline')
			{
				await Editor.callCommand(() => {
					let doc			= Api.GetDocument();
					let sdt			= Api.CreateInlineLvlSdt();
					let oParagraph  = Api.CreateParagraph();
					oParagraph.AddElement(sdt, 0);
	
					sdt.SetDataBinding({prefixMapping: "", storeItemID: Asc.scope.xmlId, xpath: Asc.scope.xPath});
					doc.InsertContent([oParagraph]);
				});
			}
			else if (value === 'checkbox')
			{
				await Editor.callCommand(() => {
					let doc			= Api.GetDocument();
					let sdt			= doc.AddCheckBoxContentControl();

					sdt.SetDataBinding({prefixMapping: "", storeItemID: Asc.scope.xmlId, xpath: Asc.scope.xPath});
				});
			}
			else if (value === 'picture')
			{
				await Editor.callCommand(() => {
					let doc			= Api.GetDocument();
					let sdt			= doc.AddPictureContentControl();
					
					sdt.SetDataBinding({prefixMapping: "", storeItemID: Asc.scope.xmlId, xpath: Asc.scope.xPath});
				});
			}
			else if (value === 'combobox' || value === 'dropdownlist')
			{
				Asc.scope.name = value;
				await Editor.callCommand(() => {
					let doc			= Api.GetDocument();
					let sdt			= (Asc.scope.name === 'combobox')
						? doc.AddComboBoxContentControl()
						: doc.AddDropDownListContentControl();

					sdt.SetDataBinding({prefixMapping: "", storeItemID: Asc.scope.xmlId, xpath: Asc.scope.xPath});
				});
			}
			else if (value === 'date')
			{
				await Editor.callCommand(() => {
					let doc			= Api.GetDocument();
					let sdt			= doc.AddDatePickerContentControl();

					sdt.SetDataBinding({prefixMapping: "", storeItemID: Asc.scope.xmlId, xpath: Asc.scope.xPath});
				});
			}
		}

	}

	window.Asc.plugin.onThemeChanged = function(theme)
	{
		window.Asc.plugin.onThemeChangedBase(theme);
		if (theme.type.indexOf("dark") !== -1)
			setTimeout(function(){codeEditor && codeEditor.setOption("theme", "bespin")});
		else
			setTimeout(function(){codeEditor && codeEditor.setOption("theme", "default")});
	};

	window.Asc.plugin.button = function() {
		this.executeCommand("close", "");
	};

})(window, undefined);
