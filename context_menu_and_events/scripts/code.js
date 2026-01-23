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

// Initialize plugin and register handlers inside init
window.Asc.plugin.init = function init() {
// Button callback (will work only if buttons are configured in config.json)
this.button = function button(id) {
	this.executeCommand("close", "");
};

// Context menu handler
this.event_onContextMenuShow = function event_onContextMenuShow(options) {
	switch (options.type) {
	case "Target":
		this.executeMethod("AddContextMenuItem", [
		{
			guid: this.guid,
			items: [
			{
				id: "onClickItem1",
				text: { en: "Item 1", de: "Menü 1" },
				items: [
				{
					id: "onClickItem1Sub1",
					text: { en: "Subitem 1", de: "Untermenü 1" },
					disabled: true
				},
				{
					id: "onClickItem1Sub2",
					text: { en: "Subitem 2", de: "Untermenü 2" },
					separator: true
				}
				]
			},
			{
				id: "onClickItem2",
				text: { en: "Item 2", de: "Menü 2" }
			}
			]
		}
		]);
		break;

	case "Selection":
		this.executeMethod("AddContextMenuItem", [
		{
			guid: this.guid,
			items: [
			{
				id: "onClickItem3",
				text: { en: "Item 3", de: "Menü 3" }
			}
			]
		}
		]);
		break;

	case "Image":
	case "Shape":
		this.executeMethod("AddContextMenuItem", [
		{
			guid: this.guid,
			items: [
			{
				id: "onClickItem4",
				text: { en: "Item 4", de: "Menü 4" }
			}
			]
		}
		]);
		break;

	default:
		break;
	}
};

// Attach click handlers
this.attachContextMenuClickEvent("onClickItem1Sub1", () => {
	this.executeMethod("InputText", ["clicked: onClickItem1Sub1"]);
});

this.attachContextMenuClickEvent("onClickItem1Sub2", () => {
	this.executeMethod("InputText", ["clicked: onClickItem1Sub2"]);
});

this.attachContextMenuClickEvent("onClickItem2", () => {
	this.executeMethod("InputText", ["clicked: onClickItem2"]);
});

this.attachContextMenuClickEvent("onClickItem3", () => {
	this.executeMethod("InputText", ["clicked: onClickItem3"]);
});

this.attachContextMenuClickEvent("onClickItem4", () => {
	console.log("clicked: onClickItem4");
});

// Target position changed event
this.event_onTargetPositionChanged = function event_onTargetPositionChanged() {
	console.log("event: onTargetPositionChanged");
};
};