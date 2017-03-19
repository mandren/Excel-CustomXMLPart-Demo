/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;
    var myControls = [];

    function controlDataObj() {
        this.Name = "";
        this.RangeAddress = "";
        this.WorksheetName = "";
        this.ID = "";
        this.Width = 12;
        this.Left = 0;
        this.ControlType = "";
    };

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $('#getxml-button-text').text("Get XML");
            $('#getxml-button-desc').text("Gets custom XML parts stored in this workbook");
            $('#getxml-button').click(getXML);

            $('#hydrate-button-text').text("Load Controls");
            $('#hydrate-button-desc').text("Loads control data into this workbook");
            $('#hydrate-button').click(hydrateWorkbook);

            $('#serialize-button-text').text("Serialize Data");
            $('#serialize-button-desc').text("Serialize the controal data into XML");
            $('#serialize-button').click(serializeData);
        });
    }

    function getXML() {
        Excel.run(function (ctx) {

            var xmlpart = ctx.workbook.customXmlParts.getByNamespace("CommonTools").getOnlyItem();
            xmlpart.load();

            return ctx.sync().then(function () {

                var xmlData = xmlpart.getXml();
                return ctx.sync().then(function () {

                    var doc = $.parseXML(xmlData.value);
                    var $items = $(doc).find("item");

                    $items.each(function () {
                        var controlData = new controlDataObj();
                        var props = this.childNodes;
                        $(props).each(function () {
                            var tmp = this.tagName;
                            switch (tmp) {
                                case "Name":
                                    controlData.Name = this.textContent;
                                    break;

                                case "RangeAddress":
                                    controlData.RangeAddress = this.textContent;
                                    break;

                                case "WorksheetName":
                                    controlData.WorksheetName = this.textContent;
                                    break;

                                case "ID":
                                    controlData.ID = this.textContent;
                                    break;

                                case "ControlType":
                                    controlData.ControlType = this.textContent;
                                    break;
                            }
                        })
                        myControls.push(controlData);
                    })

                    showNotification('Info', 'Completed.  Controls Found: ' + myControls.length);

                })

            });
        }).catch(function (error) { //...
        });

    }

    function hydrateWorkbook() {

        Excel.run(function (ctx) {

            myControls.forEach(function (value, index) {
                var item = new controlDataObj();
                item = value;

                var worksheet = ctx.workbook.worksheets.getItem(item.WorksheetName);
                var range = worksheet.getRange(item.RangeAddress);

                range.format.borders.getItem('EdgeBottom').style = 'Continuous';
                range.format.borders.getItem('EdgeLeft').style = 'Continuous';
                range.format.borders.getItem('EdgeRight').style = 'Continuous';
                range.format.borders.getItem('EdgeTop').style = 'Continuous';

                range.format.borders.getItem('EdgeBottom').color = item.ControlType;
                range.format.borders.getItem('EdgeLeft').color = item.ControlType;
                range.format.borders.getItem('EdgeRight').color = item.ControlType;
                range.format.borders.getItem('EdgeTop').color = item.ControlType;

                range.format.borders.getItem('EdgeBottom').weight = 'Medium';
                range.format.borders.getItem('EdgeLeft').weight = 'Medium';
                range.format.borders.getItem('EdgeRight').weight = 'Medium';
                range.format.borders.getItem('EdgeTop').weight = 'Medium';

                switch (item.ControlType) {
                    case "green":
                        range.format.fill.color = 'C6EFCE';
                        break;
                    case "yellow":
                        range.format.fill.color = 'FFEB9C';
                        break;
                    case "red":
                        range.format.fill.color = 'FFC7CE';
                        break;
                }

                var binding = ctx.workbook.bindings.add(range, 'Range', item.ID);

            });

            var bindings = ctx.workbook.bindings;
            bindings.load('items');

            return ctx.sync().then(function () {

                for (var i = 0; i < bindings.items.length; i++) {

                    var bindingid = bindings.items[i].id;
                    Office.select('bindings#' + bindingid).addHandlerAsync(Office.EventType.BindingSelectionChanged, onBindingSelectionChanged);

                }
            });

        }).catch(function (error) { //...
        });
    }

    function onBindingSelectionChanged(bArgs) {

        showNotification('Info', 'Binding selected: ' + bArgs.binding.id);
        //TODO: Flesh this out to add content to task pane

    }

    function serializeData() {

        var xmlData = '<?xml version="1.0"?><CommonToolsData xmlns="CommonTools">';

        Excel.run(function (ctx) {

            var bindings = ctx.workbook.bindings;
            bindings.load('items');

            return ctx.sync().then(function () {

                var ranges = [];
                var controls = [];

                for (var i = 0; i < bindings.items.length; i++) {

                    var controlData = new controlDataObj();
                    var binding = bindings.items[i];

                    controlData.ID = binding.id;
                    var range = binding.getRange();
                    range.load(["address", "format/*", "format/fill"]);

                    ranges.push(range);
                    controls.push(controlData);

                }
                return ctx.sync().then(function () {

                    for (var i = 0; i < ranges.length; i++) {

                        controls[i].RangeAddress = ranges[i].address.split("!")[1];
                        controls[i].WorksheetName = ranges[i].address.split("!")[0];

                        var format = ranges[i].format.fill.color;

                        switch (format) {
                            case "#C6EFCE":
                                controls[i].ControlType = 'green';
                                break;
                            case "#FFEB9C":
                                controls[i].ControlType = 'yellow';
                                break;
                            case "#FFC7CE":
                                controls[i].ControlType = 'red';
                                break;
                        }

                        xmlData += '<item>';
                        xmlData += '<RangeStart/>';
                        xmlData += '<RangeEnd/>';
                        xmlData += '<Name/>';
                        xmlData += '<Width/>';
                        xmlData += '<Left/>';
                        xmlData += '<ID>' + controls[i].ID + '</ID>';
                        xmlData += '<WorksheetName>' + controls[i].WorksheetName + '</WorksheetName>';
                        xmlData += '<ControlType>' + controls[i].ControlType + '</ControlType>';
                        xmlData += '<RangeAddress>' + controls[i].RangeAddress + '</RangeAddress>';
                        xmlData += '</item>';

                    }

                    xmlData += '</CommonToolsData>';

                    var xmlpart = ctx.workbook.customXmlParts.getByNamespace("CommonTools").getOnlyItem();
                    xmlpart.load();

                    return ctx.sync().then(function () {
                        xmlpart.setXml(xmlData);
                        xmlpart.load();
                        ctx.sync();

                        clearWorkbook();

                        //Should we remove bindings at this point too?
                    });
                });
            });

        }).catch(function (error) { //...
        });

    }

    function clearWorkbook() {

        Excel.run(function (ctx) {

            var bindings = ctx.workbook.bindings;
            bindings.load('items');

            return ctx.sync().then(function () {

                var ranges = [];

                for (var i = 0; i < bindings.items.length; i++) {

                    var binding = bindings.items[i];
                    var range = binding.getRange();
                    range.load(["address", "format/*", "format/fill", "format/borders"]);

                    ranges.push(range);
                }
                return ctx.sync().then(function () {

                    for (var i = 0; i < ranges.length; i++) {

                        ranges[i].format.fill.clear();
                        // ranges[i].format.borders.clear();

                        ranges[i].format.borders.getItem('EdgeBottom').style = 'None';
                        ranges[i].format.borders.getItem('EdgeLeft').style = 'None';
                        ranges[i].format.borders.getItem('EdgeRight').style = 'None';
                        ranges[i].format.borders.getItem('EdgeTop').style = 'None';

                    }

                    ctx.sync();
                });
            })
        }).catch(function (error) { //...
        });
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
