using System.Xml;

namespace CreateXslt;

public class XmlHelper
{
    public void GenerateXmlTransformationFile(ExcelWorksheet excelWorksheet, string filepath)
    {
        // Create an XmlDocument
        XmlDocument xmlDoc = new XmlDocument();

        // Create the XML declaration
        XmlDeclaration xmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
        xmlDoc.AppendChild(xmlDeclaration);

        // Create the root element (xsl:stylesheet)
        XmlElement xslStylesheet =
            xmlDoc.CreateElement("xsl", "stylesheet", "http://www.w3.org/1999/XSL/Transform");
        xslStylesheet.SetAttribute("version", "1.0");
        xslStylesheet.SetAttribute("xmlns:msxsl", "urn:schemas-microsoft-com:xslt");
        xslStylesheet.SetAttribute("exclude-result-prefixes", "msxsl");
        xmlDoc.AppendChild(xslStylesheet);

        // Create xsl:output element
        XmlElement xslOutput = xmlDoc.CreateElement("xsl", "output", "http://www.w3.org/1999/XSL/Transform");
        xslOutput.SetAttribute("method", "xml");
        xslOutput.SetAttribute("indent", "yes");
        xslOutput.SetAttribute("omit-xml-declaration", "no");
        xslStylesheet.AppendChild(xslOutput);

        // Create xsl:template element
        XmlElement xslTemplate = xmlDoc.CreateElement("xsl", "template", "http://www.w3.org/1999/XSL/Transform");
        xslTemplate.SetAttribute("match", "/");
        xslStylesheet.AppendChild(xslTemplate);

        // Create processing-instruction element
        XmlElement processingInstruction = xmlDoc.CreateElement("xsl","processing-instruction", "http://www.w3.org/1999/XSL/Transform");
        processingInstruction.SetAttribute("name", "mso-application");
        processingInstruction.InnerText = "progid=\"Excel.Sheet\"";
        xslTemplate.AppendChild(processingInstruction);

        XmlNamespaceManager nsManager = new XmlNamespaceManager(xmlDoc.NameTable);
        nsManager.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet");
        nsManager.AddNamespace("xmlns:o", "urn:schemas-microsoft-com:office:office");

        // Create Workbook element
        XmlElement workbook = xmlDoc.CreateElement("Workbook", "urn:schemas-microsoft-com:office:spreadsheet");
        workbook.SetAttribute("xmlns", "urn:schemas-microsoft-com:office:spreadsheet");
        workbook.SetAttribute("xmlns:o", "urn:schemas-microsoft-com:office:office");
        workbook.SetAttribute("xmlns:x", "urn:schemas-microsoft-com:office:excel");
        workbook.SetAttribute("xmlns:ss", "urn:schemas-microsoft-com:office:spreadsheet");
        workbook.SetAttribute("xmlns:html", "http://www.w3.org/TR/REC-html40");
        xslTemplate.AppendChild(workbook);

        // Create DocumentProperties element
        XmlElement documentProperties =
            xmlDoc.CreateElement("DocumentProperties", "urn:schemas-microsoft-com:office:office");
        workbook.AppendChild(documentProperties);

        // Create Author element
        XmlElement author = xmlDoc.CreateElement("Author", "urn:schemas-microsoft-com:office:office");
        author.InnerText = "Netcompany";
        documentProperties.AppendChild(author);

        // Create LastAuthor element
        XmlElement lastAuthor = xmlDoc.CreateElement("LastAuthor", "urn:schemas-microsoft-com:office:office");
        lastAuthor.InnerText = "jnn@netcompany.com";
        documentProperties.AppendChild(lastAuthor);

        // Create Created element
        XmlElement created = xmlDoc.CreateElement("Created", "urn:schemas-microsoft-com:office:office");
        created.InnerText = "2022-10-18T00:00:00Z";
        documentProperties.AppendChild(created);

        // Create LastSaved element
        XmlElement lastSaved = xmlDoc.CreateElement("LastSaved", "urn:schemas-microsoft-com:office:office");
        lastSaved.InnerText = "2022-10-18T00:00:00Z";
        documentProperties.AppendChild(lastSaved);

        // Create Version element
        XmlElement version = xmlDoc.CreateElement("Version", "urn:schemas-microsoft-com:office:office");
        version.InnerText = "1.00";
        documentProperties.AppendChild(version);

        // Create OfficeDocumentSettings element
        XmlElement officeDocumentSettings =
            xmlDoc.CreateElement("OfficeDocumentSettings", "urn:schemas-microsoft-com:office:office");
        workbook.AppendChild(officeDocumentSettings);

        // Create AllowPNG element
        XmlElement allowPNG = xmlDoc.CreateElement("AllowPNG", "urn:schemas-microsoft-com:office:office");
        officeDocumentSettings.AppendChild(allowPNG);

        // Create ExcelWorkbook element
        XmlElement excelWorkbook = xmlDoc.CreateElement("ExcelWorkbook", "urn:schemas-microsoft-com:office:excel");
        workbook.AppendChild(excelWorkbook);

        // Create WindowHeight element
        XmlElement windowHeight = xmlDoc.CreateElement("WindowHeight", "urn:schemas-microsoft-com:office:excel");
        windowHeight.InnerText = "8550";
        excelWorkbook.AppendChild(windowHeight);

        // Create WindowWidth element
        XmlElement windowWidth = xmlDoc.CreateElement("WindowWidth", "urn:schemas-microsoft-com:office:excel");
        windowWidth.InnerText = "19410";
        excelWorkbook.AppendChild(windowWidth);

        // Create WindowTopX element
        XmlElement windowTopX = xmlDoc.CreateElement("WindowTopX", "urn:schemas-microsoft-com:office:excel");
        windowTopX.InnerText = "0";
        excelWorkbook.AppendChild(windowTopX);

        // Create WindowTopY element
        XmlElement windowTopY = xmlDoc.CreateElement("WindowTopY", "urn:schemas-microsoft-com:office:excel");
        windowTopY.InnerText = "0";
        excelWorkbook.AppendChild(windowTopY);

        // Create ProtectStructure element
        XmlElement protectStructure =
            xmlDoc.CreateElement("ProtectStructure", "urn:schemas-microsoft-com:office:excel");
        protectStructure.InnerText = "False";
        excelWorkbook.AppendChild(protectStructure);

        // Create ProtectWindows element
        XmlElement protectWindows =
            xmlDoc.CreateElement("ProtectWindows", "urn:schemas-microsoft-com:office:excel");
        protectWindows.InnerText = "False";
        excelWorkbook.AppendChild(protectWindows);

        // Create Styles element
        XmlElement styles = xmlDoc.CreateElement("Styles", "urn:schemas-microsoft-com:office:spreadsheet");
        workbook.AppendChild(styles);

        // Create Style elements and their attributes...
        // Create Styles element
        XmlElement stylesElement = xmlDoc.CreateElement("Styles", "urn:schemas-microsoft-com:office:spreadsheet");

        var stylesData = GetStyles(xmlDoc, stylesElement);

        stylesElement.AppendChild(stylesData);
        // Finally, append Styles element to its parent
        // Replace "parentElement" with your actual parent element variable
        workbook.AppendChild(stylesElement);

        // Create Worksheet element
        XmlElement worksheet = xmlDoc.CreateElement("Worksheet", "urn:schemas-microsoft-com:office:spreadsheet");
        worksheet.SetAttribute("Name", "urn:schemas-microsoft-com:office:spreadsheet", excelWorksheet.name);
        workbook.AppendChild(worksheet);

        // Create Names element
        XmlElement names = xmlDoc.CreateElement("Names", "urn:schemas-microsoft-com:office:spreadsheet");
        worksheet.AppendChild(names);

        // Create NamedRange element
        XmlElement namedRange = xmlDoc.CreateElement("NamedRange", "urna:schemas-microsoft-com:office:spreadsheet");
        namedRange.SetAttribute("Name", "urn:schemas-microsoft-com:office:spreadsheet","_FilterDatabase");
        namedRange.SetAttribute("RefersTo","urn:schemas-microsoft-com:office:spreadsheet",
            $"={excelWorksheet.name}" +
            "!R2C1:R[{count(/ReportDto/ReportRows)}]" +
            $"C{excelWorksheet.columns.Count}"); 
        namedRange.SetAttribute("Hidden", "urn:schemas-microsoft-com:office:spreadsheet","1");
        names.AppendChild(namedRange);

        // Create Table element
        XmlElement table = xmlDoc.CreateElement("Table", "urn:schemas-microsoft-com:office:spreadsheet");
        table.SetAttribute("FullColumns","urn:schemas-microsoft-com:office:excel", "1");
        table.SetAttribute("FullRows", "urn:schemas-microsoft-com:office:excel", "1");
        table.SetAttribute("DefaultRowHeight", "urn:schemas-microsoft-com:office:spreadsheet", "15");
        worksheet.AppendChild(table);
        
        foreach (var excelWorksheetColumn in excelWorksheet.columns)
        {
            var column = ColumEXmlElement(xmlDoc);
            table.AppendChild(column);
        }

        // Report title
        XmlElement columnTitle = xmlDoc.CreateElement("Row", "urn:schemas-microsoft-com:office:spreadsheet");
        columnTitle.SetAttribute("AutoFitHeight","urn:schemas-microsoft-com:office:spreadsheet", "0");
        columnTitle.SetAttribute("Height","urn:schemas-microsoft-com:office:spreadsheet", "18.75");
        table.AppendChild(columnTitle);

        // Create Cell element inside Row
        XmlElement columnTitleCellElement =
            xmlDoc.CreateElement("Cell", "urn:schemas-microsoft-com:office:spreadsheet");
        columnTitleCellElement.SetAttribute("StyleID","urn:schemas-microsoft-com:office:spreadsheet", "s62");
        columnTitle.AppendChild(columnTitleCellElement);

        // Create Data element inside Cell
        XmlElement dataColumnTitle = xmlDoc.CreateElement("Data", "urn:schemas-microsoft-com:office:spreadsheet");
        dataColumnTitle.SetAttribute("Type","urn:schemas-microsoft-com:office:spreadsheet", "String");
        dataColumnTitle.InnerText = excelWorksheet.specialFirstLine;
        columnTitle.AppendChild(dataColumnTitle);

        // Create Row element
        XmlElement rowColTitleElement = xmlDoc.CreateElement("Row", "urn:schemas-microsoft-com:office:spreadsheet");
        rowColTitleElement.SetAttribute("AutoFitHeight", "urn:schemas-microsoft-com:office:spreadsheet", "0");
        table.AppendChild(rowColTitleElement);
        
        foreach (Column excelWorksheetColumn in excelWorksheet.columns)
        {
            var rowColTitleStylecellElement = RowColTitleElement(xmlDoc, excelWorksheetColumn);
            rowColTitleElement.AppendChild(rowColTitleStylecellElement);
        }
        
        // Create xsl:for-each element
        XmlElement xslForEach = xmlDoc.CreateElement("xsl", "for-each", "http://www.w3.org/1999/XSL/Transform");
        xslForEach.SetAttribute("select", "ReportDto/ReportRows/Row");
        table.AppendChild(xslForEach);

        // Create Rows within the xsl:for-each loop
        XmlElement rowElement = xmlDoc.CreateElement("Row");
        xslForEach.AppendChild(rowElement);

        foreach (Column excelWorksheetColumn in excelWorksheet.columns)
        {
            var cellElement = CellElement(xmlDoc, rowElement, excelWorksheetColumn, excelWorksheet.name);
            rowElement.AppendChild(cellElement);
        }
        

        // Create WorksheetOptions element
        XmlElement worksheetOptions =
            xmlDoc.CreateElement("WorksheetOptions", "urn:schemas-microsoft-com:office:excel");
        worksheet.AppendChild(worksheetOptions);

        // Create PageSetup element
        XmlElement pageSetup = xmlDoc.CreateElement("PageSetup", "urn:schemas-microsoft-com:office:excel");
        worksheetOptions.AppendChild(pageSetup);

        // Create Header element inside PageSetup
        XmlElement header = xmlDoc.CreateElement("Header", "urn:schemas-microsoft-com:office:excel");
        header.SetAttribute("Margin","urn:schemas-microsoft-com:office:excel", "0.3");
        pageSetup.AppendChild(header);

        // Create Footer element inside PageSetup
        XmlElement footer = xmlDoc.CreateElement("Footer", "urn:schemas-microsoft-com:office:excel");
        footer.SetAttribute("Margin","urn:schemas-microsoft-com:office:excel", "0.3");
        pageSetup.AppendChild(footer);

        // Create PageMargins element inside PageSetup
        XmlElement pageMargins = xmlDoc.CreateElement("PageMargins", "urn:schemas-microsoft-com:office:excel");
        pageMargins.SetAttribute("Bottom","urn:schemas-microsoft-com:office:excel", "0.75");
        pageMargins.SetAttribute("Left","urn:schemas-microsoft-com:office:excel", "0.7");
        pageMargins.SetAttribute("Right","urn:schemas-microsoft-com:office:excel", "0.7");
        pageMargins.SetAttribute("Top","urn:schemas-microsoft-com:office:excel", "0.75");
        pageSetup.AppendChild(pageMargins);

        // Create Print element inside WorksheetOptions
        XmlElement print = xmlDoc.CreateElement("Print", "urn:schemas-microsoft-com:office:excel");
        worksheetOptions.AppendChild(print);

        // Create ValidPrinterInfo element inside Print
        XmlElement validPrinterInfo =
            xmlDoc.CreateElement("ValidPrinterInfo", "urn:schemas-microsoft-com:office:excel");
        print.AppendChild(validPrinterInfo);

        // Create PaperSizeIndex element inside Print
        XmlElement paperSizeIndex =
            xmlDoc.CreateElement("PaperSizeIndex", "urn:schemas-microsoft-com:office:excel");
        paperSizeIndex.InnerText = "9";
        print.AppendChild(paperSizeIndex);

        // Create HorizontalResolution element inside Print
        XmlElement horizontalResolution =
            xmlDoc.CreateElement("HorizontalResolution", "urn:schemas-microsoft-com:office:excel");
        horizontalResolution.InnerText = "600";
        print.AppendChild(horizontalResolution);

        // Create VerticalResolution element inside Print
        XmlElement verticalResolution =
            xmlDoc.CreateElement("VerticalResolution", "urn:schemas-microsoft-com:office:excel");
        verticalResolution.InnerText = "600";
        print.AppendChild(verticalResolution);

        // Create Selected element inside WorksheetOptions
        XmlElement selected = xmlDoc.CreateElement("Selected", "urn:schemas-microsoft-com:office:excel");
        worksheetOptions.AppendChild(selected);

        // Create ProtectObjects element inside WorksheetOptions
        XmlElement protectObjects =
            xmlDoc.CreateElement("ProtectObjects", "urn:schemas-microsoft-com:office:excel");
        protectObjects.InnerText = "False";
        worksheetOptions.AppendChild(protectObjects);

        // Create ProtectScenarios element inside WorksheetOptions
        XmlElement protectScenarios =
            xmlDoc.CreateElement("ProtectScenarios", "urn:schemas-microsoft-com:office:excel");
        protectScenarios.InnerText = "False";
        worksheetOptions.AppendChild(protectScenarios);

        // Create xsl:if element
        XmlElement xslIf = xmlDoc.CreateElement("xsl", "if", "http://www.w3.org/1999/XSL/Transform");
        xslIf.SetAttribute(excelWorksheet.name, "count(/ReportDto/ReportRows/Row) != 0"); //reportname
        worksheet.AppendChild(xslIf);

        // Inside xsl:if, create AutoFilter element
        XmlElement autoFilter = xmlDoc.CreateElement("AutoFilter", "urn:schemas-microsoft-com:office:excel");
        autoFilter.SetAttribute("Range","urn:schemas-microsoft-com:office:excel", "R2C1:R[{count(/ReportDto/ReportRows/Row)}]C1");
        xslIf.AppendChild(autoFilter);

        xmlDoc.Save(filepath);
    }

    private static XmlElement ColumEXmlElement(XmlDocument xmlDoc)
    {
        // Create Column element
        XmlElement column = xmlDoc.CreateElement("Column", "urn:schemas-microsoft-com:office:spreadsheet");
        column.SetAttribute("AutoFitWidth", "urn:schemas-microsoft-com:office:spreadsheet", "0");
        column.SetAttribute("Width", "urn:schemas-microsoft-com:office:spreadsheet", "100");
        return column;
    }

    private XmlElement CellElement(XmlDocument xmlDoc, XmlElement rowElement, Column excelWorksheetColumn, string name)
    {
        if (excelWorksheetColumn.excelFilter == ExcelFilter.String ||
            excelWorksheetColumn.excelFilter == ExcelFilter.Number)
        {
            // Create Cells within the Row
            XmlElement cellElement = xmlDoc.CreateElement("Cell", "urn:schemas-microsoft-com:office:spreadsheet");
            rowElement.AppendChild(cellElement);

            // Create Data within the Cell
            XmlElement dataElement = xmlDoc.CreateElement("Data", "urn:schemas-microsoft-com:office:spreadsheet");
            dataElement.SetAttribute("Type", "urn:schemas-microsoft-com:office:spreadsheet",
                excelWorksheetColumn.excelFilter.ToString());

            // Create xsl:value-of element inside Data
            XmlElement valueOfElement = xmlDoc.CreateElement("xsl", "value-of", "http://www.w3.org/1999/XSL/Transform");
            valueOfElement.SetAttribute("select", excelWorksheetColumn._sqlQueryHeadline);
            dataElement.AppendChild(valueOfElement);

            // Create NamedCell inside Cell
            XmlElement namedCellElement =
                xmlDoc.CreateElement("NamedCell", "urn:schemas-microsoft-com:office:spreadsheet");
            namedCellElement.SetAttribute("Name", "urn:schemas-microsoft-com:office:spreadsheet", "_FilterDatabase");
            cellElement.AppendChild(dataElement);
            cellElement.AppendChild(namedCellElement);
            return cellElement;
        }

        if (excelWorksheetColumn.excelFilter == ExcelFilter.Date)
        {
            XmlElement cellElement = xmlDoc.CreateElement("Cell");
            cellElement.SetAttribute("ss:StyleID", "s69");
            xmlDoc.AppendChild(cellElement);

            XmlElement xslChooseElement = xmlDoc.CreateElement("xsl:choose");
            cellElement.AppendChild(xslChooseElement);

            XmlElement xslWhenElement = xmlDoc.CreateElement("xsl:when");
            xslWhenElement.SetAttribute(name, $"{excelWorksheetColumn._sqlQueryHeadline} != ''");
            xslChooseElement.AppendChild(xslWhenElement);

            XmlElement dataElement = xmlDoc.CreateElement("Data");
            dataElement.SetAttribute("ss:Type", "DateTime");
            xslWhenElement.AppendChild(dataElement);

            XmlElement xslValueElement = xmlDoc.CreateElement("xsl:value-of");
            xslValueElement.SetAttribute("select", "Ansvar_startdato");
            dataElement.AppendChild(xslValueElement);

            XmlElement namedCellElement = xmlDoc.CreateElement("NamedCell");
            namedCellElement.SetAttribute("ss:Name", "_FilterDatabase");
            cellElement.AppendChild(namedCellElement);
            return cellElement;
        }
        XmlElement cellErrorElement = xmlDoc.CreateElement("Cell", "urn:schemas-microsoft-com:office:spreadsheet");
        rowElement.AppendChild(cellErrorElement);
        return cellErrorElement;
    }

    private XmlElement RowColTitleElement(XmlDocument xmlDoc, Column excelWorksheetColumn)
    {
        // Create Cell element inside Row
        XmlElement rowColTitleStylecellElement =
            xmlDoc.CreateElement("Cell", "urn:schemas-microsoft-com:office:spreadsheet");
        rowColTitleStylecellElement.SetAttribute("StyleID", "urn:schemas-microsoft-com:office:spreadsheet",
            "s63"); //FIXME: STYLE?
      
        // Titel
        XmlElement rowColTitleNameElement = xmlDoc.CreateElement("Data", "urn:schemas-microsoft-com:office:spreadsheet");
        rowColTitleNameElement.SetAttribute("Type", "urn:schemas-microsoft-com:office:spreadsheet", "String");
        rowColTitleNameElement.InnerText = excelWorksheetColumn.userColumnTitle;
        rowColTitleStylecellElement.AppendChild(rowColTitleNameElement);
        
        // Create NamedCell element inside Cell
        XmlElement namedColTitleCellElement = xmlDoc.CreateElement("NamedCell", "urn:schemas-microsoft-com:office:spreadsheet");
        namedColTitleCellElement.SetAttribute("Name","urn:schemas-microsoft-com:office:spreadsheet", "_FilterDatabase");
        rowColTitleStylecellElement.AppendChild(namedColTitleCellElement);
        
        return rowColTitleStylecellElement;
    }

    private static XmlElement GetStyles(XmlDocument xmlDoc, XmlElement stylesElement)
    {
        // Create Default Style
        XmlElement defaultStyleElement = xmlDoc.CreateElement("Style", "urn:schemas-microsoft-com:office:spreadsheet");
        defaultStyleElement.SetAttribute("ID", "urn:schemas-microsoft-com:office:spreadsheet", "Default");
        defaultStyleElement.SetAttribute("Name", "urn:schemas-microsoft-com:office:spreadsheet", "Normal");

        // Add child elements to Default Style
        XmlElement alignmentElement = xmlDoc.CreateElement("Alignment", "urn:schemas-microsoft-com:office:spreadsheet");
        alignmentElement.SetAttribute("Vertical", "urn:schemas-microsoft-com:office:spreadsheet", "Bottom");
        defaultStyleElement.AppendChild(alignmentElement);

        defaultStyleElement.AppendChild(xmlDoc.CreateElement("Borders", "urn:schemas-microsoft-com:office:spreadsheet"));
        var defaultStyleFont = xmlDoc.CreateElement("Font", "urn:schemas-microsoft-com:office:spreadsheet");
        defaultStyleFont.SetAttribute("FontName", "urn:schemas-microsoft-com:office:spreadsheet", "Calibri");
        defaultStyleFont.SetAttribute("Family", "urn:schemas-microsoft-com:office:excel", "Swiss");
        defaultStyleFont.SetAttribute("Size", "urn:schemas-microsoft-com:office:spreadsheet", "14");
        defaultStyleFont.SetAttribute("Color", "urn:schemas-microsoft-com:office:spreadsheet", "#000000");
        defaultStyleFont.SetAttribute("Bold", "urn:schemas-microsoft-com:office:spreadsheet", "1");
        defaultStyleElement.AppendChild(defaultStyleFont);
        defaultStyleElement.AppendChild(xmlDoc.CreateElement("Interior", "urn:schemas-microsoft-com:office:spreadsheet"));
        defaultStyleElement.AppendChild(
            xmlDoc.CreateElement("NumberFormat", "urn:schemas-microsoft-com:office:spreadsheet"));
        defaultStyleElement.AppendChild(xmlDoc.CreateElement("Protection", "urn:schemas-microsoft-com:office:spreadsheet"));

        // Add Default Style to Styles element
        stylesElement.AppendChild(defaultStyleElement);

        // Create other Style elements similarly...
        // Style s62
        XmlElement s62Element = xmlDoc.CreateElement("Style", "urn:schemas-microsoft-com:office:spreadsheet");
        s62Element.SetAttribute("ID", "urn:schemas-microsoft-com:office:spreadsheet", "s62");
        XmlElement s62AlignmentElement = xmlDoc.CreateElement("Alignment", "urn:schemas-microsoft-com:office:spreadsheet");
        s62AlignmentElement.SetAttribute("Vertical", "urn:schemas-microsoft-com:office:spreadsheet", "Center");
        s62Element.AppendChild(s62AlignmentElement);
        XmlElement s62FontElement = xmlDoc.CreateElement("Font", "urn:schemas-microsoft-com:office:spreadsheet");
        s62FontElement.SetAttribute("FontName", "urn:schemas-microsoft-com:office:spreadsheet", "Calibri");
        s62FontElement.SetAttribute("Family", "urn:schemas-microsoft-com:office:excel", "Swiss");
        s62FontElement.SetAttribute("Size", "urn:schemas-microsoft-com:office:spreadsheet", "14");
        s62FontElement.SetAttribute("Color", "urn:schemas-microsoft-com:office:spreadsheet", "#000000");
        s62FontElement.SetAttribute("Bold", "urn:schemas-microsoft-com:office:spreadsheet", "1");
        s62Element.AppendChild(s62FontElement);
        stylesElement.AppendChild(s62Element);

        XmlElement s63Element = xmlDoc.CreateElement("Style", "urn:schemas-microsoft-com:office:spreadsheet");
        s63Element.SetAttribute("ss:ID", "s63");

        // Create and append Font element with attributes
        XmlElement fontElement = xmlDoc.CreateElement("Font", "urn:schemas-microsoft-com:office:spreadsheet");
        fontElement.SetAttribute("ss:FontName", "Calibri");
        fontElement.SetAttribute("x:Family", "Swiss");
        fontElement.SetAttribute("ss:Size", "11");
        fontElement.SetAttribute("ss:Color", "#FFFFFF");
        fontElement.SetAttribute("ss:Bold", "1");
        s63Element.AppendChild(fontElement);

        // Create and append Interior element with attributes
        XmlElement interiorElement = xmlDoc.CreateElement("Interior", "urn:schemas-microsoft-com:office:spreadsheet");
        interiorElement.SetAttribute("ss:Color", "#2F75B5");
        interiorElement.SetAttribute("ss:Pattern", "Solid");
        s63Element.AppendChild(interiorElement);
        return s63Element;
    }
}