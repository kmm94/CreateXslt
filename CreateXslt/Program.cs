using System;
using System.Collections.Generic;
using System.IO;
using DataAccess;

namespace CreateXslt
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Start");
            if (args == null || args.Length == 0)
            {
                Console.WriteLine("Please specify a filename as a parameter. ex: .\\CreateXslt.exe Mappe1.csv");
                return;
            }

            var fileContents = File.ReadAllText(args[0]);
            var dataTable = DataTable.New.ReadLazy(args[0]);

            List<Column> columns = new List<Column>();

            Console.WriteLine("Found columns:");
            foreach (string columnName in dataTable.ColumnNames)
            {
                Console.WriteLine(columnName);
                columns.Add(new Column(columnName, columnName));
            }

            Report report;
            Console.WriteLine($"Input report name: ");
            report = new Report(columns, Console.ReadLine());

            Console.WriteLine($"Total Columns: {columns.Count}");

            foreach (Column column in report.columns)
            {
                bool IsHandlingColumn = true;
                while (IsHandlingColumn)
                {
                    Console.WriteLine("\n---- o ---- o ---- o ----");
                    Console.WriteLine($"Handling column: {column.name}");
                    Console.WriteLine($"Rename column? Y/N");
                    var input = Console.ReadLine();
                    if ("y".Equals(input, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"Write the new column name:");
                        column.name = Console.ReadLine();
                    }

                    Console.WriteLine($"What is the datatype of the column? {GetAllDatatypes()}");
                    Datatype datatype;
                    bool isValid = false;
                    do
                    {
                        isValid = Enum.TryParse(Console.ReadLine(), out datatype);
                        if (!isValid)
                        {
                            Console.WriteLine("That datatype doesn't exists");
                        }
                    } while (!isValid);

                    column.datatype = datatype;
                    Console.WriteLine("\nResume");
                    Console.WriteLine($"column name: {column.name}");
                    Console.WriteLine($"column name: {column.datatype.ToString()}");
                    Console.WriteLine("Satisfied? Y/N");
                    input = Console.ReadLine(); 
                    if ("y".Equals(input, StringComparison.OrdinalIgnoreCase))
                    {
                        IsHandlingColumn = false;
                    }

                }

                string shitXsl = $"<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                                 "\n<xsl:stylesheet version=\"1.0\" xmlns:xsl=http://www.w3.org/1999/XSL/Transform xmlns:msxsl=\"urn:schemas-microsoft-com:xslt\" exclude-result-prefixes=\"msxsl\">" +
                                 "\n                <xsl:output method=\"xml\" indent=\"yes\" omit-xml-declaration=\"no\"/>" +
                                 "\n                <xsl:template match=\"/\">" +
                                 "\n                                <!--Sæt <?mso-application progid=\"Excel.Sheet\"?>-->" +
                                 "\n                                <xsl:processing-instruction name=\"mso-application\">progid=\"Excel.Sheet\"</xsl:processing-instruction>" +
                                 "\n                                <Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=http://www.w3.org/TR/REC-html40>" +
                                 "\n                                                <DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">" +
                                 "\n                                                                <Author>Netcompany</Author>" +
                                 "\n                                                                <LastAuthor>karm@netcompany.com</LastAuthor>" +
                                 "\n                                                                <Created>2024-02-19T00:00:00Z</Created>" +
                                 "\n                                                                <LastSaved>2024-02-19T00:00:00Z</LastSaved>" +
                                 "\n                                                                <Version>1.00</Version>" +
                                 "\n                                                </DocumentProperties>" +
                                 "\n                                                <OfficeDocumentSettings xmlns=\"urn:schemas-microsoft-com:office:office\">" +
                                 "\n                                                                <AllowPNG/>" +
                                 "\n                                                </OfficeDocumentSettings>" +
                                 "\n                                                <ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\">" +
                                 "\n                                                                <WindowHeight>8550</WindowHeight>" +
                                 "\n                                                                <WindowWidth>19410</WindowWidth>" +
                                 "\n                                                                <WindowTopX>0</WindowTopX>" +
                                 "\n                                                                <WindowTopY>0</WindowTopY>" +
                                 "\n                                                                <ProtectStructure>False</ProtectStructure>" +
                                 "\n                                                                <ProtectWindows>False</ProtectWindows>" +
                                 "\n                                                </ExcelWorkbook>" +
                                 "\n                                                <Styles>" +
                                 "\n                                                                <Style ss:ID=\"Default\" ss:Name=\"Normal\">" +
                                 "\n                                                                                <Alignment ss:Vertical=\"Bottom\"/>" +
                                 "\n                                                                                <Borders/>" +
                                 "\n                                                                                <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\"/>" +
                                 "\n                                                                                <Interior/>" +
                                 "\n                                                                                <NumberFormat/>" +
                                 "\n                                                                                <Protection/>" +
                                 "\n                                                                </Style>" +
                                 "\n                                                                <Style ss:ID=\"s62\">" +
                                 "\n                                                                                <Alignment ss:Vertical=\"Center\"/>" +
                                 "\n                                                                                <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"14\" ss:Color=\"#000000\" ss:Bold=\"1\"/>" +
                                 "\n                                                                </Style>" +
                                 "\n                                                                <Style ss:ID=\"s63\">" +
                                 "\n                                                                                <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#FFFFFF\" ss:Bold=\"1\"/>" +
                                 "\n                                                                                <Interior ss:Color=\"#2F75B5\" ss:Pattern=\"Solid\"/>" +
                                 "\n                                                                </Style>" +
                                 "\n                                                                <Style ss:ID=\"s64\">" +
                                 "\n                                                                                <NumberFormat ss:Format=\"Standard\"/>" +
                                 "\n                                                                </Style>" +
                                 "\n                                                                <Style ss:ID=\"s65\">" +
                                 "\n                                                                                <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\" ss:Bold=\"1\"/>" +
                                 "\n                                                                </Style>" +
                                 "\n                                                                <Style ss:ID=\"s66\">" +
                                 "\n                                                                                <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\" ss:Bold=\"1\"/>" +
                                 "\n                                                                                <NumberFormat ss:Format=\"Standard\"/>" +
                                 "\n                                                                </Style>" +
                                 "\n                                                                <Style ss:ID=\"s67\">" +
                                 "\n                                                                                <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\" ss:Bold=\"1\"/>" +
                                 "\n                                                                                <Interior ss:Color=\"#E7E6E6\" ss:Pattern=\"Solid\"/>" +
                                 "\n                                                                </Style>" +
                                 "\n                                                                <Style ss:ID=\"s68\">" +
                                 "\n                                                                                <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\" ss:Bold=\"1\"/>" +
                                 "\n                                                                                <Interior ss:Color=\"#E7E6E6\" ss:Pattern=\"Solid\"/>" +
                                 "\n                                                                                <NumberFormat ss:Format=\"Standard\"/>" +
                                 "\n                                                                </Style>" +
                                 "\n                                                                <Style ss:ID=\"s69\">" +
                                 "\n                                                                                <NumberFormat ss:Format=\"Short Date\"/>" +
                                 "\n                                                                </Style>" +
                                 "\n                                                                <Style ss:ID=\"s70\">" +
                                 "\n                                                                                <NumberFormat ss:Format=\"@\"/>" +
                                 "\n                                                                </Style>" +
                                 "\n                                                </Styles>" +
                                 $"\n                                                <Worksheet ss:Name=\"{report.name}\">" +
                                 "\n                                                                <Names>" +
                                 $"\n                                                                                <NamedRange ss:Name=\"_FilterDatabase\" ss:RefersTo=\"={report.name}!R2C1:R[{{count(/ReportDto/ReportRows)}}]C11\" ss:Hidden=\"1\"/>" +
                                 "\n                                                                </Names>" +
                                 "\n                                                                <Table  x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\">" +
                                 $"\n{GetXslColumns(report.columns)}" +
                                 "\n                                                                                <Row ss:AutoFitHeight=\"0\" ss:Height=\"18.75\">" +
                                 "\n                                                                                                <Cell ss:StyleID=\"s62\">" +
                                 "\n                                                                                                                <Data ss:Type=\"String\">Begrænsninger på medlemmer</Data>" +
                                 "\n                                                                                                </Cell>" +
                                 "\n                                                                                </Row>" +
                                 "\n                                                                                <Row ss:AutoFitHeight=\"0\">" +
                                 $"\n{GetXslHeaders(report.columns)}"+
                                 "\n                                                                                </Row>" +
                                 "\n                                                                                <xsl:for-each select=\"ReportDto/ReportRows/Row\">" +
                                 "\n                                                                                                                <Row ss:AutoFitHeight=\"0\">" +
                                 $"\n{GetSqlHeadline(report.columns)}"+
                                 "\n                                                                                                                </Row>" +
                                 "\n                                                                                </xsl:for-each>" +
                                 "\n                                                                </Table>" +
                                 "\n                                                                <WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">" +
                                 "\n                                                                                <PageSetup>" +
                                 "\n                                                                                                <Header x:Margin=\"0.3\"/>" +
                                 "\n                                                                                                <Footer x:Margin=\"0.3\"/>" +
                                 "\n                                                                                                <PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/>" +
                                 "\n                                                                                </PageSetup>" +
                                 "\n                                                                                <Unsynced/>" +
                                 "\n                                                                                <Print>" +
                                 "\n                                                                                                <ValidPrinterInfo/>" +
                                 "\n                                                                                                <PaperSizeIndex>9</PaperSizeIndex>" +
                                 "\n                                                                                                <HorizontalResolution>600</HorizontalResolution>" +
                                 "\n                                                                                                <VerticalResolution>600</VerticalResolution>" +
                                 "\n                                                                                </Print>" +
                                 "\n                                                                                <Selected/>" +
                                 "\n                                                                                <ProtectObjects>False</ProtectObjects>" +
                                 "\n                                                                                <ProtectScenarios>False</ProtectScenarios>" +
                                 "\n                                                                </WorksheetOptions>" +
                                 "\n                                                                <xsl:if test=\"count(/ReportDto/ReportRows/Row) != 0\">" +
                                 "\n                                                                                <AutoFilter x:Range=\"R2C1:R[{count(/ReportDto/ReportRows/Row)}]" +
                                 $"C{report.columns.Count}\"" +
                                 " xmlns=\"urn:schemas-microsoft-com:office:excel\">" +
                                 "\n                                                                                </AutoFilter>" +
                                 "\n                                                                </xsl:if>" +
                                 "\n                                                </Worksheet>" +
                                 "\n                                </Workbook>" +
                                 "\n                </xsl:template>" +
                                 "\n</xsl:stylesheet>";
                

                // Write the text to a new file named "WriteFile.txt".
                File.WriteAllText( "WriteFile.xml", shitXsl);
                
            }
        }
        
        private static string GetAllDatatypes()
        {
            string DatatypesText = "";

            foreach (var datatypes in Enum.GetNames(typeof(Datatype)))
            {
                DatatypesText += $"{datatypes}, ";
            }

            return DatatypesText;
        }


        private static string GetXslColumns(List<Column> reportColumns)
        {
            string columnText = "";

            foreach (Column reportColumn in reportColumns)
            {
                columnText += "                                                                                <Column ss:AutoFitWidth=\"0\" ss:Width=\"100\"/>";
            }

            return columnText;
        }
        
        private static string GetXslHeaders(List<Column> reportColumns)
        {
            string columnText = "";

            foreach (Column reportColumn in reportColumns)
            {
                columnText +=
                    "\n                                                                                                <Cell ss:StyleID=\"s63\">" +
                    $"\n                                                                                                                <Data ss:Type=\"String\">{reportColumn.name}</Data>" +
                    "\n                                                                                                                <NamedCell ss:Name=\"_FilterDatabase\"/>" +
                    "\n                                                                                                </Cell>";
            }

            return columnText;
        }
        
        
        private static string GetSqlHeadline(List<Column> reportColumns)
        {
            string columnText = "";

            foreach (Column reportColumn in reportColumns)
            {
                columnText +=
                    "\n                                                                                                                                <Cell>" +
                    $"\n                                                                                                                                                <Data ss:Type=\"{reportColumn.datatype}\">" +
                    $"\n                                                                                                                                                                <xsl:value-of select=\"{reportColumn.sqlQueryHeadline}\"/>" +
                    "\n                                                                                                                                                </Data>" +
                    "\n                                                                                                                                                <NamedCell ss:Name=\"_FilterDatabase\"/>" +
                    "\n                                                                                                                                </Cell>";
            }

            return columnText;
        }
    }
}