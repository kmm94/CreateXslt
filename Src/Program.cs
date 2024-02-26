using DataAccess;
using Terminal.Gui;

namespace CreateXslt
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            Application.Init();
            ReportController reportController = new ReportController();


            var columnNames = new Window("Columns")
            {
                X = 0,
                Y = 0,
                Width = Dim.Percent(25),
                Height = Dim.Fill() - 1,
            };
            var mainPage = new Window("DM Rapport accelerator")
            {
                X = 0,
                Y = 0,
                Width = Dim.Percent(100),
                Height = Dim.Fill() - 1,
                TextAlignment = TextAlignment.Centered,
                //TextFormatter = new TextFormatter() {Size = new Size(1, 10)}
            };
            
            var landingText = new TextView()
            {
                X = 0,
                Y = 0,
                Width = Dim.Fill(),
                Height = Dim.Fill(),
                Text = textHelper.GetLandingPageText(),
                ReadOnly = true,
            };
            
            mainPage.Add(landingText);
            

            var columnAttributes = new Window("column Attributes")
            {
                X = Pos.Right(columnNames),
                Y = 0,
                Width = Dim.Percent(60),
                Height = Dim.Fill() - 1
            };
            
            var columnData = new Window("column data")
            {
                X = Pos.Right(columnAttributes),
                Y = 0,
                Width = Dim.Percent(15),
                Height = Dim.Fill() - 1,
            };
            ListView columnDataListView = new ListView()
            {
                Width = Dim.Fill(),
                Height = Dim.Fill(),
                CanFocus = false,
                AllowsMarking = false,
            };
            
            ListView columnListView = new ListView()
            {
                Width = Dim.Fill(),
                Height = Dim.Fill(),
                CanFocus = false,
                AllowsMarking = false
            };

            var sqlColumnheaderLabel = new Label("SQL column Header: ")
            {
                X = 1,
                Y = 1,
                CanFocus = false,
            };
            
            //TODO: hotkeys to move to nex item 'KeystrokeNavigator'
            var sqlColumnheader = new Label("")
            {
                X = Pos.Right(sqlColumnheaderLabel),
                Y = 1,
                CanFocus = false,
            };

            var ColumnTitleLabel = new Label("User-friendly title: ")
            {
                X = 1,
                Y = Pos.Bottom(sqlColumnheaderLabel) + 1,
                CanFocus = false,
            };

            var inputColumnTitle = new TextField()
            {
                X = Pos.Right(ColumnTitleLabel) + 1,
                Y = Pos.Bottom(sqlColumnheaderLabel) + 1,
                Width = Dim.Fill(1),
                Height = 1
            };

            var chooseExcelFilterLabel = new Label("Choose Excel Filter:")
            {
                X = 1,
                Y = Pos.Bottom(ColumnTitleLabel)+1,
                CanFocus = false,
            };

            var excelFiltertypeRadioGroup = new RadioGroup(reportController.GetExcelfilters().ToArray())
            {
                X = 1,
                Y = Pos.Bottom(chooseExcelFilterLabel),
                DisplayMode = DisplayModeLayout.Horizontal
            };
            
            var chooseReportInputRadioGroupLabel = new Label("Choose Report input type:")
            {
                X = 1,
                Y = Pos.Bottom(excelFiltertypeRadioGroup)-1,
                CanFocus = false,
            };
            var reportInputRadioGroup = new RadioGroup(reportController.GetCrmReportInputType().ToArray())
            {
                X = 1,
                Y = Pos.Bottom(chooseReportInputRadioGroupLabel),
                DisplayMode = DisplayModeLayout.Horizontal
            };
            var exportButton = new Button("Export")
            {
                X = Pos.AnchorEnd() - 14,
                Y = Pos.AnchorEnd(1) ,
            };


            excelFiltertypeRadioGroup.SelectedItemChanged += changedArgs => reportController.SetColumnExcelFilter(changedArgs.SelectedItem);
            reportInputRadioGroup.SelectedItemChanged += changedArgs => reportController.SetColumnCrmInputType(changedArgs.SelectedItem);
            inputColumnTitle.KeyDown += ustring => reportController.selectedColumn.userColumnTitle = ustring.ToString();
            columnListView.SelectedItemChanged += (eventArgs) => HandleSelectedColumn(eventArgs, reportController, sqlColumnheader, inputColumnTitle, excelFiltertypeRadioGroup, columnDataListView);

            exportButton.Clicked += () =>
            {
                OpenDialog fileDialog = new OpenDialog()
                {
                    Title = "Choose where to save file",
                    Text = "Select a directory",
                    CanCreateDirectories = true,
                    CanChooseFiles = false,
                    CanChooseDirectories = true
                };
                Application.Run(fileDialog);
                if (fileDialog.Canceled is false)
                {
                    reportController.XmlHelper.GenerateXmlTransformationFile(reportController.GetExcelWorksheet(),(string)fileDialog.FilePath + $"\\{DateTime.Now.ToString("dd-MM-yyyy_HH-mm")}_XmlTransformationFile.xml");
                }
            };
            
            columnAttributes.Add(
                sqlColumnheaderLabel,
                sqlColumnheader,
                excelFiltertypeRadioGroup,
                chooseExcelFilterLabel,
                ColumnTitleLabel,
                inputColumnTitle,
                chooseReportInputRadioGroupLabel,
                reportInputRadioGroup);
            
            columnData.Add(columnDataListView);

            var topMenu = new MenuBar(new MenuBarItem[]
            {
                new MenuBarItem("_File", new MenuItem[]
                {
                    new MenuItem("open csv file", "", () =>
                    {
                        OpenDialog openFileDialog = new OpenDialog()
                        {
                            Title = "CSV",
                            Text = "Select the csv file produced by the sql",
                            AllowsMultipleSelection = false,
                            AllowedFileTypes = new [] { ".csv" },
                            CanChooseFiles = true,
                            CanChooseDirectories = false
                        };
                        Application.Run(openFileDialog);
                        if (openFileDialog.FilePaths.Count == 1)
                        {
                            reportController.LoadCsv(DataTable.New.ReadLazy(openFileDialog.FilePaths[0]));
                            Application.MainLoop.Invoke(() =>
                            {
                                columnListView.SetSource(reportController.GetExcelWorksheet().columns);
                                reportController.selectedColumn = reportController.GetExcelWorksheet().columns[0];
                                sqlColumnheader.Text = reportController.selectedColumn._sqlQueryHeadline;
                                inputColumnTitle.Text = reportController.selectedColumn.userColumnTitle;
                                excelFiltertypeRadioGroup.SelectedItem = reportController.GetIndexOfExcelfilters(reportController.selectedColumn.excelFilter);
                                columnDataListView.SetSource(reportController.selectedColumn._rawData);
                                
                                mainPage.Remove(landingText);
                                mainPage.Add(columnNames, columnAttributes, columnData,exportButton);
                                
                            });
                        }
                    })
                }),
            });


            //TODO: Add Column attributes:
            //TODO: Display, ColumnName validation(Display Error)
            //TODO: Validation between filter and rapport input data types
            //TODO: Logic, Guesstimate,CRM report input type
            //TODO: Logic, Generate ReportInput import to crm csv file
            //TODO: Logic, Generate select part of sql with castings for decimal numbers (is this possible)? 

            columnNames.Add(columnListView);
            Application.Top.Add(topMenu, mainPage);
            Application.Run();
        }

        private static void HandleSelectedColumn(ListViewItemEventArgs eventArgs, ReportController reportController,
            Label sqlColumnheader, TextField inputColumnTitle, RadioGroup excelFiltertypeRadioGroup,
            ListView columnDataListView)
        {
            
            Application.MainLoop.Invoke(() =>
            {
                Column selectedColumn = (Column)eventArgs.Value;
                reportController.selectedColumn = selectedColumn;
                sqlColumnheader.Text = selectedColumn._sqlQueryHeadline;
                inputColumnTitle.Text = selectedColumn.userColumnTitle;
                excelFiltertypeRadioGroup.SelectedItem = reportController.GetIndexOfExcelfilters(reportController.selectedColumn.excelFilter);
                columnDataListView.SetSource(selectedColumn._rawData);
            });
        }
    }
}
