using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;

using MetaInspector.Logic;

using Microsoft.Win32;

namespace MetaInspector
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        AssemblyLoader loader;
        public MainWindow()
        {
            InitializeComponent();
            
        }

        private void MainWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
            
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            treeView.Items.Clear();
            try
            {
               loader = new AssemblyLoader(textBox.Text);
               loader.GetAssemblyInfo();
            }
            catch (Exception)
            {
                MessageBox.Show("This is not a valid file. Please select another file", "Error", MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }
            
            treeView.Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255));
            treeView.Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E));
            treeView.FontFamily = new FontFamily("Consolas");
            treeView.FontSize = 20;
            loader.Namespaces.ForEach(y =>
            {
                var tvNamespace = new TreeViewItem
                {
                    Header = y.ClassInfo.First().Namespace,
                    Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E)),
                    Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255)),
                    FontFamily = new FontFamily("Consolas"),
                    FontSize = 20
                };

                y.ClassInfo.ForEach(item =>
                {
                    var tvitem = new TreeViewItem
                    {
                        Header = item.ClassName,
                        Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E)),
                        Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255)),
                        FontFamily = new FontFamily("Consolas"),
                        FontSize = 20
                    };
                    var tvItemProperties = new TreeViewItem
                    {
                        Header = "Properties",
                        Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E)),
                        Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255)),
                        FontFamily = new FontFamily("Consolas"),
                        FontSize = 20
                    };
                    var tvItemFields = new TreeViewItem
                    {
                        Header = "Fields",
                        Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E)),
                        Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255)),
                        FontFamily = new FontFamily("Consolas"),
                        FontSize = 20
                    };
                    var tvItemMethods = new TreeViewItem
                    {
                        Header = "Methods",
                        Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E)),
                        Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255)),
                        FontFamily = new FontFamily("Consolas"),
                        FontSize = 20
                    };
                    var tvItemMethodBody = new TreeViewItem
                    {
                        Header = "Methods",
                        Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E)),
                        Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255)),
                        FontFamily = new FontFamily("Consolas"),
                        FontSize = 20
                    };
                    var tvItemEvents = new TreeViewItem
                    {
                        Header = "Events",
                        Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E)),
                        Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255)),
                        FontFamily = new FontFamily("Consolas"),
                        FontSize = 20
                    };


                    item.Properties.ForEach(prop =>
                    {
                        tvItemProperties.Items.Add(new TreeViewItem
                        {
                            Header = prop,
                            Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E)),
                            Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255)),
                            FontFamily = new FontFamily("Consolas"),
                            FontSize = 20
                        });
                    });
                    item.Fields.ForEach(prop =>
                    {
                        tvItemFields.Items.Add(new TreeViewItem
                        {
                            Header = prop,
                            Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E)),
                            Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255)),
                            FontFamily = new FontFamily("Consolas"),
                            FontSize = 20
                        });
                    });
                    item.Methods.ForEach(prop =>
                    {
                        var methodItem = new TreeViewItem
                        {
                            Header = prop.MethodName,
                            Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E)),
                            Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255)),
                            FontFamily = new FontFamily("Consolas"),
                            FontSize = 20
                        };
                        methodItem.Items.Add(new TreeViewItem
                        {
                            Header = prop.MethodBody,
                            Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E)),
                            Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255)),
                            FontFamily = new FontFamily("Consolas"),
                            FontSize = 20
                        });

                        tvItemMethods.Items.Add(methodItem);
                    });

                   

                    item.Events.ForEach(prop =>
                    {
                        tvItemEvents.Items.Add(new TreeViewItem
                        {
                            Header = prop,
                            Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E)),
                            Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255)),
                            FontFamily = new FontFamily("Consolas"),
                            FontSize = 20
                        });
                    });

                    //tvItemMethods.Items.Add(tvItemMethodBody);
                    tvitem.Items.Add(tvItemFields);
                    tvitem.Items.Add(tvItemProperties);
                    tvitem.Items.Add(tvItemMethods);
                    tvitem.Items.Add(tvItemEvents);
                    tvNamespace.Items.Add(tvitem);
                });
                treeView.Items.Add(tvNamespace);
            });
        }

        private void BrowseFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.DefaultExt = ".dll";
            var result = dlg.ShowDialog();
            if (result == true)
            {
                textBox.Text = Path.GetFullPath(dlg.FileName);
                ButtonBase_OnClick(null, null);
            }
        }

        private void ExportClick(object sender, RoutedEventArgs e)
        {
            ShowInfoPanel("Obtaining an instance of  " + "Excel.Application", Dispatcher);
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            dynamic excelApp = Activator.CreateInstance(excelType);
            dynamic workBook = excelApp.Workbooks.Add();
            dynamic activeSheet = workBook.ActiveSheet();
            var t = new Thread(() =>
            {
                ExportData(excelApp, workBook, activeSheet,this.Dispatcher);
            });
            t.Start();            
        }

        private void ExportData(dynamic excelApp, dynamic workBook, dynamic activeSheet, Dispatcher d)
        {
            try
            {
                excelApp.Columns[1].ColumnWidth = 50;
                excelApp.Columns[2].ColumnWidth = 50;
                excelApp.Columns[3].ColumnWidth = 50;
                excelApp.Columns[4].ColumnWidth = 50;
                excelApp.Columns[5].ColumnWidth = 100;
                excelApp.Sheets["Sheet1"].Name = "Exported data";
                if (loader != null)
                {
                    var count = 1;
                    var indent = 1;
                    var isFirst = false;
                    var namespaceBuffer = String.Empty;

                    loader.Namespaces.ForEach(x =>
                    {
                        x.ClassInfo.ForEach(info =>
                        {
                            if (namespaceBuffer != info.Namespace)
                            {
                                ShowInfoPanel(String.Format("Writing namespace {0}", info.Namespace), d);
                                indent = 1;
                                activeSheet.Cells[count, indent] = info.Namespace;
                                activeSheet.Cells[count, indent].Interior.ColorIndex = 4;
                                namespaceBuffer = info.Namespace;
                                count++;
                                indent++;
                            }
                            ShowInfoPanel(String.Format("Writing class {0}", info.ClassName), d);
                            activeSheet.Cells[count, indent] = info.ClassName;
                            activeSheet.Cells[count, indent].Interior.ColorIndex = 6;
                            count++;
                            indent++;
                            activeSheet.Cells[count, indent] = "Fields";
                            activeSheet.Cells[count, indent].Interior.ColorIndex = 7;
                            count++;
                            indent++;
                            info.Fields.ForEach(field =>
                            {
                                ShowInfoPanel(String.Format("Writing field {0}", field), d);
                                activeSheet.Cells[count, indent] = field;
                                activeSheet.Cells[count, indent].Interior.ColorIndex = 19;
                                count++;
                            });
                            indent--;
                            activeSheet.Cells[count, indent] = "Events";
                            activeSheet.Cells[count, indent].Interior.ColorIndex = 4;
                            count++;
                            indent++;
                            info.Events.ForEach(events =>
                            {
                                ShowInfoPanel(String.Format("Writing event {0}", events), d);
                                activeSheet.Cells[count, indent] = events;
                                activeSheet.Cells[count, indent].Interior.ColorIndex = 46;
                                count++;
                            });
                            indent--;
                            activeSheet.Cells[count, indent] = "Properties";
                            activeSheet.Cells[count, indent].Interior.ColorIndex = 22;
                            indent++;
                            count++;
                            info.Properties.ForEach(prop =>
                            {
                                ShowInfoPanel(String.Format("Writing property {0}", prop), d);
                                activeSheet.Cells[count, indent] = prop;
                                activeSheet.Cells[count, indent].Interior.ColorIndex = 28;
                                count++;
                            });
                            indent--;
                            activeSheet.Cells[count, indent] = "Methods";
                            activeSheet.Cells[count, indent].Interior.ColorIndex = 37;
                            count++;
                            indent++;
                            info.Methods.ForEach(method =>
                            {
                                ShowInfoPanel(String.Format("Writing method {0}", method.MethodName), d);
                                activeSheet.Cells[count, indent] = method.MethodName;
                                activeSheet.Cells[count, indent].Interior.ColorIndex = 36;
                                indent++;
                                activeSheet.Cells[count, indent] = method.MethodBody;
                                activeSheet.Cells[count, indent].Interior.ColorIndex = 38;
                                indent--;
                                count++;
                            });
                            indent -= 2;
                        });
                    });
                    //////////////////////
                    dynamic statSheet = workBook.Worksheets.Add();
                    statSheet.Cells[1, 1] = "Namespace";
                    statSheet.Cells[1, 2] = "Number of containing classes In Namespaces";
                    statSheet.Name = "Statistics";
                    excelApp.Columns[1].ColumnWidth = 50;
                    excelApp.Columns[2].ColumnWidth = 50;
                    var cnt = 2;
                    loader.Namespaces.ForEach(x =>
                    {
                        statSheet.Cells[cnt, 2] = x.ClassInfo.Count;
                        statSheet.Cells[cnt, 1] = x.ClassInfo.First().Namespace;
                        cnt++;
                    });
                    dynamic chart =  workBook.Charts.Add();
                    chart.Name = "Namespace Population Statistics";
                    ///////////////////////////////
                }
                Thread.Sleep(TimeSpan.FromSeconds(5));
                string savePath = ConfigurationManager.AppSettings.Get("SavePath");
                File.Delete(savePath);
                workBook.SaveAs(savePath);
                ShowInfoPanel(String.Format("Saving {0}", savePath), d);
                if (d.Invoke(() => excelCb.IsChecked != null && !excelCb.IsChecked.Value))
                {
                    workBook.Close();
                    excelApp.Quit();
                }
                ShowInfoPanel("Excel.Application has quit", d);
                MessageBox.Show("Successfuly exported to Excel", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                HideInfoPanel(d);
                if (d.Invoke(() => excelCb.IsChecked != null && excelCb.IsChecked.Value))
                {
                    excelApp.Visible = true;
                }
            }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message, "Failed", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Failed", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            finally
            {
                if (d.Invoke(() => excelCb.IsChecked != null && !excelCb.IsChecked.Value))
                {
                    uint procId = 0;
                    GetWindowThreadProcessId((IntPtr)excelApp.Hwnd, out procId);
                    Process.GetProcessById((int)procId).Kill();
                }
                HideInfoPanel(d);
            }
        }

        private void ShowInfoPanel(string info, Dispatcher d)
        {
            d.Invoke(() => InfoPanel.Visibility = Visibility.Visible);
            d.Invoke(() => text.Text = info);
        }

        private void HideInfoPanel(Dispatcher d)
        {
            d.Invoke(() => InfoPanel.Visibility = Visibility.Collapsed);
        }
    }
}
