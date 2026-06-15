// Created by Navid Siamakmanehs
// For additional information or help please contact me at Navid.siamakmanesh@gmail.com
// Enjoy

using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;

namespace Plot_Formularizer
{
    public partial class MainWindow : Window
    {
        private BitmapImage _image;
        private bool _loaded = false;

        private bool _logX = false;
        private bool _logY = false;



        private bool _setupFinished = false;
        private int _calibStep = 0;

        private bool origins_set = false;
        private bool points_set = false;
        private int number_of_points = 0;
        private int CALIBRATION_COUNT = 3;


        private Point _x2;
        private Point _origin;
        private Point _y2;

        private int _counter = 1;


        public enum PointType
        {
            X2,
            Origin,
            Y2,
            Data
        }

        public class PlotPoint
        {
            public string Name { get; set; }

            public Point Position { get; set; }

            public string XP { get; set; }
            public string YP { get; set; }

            public string X { get; set; }
            public string Y { get; set; }

            public PointType Type { get; set; }
        }
        private double CalculateX(double pixelX)
        {
            double originPixelX = _origin.X;
            double x2PixelX = _x2.X;

            double originValueX = double.Parse(TxtX1.Text);
            double x2ValueX = double.Parse(TxtX2.Text);

            if (_logX)
            {

                double log0 = Math.Log10(originValueX);
                double log1 = Math.Log10(x2ValueX);

                double t = (pixelX - originPixelX) / (x2PixelX - originPixelX);
                double logVal = log0 + t * (log1 - log0);

                return Math.Pow(10, logVal);
            }
            else
            {
                double scale =
                    (x2ValueX - originValueX) /
                    (x2PixelX - originPixelX);

                return originValueX +
                       ((pixelX - originPixelX) * scale);
            }
        }

        private double CalculateY(double pixelY)
        {
            double originPixelY = _origin.Y;
            double y2PixelY = _y2.Y;

            double originValueY = double.Parse(TxtY1.Text);
            double y2ValueY = double.Parse(TxtY2.Text);

            if (_logY)
            {
                double log0 = Math.Log10(originValueY);
                double log1 = Math.Log10(y2ValueY);

                double t = (pixelY - originPixelY) / (y2PixelY - originPixelY);
                double logVal = log0 + t * (log1 - log0);

                return Math.Pow(10, logVal);
            }
            else
            {
                double scale =
                    (y2ValueY - originValueY) /
                    (y2PixelY - originPixelY);

                return originValueY +
                       ((pixelY - originPixelY) * scale);
            }
        }
        public ObservableCollection<PlotPoint> Points { get; set; }
            = new ObservableCollection<PlotPoint>();

        public MainWindow()
        {
            InitializeComponent();

            PointsGrid.ItemsSource = Points;

            BtnFinishSetup.Click += BtnFinishSetup_Click;

            BtnUndo.Click += BtnUndo_Click;
            BtnClearAll.Click += BtnClearAll_Click;
            BtnDeletePoint.Click += BtnDeletePoint_Click;

            BtnFinishSetup.IsEnabled = false;
        }

        private void BtnUpload_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Filter = "Images|*.png;*.jpg;*.jpeg;*.bmp"
            };

            if (dlg.ShowDialog() == true)
            {
                _image = new BitmapImage(new Uri(dlg.FileName));
                ImgPlot.Source = _image;

                _loaded = true;

                TxtStatus.Text = "Image loaded → click Finish Setup";

                PreviewPlaceholder.Visibility = Visibility.Collapsed;
                ImageScrollViewer.Visibility = Visibility.Visible;

                BtnFinishSetup.IsEnabled = true;
                BtnFinishSetup.Content = "Finish Setup";
            }
        }
        private void BtnFinishSetup_Click(object sender, RoutedEventArgs e)
        {
            if (!_setupFinished)
            {
                StartSetup();
            }
            else if (!origins_set && _calibStep >= 4)
            {
                set_origins();
            }
            else if (origins_set && !points_set)
            {
                set_points();
            }
            else if(origins_set && points_set)
            {
                SaveToExcel();
            }
        }
        public void set_points()
        {
            points_set = true;
            BtnFinishSetup.Content = "Save to Excel";
            number_of_points = Points.Count;
            BtnInterpolate.IsEnabled = true;
            TxtStatus.Text = "Interpolate or save to excel";
        }
        private void SaveToExcel()
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet sheet = null;

            try
            {
                if (_image == null)
                {
                    TxtStatus.Text = "No image loaded.";
                    return;
                }


                string imagePath = new Uri(_image.UriSource.ToString()).LocalPath;

                if (!File.Exists(imagePath))
                {
                    MessageBox.Show("Image file not found.");
                    return;
                }


                string fileName = System.IO.Path.GetFileNameWithoutExtension(imagePath);
                string folder = System.IO.Path.GetDirectoryName(imagePath);
                string excelPath = System.IO.Path.Combine(folder, fileName + ".xlsx");


                if (File.Exists(excelPath))
                {
                    try
                    {
                        using (FileStream fs = File.Open(
                            excelPath,
                            FileMode.Open,
                            FileAccess.ReadWrite,
                            FileShare.None))
                        { }
                    }
                    catch
                    {
                        MessageBox.Show("File is currently open in Excel. Close it first.");
                        return;
                    }

                    MessageBoxResult res = MessageBox.Show(
                        "A file with this name already exists.\n\n" +
                        "YES = Replace\nNO = New numbered file\nCANCEL = Abort",
                        "File Exists",
                        MessageBoxButton.YesNoCancel,
                        MessageBoxImage.Question);

                    if (res == MessageBoxResult.Cancel)
                        return;

                    if (res == MessageBoxResult.No)
                    {
                        int suffix = 1;
                        string candidate;

                        do
                        {
                            candidate = System.IO.Path.Combine(folder, $"{fileName}_{suffix}.xlsx");
                            suffix++;
                        }
                        while (File.Exists(candidate));

                        excelPath = candidate;
                    }
                }


                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                workbook = excelApp.Workbooks.Add();
                sheet = (Excel.Worksheet)workbook.Worksheets[1];

                sheet.Cells[1, 1] = "Name";
                sheet.Cells[1, 2] = "XP";
                sheet.Cells[1, 3] = "YP";
                sheet.Cells[1, 4] = "X";
                sheet.Cells[1, 5] = "Y";

                sheet.Cells[1, 9] = "X Log";
                sheet.Cells[1, 10] = _logX ? "True" : "False";

                sheet.Cells[2, 9] = "Y Log";
                sheet.Cells[2, 10] = _logY ? "True" : "False";


                int row = 2;

                foreach (var p in Points)
                {
                    sheet.Cells[row, 1] = p.Name;
                    sheet.Cells[row, 2] = p.XP;
                    sheet.Cells[row, 3] = p.YP;
                    sheet.Cells[row, 4] = p.X;
                    sheet.Cells[row, 5] = p.Y;

                    row++;
                }

                float left = 300f;
                float top = 10f;

                var pic = sheet.Shapes.AddPicture(
                    imagePath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoCTrue,
                    left,
                    top,
                    _image.PixelWidth,
                    _image.PixelHeight
                );

                pic.LockAspectRatio =
                    Microsoft.Office.Core.MsoTriState.msoTrue;

                sheet.Columns.AutoFit();

                workbook.SaveAs(excelPath);
                workbook.Close(false);
                excelApp.Quit();

                TxtStatus.Text =
                    $"Saved successfully:\n{System.IO.Path.GetFileName(excelPath)}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
            finally
            {
                if (sheet != null)
                    Marshal.ReleaseComObject(sheet);

                if (workbook != null)
                    Marshal.ReleaseComObject(workbook);

                if (excelApp != null)
                    Marshal.ReleaseComObject(excelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        private void set_origins()
        {
            if (_calibStep < 4)
            {
                TxtStatus.Text = "Finish calibration first (X2, Origin, Y2)";
                return;
            }

            origins_set = true;
            BtnDeletePoint.IsEnabled = true;
            TxtStatus.Text = "Origins locked → digitize data points";

            BtnFinishSetup.Content = "Set Points";
        }

        private void StartSetup()
        {


            if (!_loaded)
            {
                TxtStatus.Text = "Upload image first.";
                return;
            }

            if (string.IsNullOrWhiteSpace(TxtX2.Text) ||
                string.IsNullOrWhiteSpace(TxtY2.Text))
            {
                TxtStatus.Text = "Fill X2 and Y2 values first.";
                return;
            }

            _setupFinished = true;
            _calibStep = 1;


            ChkLogX.IsEnabled = false;
            ChkLogY.IsEnabled = false;

            _logX = ChkLogX.IsChecked == true;
            _logY = ChkLogY.IsChecked == true;

            TxtX1.IsEnabled = false;
            TxtX2.IsEnabled = false;
            TxtY1.IsEnabled = false;
            TxtY2.IsEnabled = false;

            BtnUpload.IsEnabled = false;
            ExcelUpload.IsEnabled = false;
            BtnFinishSetup.Content = "Set Origin Points";

            PointsGrid.IsEnabled = true;
            PointsGrid.IsReadOnly = true;
            BtnUndo.IsEnabled = true;
            BtnClearAll.IsEnabled = true;

            TxtStatus.Text = "Click X2 (blue)";
        }

        private void RestartAll()
        {
            MessageBoxResult res = MessageBox.Show(
                "Reset everything?",
                "Confirm Reset",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (res != MessageBoxResult.Yes)
                return;

            _setupFinished = false;
            _loaded = false;
            _calibStep = 0;
            _counter = 1;

            Points.Clear();
            OverlayCanvas.Children.Clear();

            ImgPlot.Source = null;
            ImgMagnifier.Source = null;

            TxtX2.Text = "";
            TxtY2.Text = "";

            TxtX2.IsEnabled = true;
            TxtY2.IsEnabled = true;

            BtnUpload.IsEnabled = true;
            BtnFinishSetup.IsEnabled = false;
            BtnFinishSetup.Content = "Finish Setup";

            PointsGrid.IsEnabled = false;
            BtnUndo.IsEnabled = false;
            BtnClearAll.IsEnabled = false;
            BtnDeletePoint.IsEnabled = false;

            PreviewPlaceholder.Visibility = Visibility.Visible;
            ImageScrollViewer.Visibility = Visibility.Collapsed;

            TxtStatus.Text = "Reset complete";
        }

        private void Img_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (!points_set)
            {

           
            if (!_loaded || !_setupFinished)
                return;

            Point p = e.GetPosition(OverlayCanvas);

            if (!origins_set && _calibStep >= 1 && _calibStep <= 3)
            {
                HandleCalibration(p);
                return;
            }

            if (origins_set)
            {
                AddDataPoint(p);
            }
            }
        }

        private void HandleCalibration(Point p)
        {
            if (origins_set)
                return;

            if (_calibStep == 1)
            {
                _x2 = p;
                AddPoint("X2", p, PointType.X2);
                TxtStatus.Text = "Click Origin";
            }
            else if (_calibStep == 2)
            {
                _origin = p;
                AddPoint("Origin", p, PointType.Origin);
                TxtStatus.Text = "Click Y2";
            }
            else if (_calibStep == 3)
            {
                _y2 = p;
                AddPoint("Y2", p, PointType.Y2);
                TxtStatus.Text = "Click Set Origins button";
            }

            _calibStep++;
        }

        private void AddDataPoint(Point p)
        {
            AddPoint("P" + _counter, p, PointType.Data);
            _counter++;
        }

        private double CalcX(double px)
        {
            double x0 = _origin.X;
            double x1 = _x2.X;

            double v0 = double.Parse(TxtX1.Text);
            double v1 = double.Parse(TxtX2.Text);

            return v0 + (px - x0) * (v1 - v0) / (x1 - x0);
        }

        private double CalcY(double py)
        {
            double y0 = _origin.Y;
            double y1 = _y2.Y;

            double v0 = double.Parse(TxtY1.Text);
            double v1 = double.Parse(TxtY2.Text);

            return v0 + (py - y0) * (v1 - v0) / (y1 - y0);
        }

        private void AddPoint(string name, Point p, PointType type)
        {
            double calcX = 0;
            double calcY = 0;

            if (type == PointType.X2)
            {
                calcX = double.Parse(TxtX2.Text);
                calcY = double.Parse(TxtY1.Text);
            }
            else if (type == PointType.Origin)
            {
                calcX = double.Parse(TxtX1.Text);
                calcY = double.Parse(TxtY1.Text);
            }
            else if (type == PointType.Y2)
            {
                calcX = double.Parse(TxtX1.Text);
                calcY = double.Parse(TxtY2.Text);
            }
            else
            {
                calcX = CalculateX(p.X);
                calcY = CalculateY(p.Y);
            }

            Points.Add(new PlotPoint
            {
                Name = name,
                Position = p,

                XP = ((int)p.X).ToString(),
                YP = ((int)p.Y).ToString(),

                X = calcX.ToString("0.###"),
                Y = calcY.ToString("0.###"),

                Type = type
            });

            DrawDot(p, GetColor(type));
        }

        private Brush GetColor(PointType type)
        {
            return type == PointType.Data ? Brushes.Red : Brushes.Blue;
        }

        private void DrawDot(Point p, Brush color)
        {
            Ellipse el = new Ellipse
            {
                Width = 7,
                Height = 7,
                Fill = color
            };

            Canvas.SetLeft(el, p.X - 3.5);
            Canvas.SetTop(el, p.Y - 3.5);

            OverlayCanvas.Children.Add(el);
        }

        private void Img_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_loaded || _image == null)
                return;

            Point p = e.GetPosition(ImgPlot);
            UpdateMagnifier(p);
        }

        private void UpdateMagnifier(Point p)
        {
            try
            {
                int x = (int)p.X;
                int y = (int)p.Y;

                int size = 20;

                int x0 = Math.Max(0, x - size);
                int y0 = Math.Max(0, y - size);

                int w = size * 2;
                int h = size * 2;

                if (x0 + w > _image.PixelWidth)
                    w = _image.PixelWidth - x0;

                if (y0 + h > _image.PixelHeight)
                    h = _image.PixelHeight - y0;

                CroppedBitmap crop = new CroppedBitmap(
                    _image,
                    new Int32Rect(x0, y0, w, h)
                );

                ImgMagnifier.Source = new TransformedBitmap(
                    crop,
                    new ScaleTransform(4.0, 4.0)
                );
            }
            catch { }
        }

        private void BtnClearAll_Click(object sender, RoutedEventArgs e)
        {

            if (!origins_set)
            {
                Points.Clear();
                OverlayCanvas.Children.Clear();

                _calibStep = 1;

                TxtStatus.Text = "Click X2 (blue)";

                return;
            }

            if (origins_set && !points_set)
            {
                for (int i = Points.Count - 1; i >= CALIBRATION_COUNT; i--)
                {
                    OverlayCanvas.Children.RemoveAt(i);
                    Points.RemoveAt(i);
                }

                _counter = 1;
                return;
            }


            if (origins_set && points_set)
            {
                for (int i = Points.Count - 1; i >= number_of_points; i--)
                {
                    OverlayCanvas.Children.RemoveAt(i);
                    Points.RemoveAt(i);
                }

                _counter = 1;
                return;
            }
        }

        private void BtnUndo_Click(object sender, RoutedEventArgs e)
        {
            if (Points.Count == 0)
                return;

            if (!origins_set)
            {
                int lastIndex = Points.Count - 1;

                Points.RemoveAt(lastIndex);
                OverlayCanvas.Children.RemoveAt(lastIndex);

                if (_calibStep > 1)
                    _calibStep--;

                switch (_calibStep)
                {
                    case 1:
                        TxtStatus.Text = "Click X2 (blue)";
                        break;

                    case 2:
                        TxtStatus.Text = "Click Origin";
                        break;

                    case 3:
                        TxtStatus.Text = "Click Y2";
                        break;

                    case 4:
                        TxtStatus.Text = "Click Set Origins button";
                        break;
                }

                return;
            }


            if (points_set)
            {
                if (Points.Count <= number_of_points)
                    return;
            }
            else
            {
                if (Points.Count <= CALIBRATION_COUNT)
                    return;
            }

            int index = Points.Count - 1;

            Points.RemoveAt(index);
            OverlayCanvas.Children.RemoveAt(index);

            if (_counter > 1)
                _counter--;
        }
        private void BtnDeletePoint_Click(object sender, RoutedEventArgs e)
        {
            if (!(PointsGrid.SelectedItem is PlotPoint item))
                return;

            int index = Points.IndexOf(item);

            if (points_set)
            {
                if (index < number_of_points)
                    return;
            }
            else if (origins_set)
            {
                if (index < CALIBRATION_COUNT)
                    return;
            }

            Points.RemoveAt(index);

            OverlayCanvas.Children.Clear();

            foreach (var p in Points)
                DrawDot(p.Position, GetColor(p.Type));
        }
        private void LoadFromExcel(string excelPath)
{
    Excel.Application excelApp = null;
    Excel.Workbook workbook = null;
    Excel.Worksheet pointsSheet = null;
    Excel.Worksheet settingsSheet = null;

    try
    {
        excelApp = new Excel.Application();
        workbook = excelApp.Workbooks.Open(excelPath);

        pointsSheet =
            (Excel.Worksheet)workbook.Worksheets["Sheet1"];

        settingsSheet =
            (Excel.Worksheet)workbook.Worksheets["Settings"];

        RestartAll();


        TxtX1.Text = settingsSheet.Cells[1, 2].Value?.ToString();
        TxtX2.Text = settingsSheet.Cells[2, 2].Value?.ToString();
        TxtY1.Text = settingsSheet.Cells[3, 2].Value?.ToString();
        TxtY2.Text = settingsSheet.Cells[4, 2].Value?.ToString();

        ChkLogX.IsChecked =
            Convert.ToBoolean(settingsSheet.Cells[5, 2].Value);

        ChkLogY.IsChecked =
            Convert.ToBoolean(settingsSheet.Cells[6, 2].Value);

        string imagePath =
            settingsSheet.Cells[7, 2].Value?.ToString();

        if (!File.Exists(imagePath))
            throw new Exception(
                "Original image file is missing."
            );

        _image = new BitmapImage(new Uri(imagePath));

        ImgPlot.Source = _image;

        PreviewPlaceholder.Visibility =
            Visibility.Collapsed;

        ImageScrollViewer.Visibility =
            Visibility.Visible;

        _loaded = true;


        int row = 2;

        while (pointsSheet.Cells[row, 1].Value != null)
        {
            string name =
                pointsSheet.Cells[row, 1].Value.ToString();

            double xp =
                Convert.ToDouble(
                    pointsSheet.Cells[row, 2].Value);

            double yp =
                Convert.ToDouble(
                    pointsSheet.Cells[row, 3].Value);

            Point p = new Point(xp, yp);

            PointType type;

            if (name == "X2")
                type = PointType.X2;
            else if (name == "Origin")
                type = PointType.Origin;
            else if (name == "Y2")
                type = PointType.Y2;
            else
                type = PointType.Data;

            Points.Add(new PlotPoint
            {
                Name = name,
                Position = p,
                XP = xp.ToString(),
                YP = yp.ToString(),
                X = pointsSheet.Cells[row, 4].Value?.ToString(),
                Y = pointsSheet.Cells[row, 5].Value?.ToString(),
                Type = type
            });

            DrawDot(p, GetColor(type));

            if (type == PointType.X2)
                _x2 = p;

            if (type == PointType.Origin)
                _origin = p;

            if (type == PointType.Y2)
                _y2 = p;

            row++;
        }


        origins_set = true;
        _setupFinished = true;
        _calibStep = 4;

        _counter = 1;

        foreach (var p in Points)
        {
            if (p.Type == PointType.Data)
                _counter++;
        }

        BtnUpload.IsEnabled = false;

        TxtX1.IsEnabled = false;
        TxtX2.IsEnabled = false;
        TxtY1.IsEnabled = false;
        TxtY2.IsEnabled = false;

        ChkLogX.IsEnabled = false;
        ChkLogY.IsEnabled = false;

        BtnFinishSetup.Content = "Save to Excel";

        TxtStatus.Text =
            "Project restored successfully.";
    }
    catch (Exception ex)
    {
        MessageBox.Show(
            "Invalid or corrupted project file.\n\n" +
            ex.Message,
            "Load Error",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
    }
    finally
    {
        if (pointsSheet != null)
            Marshal.ReleaseComObject(pointsSheet);

        if (settingsSheet != null)
            Marshal.ReleaseComObject(settingsSheet);

        if (workbook != null)
        {
            workbook.Close(false);
            Marshal.ReleaseComObject(workbook);
        }

        if (excelApp != null)
        {
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
        }

        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}
        private void ExcelUpload_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx"
            };

            if (dlg.ShowDialog() == true)
            {
                LoadFromExcel(dlg.FileName);
            }
        }

        private void BtnInterpolate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Txtinterpol.Text = Txtinterpol.Text.Trim();
                if (!points_set)
                {
                    MessageBox.Show("Points are not finalized yet.");
                    return;
                }


                if (string.IsNullOrWhiteSpace(Txtinterpol.Text))
                {
                    MessageBox.Show("Enter the number of interpolation points.");
                    return;
                }

                if (!int.TryParse(Txtinterpol.Text, out int resolution))
                {
                    MessageBox.Show("Interpolation points must be a whole number.");
                    return;
                }

                if (resolution < 2 || resolution>100000)
                {
                    MessageBox.Show(
                        "Interpolation points must be between 2 and 100000.");
                    return;
                }

           

                List<PlotPoint> data = Points
                    .Skip(CALIBRATION_COUNT)
                    .Where(p => p.Type == PointType.Data)
                    .ToList();

                if (data.Count < 2)
                {
                    MessageBox.Show("Not enough data points.");
                    return;
                }


                data = data
                    .OrderBy(p => double.Parse(p.X))
                    .ToList();

                List<double> xs = data.Select(p => double.Parse(p.X)).ToList();
                List<double> ys = data.Select(p => double.Parse(p.Y)).ToList();


                List<PlotPoint> new_rows = new List<PlotPoint>();

                for (int i = 0; i < resolution; i++)
                {
                    double t = (double)i / (resolution - 1);

                    double x = CatmullRom(xs, t);
                    double y = CatmullRom(ys, t);

                    double px = InverseX(x);
                    double py = InverseY(y);

                    var p = new PlotPoint
                    {
                        Name = $"C{i + 1}",

                        X = x.ToString("0.######"),
                        Y = y.ToString("0.######"),

                        XP = px.ToString("0"),
                        YP = py.ToString("0"),

                        Type = PointType.Data
                    };

                    new_rows.Add(p);
                    Points.Add(p);

                    DrawDot(new Point(px, py), Brushes.LimeGreen, 2.0);
                }

                number_of_points = Points.Count;
                points_set = true;

                TxtStatus.Text =
                    $"Interpolation complete ({resolution} curve points).";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Interpolation error:\n" + ex.Message);
            }

            BtnInterpolate.IsEnabled = false;
            BtnUndo.IsEnabled = false;
            BtnClearAll.IsEnabled = false;
            BtnDeletePoint.IsEnabled = false;
            Txtinterpol.IsEnabled = false;
        }



        private double CatmullRom(List<double> v, double t)
        {
            int n = v.Count;
            if (n < 2)
                return v.First();

            double scaledT = t * (n - 1);
            int i1 = (int)scaledT;

            int i0 = Math.Max(i1 - 1, 0);
            int i2 = Math.Min(i1 + 1, n - 1);
            int i3 = Math.Min(i1 + 2, n - 1);

            double lt = scaledT - i1;

            double p0 = v[i0];
            double p1 = v[i1];
            double p2 = v[i2];
            double p3 = v[i3];

            return 0.5 * (
                (2 * p1) +
                (-p0 + p2) * lt +
                (2 * p0 - 5 * p1 + 4 * p2 - p3) * lt * lt +
                (-p0 + 3 * p1 - 3 * p2 + p3) * lt * lt * lt
            );
        }



        private double InverseX(double x)
        {
            double x0p = _origin.X;
            double x1p = _x2.X;

            double x0v = double.Parse(TxtX1.Text);
            double x1v = double.Parse(TxtX2.Text);

            double t;

            if (_logX)
            {
                double lx = Math.Log10(x);
                double l0 = Math.Log10(x0v);
                double l1 = Math.Log10(x1v);

                t = (lx - l0) / (l1 - l0);
            }
            else
            {
                t = (x - x0v) / (x1v - x0v);
            }

            return x0p + t * (x1p - x0p);
        }

        private double InverseY(double y)
        {
            double y0p = _origin.Y;
            double y1p = _y2.Y;

            double y0v = double.Parse(TxtY1.Text);
            double y1v = double.Parse(TxtY2.Text);

            double t;

            if (_logY)
            {
                double ly = Math.Log10(y);
                double l0 = Math.Log10(y0v);
                double l1 = Math.Log10(y1v);

                t = (ly - l0) / (l1 - l0);
            }
            else
            {
                t = (y - y0v) / (y1v - y0v);
            }

            return y0p + t * (y1p - y0p);
        }


        private void DrawDot(Point p, Brush color, double size = 7)
        {
            Ellipse el = new Ellipse
            {
                Width = size,
                Height = size,
                Fill = color
            };

            Canvas.SetLeft(el, p.X - size / 2);
            Canvas.SetTop(el, p.Y - size / 2);

            OverlayCanvas.Children.Add(el);
        }




    }
}
