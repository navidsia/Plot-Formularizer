# Plot Formularizer

**Turn a plot image into numerical data.** Plot Formularizer is a Windows desktop application for manually digitizing curves from images of graphs. Calibrate the axes, click points along a plotted curve, optionally generate a smoother and denser series with interpolation, and export the results to an Excel workbook.

Built with **C#**, **WPF**, and **.NET Framework 4.8**.

## Features

- **Manual plot digitization:** load an image and select data points directly on the graph.
- **Three-point axis calibration:** identify the X-axis reference, origin, and Y-axis reference to convert image pixels into plot coordinates.
- **Linear and logarithmic axes:** configure the X and Y axes independently.
- **Precision magnifier:** inspect an enlarged region around the cursor with crosshairs while placing points.
- **Point management:** review image coordinates and calculated values, undo clicks, clear points, or delete a selected point before finalizing the data.
- **Optional curve interpolation:** generate a user-specified number of points using Catmull–Rom cubic interpolation.
- **Excel export:** save the selected and interpolated points together with a copy of the source image and the axis-scale flags.

## Requirements

- Windows
- [.NET Framework 4.8](https://dotnet.microsoft.com/en-us/download/dotnet-framework/net48)
- A locally installed version of **Microsoft Excel** for `.xlsx` export (the application uses Excel COM automation)

To build from source, use Visual Studio with .NET desktop development support and the .NET Framework 4.8 targeting pack. The solution uses the `Microsoft.Office.Interop.Excel` NuGet package and a Microsoft Office Core COM reference, so an Office installation may also be needed to resolve the reference during a build.

## Build and run

1. Clone the repository:

   ```bash
   git clone https://github.com/navidsia/Plot-Formularizer.git
   ```

2. Open `Plot Formularizer.sln` in Visual Studio on Windows.
3. Restore NuGet packages if prompted, then build the solution.
4. Run the project from Visual Studio, or launch the compiled `Plot Formularizer.exe` from the build output folder.

The repository is a Visual Studio/WPF project, not a cross-platform command-line utility.

## How to use

### 1. Load a plot image

Click **Upload Plot Image** and choose a `.png`, `.jpg`, `.jpeg`, or `.bmp` file. The image is displayed in the right-hand panel; scrollbars allow you to navigate a large image.

### 2. Enter the axis values

Specify four known values:

| Field | Meaning |
| --- | --- |
| `X1` | X value at the plot origin/reference point |
| `X2` | X value at a second point along the X axis |
| `Y1` | Y value at the origin/reference point |
| `Y2` | Y value at a second point along the Y axis |

Select **Log X axis** and/or **Log Y axis** before starting calibration if the corresponding axis is logarithmic. Logarithmic reference values must be positive. Use distinct reference values on each axis, and choose separate pixel positions for its two calibration points.

### 3. Calibrate the plot

Click **Finish Setup**, then click these locations on the image **in this exact order**:

1. **X2** — the position corresponding to the `X2` value.
2. **Origin** — the intersection/reference point corresponding to (`X1`, `Y1`).
3. **Y2** — the position corresponding to the `Y2` value.

The application marks calibration points in **blue**. Once all three have been placed, click **Set Origin Points** to lock calibration.

### 4. Digitize the curve

Click points along the plotted curve. Each click creates a **red** marker and a row in the **Digitized Points** grid.

| Column | Description |
| --- | --- |
| `Name` | `X2`, `Origin`, `Y2`, or a data point such as `P1` |
| `XP`, `YP` | Pixel coordinates relative to the loaded image |
| `X`, `Y` | Converted numerical coordinates |

While digitizing, use **Undo**, **Clear All**, or **Delete Selected** to correct mistakes. The magnifier shows a zoomed view of the region under the cursor for more precise placement.

### 5. Finalize and optionally interpolate

When you are satisfied with the selected data points, click **Set Points**. This finalizes manual point selection and enables **Interpolate**.

Interpolation is optional. To use it:

1. Select at least two curve points before finalizing.
2. Enter the **Number of Points for Interpolation** (from `2` to `100000`; default `1000`).
3. Click **Interpolate**.

The application sorts the selected data points by their calculated X values and applies **Catmull–Rom cubic interpolation** to produce a denser curve. Interpolated points are named `C1`, `C2`, … and drawn as small **green** dots. They are appended to the existing data; they do not replace the manually selected points.

**Note:** Interpolation is a one-time operation in the current interface. After it runs, the point-editing controls are disabled. Review your manually selected points before interpolating.

### 6. Export to Excel

Click **Save to Excel**. The application creates a workbook in the **same directory as the source image**, using the image's base name:

```text
my_plot.png
my_plot.xlsx
```

If a workbook with that name already exists, you can replace it, choose a numbered filename such as `my_plot_1.xlsx`, or cancel. Close the existing workbook in Excel before trying to overwrite it.

The export includes:

- `Name`, `XP`, `YP`, `X`, and `Y` columns, including calibration, manually digitized, and any interpolated points.
- `X Log` and `Y Log` settings in columns I–J.
- The original plot image embedded in the worksheet for reference.

## How coordinate conversion works

For a linear X axis, the application maps the clicked pixel coordinate `px` onto calibrated data values:

```text
X = X1 + (px - originPixelX) * (X2 - X1) / (x2PixelX - originPixelX)
```

The Y calculation follows the same principle with the corresponding Y coordinates. For logarithmic axes, interpolation is performed in `log10` space and converted back to the original scale.

The displayed values are rounded for readability. As with any image-based digitizer, the result depends on accurate calibration, image quality, and point placement.

## Project structure

```text
Plot Formularizer.sln
Plot Formularizer/
  MainWindow.xaml        # WPF interface
  MainWindow.xaml.cs     # Calibration, digitization, interpolation, Excel export
  Plot Formularizer.csproj
  packages.config        # Excel interop dependency
```

## Current limitations

- **Manual digitization:** the program does not automatically identify plotted lines or curve points.
- **Axis assumptions:** calibration uses an axis-aligned X/Y mapping; it does not correct for rotated plots, perspective distortion, or arbitrary nonlinear axes.
- **Excel dependency:** exporting requires desktop Microsoft Excel; the project does not write `.xlsx` files independently of Office.
- **Interpolation behavior:** interpolation follows increasing X values and may not suit loops, multi-valued curves, or curves with duplicate X coordinates.
- **Project reloading:** the source contains Excel-import code, but the **Upload Excel File** control is hidden and the importer expects a workbook layout different from the current export. Treat exported workbooks as results, not as reliably reloadable project files.
- **Windows-only:** this version uses WPF and .NET Framework 4.8.

## Author

Created by **Navid Siamakmanesh**. For questions or feedback, open a GitHub issue.
