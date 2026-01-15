# -*- coding: utf-8 -*-
"""Export active sheet to DWFx with all quality parameter combinations.

This script exports the active sheet view multiple times, exploring all
combinations of DWFx quality parameters. Output filenames include parameter
values for comparison of file size and quality.
"""

__title__ = "Test DWFx Preferences"

import os
import datetime
import time
from pyrevit import script, forms
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    DWFXExportOptions,
    DWFImageFormat,
    DWFImageQuality,
    ViewSet,
    ViewSheet,
    Transaction,
    ColorDepthType,
    RasterQualityType,
    PrintRange,
    PageOrientationType,
    ZoomType,
    PrintSetting
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Events import DialogBoxShowingEventArgs

# Event handler to suppress rasterization warning dialog
def on_dialog_showing(sender, args):
    """Suppress the rasterization warning dialog by auto-closing it."""
    try:
        # Check if this is the rasterization warning dialog
        if hasattr(args, 'DialogId'):
            dialog_id = args.DialogId
            # Suppress printing/rasterization related dialogs
            if 'Printing' in str(dialog_id) or 'Raster' in str(dialog_id) or 'Shaded' in str(dialog_id):
                args.OverrideResult(1)  # Close/OK button
                return
        # Also check message content
        if hasattr(args, 'Message'):
            msg = str(args.Message).lower()
            if 'raster' in msg or 'shaded' in msg or 'vector' in msg:
                args.OverrideResult(1)
    except:
        pass

def export_sheet_dwfx(doc, sheet, output_folder, filename, img_format, img_quality, color_depth, raster_quality):
    """Export a sheet to DWFx with specified quality parameters.
    
    Args:
        doc: Revit Document
        sheet: ViewSheet to export
        output_folder: Output directory path
        filename: Output filename (without extension)
        img_format: DWFImageFormat enum value
        img_quality: DWFImageQuality enum value (ignored for Lossless)
        color_depth: ColorDepthType enum value
        raster_quality: RasterQualityType enum value
    
    Returns:
        bool: True if export succeeded, False otherwise
    """
    print_manager = doc.PrintManager
    views = ViewSet()
    views.Insert(sheet)
    
    try:
        # Transaction: Set and save print parameters
        t = Transaction(doc, "pyAreaDWFxPrintSetup")
        t.Start()
        
        # Configure print setup parameters
        print_setup = print_manager.PrintSetup
        print_setup.CurrentPrintSetting = print_setup.InSession
        
        print_params = print_setup.InSession.PrintParameters
        print_params.ColorDepth = color_depth
        print_params.RasterQuality = raster_quality
        
        # Save the InSession settings as a named print setup
        setup_name = "_TempDWFExportSetup"
        try:
            print_setup.SaveAs(setup_name)
        except:
            try:
                print_setup.Revert()
                print_params.ColorDepth = color_depth
                print_params.RasterQuality = raster_quality
                print_setup.Save()
            except:
                pass
        
        t.Commit()
        
        # Transaction: Export DWFx
        t2 = Transaction(doc, "Export DWFx")
        t2.Start()
        
        dwfx_options = DWFXExportOptions()
        dwfx_options.ImageFormat = img_format
        dwfx_options.ImageQuality = img_quality
        
        result = doc.Export(output_folder, filename, views, dwfx_options)
        
        t2.Commit()
        
        return result
        
    except Exception as e:
        if 't' in locals() and t.HasStarted():
            t.RollBack()
        if 't2' in locals() and t2.HasStarted():
            t2.RollBack()
        raise e


def _get_temp_print_setup_id(doc, setup_name):
    for setting in DB.FilteredElementCollector(doc).OfClass(PrintSetting):
        if setting.Name == setup_name:
            return setting.Id
    return None


uiapp = __revit__
doc = uiapp.ActiveUIDocument.Document
uidoc = uiapp.ActiveUIDocument
active_view = doc.ActiveView

print_manager = doc.PrintManager
print_setup = print_manager.PrintSetup
original_print_setting = print_setup.CurrentPrintSetting
original_print_setting_id = getattr(original_print_setting, 'Id', None)
original_in_session = (original_print_setting_id is None) or (original_print_setting == print_setup.InSession)
original_color_depth = original_print_setting.PrintParameters.ColorDepth
original_raster_quality = original_print_setting.PrintParameters.RasterQuality

# Subscribe to dialog event to suppress rasterization warnings
uiapp.DialogBoxShowing += on_dialog_showing

# Check if active view is a sheet
if not isinstance(active_view, ViewSheet):
    TaskDialog.Show("Error", "Active view must be a Sheet. Please open a sheet and run again.")
else:
    # Get sheet info for filename
    sheet_number = active_view.SheetNumber
    sheet_name = active_view.Name
    
    # Define output folder - Downloads\YYYYMMDD-<modelFileName>-<sheetNumber>
    model_path = doc.PathName
    if model_path:
        model_name = os.path.splitext(os.path.basename(model_path))[0]
    else:
        model_name = doc.Title
    date_stamp = datetime.datetime.now().strftime("%Y%m%d")
    folder_name = "{date}-{model}-{sheet}".format(
        date=date_stamp,
        model=model_name,
        sheet=sheet_number.replace(" ", "_")
    )
    export_folder = os.path.join(os.environ['USERPROFILE'], 'Downloads', folder_name)
    if not os.path.exists(export_folder):
        os.makedirs(export_folder)
    
    # Define parameter combinations to test
    image_options = [
        (DWFImageFormat.Lossless, DWFImageQuality.High, "Lossless"),
        (DWFImageFormat.Lossy, DWFImageQuality.High, "LossyHigh"),
        (DWFImageFormat.Lossy, DWFImageQuality.Medium, "LossyMedium"),
        (DWFImageFormat.Lossy, DWFImageQuality.Low, "LossyLow")
    ]
    
    color_depths = [
        (ColorDepthType.Color, "Color")
    ]
    
    raster_qualities = [
        (RasterQualityType.Low, "RasterLow"),
        (RasterQualityType.Medium, "RasterMed"),
        (RasterQualityType.High, "RasterHigh"),
        (RasterQualityType.Presentation, "RasterPresentation")
    ]
    
    total_combinations = len(image_options) * len(color_depths) * len(raster_qualities)
    export_count = 0
    failed_count = 0
    size_table = {}
    time_table = {}
    for _, _, img_option_name in image_options:
        for _, raster_quality_name in raster_qualities:
            size_table.setdefault(raster_quality_name, {})
            size_table[raster_quality_name].setdefault(img_option_name, {})
            time_table.setdefault(raster_quality_name, {})
            time_table[raster_quality_name].setdefault(img_option_name, {})
    
    # Iterate through all combinations
    progress_title = "Exporting DWFx combinations"
    with forms.ProgressBar(title=progress_title, max_value=total_combinations, cancellable=True) as progress:
        step = 0
        for img_format, img_quality, img_option_name in image_options:
            for color_depth, color_depth_name in color_depths:
                for raster_quality, raster_quality_name in raster_qualities:
                    if progress.cancelled:
                        raise Exception("Export cancelled by user.")

                    step += 1
                    progress.update_progress(step, total_combinations)
                    # Build filename with parameters
                    filename = "Img-{img}__Raster-{raster}".format(
                        img=img_option_name,
                        raster=raster_quality_name
                    )
                    
                    try:
                        start_time = time.time()
                        result = export_sheet_dwfx(
                            doc=doc,
                            sheet=active_view,
                            output_folder=export_folder,
                            filename=filename,
                            img_format=img_format,
                            img_quality=img_quality,
                            color_depth=color_depth,
                            raster_quality=raster_quality
                        )
                        elapsed_sec = time.time() - start_time
                        
                        if result:
                            export_count += 1
                            size_mb = None
                            output_path = os.path.join(export_folder, filename + ".dwfx")
                            if os.path.exists(output_path):
                                size_mb = os.path.getsize(output_path) / (1024.0 * 1024.0)
                                size_table[raster_quality_name][img_option_name][color_depth_name] = size_mb
                                time_table[raster_quality_name][img_option_name][color_depth_name] = elapsed_sec
                            if size_mb is not None:
                                print("Exported: {}.dwfx - {:.2f} MB @ {:.2f}s".format(filename, size_mb, elapsed_sec))
                            else:
                                print("Exported: {}.dwfx - {:.2f}s".format(filename, elapsed_sec))
                        else:
                            failed_count += 1
                            print("Failed: {}".format(filename))
                    except Exception as e:
                        failed_count += 1
                        print("Error exporting {}: {}".format(filename, str(e)))
    
    # Summary
    summary = """DWFx Export Test Complete!

Sheet: {} - {}
Output Folder: {}

Total Combinations: {}
Successful Exports: {}
Failed Exports: {}

Parameter Legend:
- Img: Image Format/Quality (Lossless/LossyHigh/LossyMedium/LossyLow)
- Color: Color Depth (Color/Grayscale)
- Raster: Raster Quality (RasterHigh/RasterMed/RasterLow/RasterPresentation)
""".format(sheet_number, sheet_name, export_folder, total_combinations, export_count, failed_count)
    
    # Restore original print setup
    try:
        t_restore = Transaction(doc, "pyAreaDWFxPrintSetup(Revert)")
        t_restore.Start()

        print_manager = doc.PrintManager
        print_setup = print_manager.PrintSetup
        if original_in_session:
            print_setup.CurrentPrintSetting = print_setup.InSession
            restore_params = print_setup.InSession.PrintParameters
            restore_params.ColorDepth = original_color_depth
            restore_params.RasterQuality = original_raster_quality
        else:
            original_setting = doc.GetElement(original_print_setting_id)
            if original_setting:
                print_setup.CurrentPrintSetting = original_setting

        temp_setting_id = _get_temp_print_setup_id(doc, "_TempDWFExportSetup")
        if temp_setting_id:
            doc.Delete(temp_setting_id)

        t_restore.Commit()
    except Exception as e:
        if 't_restore' in locals() and t_restore.HasStarted():
            t_restore.RollBack()
        print("WARNING: Failed to restore print setup: {}".format(str(e)))

    # Unsubscribe from dialog event
    uiapp.DialogBoxShowing -= on_dialog_showing
    
    print(summary)

    # File size table (MB) using pyRevit output
    output = script.get_output()
    table_rows = []
    for _, raster_quality_name in raster_qualities:
        row = [raster_quality_name]
        for _, _, img_option_name in image_options:
            color_size = size_table.get(raster_quality_name, {}).get(img_option_name, {}).get("Color")
            elapsed_sec = time_table.get(raster_quality_name, {}).get(img_option_name, {}).get("Color")
            if color_size is not None and elapsed_sec is not None:
                cell = "{:.2f} MB @ {:.2f}s".format(color_size, elapsed_sec)
            else:
                cell = "-"
            row.append(cell)
        table_rows.append(row)

    output.print_table(
        table_data=table_rows,
        title="File Size Table (MB)",
        columns=["Raster"] + [name for _, _, name in image_options],
        formats=[""] * (len(image_options) + 1)
    )

    # Open output folder in Windows Explorer
    try:
        os.startfile(export_folder)
    except Exception as e:
        print("WARNING: Failed to open output folder: {}".format(str(e)))
