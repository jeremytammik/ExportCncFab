#region Namespaces
using System.IO;
using System.Windows.Media.Imaging;
using System.Reflection;
using Autodesk.Revit.UI;
#endregion

namespace ExportCncFab
{
  class App : IExternalApplication
  {
    public const string Caption
      = "Export to CNC Fabrication";

    static string _namespace_prefix
      = typeof( App ).Namespace + ".";

    const string _name = "DXF";
    const string _name2 = "SAT";
    const string _name3 = "Create\r\nShared\r\nParameters";

    const string _class_name = "CmdDxf";
    const string _class_name2 = "CmdSat";
    const string _class_name3 = "CmdCreateSharedParameters";

    const string _tooltip_format
      = "Export to CNC Fabrication in {0} format";

    const string _tooltip_long_description_format
      = "Export Revit parts to CNC Fabrication in {0} format.";

    /// <summary>
    /// Load a new icon bitmap from embedded resources.
    /// For the BitmapImage, make sure you reference 
    /// WindowsBase and PresentationCore, and import 
    /// the System.Windows.Media.Imaging namespace. 
    /// </summary>
    BitmapImage NewBitmapImage(
      Assembly a,
      string imageName )
    {
      // to read from an external file:
      //return new BitmapImage( new Uri(
      //  Path.Combine( _imageFolder, imageName ) ) );

      Stream s = a.GetManifestResourceStream(
        _namespace_prefix + imageName );

      BitmapImage img = new BitmapImage();

      img.BeginInit();
      img.StreamSource = s;
      img.EndInit();

      return img;
    }

    public Result OnStartup(
      UIControlledApplication a )
    {
      Assembly exe = Assembly.GetExecutingAssembly();
      string path = exe.Location;

      // Create ribbon panel

      RibbonPanel p = a.CreateRibbonPanel( Caption );

      // Create DXF button

      PushButtonData d = new PushButtonData(
        _name, _name, path,
        _namespace_prefix + _class_name );

      d.ToolTip = string.Format( _tooltip_format, _name );
      d.Image = NewBitmapImage( exe, "cnc_icon_16x16_size.png" );
      d.LargeImage = NewBitmapImage( exe, "cnc_icon_32x32_size.png" );
      d.LongDescription = string.Format( _tooltip_long_description_format, _name );
      d.ToolTipImage = NewBitmapImage( exe, "cnc_icon_full_size.png" );

      p.AddItem( d );

      // Create SAT button

      d = new PushButtonData(
        _name2, _name2, path,
        _namespace_prefix + _class_name2 );

      d.ToolTip = string.Format( _tooltip_format, _name2 );
      d.Image = NewBitmapImage( exe, "cnc_icon_16x16_size.png" );
      d.LargeImage = NewBitmapImage( exe, "cnc_icon_32x32_size.png" );
      d.LongDescription = string.Format( _tooltip_long_description_format, _name2 );
      d.ToolTipImage = NewBitmapImage( exe, "cnc_icon_full_size.png" );

      p.AddItem( d );

      // Create shared parameters button

      d = new PushButtonData(
        _name3, _name3, path,
        _namespace_prefix + _class_name3 );

      d.ToolTip
        = "Create shared parameters for tracking export history";

      d.LongDescription
        = "Create and bind shared parameters to the "
        + "Parts category for tracking export history:\r\n\r\n"
        + " * CncFabIsExported - Boolean\r\n"
        + " * CncFabExportedFirst - Text timestamp ISO 8601\r\n"
        + " * CncFabExportedLast - Text timestamp ISO 8601";

      d.ToolTipImage = NewBitmapImage( exe,
        "cnc_icon_full_size.png" );

      p.AddItem( d );

      return Result.Succeeded;
    }

    public Result OnShutdown( UIControlledApplication a )
    {
      return Result.Succeeded;
    }
  }
}
