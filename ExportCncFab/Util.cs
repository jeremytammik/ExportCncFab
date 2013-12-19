#region Namespaces
using System;
using System.Windows.Forms;
#endregion // Namespaces

namespace ExportCncFab
{
  class Util
  {
    /// <summary>
    /// Prompt user to interactively select a directory.
    /// </summary>
    /// <param name="path">Input initial path and return selected value</param>
    /// <param name="allowCreate">Enable creation of new folder</param>
    /// <returns>True on successful selection</returns>
    public static bool BrowseDirectory(
      ref string path,
      bool allowCreate )
    {
      FolderBrowserDialog browseDlg
        = new FolderBrowserDialog();

      browseDlg.SelectedPath = path;
      browseDlg.ShowNewFolderButton = allowCreate;

      bool rc = ( DialogResult.OK
        == browseDlg.ShowDialog() );

      if( rc )
      {
        path = browseDlg.SelectedPath;
      }

      return rc;
    }
  }
}
