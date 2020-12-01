#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.UI.Selection;
using System.IO;
#endregion // Namespaces

namespace ExportCncFab
{
  #region CmdDxf - Export to DXF
  [Transaction( TransactionMode.Manual )]
  public class CmdDxf : IExternalCommand
  {
    #region WallPartSelectionFilter
    class WallPartSelectionFilter : ISelectionFilter
    {
      public bool AllowElement( Element e )
      {
        bool rc = false;

        Part part = e as Part;

        if( null != part )
        {
          ICollection<ElementId> cids
            = part.GetSourceElementOriginalCategoryIds();

          if( 1 == cids.Count )
          {
            ElementId cid = cids.First<ElementId>();

            BuiltInCategory bic
              = (BuiltInCategory) cid.IntegerValue;

            rc = ( BuiltInCategory.OST_Walls == bic );
          }
        }
        return rc;
      }

      public bool AllowReference( Reference r, XYZ p )
      {
        return true;
      }
    }
    #endregion // WallPartSelectionFilter

    /// <summary>
    /// Default folder for test files, RVT project
    /// to load and output folder for DXF and SAT.
    /// </summary>
    static string _folder
      = "Z:\\a\\src\\revit\\export_cnc_fab\\test";

    static void InfoMsg( string msg )
    {
      Debug.Print( msg );
      TaskDialog.Show( App.Caption, msg );
    }

    static void ErrorMsg( string msg )
    {
      Debug.Print( msg );
      TaskDialog dlg = new TaskDialog( App.Caption );
      dlg.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
      dlg.MainInstruction = msg;
      dlg.Show();
    }

    static void OnDialogBoxShowing(
      object sender,
      DialogBoxShowingEventArgs e )
    {
      TaskDialogShowingEventArgs e2
        = e as TaskDialogShowingEventArgs;

      if( null != e2 && e2.DialogId.Equals(
        "TaskDialog_Really_Print_Or_Export_Temp_View_Modes" ) )
      {
        int cmdLink
          = (int) TaskDialogResult.CommandLink2;

        e.OverrideResult( cmdLink );
      }
    }

    public static Result Execute2(
      ExternalCommandData commandData,
      bool exportToSatFormat )
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Application app = uiapp.Application;
      Document doc = uidoc.Document;

      if( null == doc )
      {
        ErrorMsg( "Please run this command in a valid"
          + " Revit project document." );
        return Result.Failed;
      }

      View view = doc.ActiveView;

      if( null == view || !( view is View3D ) )
      {
        ErrorMsg( "Please run this command in a valid"
          + " 3D view." );
        return Result.Failed;
      }

      if( PartsVisibility.ShowPartsOnly
        != view.PartsVisibility )
      {
        ErrorMsg( "Please run this command in a view"
          + " displaying parts and not source elements." );
        return Result.Failed;
      }

      // Define the list of views to export, 
      // including only the current 3D view

      List<ElementId> viewIds = new List<ElementId>( 1 );

      viewIds.Add( view.Id );

      // Iterate over all pre-selected parts

      List<ElementId> ids = null;

      Selection sel = uidoc.Selection;
      ICollection<ElementId> selIds = sel.GetElementIds(); // 2015

      //if( 0 < sel.Elements.Size ) // 2014
        
      if( 0 < selIds.Count ) // 2015
      {
        //foreach( Element e in sel.Elements ) // 2014

        foreach( ElementId id in selIds ) // 2015
        {
          Element e = doc.GetElement( id );

          if( !( e is Part ) )
          {
            ErrorMsg( "Please pre-select only gyp wallboard"
              + " parts before running this command." );
            return Result.Failed;
          }

          Part part = e as Part;

          ICollection<LinkElementId> lids
            = part.GetSourceElementIds();

          if( 1 != lids.Count )
          {
            ErrorMsg( "Gyp wallboard part has multiple"
              + " source elements." );
            return Result.Failed;
          }

          LinkElementId lid = lids.First<LinkElementId>();
          ElementId hostId = lid.HostElementId;
          ElementId linkedId = lid.LinkedElementId;
          ElementId parentId = hostId;
          ElementId partId = e.Id;

          // Determine parent category

          Element parent = doc.GetElement( parentId );
          Category cat = parent.Category;

          ICollection<ElementId> cids
            = part.GetSourceElementOriginalCategoryIds();

          if( 1 != cids.Count )
          {
            ErrorMsg( "Gyp wallboard part has multiple"
              + " source element categories." );
            return Result.Failed;
          }

          ElementId cid = cids.First<ElementId>();

          //cat = doc.GetElement( id ) as Category;

          // Expected parent category is OST_Walls

          BuiltInCategory bic
            = (BuiltInCategory) cid.IntegerValue;

          if( BuiltInCategory.OST_Walls != bic )
          {
            ErrorMsg( "Please pre-select only "
              + " gyp wallboard parts." );

            return Result.Failed;
          }

          if( null == ids )
          {
            ids = new List<ElementId>( 1 );
          }

          ids.Add( partId );
        }

        if( null == ids )
        {
          ErrorMsg( "Please pre-select only gyp wallboard"
            + " parts before running this command." );
          return Result.Failed;
        }
      }

      // If no parts were pre-selected, 
      // prompt for post-selection

      if( null == ids )
      {
        IList<Reference> refs = null;

        try
        {
          refs = sel.PickObjects( ObjectType.Element,
            new WallPartSelectionFilter(),
            "Please select wall parts." );
        }
        catch( Autodesk.Revit.Exceptions
          .OperationCanceledException )
        {
          return Result.Cancelled;
        }
        ids = new List<ElementId>(
          refs.Select<Reference, ElementId>(
            r => r.ElementId ) );
      }

      if( 0 == ids.Count )
      {
        ErrorMsg( "No valid parts selected." );

        return Result.Failed;
      }

      // Check for shared parameters 
      // to record export history

      ExportParameters exportParameters
        = new ExportParameters(
          doc.GetElement( ids[0] ) );

      if( !exportParameters.IsValid )
      {
        ErrorMsg( "Please initialise the CNC fabrication "
          + "export history shared parameters before "
          + "launching this command." );

        return Result.Failed;
      }

      if( !Util.BrowseDirectory( ref _folder, true ) )
      {
        return Result.Cancelled;
      }

      try
      {
        // Register event handler for 
        // "TaskDialog_Really_Print_Or_Export_Temp_View_Modes" 
        // dialogue

        uiapp.DialogBoxShowing
          += new EventHandler<DialogBoxShowingEventArgs>(
            OnDialogBoxShowing );

        object opt = exportToSatFormat
          ? (object) new SATExportOptions()
          : (object) new DXFExportOptions();

        //opt.FileVersion = ACADVersion.R2000;

        string filename, sort_mark;

        using( TransactionGroup txg = new TransactionGroup( doc ) )
        {
          txg.Start( "Export Wall Parts" );

          foreach( ElementId id in ids )
          {
            Element e = doc.GetElement( id );

            Debug.Assert( e is Part,
              "expected parts only" );

            Part part = e as Part;

            ICollection<LinkElementId> lids
              = part.GetSourceElementIds();

            Debug.Assert( 1 == lids.Count,
              "unexpected multiple part source elements." );

            LinkElementId lid = lids.First<LinkElementId>();
            ElementId hostId = lid.HostElementId;
            ElementId linkedId = lid.LinkedElementId;
            ElementId parentId = hostId;
            ElementId partId = e.Id;

            sort_mark = exportParameters.GetSortMarkFor( e );

            filename = (null == sort_mark)
              ? string.Empty
              : sort_mark + '_';

            filename += string.Format( "{0}_{1}",
              parentId, partId );

            Element host = doc.GetElement( hostId );

            Debug.Assert( null != host, "expected to be able to access host element" );
            //Debug.Assert( ( host is Wall ), "expected host element to be a wall" ); 
            Debug.Assert( ( host is Wall ) || ( host is Part ), "expected host element to be a wall or part" );
            Debug.Assert( null != host.Category, "expected host element to have a valid category" );
            //Debug.Assert( host.Category.Id.IntegerValue.Equals( (int) BuiltInCategory.OST_Walls ), "expected host element to have wall category" );
            Debug.Assert( host.Category.Id.IntegerValue.Equals( (int) BuiltInCategory.OST_Walls ) || host.Category.Id.IntegerValue.Equals( (int) BuiltInCategory.OST_Parts ), "expected host element to have wall or part category" );
            Debug.Assert( ElementId.InvalidElementId != host.LevelId, "expected host element to have a valid level id" );

            if( ElementId.InvalidElementId != host.LevelId )
            {
              Element level = doc.GetElement( host.LevelId );

              filename = level.Name.Replace( ' ', '_' )
                + "_" + filename;
            }

            if( view.IsTemporaryHideIsolateActive() )
            {
              using( Transaction tx = new Transaction( doc ) )
              {
                tx.Start( "Disable Temporary Isolate" );

                view.DisableTemporaryViewMode(
                  TemporaryViewMode.TemporaryHideIsolate );

                tx.Commit();
              }

              Debug.Assert( !view.IsTemporaryHideIsolateActive(),
                "expected to turn off temporary hide/isolate" );
            }

            using( Transaction tx = new Transaction( doc ) )
            {
              tx.Start( "Export Wall Part "
                + partId.ToString() );

              // This call requires a transaction.

              view.IsolateElementTemporary( partId );

              //List<ElementId> unhideIds = new List<ElementId>( 1 );
              //unhideIds.Add( partId );
              //view.UnhideElements( unhideIds );

              //doc.Regenerate(); // this is insufficient

              tx.Commit();
            }

            if( exportToSatFormat )
            {
              //ViewSet viewSet = new ViewSet();
              //
              //foreach( ElementId vid in viewIds )
              //{
              //  viewSet.Insert( doc.GetElement( vid ) 
              //    as View );
              //}
              //
              //doc.Export( _folder, filename, viewSet, 
              //  (SATExportOptions) opt ); // 2013

              doc.Export( _folder, filename, viewIds,
                (SATExportOptions) opt ); // 2014
            }
            else
            {
              doc.Export( _folder, filename, viewIds,
                (DXFExportOptions) opt );
            }

            // Update CNC fabrication 
            // export shared parameters -- oops, 
            // cannot do this immediately, since 
            // this transaction group will be
            // rolled back ... just save the 
            // element id and do it later
            // searately.

            //exportParameters.UpdateExportHistory( e );
            exportParameters.Add( e.Id );
          }

          // We do not commit the transaction group, 
          // because no modifications should be saved.
          // The transaction group is only created and 
          // started to encapsulate the transactions 
          // required by the IsolateElementTemporary 
          // method. Since the transaction group is not 
          // committed, the changes are automatically 
          // discarded.

          //txg.Commit();
        }
      }
      finally
      {
        uiapp.DialogBoxShowing
          -= new EventHandler<DialogBoxShowingEventArgs>(
            OnDialogBoxShowing );
      }

      using( Transaction tx = new Transaction( doc ) )
      {
        tx.Start( "Update CNC Fabrication Export "
          + "History Shared Parameters" );

        exportParameters.UpdateExportHistory();

        tx.Commit();
      }
      return Result.Succeeded;
    }

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      return Execute2( commandData, false );
    }
  }
  #endregion // CmdDxf - Export to DXF

  #region CmdSat - Export to SAT
  [Transaction( TransactionMode.Manual )]
  public class CmdSat : IExternalCommand
  {
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      return CmdDxf.Execute2( commandData, true );
    }
  }
  #endregion // CmdSat - Export to SAT

  #region CmdCreateSharedParameters
  [Transaction( TransactionMode.Manual )]
  public class CmdCreateSharedParameters
    : IExternalCommand
  {
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      UIApplication uiapp = commandData.Application;
      Document doc = uiapp.ActiveUIDocument.Document;

      ExportParameters.Create( doc );

      return Result.Succeeded;
    }
  }
  #endregion // CmdCreateSharedParameters
}
