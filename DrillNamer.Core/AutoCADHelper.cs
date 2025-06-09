using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System;

global using static Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace DrillNamer.Core;

/// <summary>
/// Utility helpers for AutoCAD related operations.
/// </summary>
public static class AutoCADHelper
{
    /// <summary>
    /// Starts a transaction for the active document.
    /// </summary>
    public static Transaction StartTransaction()
    {
        Document doc = DocumentManager.MdiActiveDocument;
        return doc.Database.TransactionManager.StartTransaction();
    }
}

/// <summary>
/// Provides layer lock management utilities.
/// </summary>
public static class LayerState
{
    /// <summary>
    /// Executes an action with the specified layer unlocked.
    /// </summary>
    public static void WithUnlocked(ObjectId layerId, Action action)
    {
        using Transaction tr = StartTransaction();
        var layerTblRec = tr.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
        if (layerTblRec == null)
        {
            action();
            tr.Commit();
            return;
        }
        bool relock = layerTblRec.IsLocked;
        if (relock)
        {
            Logging.Info($"Temporarily unlocking layer '{layerTblRec.Name}'.");
            layerTblRec.IsLocked = false;
            Logging.Debug($"Temporarily unlocked layer '{layerTblRec.Name}'");
        }
        action();
        if (relock)
        {
            layerTblRec.IsLocked = true;
        }
        tr.Commit();
    }

    internal static void WithUnlocked(ILayerLock layer, Action action)
    {
        bool relock = layer.IsLocked;
        if (relock)
        {
            layer.IsLocked = false;
        }
        action();
        if (relock)
        {
            layer.IsLocked = true;
        }
    }
}

internal interface ILayerLock
{
    bool IsLocked { get; set; }
    string Name { get; }
}
