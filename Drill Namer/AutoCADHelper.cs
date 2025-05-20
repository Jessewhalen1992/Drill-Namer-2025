using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;

namespace Drill_Namer
{
    public static class AutoCADHelper
    {
        /// <summary>
        /// Start a transaction in the current AutoCAD document.
        /// </summary>
        public static Transaction StartTransaction()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            return doc.Database.TransactionManager.StartTransaction();
        }

        /// <summary>
        /// Prompts the user to select an insertion point. Returns true if a point was chosen.
        /// </summary>
        /// <param name="promptMessage">Message to display.</param>
        /// <param name="insertionPoint">The point selected by the user.</param>
        /// <returns>True if user picked a point; false if canceled.</returns>
        public static bool GetInsertionPoint(string promptMessage, out Point3d insertionPoint)
        {
            insertionPoint = Point3d.Origin;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return false;

            Editor ed = doc.Editor;
            PromptPointOptions opts = new PromptPointOptions("\n" + promptMessage);
            PromptPointResult res = ed.GetPoint(opts);

            if (res.Status == PromptStatus.OK)
            {
                insertionPoint = res.Value;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Insert a block reference at a specified point, with given attributes and scale,
        /// on layer "CG-NOTES". If missing, tries to load the block from 
        /// C:\AUTOCAD-SETUP\Lisp_2000\Drill Properties\[blockName].dwg
        /// </summary>
        public static ObjectId InsertBlock(
            string blockName,
            Point3d insertionPoint,
            Dictionary<string, string> attributes,
            double scale = 1.0)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId brId = ObjectId.Null;

            // Make sure we lock the doc for changes
            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 1) Ensure the block is loaded into this DWG
                    if (!EnsureBlockIsLoaded(db, blockName))
                    {
                        // Could not find or load it from the external .dwg
                        return ObjectId.Null;
                    }

                    // 2) Get the block table record for that block
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    if (!bt.Has(blockName))
                    {
                        // Even after loading, not found
                        return ObjectId.Null;
                    }

                    ObjectId btrId = bt[blockName];
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                    // 3) Open ModelSpace (or whichever space you need)
                    //    For standard insertion, usually model space:
                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    // 4) Create the block reference
                    BlockReference br = new BlockReference(insertionPoint, btrId);

                    // Apply the scale
                    br.ScaleFactors = new Scale3d(scale, scale, scale);

                    // Set layer to "CG-NOTES" 
                    // (make sure the layer exists; if not, create it or set to another fallback)
                    if (!CreateLayerIfMissing(db, "CG-NOTES"))
                    {
                        // If for any reason we fail to create/find the layer, fallback to "0"
                        br.Layer = "0";
                    }
                    else
                    {
                        br.Layer = "CG-NOTES";
                    }

                    // 5) Append the block reference to model space
                    ms.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);

                    // 6) Add the attribute references
                    foreach (ObjectId id in btr)
                    {
                        DBObject obj = tr.GetObject(id, OpenMode.ForRead);
                        if (obj is AttributeDefinition attDef)
                        {
                            // Create an AttributeReference from the def
                            AttributeReference attRef = new AttributeReference();
                            attRef.SetAttributeFromBlock(attDef, br.BlockTransform);

                            // If user provided an attribute matching this Tag
                            if (attributes != null && attributes.ContainsKey(attDef.Tag))
                            {
                                attRef.TextString = attributes[attDef.Tag];
                            }
                            else
                            {
                                // If not provided, just keep the attDef's default
                                attRef.TextString = attDef.TextString;
                            }

                            // set the same layer as the block reference
                            attRef.Layer = br.Layer;

                            // Add it to the blockRef
                            br.AttributeCollection.AppendAttribute(attRef);
                            tr.AddNewlyCreatedDBObject(attRef, true);
                        }
                    }

                    tr.Commit();
                    brId = br.ObjectId;
                }
            }

            return brId;
        }

        /// <summary>
        /// Attempt to ensure the block definition is present in the current DB.
        /// If missing, tries to load it from "C:\AUTOCAD-SETUP\Lisp_2000\Drill Properties\blockName.dwg".
        /// </summary>
        private static bool EnsureBlockIsLoaded(Database db, string blockName)
        {
            bool exists = false;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                exists = bt.Has(blockName);
                tr.Commit();
            }
            if (exists) return true;

            // If not found, try to import from external .dwg
            string baseFolder = @"C:\AUTOCAD-SETUP\Lisp_2000\Drill Properties";
            string dwgPath = Path.Combine(baseFolder, blockName + ".dwg");
            if (!File.Exists(dwgPath))
            {
                // can't find the external block .dwg
                return false;
            }

            return (ImportBlockDefinition(db, dwgPath, blockName) != ObjectId.Null);
        }

        /// <summary>
        /// Imports a DWG file (which is effectively a single block definition) into the current DB
        /// with DuplicateRecordCloning.Replace (so we re-use the name).
        /// </summary>
        private static ObjectId ImportBlockDefinition(Database destDb, string sourceDwgPath, string blockName)
        {
            if (destDb == null || string.IsNullOrEmpty(sourceDwgPath) || !File.Exists(sourceDwgPath))
                return ObjectId.Null;

            ObjectId result = ObjectId.Null;
            using (Database tempDb = new Database(false, true))
            {
                try
                {
                    // read the .dwg containing the block
                    tempDb.ReadDwgFile(sourceDwgPath, FileShare.Read, true, "");
                    using (Transaction tr = tempDb.TransactionManager.StartTransaction())
                    {
                        BlockTable sourceBT = (BlockTable)tr.GetObject(tempDb.BlockTableId, OpenMode.ForRead);
                        if (!sourceBT.Has(blockName))
                        {
                            // blockName does not exist in that DWG
                            return ObjectId.Null;
                        }

                        // get the block's btr
                        ObjectId sourceBtrId = sourceBT[blockName];
                        ObjectIdCollection idsToClone = new ObjectIdCollection();
                        idsToClone.Add(sourceBtrId);

                        tr.Commit();

                        // now clone it into destDb
                        IdMapping mapping = new IdMapping();
                        destDb.WblockCloneObjects(idsToClone, destDb.BlockTableId, mapping,
                            DuplicateRecordCloning.Replace, false);
                    }

                    // verify we have it
                    using (Transaction destTr = destDb.TransactionManager.StartTransaction())
                    {
                        BlockTable destBT = (BlockTable)destTr.GetObject(destDb.BlockTableId, OpenMode.ForRead);
                        if (destBT.Has(blockName))
                        {
                            result = destBT[blockName];
                        }
                        destTr.Commit();
                    }
                }
                catch
                {
                    // do nothing
                }
            }

            return result;
        }

        /// <summary>
        /// Creates or verifies that layerName exists. Returns true if found or created successfully.
        /// </summary>
        private static bool CreateLayerIfMissing(Database db, string layerName)
        {
            bool success = false;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (lt.Has(layerName))
                {
                    success = true;
                }
                else
                {
                    // create the layer
                    lt.UpgradeOpen();
                    LayerTableRecord ltr = new LayerTableRecord();
                    ltr.Name = layerName;
                    // optionally set color, line type, etc.
                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);

                    success = true;
                }
                tr.Commit();
            }
            return success;
        }
    }
}
