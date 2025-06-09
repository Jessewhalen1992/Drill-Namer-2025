using NUnit.Framework;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System.Reflection;
using Drill_Namer;

namespace Drill_Namer.Tests
{
    [TestFixture]
    public class LayerUnlockTests
    {
        [Test]
        public void UpdateAttributes_RelocksLayerAndUpdatesValue()
        {
            const string layerName = "LOCKED";
            const string oldVal = "OLD";
            const string newVal = "NEW";

            using (var db = new Database(true, true))
            {
                HostApplicationServices.WorkingDatabase = db;
                ObjectId layerId;
                ObjectId btrId;
                ObjectId brId;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);
                    var ltr = new LayerTableRecord { Name = layerName, IsLocked = true };
                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                    layerId = ltr.ObjectId;

                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                    var btr = new BlockTableRecord { Name = "TESTBLK" };
                    bt.Add(btr);
                    tr.AddNewlyCreatedDBObject(btr, true);
                    btrId = btr.ObjectId;

                    var attDef = new AttributeDefinition(Point3d.Origin, oldVal, "TAG1", "", 0)
                    {
                        LayerId = layerId
                    };
                    btr.AppendEntity(attDef);
                    tr.AddNewlyCreatedDBObject(attDef, true);

                    tr.Commit();
                }

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    var br = new BlockReference(Point3d.Origin, btrId) { LayerId = layerId };
                    ms.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    foreach (ObjectId id in btr)
                    {
                        if (tr.GetObject(id, OpenMode.ForRead) is AttributeDefinition ad)
                        {
                            var ar = new AttributeReference();
                            ar.SetAttributeFromBlock(ad, br.BlockTransform);
                            ar.TextString = oldVal;
                            ar.LayerId = layerId;
                            br.AttributeCollection.AppendAttribute(ar);
                            tr.AddNewlyCreatedDBObject(ar, true);
                        }
                    }
                    tr.Commit();
                    brId = br.ObjectId;
                }

                var form = new FindReplaceForm();
                var mi = typeof(FindReplaceForm).GetMethod("UpdateAttributesWithMatchingValue", BindingFlags.Instance | BindingFlags.NonPublic);
                mi.Invoke(form, new object[] { oldVal, newVal });

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                    Assert.IsTrue(ltr.IsLocked, "Layer should be relocked");
                    var br = (BlockReference)tr.GetObject(brId, OpenMode.ForRead);
                    foreach (ObjectId id in br.AttributeCollection)
                    {
                        var ar = (AttributeReference)tr.GetObject(id, OpenMode.ForRead);
                        Assert.AreEqual(newVal, ar.TextString.Trim());
                    }
                }
            }
        }
    }
}
