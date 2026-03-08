using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NHBlockToBlock
{
    public class Commands
    {
        private const string QuantityTag = "Quantity";
        private const double CellWidth = 10000;
        private const double CellHeight = 5000;
        private const int ColumnsPerRow = 10;
        private const string IgnoredAttributePrefix = "I_";

        [CommandMethod("NHBlockToBlock")]
        public void NHBlock()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // --- Step 1: select objects, keep only block references ---
            var psr = ed.GetSelection();
            if (psr.Status != PromptStatus.OK)
                return;

            var blockIds = new List<ObjectId>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var selObj in psr.Value)
                {
                    var so = (SelectedObject)selObj;
                    if (!so.ObjectId.IsValid) continue;
                    var ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (ent != null)
                        blockIds.Add(so.ObjectId);
                }
                tr.Commit();
            }

            if (blockIds.Count == 0)
            {
                ed.WriteMessage("\nNo block references found in selection.");
                return;
            }

            ed.WriteMessage($"\n{blockIds.Count} block reference(s) selected.");

            // --- Step 2: group source blocks by effective name and collect their data ---
            // Key = effective block name, Value = list of source block data
            var groupedSources = new Dictionary<string, List<SourceBlockData>>(StringComparer.OrdinalIgnoreCase);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var id in blockIds)
                {
                    var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                    string name = GetBlockEffectiveName(tr, br);

                    var data = new SourceBlockData
                    {
                        Attributes = CollectAttributeValues(tr, br),
                        DynamicProps = CollectDynamicPropertyValues(br)
                    };

                    if (!groupedSources.ContainsKey(name))
                        groupedSources[name] = new List<SourceBlockData>();
                    groupedSources[name].Add(data);
                }
                tr.Commit();
            }

            // --- Step 3: select allowed template blocks (both source-type and -Z) ---
            ed.WriteMessage("\nSelect allowed template blocks (include both e.g. \"P-1\" and \"P-1-Z\"): ");
            var psrTemplates = ed.GetSelection();
            if (psrTemplates.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nCancelled. Aborting.");
                return;
            }

            // Key = block name, Value = first ObjectId seen for that name
            var allowedPool = new Dictionary<string, ObjectId>(StringComparer.OrdinalIgnoreCase);
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var selObj in psrTemplates.Value)
                {
                    var so = (SelectedObject)selObj;
                    if (!so.ObjectId.IsValid) continue;
                    var br = tr.GetObject(so.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (br == null) continue;
                    string name = GetBlockEffectiveName(tr, br);
                    if (!allowedPool.ContainsKey(name))
                        allowedPool[name] = so.ObjectId;
                }
                tr.Commit();
            }

            // --- Step 3b: resolve valid pairs and validate compatibility ---
            // Key = source block name, Value = ObjectId of the -Z template block reference
            var templateIds = new Dictionary<string, ObjectId>(StringComparer.OrdinalIgnoreCase);

            foreach (var name in groupedSources.Keys.OrderBy(n => n, StringComparer.OrdinalIgnoreCase))
            {
                string zName = $"{name}-Z";

                bool hasSrc = allowedPool.ContainsKey(name);
                bool hasZ   = allowedPool.ContainsKey(zName);

                if (!hasSrc || !hasZ)
                {
                    if (!hasSrc && !hasZ)
                        ed.WriteMessage($"\nNo template blocks found for \"{name}\" (missing \"{name}\" and \"{zName}\") — skipping.");
                    else if (!hasSrc)
                        ed.WriteMessage($"\nNo template block found for \"{name}\" (missing \"{name}\") — skipping.");
                    else
                        ed.WriteMessage($"\nNo template block found for \"{name}\" (missing \"{zName}\") — skipping.");
                    continue;
                }

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var brZ = (BlockReference)tr.GetObject(allowedPool[zName], OpenMode.ForRead);

                    var sampleSource = groupedSources[name][0];
                    var srcAttrNames = new HashSet<string>(sampleSource.Attributes.Keys, StringComparer.OrdinalIgnoreCase);
                    var srcPropNames = new HashSet<string>(sampleSource.DynamicProps.Keys, StringComparer.OrdinalIgnoreCase);

                    var zAttrNames = CollectAttributeNames(tr, brZ);
                    var zPropNames = CollectDynamicPropertyNames(brZ);

                    var zAttrNamesForCompare = new HashSet<string>(zAttrNames, StringComparer.OrdinalIgnoreCase);
                    zAttrNamesForCompare.Remove(QuantityTag);

                    if (!srcAttrNames.SetEquals(zAttrNamesForCompare) || !srcPropNames.SetEquals(zPropNames))
                    {
                        ed.WriteMessage($"\n\"{name}\" and \"{zName}\" do not have matching attributes/properties — skipping.");

                        var missingInZ    = srcAttrNames.Except(zAttrNamesForCompare, StringComparer.OrdinalIgnoreCase).ToList();
                        var extraInZ      = zAttrNamesForCompare.Except(srcAttrNames, StringComparer.OrdinalIgnoreCase).ToList();
                        var missingPropsInZ = srcPropNames.Except(zPropNames, StringComparer.OrdinalIgnoreCase).ToList();
                        var extraPropsInZ   = zPropNames.Except(srcPropNames, StringComparer.OrdinalIgnoreCase).ToList();

                        if (missingInZ.Count > 0)    ed.WriteMessage($"\n  Attributes in \"{name}\" but not in \"{zName}\": {string.Join(", ", missingInZ)}");
                        if (extraInZ.Count > 0)      ed.WriteMessage($"\n  Attributes in \"{zName}\" but not in \"{name}\": {string.Join(", ", extraInZ)}");
                        if (missingPropsInZ.Count > 0) ed.WriteMessage($"\n  Dynamic props in \"{name}\" but not in \"{zName}\": {string.Join(", ", missingPropsInZ)}");
                        if (extraPropsInZ.Count > 0)   ed.WriteMessage($"\n  Dynamic props in \"{zName}\" but not in \"{name}\": {string.Join(", ", extraPropsInZ)}");

                        tr.Commit();
                        continue;
                    }

                    if (!zAttrNames.Contains(QuantityTag))
                    {
                        ed.WriteMessage($"\n\"{zName}\" is missing the \"{QuantityTag}\" attribute — skipping.");
                        tr.Commit();
                        continue;
                    }

                    templateIds[name] = allowedPool[zName];
                    tr.Commit();
                }
            }

            if (templateIds.Count == 0)
            {
                ed.WriteMessage("\nNo valid block pairs found. Aborting.");
                return;
            }

            // --- Step 3b: select the NET block ---
            var peoNet = new PromptEntityOptions("\nSelect the NET block: ");
            peoNet.SetRejectMessage("\nOnly block references are allowed.");
            peoNet.AddAllowedClass(typeof(BlockReference), exactMatch: false);

            var perNet = ed.GetEntity(peoNet);
            if (perNet.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nCancelled. Aborting.");
                return;
            }

            Point3d netOrigin;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var netBr = (BlockReference)tr.GetObject(perNet.ObjectId, OpenMode.ForRead);
                // NET origin = top-left corner of the grid (left edge X, top edge Y)
                var extents = netBr.GeometricExtents;
                netOrigin = new Point3d(extents.MinPoint.X, extents.MaxPoint.Y, extents.MinPoint.Z);
                tr.Commit();
            }

            ed.WriteMessage($"\nNET origin: ({netOrigin.X:0.##}, {netOrigin.Y:0.##})");

            // --- Step 4: for each group, deduplicate by fingerprint and create -Z blocks ---
            int cellIndex = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                foreach (var kvp in groupedSources.Where(k => templateIds.ContainsKey(k.Key)))
                {
                    string srcName = kvp.Key;
                    var sources = kvp.Value;
                    var templateId = templateIds[srcName];

                    var templateBr = (BlockReference)tr.GetObject(templateId, OpenMode.ForRead);

                    // Group sources by fingerprint (identical attr+prop values → same -Z block with higher Quantity)
                    var fingerGroups = new Dictionary<string, List<SourceBlockData>>(StringComparer.Ordinal);
                    foreach (var src in sources)
                    {
                        string fp = src.GetFingerprint();
                        if (!fingerGroups.ContainsKey(fp))
                            fingerGroups[fp] = new List<SourceBlockData>();
                        fingerGroups[fp].Add(src);
                    }

                    foreach (var group in fingerGroups.Values)
                    {
                        int quantity = group.Count;
                        var representative = group[0];

                        // Calculate grid cell position (center of cell)
                        int col = cellIndex % ColumnsPerRow;
                        int row = cellIndex / ColumnsPerRow;
                        double cellX = netOrigin.X + (col * CellWidth) + (CellWidth / 2.0);
                        double cellY = netOrigin.Y - (row * CellHeight) - (CellHeight / 2.0);
                        var cellPos = new Point3d(cellX, cellY, netOrigin.Z);
                        cellIndex++;

                        // Create new -Z block reference
                        // Use DynamicBlockTableRecord so dynamic properties are available on the new instance
                        var btrId = templateBr.IsDynamicBlock
                            ? templateBr.DynamicBlockTableRecord
                            : templateBr.BlockTableRecord;

                        var newBr = new BlockReference(cellPos, btrId);
                        newBr.ScaleFactors = templateBr.ScaleFactors;
                        newBr.Rotation = templateBr.Rotation;
                        newBr.Normal = templateBr.Normal;

                        ms.AppendEntity(newBr);
                        tr.AddNewlyCreatedDBObject(newBr, true);

                        // Add attributes from the BTR definition
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        foreach (ObjectId entId in btr)
                        {
                            var dbObj = tr.GetObject(entId, OpenMode.ForRead);
                            if (dbObj is AttributeDefinition attDef && !attDef.Constant
                                && !attDef.Tag.StartsWith(IgnoredAttributePrefix, StringComparison.OrdinalIgnoreCase))
                            {
                                var attRef = new AttributeReference();
                                attRef.SetAttributeFromBlock(attDef, newBr.BlockTransform);

                                if (string.Equals(attDef.Tag, QuantityTag, StringComparison.OrdinalIgnoreCase))
                                {
                                    attRef.TextString = quantity.ToString();
                                }
                                else if (representative.Attributes.TryGetValue(attDef.Tag, out string val))
                                {
                                    attRef.TextString = val;
                                }

                                newBr.AttributeCollection.AppendAttribute(attRef);
                                tr.AddNewlyCreatedDBObject(attRef, true);
                            }
                        }

                        // Set dynamic properties
                        if (newBr.IsDynamicBlock)
                        {
                            try
                            {
                                var dynProps = newBr.DynamicBlockReferencePropertyCollection;
                                if (dynProps != null)
                                {
                                    foreach (DynamicBlockReferenceProperty p in dynProps)
                                    {
                                        if (p.ReadOnly) continue;
                                        if (string.Equals(p.PropertyName, "Origin", StringComparison.OrdinalIgnoreCase)) continue;
                                        if (representative.DynamicProps.TryGetValue(p.PropertyName, out object srcVal))
                                        {
                                            try { p.Value = srcVal; }
                                            catch { }
                                        }
                                    }
                                }
                            }
                            catch { }
                        }

                        ed.WriteMessage($"\nCreated \"{srcName}-Z\" (Quantity={quantity}) at cell [{row},{col}] ({cellPos.X:0.##}, {cellPos.Y:0.##})");
                    }
                }

                tr.Commit();
            }

            ed.WriteMessage($"\n{cellIndex} block(s) placed in the NET grid.");
            ed.WriteMessage("\n--- NHBlock complete ---\n");
        }

        #region Data Classes

        private class SourceBlockData
        {
            public Dictionary<string, string> Attributes { get; set; } = new(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, object> DynamicProps { get; set; } = new(StringComparer.OrdinalIgnoreCase);

            public string GetFingerprint()
            {
                var parts = new List<string>();

                foreach (var kv in Attributes.OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
                    parts.Add($"A:{kv.Key}={kv.Value}");

                foreach (var kv in DynamicProps.OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
                    parts.Add($"P:{kv.Key}={ValueToString(kv.Value)}");

                return string.Join("|", parts);
            }
        }

        #endregion

        #region Helpers

        private static string GetBlockEffectiveName(Transaction tr, BlockReference br)
        {
            try
            {
                return ((BlockTableRecord)br.DynamicBlockTableRecord.GetObject(OpenMode.ForRead)).Name;
            }
            catch
            {
                // ignore
            }

            var btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
            return btr.Name;
        }

        private static Dictionary<string, string> CollectAttributeValues(Transaction tr, BlockReference br)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (br.AttributeCollection == null || br.AttributeCollection.Count == 0)
                return dict;

            foreach (ObjectId attId in br.AttributeCollection)
            {
                if (!attId.IsValid) continue;
                var obj = tr.GetObject(attId, OpenMode.ForRead, false);
                if (obj is AttributeReference ar && !string.IsNullOrWhiteSpace(ar.Tag) && !IsIgnoredName(ar.Tag))
                    dict[ar.Tag] = ar.TextString ?? "";
            }

            return dict;
        }

        private static Dictionary<string, object> CollectDynamicPropertyValues(BlockReference br)
        {
            var dict = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            if (!br.IsDynamicBlock)
                return dict;

            DynamicBlockReferencePropertyCollection props;
            try
            {
                props = br.DynamicBlockReferencePropertyCollection;
            }
            catch
            {
                return dict;
            }

            if (props == null || props.Count == 0)
                return dict;

            foreach (DynamicBlockReferenceProperty p in props)
            {
                if (p.ReadOnly) continue;
                if (string.Equals(p.PropertyName, "Origin", StringComparison.OrdinalIgnoreCase)) continue;
                if (!string.IsNullOrWhiteSpace(p.PropertyName) && !IsIgnoredName(p.PropertyName))
                {
                    try { dict[p.PropertyName] = p.Value; }
                    catch { }
                }
            }

            return dict;
        }

        private static HashSet<string> CollectAttributeNames(Transaction tr, BlockReference br)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (br.AttributeCollection == null || br.AttributeCollection.Count == 0)
                return names;

            foreach (ObjectId attId in br.AttributeCollection)
            {
                if (!attId.IsValid) continue;
                var obj = tr.GetObject(attId, OpenMode.ForRead, false);
                if (obj is AttributeReference ar && !string.IsNullOrWhiteSpace(ar.Tag) && !IsIgnoredName(ar.Tag))
                    names.Add(ar.Tag);
            }

            return names;
        }

        private static HashSet<string> CollectDynamicPropertyNames(BlockReference br)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (!br.IsDynamicBlock)
                return names;

            DynamicBlockReferencePropertyCollection props;
            try
            {
                props = br.DynamicBlockReferencePropertyCollection;
            }
            catch
            {
                return names;
            }

            if (props == null || props.Count == 0)
                return names;

            foreach (DynamicBlockReferenceProperty p in props)
            {
                if (p.ReadOnly) continue;
                if (string.Equals(p.PropertyName, "Origin", StringComparison.OrdinalIgnoreCase)) continue;
                if (!string.IsNullOrWhiteSpace(p.PropertyName) && !IsIgnoredName(p.PropertyName))
                    names.Add(p.PropertyName);
            }

            return names;
        }

        private static bool IsIgnoredName(string name)
            => name.StartsWith(IgnoredAttributePrefix, StringComparison.OrdinalIgnoreCase);

        private static string ValueToString(object v)
        {
            if (v == null) return "null";
            if (v is double d) return d.ToString("0.########");
            if (v is string s) return s;
            return v.ToString() ?? "";
        }

        #endregion
    }
}