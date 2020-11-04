namespace mpTxtCopyPaste
{
    using System;
    using System.Globalization;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.Runtime;
    using ModPlusAPI;
    using ModPlusAPI.Windows;
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

    /// <summary>
    /// Main class of plugin
    /// </summary>
    public class TextCopyPaste
    {
        /// <summary>
        /// Command start
        /// </summary>
        [CommandMethod("ModPlus", "mpTxtCopyPaste", CommandFlags.UsePickSet)]
        public static void MainFunction()
        {
#if !DEBUG
            Statistic.SendCommandStarting(new ModPlusConnector());
#endif

            try
            {
                var deleteSource = false; // Удаление исходника

                // Значение по-умолчанию из расширенных данных чертеже
                var defVal = ModPlus.Helpers.XDataHelpers.GetStringXData("mpTxtCopyPaste");
                if (!string.IsNullOrEmpty(defVal))
                {
                    if (bool.TryParse(defVal, out var tmp))
                        deleteSource = tmp;
                }

                var doc = AcApp.DocumentManager.MdiActiveDocument;
                var db = doc.Database;
                var ed = doc.Editor;

                var per = PromptSourceEntity(doc, ref deleteSource);

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var str = GetString(per, deleteSource, ed);

                    if (!string.IsNullOrEmpty(str))
                    {
                        while (true)
                        {
                            per = ed.GetEntity(GetDestinationEntityOptions());
                            if (per.Status == PromptStatus.Cancel)
                                break;
                            if (per.Status != PromptStatus.OK)
                                continue;

                            var destinationEntity = tr.GetObject(per.ObjectId, OpenMode.ForWrite) as Entity;
                            SetString(destinationEntity, str, ed);
                            db.TransactionManager.QueueForGraphicsFlush();
                        }
                    }

                    tr.Commit();
                }
            }
            catch (OperationCanceledException)
            {
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        private static PromptEntityOptions GetDestinationEntityOptions()
        {
            // Выберите объект (текст, выноска или таблица) для замены содержимого:
            var peo = new PromptEntityOptions($"\n{Language.GetItem("msg5")}");
            peo.SetRejectMessage($"\n{Language.GetItem("msg3")}");
            peo.AddAllowedClass(typeof(DBText), false);
            peo.AddAllowedClass(typeof(MText), false);
            peo.AddAllowedClass(typeof(MLeader), false);
            peo.AddAllowedClass(typeof(Table), false);
            peo.AllowNone = false;
            return peo;
        }

        private static PromptEntityResult PromptSourceEntity(Document doc, ref bool deleteSource)
        {
            // Выберите объект-исходник (текст, выноска, таблица или размер):
            var peo = new PromptEntityOptions($"\n{Language.GetItem("msg1")}");
            peo.SetMessageAndKeywords($"\n{Language.GetItem("msg2")}", "Delete");
            peo.AppendKeywordsToMessage = true;
            peo.SetRejectMessage($"\n{Language.GetItem("msg3")}");
            peo.AddAllowedClass(typeof(DBText), false);
            peo.AddAllowedClass(typeof(MText), false);
            peo.AddAllowedClass(typeof(MLeader), false);
            peo.AddAllowedClass(typeof(Table), false);
            peo.AddAllowedClass(typeof(Dimension), false);
            peo.AllowNone = true;
            var per = doc.Editor.GetEntity(peo);

            if (per.Status == PromptStatus.Keyword)
            {
                // Удалять объект-исходник (для таблиц - содержимое ячейки)?
                deleteSource = MessageBox.ShowYesNo(Language.GetItem("msg4"), MessageBoxIcon.Question);

                // Сохраняем текущее значение как значение по умолчанию
                ModPlus.Helpers.XDataHelpers.SetStringXData("mpTxtCopyPaste", deleteSource.ToString());
                return PromptSourceEntity(doc, ref deleteSource);
            }

            if (per.Status == PromptStatus.OK)
            {
                return per;
            }

            throw new OperationCanceledException();
        }

        private static string GetString(PromptEntityResult prompt, bool deleteSource, Editor ed)
        {
            var mode = deleteSource ? OpenMode.ForWrite : OpenMode.ForRead;
            var ent = prompt.ObjectId.GetObject(mode);

            // Delete source
            if (deleteSource && !(ent is Table))
                ent.Erase(true);

            switch (ent)
            {
                case DBText dbText:
                    return dbText.TextString;
                case MText mText:
                    return mText.Contents;
                case MLeader mLeader:
                    return mLeader.MText.Contents;
                case Dimension dimension:
                    return !string.IsNullOrEmpty(dimension.DimensionText)
                        ? dimension.DimensionText
                        : Math.Round(dimension.Measurement, dimension.Dimdec)
                            .ToString(CultureInfo.InvariantCulture)
                            .Replace(".", dimension.Dimdsep.ToString());
                case Table table:
                    var cell = GetCell(table, ed);
                    if (cell != null)
                    {
                        var str = cell.TextString;
                        if (deleteSource)
                            cell.TextString = string.Empty;
                        return str;
                    }

                    break;
            }

            return null;
        }

        private static void SetString(Entity ent, string str, Editor ed)
        {
            switch (ent)
            {
                case DBText dbText:
                    dbText.TextString = str;
                    break;
                case MText mText:
                    mText.Contents = str;
                    break;
                case MLeader mLeader:
                    var text = mLeader.MText;
                    if (text != null)
                    {
                        text.Contents = str;
                        mLeader.MText = text;
                    }

                    break;
                case Table table:
                    var cell = GetCell(table, ed);
                    if (cell != null)
                        cell.TextString = str;
                    table.RecomputeTableBlock(true);
                    break;
            }
        }

        private static Cell GetCell(Table table, Editor ed)
        {
            // Укажите ячейку таблицы:
            var r = ed.GetPoint($"\n{Language.GetItem("msg6")}");
            if (r.Status == PromptStatus.OK)
            {
                var hit = table.HitTest(r.Value, Vector3d.ZAxis);
                if (hit.Type == TableHitTestType.Cell)
                    return table.Cells[hit.Row, hit.Column];
            }

            return null;
        }
    }
}