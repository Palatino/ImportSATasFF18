using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace ImportSATasFF18
{
    [TransactionAttribute(TransactionMode.Manual)]
    class ImportSATasFF18 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            //Create folder dialog path
            FileOpenDialog file_dia = new FileOpenDialog("SAT file (*.sat)|*.sat");
            file_dia.Title = "Select SAT file to import";
            file_dia.Show();
            ModelPath path = file_dia.GetSelectedModelPath();
            file_dia.Dispose();

            //Convert file path to a string
            string path_str = ModelPathUtils.ConvertModelPathToUserVisiblePath(path);

            SATImportOptions satOpt = new SATImportOptions();

            FilteredElementCollector filEle = new FilteredElementCollector(doc);

            IList<Element> views = filEle.OfClass(typeof(View)).ToElements();

            View import_view = views[0] as View;
            try
            {
                using (Transaction trans = new Transaction(doc, "Import SAT"))
                {   // Start transaction, import SAT file and get the element
                    trans.Start();
                    ElementId importedElementId = doc.Import(path_str, satOpt, import_view);
                    Element importedElement = doc.GetElement(importedElementId);

                    //Extract geometry element from the imported element
                    Options geoOptions = new Options();
                    GeometryElement importedGeometry = importedElement.get_Geometry(geoOptions);


                    //Iterate through the geometry elements extracting the geometry as individual elements
                    foreach (GeometryObject geoObj in importedGeometry)
                    {
                        GeometryInstance instance = geoObj as GeometryInstance;
                        foreach (GeometryObject instObj in instance.SymbolGeometry)
                        {
                            Solid solid = instObj as Solid;
                            FreeFormElement.Create(doc, solid);
                        }

                    }

                    //Delete SAT file

                    doc.Delete(importedElementId);
                    trans.Commit();

                }

                return Result.Succeeded;
            }

            catch
            {
                TaskDialog.Show("Error Importing", "Something went wrong");
                return Result.Failed;
            }

        }
    }
}
