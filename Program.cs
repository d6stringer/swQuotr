/*
 Created by Daniel Woodson
 Started On: 12/5/2023
 From Sources:
 https://stackoverflow.com/questions/53682824/convert-all-solidworks-files-in-folder-to-step-files-macro
 https://help.solidworks.com/2022/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDocExtension~SaveAs3.html
 V0.1: 12/6/2023
 Init functional copy: creates PDFs from solidworks drawing files and step files from solidworks part and assembly files. Very little testing done,
 probably still very buggy.
    - Assms not tested
    - Long file paths may be a problem
 */

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System;


namespace CreateQuoteFiles
{
    class Program
    {
        static SldWorks swApp;

        static void Main(string[] args)
        {
            string directoryName = GetDirectoryName();

            if (!GetSolidWorks())
            {
                return;
            }

            int i = 0;

            foreach (string fileName in Directory.GetFiles(directoryName))
            {
                if (Path.GetExtension(fileName).ToLower() == ".sldprt")
                {
                    CreateStepFile(fileName, 1);
                    i += 1;
                }
                else if (Path.GetExtension(fileName).ToLower() == ".sldasm")
                {
                    CreateStepFile(fileName, 2);
                    i += 1;
                }
                else if (Path.GetExtension(fileName).ToLower() == ".slddrw")
                {
                    CreatePDFFile(fileName, 3);
                    i += 1;
                }
                else
                {
                    Console.WriteLine("File type not supported");
                }
            }

            Console.WriteLine("Finished converting {0} files", i);

        }

        static void CreateStepFile(string fileName, int docType)
        {
            int errors = 0;
            int warnings = 0;

            ModelDoc2 swModel = swApp.OpenDoc6(fileName, docType, 1, "", ref errors, ref warnings);

            string stepFile = Path.Combine(Path.GetDirectoryName(fileName), Path.GetFileNameWithoutExtension(fileName) + ".STEP");

            if (swModel.Extension.SaveAs3(stepFile, 0, 1, null, null, ref errors, ref warnings))
            {
                Console.WriteLine("Created STEP file: " + stepFile); ;

                swApp.CloseDoc(fileName);
            }
            else
            {
                Console.WriteLine("Failed to created STEP file: " + stepFile); ;

                swApp.CloseDoc(fileName);
            }

        }

        static void CreatePDFFile(string fileName, int docType)
        {
            //Variable declaration
            int errors = 0;
            int warnings = 0;
            bool boolstatus = false;
            string[] obj = null;
            Sheet swSheet = default(Sheet);

            //Get PDF Data
            ModelDoc2 swModel = (ModelDoc2)swApp.OpenDoc6(fileName, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
            ModelDocExtension swModExt = (ModelDocExtension)swModel.Extension;
            ExportPdfData swExportPDFData = (ExportPdfData)swApp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);
            DrawingDoc swDrawDoc = (DrawingDoc)swModel;

            //honestly not sure what this part does but it works!
            obj = (string[])swDrawDoc.GetSheetNames();
            int count = 0;
            count = obj.Length;
            int i = 0;
            object[] objs = new object[count - 1];
            DispatchWrapper[] arrObjIn = new DispatchWrapper[count - 1];

            for (i = 0; i < count - 1; i++)
            {
                boolstatus = swDrawDoc.ActivateSheet((obj[i]));
                swSheet = (Sheet)swDrawDoc.GetCurrentSheet();
                objs[i] = swSheet;
                arrObjIn[i] = new DispatchWrapper(objs[i]);
            }

            string pdfFile = Path.Combine(Path.GetDirectoryName(fileName), Path.GetFileNameWithoutExtension(fileName) + ".PDF");

            if (swExportPDFData.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportSpecifiedSheets, (arrObjIn)))
            {
                swExportPDFData.ViewPdfAfterSaving = false;
                if (swModExt.SaveAs(pdfFile, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, swExportPDFData, ref errors, ref warnings))
                {
                    Console.WriteLine("Created PDF file: " + pdfFile); ;

                    swApp.CloseDoc(fileName);
                }
                else
                {
                    Console.WriteLine("Failed to created PDF file: " + pdfFile); ;

                    swApp.CloseDoc(fileName);
                }
            }


        }

        static string GetDirectoryName()
        {

            Console.WriteLine("Enter Directory to Convert to Step Files!");
            string s = Console.ReadLine();

            if (Directory.Exists(s))
            {
                return s.Remove(s.Length - 1, 1);
            }

            Console.WriteLine("Directory does not exists, try again");
            return GetDirectoryName();
        }

        static bool GetSolidWorks()
        {
            try
            {
                swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));

                if (swApp == null)
                {
                    throw new NullReferenceException(nameof(swApp));
                }

                if (!swApp.Visible)
                {
                    swApp.Visible = true;
                }

                Console.WriteLine("SolidWorks Loaded");
                return true;
            }
            catch (Exception)
            {
                Console.WriteLine("Could not launch SolidWorks");
                return false;
            }
        }
    }
}