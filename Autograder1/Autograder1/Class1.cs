/* Autograder for ME1770
 * Author:Vaibhav Kumar
 * Author email:Vaibhavkumar@gatech.edu
 * code git link - github.com/vaibhav-kumar/autograder-Version1
 * */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;

[assembly: CommandClass(typeof(Autograder1.Class1))]

namespace Autograder1
{
    public class Class1
    {
        [CommandMethod("Autograder", CommandFlags.Session)]
        public static void Autograder()
        {
            const string systemVar_DwgCheck = "DWGCHECK";
            Int16 dwgCheckPrevious = (Int16)Application.GetSystemVariable(systemVar_DwgCheck);
            Application.SetSystemVariable(systemVar_DwgCheck, 2);
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            PromptStringOptions pStrOpts = new PromptStringOptions("\nEnter entire path to the folder containing all the submissions");
            pStrOpts.AllowSpaces = true;
            PromptResult pStrRes = acDoc.Editor.GetString(pStrOpts);
            /*string[] folderpaths = Directory.GetDirectories(@"C:\Users\Vaibhav\Desktop\2D\2D Cad project\");*/
            string[] folderpaths = Directory.GetDirectories(pStrRes.StringResult);
            int i = 0;
            foreach (string strFolderStudname in folderpaths)
            {
                i = i + 1;
                //define dictionary for storing data.
                Dictionary<string, string> dict1 = new Dictionary<string, string>();
                dict1["auxinfodim"] = "";
                dict1["blocksinlastlayout"] = "";
                dict1["filename"] = i.ToString() + ": " + strFolderStudname;
                //writestuff(strwrite,i.ToString()+": "+strFolderStudname,1);
                Directory.SetCurrentDirectory(strFolderStudname + "/Submission attachment(s)/");
                string[] filepaths = Directory.GetFiles(Directory.GetCurrentDirectory());
                bool checkfilepass = false;
                bool checkfilepdf = false;
                bool checkfilepdfmulti = false;
                int checkblockexistance = 0;
                foreach (string strFileName in filepaths)
                {
                    string newStr = strFileName.Substring(strFileName.Length - 3, 3);
                    if (newStr.Equals("dwg"))
                    {
                        /*Application.DocumentManager.Open(strFileName,false); //the boolean is for read only
                        Document acDoc = Application.DocumentManager.MdiActiveDocument;
                        Database acCurDb = acDoc.Database;*/
                        Database acCurDb = new Database(false, true);
                        acCurDb.ReadDwgFile(strFileName, FileOpenMode.OpenForReadAndAllShare, false, null);
                        acCurDb.CloseInput(true);
                        checkfilepass = true;
                        Layout taborderlayout = null;
                        int vp2 = 0;
                        string[] Layerlist = { };

                        using (Transaction acTrans2 = acCurDb.TransactionManager.StartTransaction())
                        {
                            DBDictionary lays = acTrans2.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                            LayoutManager Lm = LayoutManager.Current;
                            string currentLayout = Lm.CurrentLayout;
                            int layoutnum = 0;
                            string slayouts = "NA";
                            int tabordercheck = 0;
                            int tabviewports = 0;
                            bool tabcheck = false;
                            Viewport taborderviewport=null;
                            Viewport tempviewport = null;
                            Layout firsttaborderlayout=null;
                            Viewport firsttaporderviewport=null;
                            int firsttabviewports = 0;
                            foreach (DBDictionaryEntry de in lays)
                            {
                                layoutnum += 1;
                                Layout layout = de.Value.GetObject(OpenMode.ForRead) as Layout;
                                string layoutName = layout.LayoutName;
                                if (layoutName != "Model")
                                {
                                    if (slayouts.Equals("NA"))
                                    {
                                        slayouts = layoutName;
                                    }
                                    else
                                    {
                                        slayouts = layoutName + " , " + slayouts;
                                    }
                                    ObjectIdCollection vpIdCol = layout.GetViewports();
                                    if (tabordercheck < layout.TabOrder)
                                    {
                                        tabordercheck = layout.TabOrder;
                                        taborderlayout = layout;
                                        tabcheck = true;
                                    }
                                    else
                                    {
                                        tabcheck = false;
                                    }

                                    int nviewports = -1;
                                    foreach (ObjectId vpId in vpIdCol)
                                    {
                                        Viewport vp = vpId.GetObject(OpenMode.ForRead) as Viewport;
                                        nviewports += 1;
                                        tempviewport = vp;
                                        if (tabcheck == true)
                                        {
                                            tabviewports = nviewports;
                                        }
                                        if (layout.TabOrder == 2)
                                        {
                                            firsttaborderlayout = layout;
                                            firsttaporderviewport = vp;
                                            firsttabviewports = nviewports;
                                        }

                                        // now check for dimensions inside the viewports
                                        int numdin = 0;
                                        LayerTable lyrtbl1 = acTrans2.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
                                        foreach (ObjectId acobj1 in lyrtbl1)
                                        {
                                            LayerTableRecord Lyrtblrec1 = acTrans2.GetObject(acobj1, OpenMode.ForRead) as LayerTableRecord;
                                            bool isfrozenlayer1 = false;
                                            foreach (ObjectId acobj3 in vp.GetFrozenLayers())
                                            {
                                                if (Lyrtblrec1.Id == acobj3)
                                                {
                                                    isfrozenlayer1 = true;
                                                }
                                            }
                                            if (isfrozenlayer1 == false)
                                            {
                                                BlockTable BlckTbl1 = acTrans2.GetObject(acCurDb.BlockTableId,
                                                                             OpenMode.ForRead) as BlockTable;
                                                BlockTableRecord acBlkTblRec1 = acTrans2.GetObject(BlckTbl1[BlockTableRecord.ModelSpace],
                                                                                OpenMode.ForRead) as BlockTableRecord;
                                                foreach (ObjectId BlockId1 in acBlkTblRec1)
                                                {
                                                    try
                                                    {
                                                        Dimension Diment = acTrans2.GetObject(BlockId1, OpenMode.ForRead) as Dimension;
                                                        if (Diment.Layer == Lyrtblrec1.Name)
                                                        {
                                                            numdin += 1;
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        //catch nothing - the exception is that its not a dim, all i care about is it being a dim
                                                    }
                                                }
                                            }
                                        }
                                        if (nviewports > 0)
                                        {
                                            //writestuff(strwrite,layoutName + " The number of dimensions in viewport number "+ nviewports.ToString() + " are" + numdin.ToString(),0);
                                            //dict1.Add("auxinfodim", dict1["auxinfodim"] + layoutName + " The number of dimensions in viewport number " + nviewports.ToString() + " are" + numdin.ToString() + "/n");
                                            dict1["auxinfodim"]=dict1["auxinfodim"] +"    " +layoutName + ": The number of dimensions in viewport number " + nviewports.ToString() + " are " + numdin.ToString() + "\r\n";
                                        }
                                        //end of this code block for the viewports

                                    }
                                    if (nviewports >= 2)
                                    {
                                        vp2 += 1;
                                    }
                                }
                                if (tabviewports >= 1 && layoutName.Equals(taborderlayout.LayoutName))
                                {
                                    taborderviewport = null;
                                }
                                if (tabviewports == 1 && layoutName.Equals(taborderlayout.LayoutName))
                                {
                                    taborderviewport = tempviewport;
                                }

                            }
                            Lm.CurrentLayout = currentLayout;
                            //writestuff("Number of layouts: " + (layoutnum - 1).ToString());
                            //dict1.Add("layoutnum", "Number of layouts: " + (layoutnum - 1).ToString());
                            dict1.Add("layoutnum","Number of layouts: " + (layoutnum - 1).ToString());
                            //writestuff("Names of layouts: " + slayouts);
                            dict1.Add("layoutnames", "Names of layouts: " + slayouts);
                            //writestuff("Number of layouts containing two or greater viewports" + vp2.ToString());
                            dict1.Add("Numviewports2", "Number of layouts containing two or greater viewports: " + vp2.ToString());
                            if (firsttabviewports != 1)
                            {
                                //writestuff("First layout scale: The First layout has more than one viewports");
                                dict1.Add("firstlayoutscale", "First layout scale: The First layout has more than one viewports");
                            }
                            else
                            {
                                //writestuff("First layout scale: 1:" + (1 / firsttaporderviewport.AnnotationScale.Scale).ToString());
                                dict1.Add("firstlayoutscale", "First layout scale: 1:" + (1 / firsttaporderviewport.AnnotationScale.Scale).ToString());
                            }
                            //writestuff("Name of last layout: " + taborderlayout.LayoutName);
                            dict1.Add("lastlayoutname", "Name of last layout: " + taborderlayout.LayoutName);
                            if (taborderviewport == null)
                            {
                                //writestuff("The last layout has more than one viewports. But this can also be a bug.please check manually");
                                dict1.Add("lastlayoutscale", "The last layout has more than one viewports. But this can also be a bug.please check manually");
                                //writestuff("As there are multiple viewports, the software will not check for presence of blocks");
                                //dict1.Add("blocksinlastlayout", "As there are multiple viewports, the software will not check for presence of blocks");
                                dict1["blocksinlastlayout"] = "As there are multiple viewports, the software will not check for presence of blocks";
                            }
                            else
                            {
                                //writestuff("Last Layout scale: 1:" + (1 / taborderviewport.AnnotationScale.Scale).ToString());
                                dict1.Add("lastlayoutscale", "Last Layout scale: 1:" + (1 / taborderviewport.AnnotationScale.Scale).ToString());
                                // The following codeblock is for the black presence in the last layout.
                                LayerTable lyrtbl = acTrans2.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
                                foreach (ObjectId acobj in lyrtbl)
                                {
                                    LayerTableRecord Lyrtblrec = acTrans2.GetObject(acobj, OpenMode.ForRead) as LayerTableRecord;
                                    bool isfrozenlayer = false;
                                    foreach (ObjectId acobj2 in taborderviewport.GetFrozenLayers())
                                    {
                                        if (Lyrtblrec.Id == acobj2)
                                        {
                                            isfrozenlayer = true;
                                        }
                                    }
                                    if (isfrozenlayer==false)
                                    {
                                        BlockTable BlckTbl=acTrans2.GetObject(acCurDb.BlockTableId,
                                                                     OpenMode.ForRead) as BlockTable;
                                        BlockTableRecord acBlkTblRec=acTrans2.GetObject(BlckTbl[BlockTableRecord.ModelSpace],
                                                                        OpenMode.ForRead) as BlockTableRecord;
                                        foreach (ObjectId BlockId in acBlkTblRec)
                                        {
                                            if ((BlockId.ObjectClass.DxfName).Equals("INSERT"))
                                            {
                                                Entity ent = acTrans2.GetObject(BlockId, OpenMode.ForRead) as Entity;
                                                if (ent.Layer == Lyrtblrec.Name)
                                                {
                                                    //writestuff("Block present in layer" + Lyrtblrec.Name+" in the last layout");
                                                    //dict1.Add("blocksinlastlayout", dict1["blocksinlastlayout"] + "Block present in layer" + Lyrtblrec.Name + " in the last layout");
                                                    dict1["blocksinlastlayout"]=dict1["blocksinlastlayout"] + "    "+ "Block present in layer" + Lyrtblrec.Name + " in the last layout."+"\r\n";
                                                }

                                            }
                                        }
                                    }
                                }
                            }
                        }
                        //starting a transcation for layers
                        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                        {
                            LayerTable acLyrTbl;
                            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId,OpenMode.ForRead) as LayerTable;
                            string sLayerNames = "NA";
                            int number1 = 0;
                            int uniquelayers = 0;
                            foreach (ObjectId acObjId in acLyrTbl)
                            {
                                LayerTableRecord acLyrTblRec;
                                acLyrTblRec = acTrans.GetObject(acObjId,
                                                                OpenMode.ForRead) as LayerTableRecord;
                                if (acLyrTblRec.Name.Equals("dimension") || acLyrTblRec.Name.Equals("center_line") || acLyrTblRec.Name.Equals("semester") || acLyrTblRec.Name.Equals("cuttingplane_line") || acLyrTblRec.Name.Equals("hidden_line") || acLyrTblRec.Name.Equals("Defpoints") || acLyrTblRec.Name.Equals("mview1") || acLyrTblRec.Name.Equals("0"))
                                {
                                    //deprecated as of 9/20/2013. dont need this part
                                }
                                else
                                {
                                    Layerlist = new string[Layerlist.Length + 1];
                                    Layerlist[Layerlist.Length - 1] = acLyrTblRec.Name;
                                    uniquelayers = uniquelayers + 1;
                                    if (sLayerNames.Equals("NA"))
                                    {
                                        sLayerNames = acLyrTblRec.Name;
                                    }
                                    else
                                    {
                                        sLayerNames = acLyrTblRec.Name + " , " + sLayerNames;
                                    }
                                }
                                number1 = number1 + 1;
                            }
                            //writestuff("Total Number of Layers: " + number1.ToString() + "\n");
                            dict1.Add("layernum", "Total Number of Layers: " + number1.ToString() + "\n");
                            //writestuff("Total Number of User Defined layers: " + uniquelayers + "\n");
                            dict1.Add("userdefinedlayernum", "Total Number of User Defined layers: " + uniquelayers + "\n");
                            //writestuff("The User defined layers are:" + sLayerNames + "\n");
                            dict1.Add("userdefinedlayernames", "The User defined layers are:" + sLayerNames + "\n");
                        }
                        //starting a new transaction for blocks
                        using (Transaction acTrans3 = acCurDb.TransactionManager.StartTransaction())
                        {
                            BlockTable BlckTbl;
                            BlckTbl = acTrans3.GetObject(acCurDb.BlockTableId,OpenMode.ForRead) as BlockTable;
                            BlockTableRecord acBlkTblRec = acTrans3.GetObject(BlckTbl[BlockTableRecord.ModelSpace],OpenMode.ForRead) as BlockTableRecord;
                            foreach (ObjectId BlockId in acBlkTblRec)
                            {
                                if ((BlockId.ObjectClass.DxfName).Equals("INSERT"))
                                {
                                    checkblockexistance += 1;
                                    
                                }
                            }
                            if (checkblockexistance > 0)
                            {
                                //writestuff("There are " + checkblockexistance + " Blocks inserted in the model space");
                                dict1.Add("blocksexist", "There are " + checkblockexistance + " Blocks inserted in the model space");
                            }
                            else
                            {
                                //writestuff("There are no blocks inserted in the modelspace");
                                dict1.Add("blocksexist","There are no blocks inserted in the modelspace");
                            }
                        }

                        //check bock existance in last layout
                        using (Transaction acTrans7 = acCurDb.TransactionManager.StartTransaction())
                        {
                            BlockTable bt = acTrans7.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                            foreach (ObjectId objid in bt)
                            {
                                BlockTableRecord Btr = acTrans7.GetObject(objid, OpenMode.ForRead) as BlockTableRecord;
                                if (Btr.IsLayout)
                                {
                                    ObjectId lid = Btr.LayoutId;
                                    Layout lt = acTrans7.GetObject(lid, OpenMode.ForRead) as Layout;
                                    if (lt==taborderlayout)
                                    {
                                        foreach (ObjectId eid in Btr)
                                        {
                                            Entity ent;
                                            if ((eid.ObjectClass.DxfName).Equals("INSERT"))
                                            {
                                                Application.ShowAlertDialog("block present");
                                            }
                                            ent = acTrans7.GetObject(eid, OpenMode.ForRead) as Entity;
                                        }
                                    }
                                }
                            }
                        }
                        /*acDoc.CloseAndDiscard();*/
                    } //if close

                    if (newStr.Equals("pdf"))
                    {
                        if (checkfilepdf)
                        {
                            //there is more than one pdf file
                            checkfilepdfmulti = true;
                            break;
                        }
                        checkfilepdf = true;
                    }
                } //foreach for the document finder close
                if (checkfilepass == false)
                {
                    writestuff("No .dwg file could be found.Please check manually though and if the file exists, send bug report to vaibhavkumar@gatech.edu", dict1, pStrRes.StringResult);
                }
                else
                {
                    if (checkfilepdf == false)
                    {
                        //writestuff("PDF:No PDF file can be found");
                        dict1.Add("pdfpresent", "PDF:No PDF file can be found");
                        writestuff("0", dict1, pStrRes.StringResult);
                    }
                    if (checkfilepdf == true && checkfilepdfmulti == true)
                    {
                        //writestuff("PDF:Multiple PDF files present");
                        dict1.Add("pdfpresent", "PDF:Multiple PDF files present");
                        writestuff("0", dict1,pStrRes.StringResult);
                    }
                    if (checkfilepdf == true && checkfilepdfmulti == false)
                    {
                        //writestuff("PDF: Only one(1) PDF file present");
                        dict1.Add("pdfpresent", "PDF: Only one(1) PDF file present");
                        writestuff("0", dict1, pStrRes.StringResult);
                    }
                }
                //writestuff("\r\n"); //terminating space
                dict1.Clear();
            } //foreach for the folders
            Application.SetSystemVariable(systemVar_DwgCheck, dwgCheckPrevious);
            Application.ShowAlertDialog("Program Execution Over. Please ceck your grade file.Any bugs must be reported to vaibhavkumar@gatech.edu");
        } //function close

        //function to write stuff.removed all individual calls to the writeline function for redundancy. 9/20/2013
        static void writestuff(string s1, Dictionary<string,string> dict1, string s2)
        {
            s2 = s2 + "/Grade.txt";
            if (s1.Equals("0"))
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(s2, true))
                {
                    file.WriteLine(dict1["filename"] + "\n");
                    file.WriteLine("A: Planning notes and Document PDF \n");
                    file.WriteLine("    "+dict1["pdfpresent"] + "\n");
                    file.WriteLine("B: Layouts: \n");
                    file.Write("    " + dict1["layoutnum"] + "\r\n");
                    file.WriteLine("    " + dict1["layoutnames"] + "\n");
                    file.WriteLine("C:" + dict1["Numviewports2"]+"\n");
                    file.WriteLine("D:Blocks in last layout \n");
                    file.WriteLine("    " + dict1["blocksexist"] + "\n");
                    file.WriteLine(dict1["blocksinlastlayout"]);
                    file.WriteLine("E: Proper Use of layers \n");
                    file.Write("    " + dict1["layernum"]);
                    file.WriteLine("\r");
                    file.WriteLine("    " + dict1["userdefinedlayernum"] + "\n");
                    file.WriteLine("    " + dict1["userdefinedlayernames"] + "\n");
                    file.WriteLine("F: Auxillary Information:" + "\n");
                    file.WriteLine("F1: Firstlayoutscale \n");
                    file.WriteLine("    " + dict1["firstlayoutscale"] + "\n");
                    file.WriteLine("F2: LastLayoutscale \n");
                    file.WriteLine("    " + dict1["lastlayoutscale"] + "\n");
                    file.WriteLine("F3: Dimensions in each viewport \n");
                    file.WriteLine(dict1["auxinfodim"]);
                    file.WriteLine("\n");
                }

            }
            else
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(s2, true))
                {
                    file.WriteLine(dict1["filename"] + "\n");
                    file.WriteLine(s1);
                    file.WriteLine("\n");
                }
            }
        }
    } //class close
} //namespace close






 
