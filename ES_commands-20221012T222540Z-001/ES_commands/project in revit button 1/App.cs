using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Architecture;
using System.Diagnostics;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Media.Imaging;
using System.Windows.Forms;

using AJC_Commands_1;

using System.Reflection;
using OfficeOpenXml;
using Rhino.FileIO;
using Rhino.Geometry;

using RhinoCon;
using winform = System.Windows.Forms;
using System.Windows.Annotations;
using System.Windows.Media;

using adWin = Autodesk.Windows;
using Autodesk.Revit.DB.Events;
using System.Net.NetworkInformation;
using System.Runtime.CompilerServices;
using Microsoft.Win32;
using System.IO.Packaging;
using ES_commands;
using System.Windows.Controls;
using System.Threading.Tasks;
using Rhino.UI;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Timers;

namespace BoostYourBIM
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class PlaceView_CrateSheet : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("F5FB1A7F-8110-4862-8820-04AE05C1239E"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
           




            //---------------------------------------- FILTERS ------------------------------------
            IEnumerable<FamilySymbol> familyList = from elem in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_TitleBlocks)
                                                   let type = elem as FamilySymbol
                                                   select type;

            List<Element> viewElems = new List<Element>();
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            viewElems.AddRange(collector.OfClass(typeof(Autodesk.Revit.DB.View)).ToElements());

            List<ElementId> lista_de_views = new List<ElementId>();
            List<ElementId> tittleblocks_list = new List<ElementId>();
            List<String> tittle_block_list_name = new List<String>();
            List<string> nombres_hojas = new List<string>();
            List<string> nameofviews = new List<string>();
            List<Element> hojitas = new List<Element>();
            List<ViewSchedule> ViewSchedule_LIST = new List<ViewSchedule>();
            List<string> gourpheader_list = new List<string>();

            
            Form13 form2 = new Form13();

            form2.comboBox1.Items.Add("Plans");
            form2.comboBox1.Items.Add("3D views");
            form2.comboBox1.Items.Add("Section & Elevation views");
            form2.comboBox1.Items.Add("Drafting");

            FilteredElementCollector collector2 = new FilteredElementCollector(doc);
            ICollection<Element> hojas = collector2.OfClass(typeof(Autodesk.Revit.DB.View)).ToElements();

            FilteredElementCollector fec = new FilteredElementCollector(doc);
            fec.OfClass(typeof(ViewSection));

            IEnumerable<ViewSheet> viewSheet = from elem in new FilteredElementCollector(doc)
                                                .OfClass(typeof(ViewSheet))
                                                .OfCategory(BuiltInCategory.OST_Sheets)
                                               let type = elem as ViewSheet
                                               where type.Name != null
                                               select type;


            IEnumerable<Autodesk.Revit.DB.View> view_list = from elem in new FilteredElementCollector(doc)
                                                .OfClass(typeof(Autodesk.Revit.DB.View))
                                                .OfCategory(BuiltInCategory.OST_Views)
                                                            let type = elem as Autodesk.Revit.DB.View
                                                            where type.Name != null
                                                            select type;

            List<ElementId> ids_ = new List<ElementId>();


            foreach (var sheet_ in viewSheet)
            {
                ICollection<ElementId> views_ = sheet_.GetAllPlacedViews();
                foreach (var item in views_)
                {
                    ids_.Add(item);
                }
            }


            form2.ShowDialog();

            if (form2.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }
            if (form2.comboBox1.SelectedItem ==  null)
            {

                TaskDialog.Show("Instruction", "You must select a type of view");
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            if (form2.comboBox1.SelectedItem.ToString() == "Section & Elevation views")
            {
                var viewPlans = fec.Cast<ViewSection>().Where<ViewSection>(vp => vp.IsTemplate);
                IEnumerable<ViewSection> viewList = from elem in new FilteredElementCollector(doc)
                                                     .OfClass(typeof(ViewSection))
                                                     .OfCategory(BuiltInCategory.OST_Views)
                                                    let type = elem as ViewSection
                                                    where type.Name != null
                                                    select type;

                List<ViewSection> SecView2 = new List<ViewSection>();
                List<ViewSection> Views = new List<ViewSection>();

                foreach (var view in viewList)
                {
                    if (!view.IsTemplate)
                    {
                        Views.Add(view);

                    }
                }
                foreach (var id1 in viewList)
                {
                    foreach (var viewid in ids_)
                    {
                        if (!id1.Id.IntegerValue.Equals(viewid.IntegerValue))
                        {

                        }
                        else
                        {
                            if (!SecView2.Contains(id1))
                            {
                                SecView2.Add(id1);

                            }
                        }
                    }
                }

                for (int i = 0; i < Views.ToArray().Length; i++)
                {
                    foreach (var item in SecView2)
                    {
                        if (Views.ToArray()[i].Name == item.Name)
                        {
                            Views.RemoveAt(i);
                        }

                    }
                }
                foreach (Autodesk.Revit.DB.View i in Views) //add Name & id to project views list
                {
                    if (!i.IsTemplate)
                    {

                        lista_de_views.Add(i.Id);

                        nameofviews.Add(i.Name);

                        hojitas.Add(i);

                        nombres_hojas.Add(i.Name);

                    }

                }
            }

            if (form2.comboBox1.SelectedItem.ToString() == "3D views")
            {
                IEnumerable<View3D> viewList3d = from elem in new FilteredElementCollector(doc)
                                                .OfClass(typeof(View3D))
                                                .OfCategory(BuiltInCategory.OST_Views)
                                                 let type = elem as View3D
                                                 //where type.Name.Contains("SCHEDULE")
                                                 select type;

                List<Autodesk.Revit.DB.View> SecView2 = new List<Autodesk.Revit.DB.View>();
                List<Autodesk.Revit.DB.View> Views = new List<Autodesk.Revit.DB.View>();

                foreach (var view in viewList3d)
                {
                    if (!view.IsTemplate)
                    {
                        Views.Add(view);

                    }
                }
                foreach (var id1 in viewList3d)
                {
                    foreach (var viewid in ids_)
                    {
                        if (!id1.Id.IntegerValue.Equals(viewid.IntegerValue))
                        {

                        }
                        else
                        {
                            if (!SecView2.Contains(id1))
                            {
                                SecView2.Add(id1);

                            }
                        }
                    }
                }

                for (int i = 0; i < Views.ToArray().Length; i++)
                {
                    foreach (var item in SecView2)
                    {
                        if (Views.ToArray()[i].Name == item.Name)
                        {
                            Views.RemoveAt(i);
                        }

                    }
                }
                foreach (Autodesk.Revit.DB.View3D i in Views) //add Name & id to project views list
                {
                    if (!i.IsTemplate)
                    {
                        lista_de_views.Add(i.Id);

                        nameofviews.Add(i.Name);

                        hojitas.Add(i);

                        nombres_hojas.Add(i.Name);
                    }

                }
            }

            if (form2.comboBox1.SelectedItem.ToString() == "Plans")
            {
                IEnumerable<ViewPlan> viewListViewPlan = from elem in new FilteredElementCollector(doc)
                                                .OfClass(typeof(ViewPlan))
                                                .OfCategory(BuiltInCategory.OST_Views)
                                                         let type = elem as ViewPlan
                                                         //where type.Name.Contains("SCHEDULE")
                                                         select type;

                List<Autodesk.Revit.DB.View> SecView2 = new List<Autodesk.Revit.DB.View>();
                List<Autodesk.Revit.DB.View> Views = new List<Autodesk.Revit.DB.View>();

                foreach (var view in viewListViewPlan)
                {
                    if (!view.IsTemplate)
                    {
                        Views.Add(view);

                    }
                }
                foreach (var id1 in viewListViewPlan)
                {
                    foreach (var viewid in ids_)
                    {
                        if (!id1.Id.IntegerValue.Equals(viewid.IntegerValue))
                        {

                        }
                        else
                        {
                            if (!SecView2.Contains(id1))
                            {
                                SecView2.Add(id1);

                            }
                        }
                    }
                }
                for (int i = 0; i < Views.ToArray().Length; i++)
                {
                    foreach (var item in SecView2)
                    {
                        if (Views.ToArray()[i].Name == item.Name)
                        {
                            Views.RemoveAt(i);
                        }

                    }
                }
                foreach (ViewPlan i in Views) //add Name & id to project views list
                {
                    if (!i.IsTemplate)
                    {
                        lista_de_views.Add(i.Id);

                        nameofviews.Add(i.Name);

                        hojitas.Add(i);

                        nombres_hojas.Add(i.Name);
                    }

                }
            }
            if (form2.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }
            if (form2.comboBox1.SelectedItem.ToString() == "Drafting")
            {
                IEnumerable<ViewDrafting> viewdrafting_ = from elem in new FilteredElementCollector(doc)
                                                .OfClass(typeof(ViewDrafting))
                                                .OfCategory(BuiltInCategory.OST_Views)
                                                          let type = elem as ViewDrafting
                                                          //where type.Name.Contains("SCHEDULE")
                                                          select type;

                List<Autodesk.Revit.DB.View> SecView2 = new List<Autodesk.Revit.DB.View>();
                List<Autodesk.Revit.DB.View> Views = new List<Autodesk.Revit.DB.View>();

                foreach (var view in viewdrafting_)
                {
                    if (!view.IsTemplate)
                    {
                        Views.Add(view);

                    }
                }
                foreach (var id1 in viewdrafting_)
                {
                    foreach (var viewid in ids_)
                    {
                        if (!id1.Id.IntegerValue.Equals(viewid.IntegerValue))
                        {

                        }
                        else
                        {
                            if (!SecView2.Contains(id1))
                            {
                                SecView2.Add(id1);

                            }
                        }
                    }
                }
                for (int i = 0; i < Views.ToArray().Length; i++)
                {
                    foreach (var item in SecView2)
                    {
                        if (Views.ToArray()[i].Name == item.Name)
                        {
                            Views.RemoveAt(i);
                        }

                    }
                }
                foreach (ViewDrafting i in Views) //add Name & id to project views list
                {
                    if (!i.IsTemplate)
                    {
                        lista_de_views.Add(i.Id);

                        nameofviews.Add(i.Name);

                        hojitas.Add(i);

                        nombres_hojas.Add(i.Name);
                    }

                }
            }
            if (form2.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }
            foreach (FamilySymbol T in familyList) //add Name & id to tittle block list
            {
                ElementId HOJAS_ID = T.Id;
                tittleblocks_list.Add(HOJAS_ID);

                string nombre_hojas_each = T.Name;
                tittle_block_list_name.Add(T.FamilyName + " - " + nombre_hojas_each);
            }

            Form6 form = new Form6();

            foreach (var item in tittle_block_list_name)
            {
                form.comboBox1.Items.Add(item);
            }

            List<Element> store_Selected = new List<Element>(); //nombre hojas_copy

            foreach (var item in tittle_block_list_name)
            {
                form.comboBox1.Items.Add(item);
            }
            foreach (var item in nombres_hojas)
            {
                form.listBox1.Items.Add(item);
            }

            foreach (Autodesk.Revit.DB.View view in viewElems)
            {
                if (!view.IsTemplate && view.CanBePrinted && view.ViewType == ViewType.DrawingSheet)
                {
                    Debug.Print(view.Name);
                    
                    BrowserOrganization org = BrowserOrganization.GetCurrentBrowserOrganizationForSheets(doc);
                    IList<Parameter> param = org.GetOrderedParameters();


                    List<FolderItemInfo> folderfields = org.GetFolderItems(view.Id).ToList();

                    foreach (FolderItemInfo info in folderfields)
                    {
                        string groupheader = info.Name;

                        if (!gourpheader_list.Contains(groupheader))
                        {
                            gourpheader_list.Add(groupheader);
                        }
                        ElementId parameterId = info.ElementId;
                    }
                }
            }


            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            if (nombres_hojas == null)
            {
                TaskDialog.Show("Instruction!", "no views found");
                form.Close();
            }

            if (form.Equals(false))
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            int index_tBox = form.comboBox1.SelectedIndex;
            ElementId choose_Tblock = tittleblocks_list.ElementAtOrDefault(index_tBox);

            if (index_tBox == -1)
            {
                choose_Tblock = tittleblocks_list.ElementAtOrDefault(0);
            }

            List<ElementId> viewport_views = new List<ElementId>();
            List<Element> viewportEle = new List<Element>();
            List<Element> ScheduleEle = new List<Element>();

            foreach (string item in form.listBox2.Items)
            {
                string nam = item;
                foreach (Element i in hojitas)
                {
                    if (i.Name == nam)
                    {
                        Element a = i;
                        viewportEle.Add(a);
                        ElementId BC = i.Id;
                        viewport_views.Add(BC);
                        string name = i.Name;

                    }
                }
            }

            List<string> fromSelected = new List<string>();

            //--------------------------------------------------------------------------------------------------------------------
            List<XYZ> Bboxcenter = new List<XYZ>();
            List<XYZ> locationInsheet = new List<XYZ>();
            IEnumerable<int> sequencePoints = Enumerable.Range(0, 2);
            List<double> count = new List<double>();
            List<Viewport> countPorts = new List<Viewport>();


            List<ElementId> sobrasId = new List<ElementId>();
            List<ElementId> sobrasId2 = new List<ElementId>();
            List<ElementId> sobrasId3 = new List<ElementId>();
            List<Element> sobras = new List<Element>();
            List<double> numHoj = new List<double>();

            double uno1 = 1;
            numHoj.Add(uno1);
            int j = 0;

          
          

            List<ViewSheet> CreatedSheets = new List<ViewSheet>();
            Parameter p = null;

            if (viewportEle.ToArray().Length == 0)
            {
                TaskDialog.Show("Fail", "You must chose a Revit view to place in a sheet");
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            //if (form.radioButton2.Checked == false && form.radioButton2.Checked == false)
            //{
            //    TaskDialog.Show("Fail", "You must select a type of arrangement for the Revit views");
            //    return Autodesk.Revit.UI.Result.Cancelled;
            //}


            if (form.radioButton1.Checked == true)
            {
                //--------------------------------CREATION OF SHEET AND VIEW TO POINT 0,0,0----------------------------------
                using (Transaction t = new Transaction(doc, "Create one RevitSheet"))
                {
                    t.Start();
                    double yy1 = 0.5;
                    double xx1 = 2.2;
                    foreach (int item in numHoj)
                    {
                        try
                        {
                            if (viewportEle.ToArray().Length == 0)
                            {
                                TaskDialog.Show("Fail", "You must chose a Revit view to place in a sheet");
                                return Autodesk.Revit.UI.Result.Cancelled;
                            }

                            start1: ViewSheet sheet2 = ViewSheet.Create(doc, choose_Tblock);
                            foreach (Element v in viewportEle)
                            {
                                if (yy1 >= 0.5)
                                {
                                    if (xx1 == 2.2 && yy1 == 2.0)
                                    {

                                        goto start1;
                                    }
                                    while (yy1 >= 2.5)
                                    {
                                        yy1 = 0.5;
                                        xx1 = 2.2;
                                    }

                                    if (Viewport.CanAddViewToSheet(doc, sheet2.Id, v.Id))
                                    {
                                        Viewport viewport = Viewport.Create(doc, sheet2.Id, v.Id, new XYZ(xx1, yy1, 0));
                                        //p = viewport.LookupParameter("Detail Number");
                                        //p.SetValueString("100");
                                    }
                                    else
                                    {
                                        TaskDialog.Show("Warning", "The view is already placed on a sheet");
                                       
                                    }
                                }
                                xx1 = xx1 - 0.5;
                                while (xx1 <= 0)
                                {
                                    yy1 = yy1 + 0.5;
                                    xx1 = 2.2;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            TaskDialog.Show("Tittle block", "You must chose a title block");
                            throw;
                        }
                    }
                    doc.Regenerate();
                    t.Commit();
                    TaskDialog.Show("Done","One sheet was created");
                   
                }
            }

            

            if (form.radioButton2.Checked == true)
            {
                using (Transaction t = new Transaction(doc, "CreateSheetByPlan"))
                {
                    int numero = 0;
                    t.Start();
                    foreach (var view in hojitas)
                    {
                        if (viewportEle.ToArray().Length == 0)
                        {
                            TaskDialog.Show("Fail", "You must chose a Revit view to place in a sheet");
                            return Autodesk.Revit.UI.Result.Cancelled;
                        }

                        foreach (var VP in viewportEle)
                        {
                            if (VP.Id == view.Id)
                            {
                                ViewSheet sheet2 = ViewSheet.Create(doc, choose_Tblock);
                                numero++;
                                if (Viewport.CanAddViewToSheet(doc, sheet2.Id, view.Id))
                                {
                                    Viewport viewport = Viewport.Create(doc, sheet2.Id, view.Id, new XYZ(0, 0, 0));
                                }
                            }
                        }
                    }
                    t.Commit();
                    TaskDialog.Show("Done", numero + " sheets were created");
                    
                }
            }
           
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Duplicate_0ne_sheet : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("F5FB1A7F-8410-6562-8420-04AE05C1239E"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            

            //---------------------------------------- FILTERS ------------------------------------
            IEnumerable<FamilySymbol> familyList = from elem in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_TitleBlocks)
                                                   let type = elem as FamilySymbol
                                                   select type;

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> hojas = collector.OfClass(typeof(Autodesk.Revit.DB.View)).ToElements();

            FilteredElementCollector fec = new FilteredElementCollector(doc);
            fec.OfClass(typeof(ViewSection));

            FilteredElementCollector fec2 = new FilteredElementCollector(doc);
            fec2.OfClass(typeof(ViewSheet));



            var viewPlans = fec.Cast<ViewSection>().Where<ViewSection>(vp => vp.IsTemplate);
            IEnumerable<Autodesk.Revit.DB.View> viewList = from elem in new FilteredElementCollector(doc)
                                                 .OfClass(typeof(Autodesk.Revit.DB.View))
                                                 .OfCategory(BuiltInCategory.OST_Views)
                                                           let type = elem as Autodesk.Revit.DB.View
                                                           where type.Name != null
                                                           select type;

            IEnumerable<ViewSheet> viewSheet = from elem in new FilteredElementCollector(doc)
                                                 .OfClass(typeof(ViewSheet))
                                                 .OfCategory(BuiltInCategory.OST_Sheets)
                                               let type = elem as ViewSheet
                                               where type.Name != null
                                               select type;

            IEnumerable<TextNote> textnoteslist = from elem in new FilteredElementCollector(doc)
                                               .OfClass(typeof(TextNote)).OfCategory(BuiltInCategory.OST_TextNotes)
                                                  let type = elem as TextNote
                                                  where type.Name != null
                                                  select type;




            //TaskDialog.Show("point1", "point1");

            FamilySymbol FamilySymbol = null;
            IList<Element> m_alltitleblocks = new List<Element>();
            IList<Element> ElementsOnSheet = new List<Element>();
            List<ViewSheet> ViewSchedule_LIST = new List<ViewSheet>();


            List<Viewport> ViewPorts_ = new List<Viewport>();
            List<ElementId> Ids_ = new List<ElementId>();
            List<ElementId> vtype_ = new List<ElementId>();

            ICollection<ElementId> Ids2_ = new List<ElementId>();

            List<Autodesk.Revit.DB.View> Total_viewcount_onproject = new List<Autodesk.Revit.DB.View>();
            List<Autodesk.Revit.DB.View> view_to_copy = new List<Autodesk.Revit.DB.View>();
            List<ElementId> views_already_copied = new List<ElementId>();
            List<BoundingBoxXYZ> sectionBox = new List<BoundingBoxXYZ>();
            List<XYZ> vp_centers = new List<XYZ>();
            List<ViewSchedule> schedule_existing_insheet = new List<ViewSchedule>();

            Form18 form = new Form18();

            ViewFamilyType vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Section == x.ViewFamily);

            ViewSheet activeViewSheet = doc.ActiveView as ViewSheet;

            FilteredElementCollector col_sheetele = new FilteredElementCollector(doc, activeViewSheet.Id);
            var scheduleSheetInstances = col_sheetele.OfClass(typeof(ScheduleSheetInstance)).ToElements().OfType<ScheduleSheetInstance>();

            IList<ElementId> rev_Id = activeViewSheet.GetAllRevisionIds();



            List<Viewport> viewports = new List<Viewport>();
            List<ElementId> views = new List<ElementId>();


            List<XYZ> schudule_pt = new List<XYZ>();
            foreach (var scheduleSheetInstance in scheduleSheetInstances)
            {
                if (!scheduleSheetInstance.Name.Contains("Revision"))
                {
                    //pt = scheduleSheetInstance.Point;

                    var scheduleId = scheduleSheetInstance.ScheduleId;
                    if (scheduleId == ElementId.InvalidElementId)
                        continue;
                    var viewSchedule_ = doc.GetElement(scheduleId) as ViewSchedule;
                    schedule_existing_insheet.Add(viewSchedule_);
                }
            }
            foreach (var scheduleSheetInstance in scheduleSheetInstances)
            {
                if (!scheduleSheetInstance.Name.Contains("Revision"))
                {
                    schudule_pt.Add(scheduleSheetInstance.Point);
                }
            }

            List<ElementId> text_type = new List<ElementId>();
            List<XYZ> text_pt = new List<XYZ>();
            List<TextNote> textnotessearch = new List<TextNote>();
            foreach (var item in textnoteslist)
            {
                if (item.OwnerViewId == activeViewSheet.Id)
                {
                    text_type.Add(item.TextNoteType.Id);
                    text_pt.Add(item.Coord /*as XYZ*/);
                    textnotessearch.Add(item);
                }

            }

            try
            {
                ICollection<ElementId> views_ = activeViewSheet.GetAllPlacedViews();
                foreach (var item in views_)
                {
                    views.Add(item);
                }


                IList<Viewport> viewports__ = new FilteredElementCollector(doc).OfClass(typeof(Viewport)).Cast<Viewport>()
            .Where(q => q.SheetId == activeViewSheet.Id).ToList();
                foreach (var item in viewports__)
                {
                    viewports.Add(item);
                }
            }
            catch (Exception)
            {
                TaskDialog.Show("Warning", "Active View must be a sheet");
                throw;
            }



            //TaskDialog.Show("point2", "point2");


            foreach (var VID in views)
            {
                foreach (var VP in viewports)
                {
                    if (VP.ViewId == VID)
                    {
                        XYZ center = VP.GetBoxCenter();
                        vp_centers.Add(center);

                    }
                }
            }

            // foreach (var item in viewList)
            // {
            //     foreach (var item2 in new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.View)).Cast<Autodesk.Revit.DB.View>()
            //.Where(q => q.Id == item.Id).ToList())
            //     {
            //         Total_viewcount_onproject.Add(item2);
            //     }
            // }

            //TaskDialog.Show("point3", "point3");

            foreach (var view_onproject in viewList)
            {
                foreach (var view_ID_oncurrentpage in views)
                {
                    if (view_onproject.Id == view_ID_oncurrentpage)
                    {
                        view_to_copy.Add(view_onproject);
                        ViewType vt = view_onproject.ViewType;



                        BoundingBoxXYZ room_box = view_onproject.get_BoundingBox(null);
                        sectionBox.Add(room_box);
                        vtype_.Add(view_onproject.ViewTemplateId);

                    }
                }
            }

            //TaskDialog.Show("point4", "point4");

            foreach (Element e in new FilteredElementCollector(doc).OwnedByView(/*sheet_.Id*/activeViewSheet.Id))
            {
                if (e.Category != null && e.Category.Name == "Viewports")
                {
                    ViewPorts_.Add(e as Viewport);
                }

                ElementsOnSheet.Add(e);
            }



            foreach (Element el in ElementsOnSheet)
            {
                foreach (FamilySymbol Fs in /*m_alltitleblocks*/ familyList)
                {
                    if (el.GetTypeId().IntegerValue == Fs.Id.IntegerValue)
                    {
                        FamilySymbol = Fs;
                    }
                }
            }

            form.ShowDialog();

            //TaskDialog.Show("point5", "point5");

            using (Transaction t = new Transaction(doc, "Create RevitSheet"))
            {
                t.Start();




                foreach (var item in view_to_copy)
                {

                    ElementId view_ = item.Duplicate(ViewDuplicateOption.WithDetailing);
                    views_already_copied.Add(view_);

                }

                //TaskDialog.Show("point6", "point6");

                IEnumerable<Autodesk.Revit.DB.View> viewList_new = from elem in new FilteredElementCollector(doc)
                                                 .OfClass(typeof(Autodesk.Revit.DB.View))
                                                 .OfCategory(BuiltInCategory.OST_Views)
                                                                   let type = elem as Autodesk.Revit.DB.View
                                                                   where type.Name != null
                                                                   select type;

                ViewSheet sheet2 = ViewSheet.Create(doc, FamilySymbol.Id);
                try
                {
                    sheet2.Name = activeViewSheet.Name + "copy";
                    sheet2.SheetNumber = activeViewSheet.SheetNumber + "copy";
                }
                catch (Exception)
                {

                    MessageBox.Show("Name or Number might be already in use!", "");
                }

                try
                {
                    for (int i = 0; i < schedule_existing_insheet.ToArray().Length; i++)
                    {
                        ScheduleSheetInstance.Create(doc, sheet2.Id, schedule_existing_insheet.ToArray()[i].Id, schudule_pt.ToArray()[i]);
                    }
                }
                catch (Exception)
                {

                    MessageBox.Show("The program can not copy schedule, try hidding them before copying the sheet", "");
                }

                /*ElementId defaultTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);*/
                try
                {
                    for (int i = 0; i < textnotessearch.ToArray().Length; i++)
                    {
                        TextNote.Create(doc, sheet2.Id, text_pt.ToArray()[i], textnotessearch.ToArray()[i].Text, /*defaultTypeId*/ text_type.ToArray()[i]);
                    }
                }
                catch (Exception)
                {

                    MessageBox.Show("The program can not copy textnotes, try hidding them before copying the sheet", "");
                }




                if (form.radioButton1.Checked)
                {

                    sheet2.SetAdditionalRevisionIds(rev_Id);

                }



                if (sheet2.LookupParameter("View Organization") != null)
                {
                    string parametro = activeViewSheet.LookupParameter("View Organization").AsString();
                    Parameter param = sheet2.LookupParameter("View Organization");
                    param.Set(parametro);
                }

                if (sheet2.LookupParameter("Drawing Series") != null)
                {
                    string parametro = activeViewSheet.LookupParameter("Drawing Series").AsString();
                    Parameter param2 = sheet2.LookupParameter("Drawing Series");
                    param2.Set(parametro);
                }

                int centerpt = 0;
                foreach (var view_onproject in viewList_new)
                {
                    foreach (var view_ID_oncurrentpage in views_already_copied)
                    {
                        if (view_onproject.Id == view_ID_oncurrentpage)
                        {


                            if (Viewport.CanAddViewToSheet(doc, sheet2.Id, view_onproject.Id))
                            {
                                Viewport viewport = Viewport.Create(doc, sheet2.Id, view_onproject.Id, vp_centers.ToArray()[centerpt]);
                            }
                            centerpt++;
                        }
                    }
                }
                //TaskDialog.Show("point8", "point8");

                doc.Regenerate();
                t.Commit();
                uidoc.ActiveView = sheet2;
                TaskDialog.Show("Completed", "1 Sheet copied");
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class CreatMultipleSheet : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("F5FB1A7F-8220-4862-8820-04AE05C1239E"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;

            //try
            //{
            //    string filename = @"T:\Transfer\lopez\Book1.xlsx";
            //    using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
            //    {
            //        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

            //        int column = 7;
            //        int number = Convert.ToInt32(sheet.Cells[2, column].Value);
            //        sheet.Cells[2, column].Value = (number + 1); ;
            //        package.Save();
            //    }
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Excel file not found", "");
            //}

            //---------------------------------------- FILTERS ------------------------------------
            IEnumerable<FamilySymbol> TittleBlock_List = from elem in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_TitleBlocks)
                                                         let type = elem as FamilySymbol select type;

            List<Element> viewElems = new List<Element>();
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            viewElems.AddRange(collector.OfClass(typeof(Autodesk.Revit.DB.View)).ToElements());

            List<ElementId> tittleblocks_list = new List<ElementId>();
            List<String> tittle_block_list_name = new List<String>();
            List<string> gourpheader_list = new List<string>();

            int distance = 0;
            Form9 form = new Form9();

            foreach (FamilySymbol T in TittleBlock_List) 
            {
                ElementId HOJAS_ID = T.Id;
                tittleblocks_list.Add(HOJAS_ID);

                string nombre_hojas_each = T.Name;
                tittle_block_list_name.Add(T.FamilyName + " - " + nombre_hojas_each);
            }

            foreach (var item in tittle_block_list_name)
            {
                form.comboBox1.Items.Add(item);
            }

            List<Element> store_Selected = new List<Element>(); //nombre hojas_copy

            foreach (var item in tittle_block_list_name)
            {
                form.comboBox1.Items.Add(item);
            }

            int numero_hojas = 0;
            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            if (form.DialogResult == System.Windows.Forms.DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            if (form.Equals(false))
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            int index_tBox = form.comboBox1.SelectedIndex;
            ElementId choose_Tblock = tittleblocks_list.ElementAtOrDefault(index_tBox);

            numero_hojas = (int)form.numericUpDown1.Value;

            using (Transaction t = new Transaction(doc, "Create RevitSheet"))
            {
                t.Start();

                {
                    for (int i = 0; i < numero_hojas; i++)
                    {

                        ViewSheet sheet2 = ViewSheet.Create(doc, choose_Tblock);
                        try
                        {
                            //sheet2.SheetNumber = form.comboBox2.SelectedItem.ToString() + (distance + i);
                            sheet2.Name = "New sheet " + i.ToString();
                        }
                        catch
                        {
                            TaskDialog.Show("warning", "name Exists");
                        }
                    }
                }
                doc.Regenerate();
                t.Commit();

                return Autodesk.Revit.UI.Result.Succeeded;
            }
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class copy_schedule : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("6C22CC72-A167-4819-AAF1-A178F6B44BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            //try
            //{
            //    string filename = @"T:\Transfer\lopez\Book1.xlsx";
            //    using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
            //    {
            //        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

            //        int column = 20;
            //        int number = Convert.ToInt32(sheet.Cells[2, column].Value);
            //        sheet.Cells[2, column].Value = (number + 1); ;
            //        package.Save();
            //    }
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Excel file not found", "");
            //}

            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule));

            IEnumerable<ViewSchedule> views = from Autodesk.Revit.DB.ViewSchedule f in collector
                                              where (f.ViewType == ViewType.Schedule /*&& f.ViewName.Equals("Unit Default")*/)
                                              select f;

            IEnumerable<ViewSheet> viewSheet = from elem in new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)).OfCategory(BuiltInCategory.OST_Sheets)
                                               let type = elem as ViewSheet
                                               where type.Name != null
                                               select type;

            IEnumerable<Autodesk.Revit.DB.View> legends = from elem in new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.View)).OfCategory(BuiltInCategory.OST_Views)
                                                          let type = elem as Autodesk.Revit.DB.View
                                                          where type.ViewType == ViewType.Legend
                                                          select type;

            Form19 form = new Form19();
            ViewSheet activeViewSheet = null;

            if (doc.ActiveView.ViewType == ViewType.DrawingSheet)
            {
                activeViewSheet = doc.ActiveView as ViewSheet;
            }
            else
            {
                TaskDialog.Show("warning", "The active view must be a view sheet");
                return Autodesk.Revit.UI.Result.Cancelled;
            }
            List<ViewSchedule> schedule_existing_insheet = new List<ViewSchedule>();
            List<ImageType> img_existing_insheet = new List<ImageType>();
            List<Autodesk.Revit.DB.View> Legend_list = new List<Autodesk.Revit.DB.View>();

            FilteredElementCollector col_sheetele = new FilteredElementCollector(doc, activeViewSheet.Id);
            var scheduleSheetInstances = col_sheetele.OfClass(typeof(ScheduleSheetInstance)).ToElements().OfType<ScheduleSheetInstance>();
            var ImageType_ = col_sheetele.OfClass(typeof(ImageType)).ToElements().OfType<ImageType>();
            XYZ schudule_pt = null;
            XYZ legend_pt = null;

            List<Viewport> viewports = new List<Viewport>();
            List<ElementId> views_list = new List<ElementId>();

            try
            {
                ICollection<ElementId> views_ = activeViewSheet.GetAllPlacedViews();
                foreach (var item in views_)
                {
                    views_list.Add(item);
                }


                IList<Viewport> viewports__ = new FilteredElementCollector(doc).OfClass(typeof(Viewport)).Cast<Viewport>()
            .Where(q => q.SheetId == activeViewSheet.Id).ToList();
                foreach (var item in viewports__)
                {

                    viewports.Add(item);
                }
            }
            catch (Exception)
            {
                TaskDialog.Show("Warning", "Active View must be a sheet");
                throw;
            }

            foreach (var VP in viewports)
            {
                foreach (var item in legends)
                {
                    if (VP.ViewId == item.Id)
                    {
                        form.listBox1.Items.Add(item.Name);
                    }
                }
            }

            foreach (var scheduleSheetInstance in scheduleSheetInstances)
            {
                if (!scheduleSheetInstance.Name.Contains("Revision"))
                {
                    var scheduleId = scheduleSheetInstance.ScheduleId;
                    if (scheduleId == ElementId.InvalidElementId)
                        continue;
                    var viewSchedule_ = doc.GetElement(scheduleId) as ViewSchedule;
                    schedule_existing_insheet.Add(viewSchedule_);
                }
            }

            foreach (var item in schedule_existing_insheet)
            {
                if (!item.Name.Contains("Revision"))
                {
                    form.listBox1.Items.Add(item.Name);
                }
            }

            foreach (var item in viewSheet)
            {
                form.listBox2.Items.Add(item.Name);
            }

            form.ShowDialog();



            Autodesk.Revit.DB.View vID = null;
            foreach (var item in legends)
            {
                if (item.Name == form.listBox1.SelectedItem.ToString())
                {
                    vID = item;
                }
            }



            if (vID != null)
            {
                foreach (var VP in viewports)
                {

                    if (VP.ViewId == vID.Id)
                    {
                        legend_pt = VP.GetBoxCenter();
                    }
                }
            }

            foreach (var scheduleSheetInstance in scheduleSheetInstances)
            {
                if (!scheduleSheetInstance.Name.Contains("Revision"))
                {
                    if (form.listBox1.SelectedItem.ToString() == scheduleSheetInstance.Name)
                    {
                        schudule_pt = scheduleSheetInstance.Point;
                    }
                }
            }

            ViewSchedule schedule_selected = null;
            foreach (string item in form.listBox1.SelectedItems)
            {
                string nam = item;
                foreach (Element i in schedule_existing_insheet)
                {
                    if (i.Name == nam)
                    {

                        schedule_selected = i as ViewSchedule;
                    }
                }
            }

            List<ViewSheet> viewsheet_selected = new List<ViewSheet>();
            foreach (string item in form.listBox2.SelectedItems)
            {

                string nam = item;
                foreach (ViewSheet i in viewSheet)
                {
                    if (i.Name == nam)
                    {
                        if (!viewsheet_selected.Contains(i as ViewSheet))
                        {
                            viewsheet_selected.Add(i as ViewSheet);
                        }
                    }
                }
            }

            using (Transaction trans = new Transaction(doc, "ViewDuplicate"))
            {
                trans.Start();

                if (schedule_selected != null)
                {
                    for (int i = 0; i < viewsheet_selected.ToArray().Length; i++)
                    {
                        try
                        {
                            ScheduleSheetInstance.Create(doc, viewsheet_selected.ToArray()[i].Id, schedule_selected.Id, schudule_pt);
                        }
                        catch (Exception)
                        {

                            TaskDialog.Show("warning", "The selected sheet template does not contain a schedule");
                        }
                    }
                    trans.Commit();
                }

                if (vID != null)
                {
                    for (int i = 0; i < viewsheet_selected.ToArray().Length; i++)
                    {
                        try
                        {
                            Viewport.Create(doc, viewsheet_selected.ToArray()[i].Id, vID.Id, legend_pt);
                        }
                        catch (Exception)
                        {

                            TaskDialog.Show("warning", "The selected sheet template does not contain a legend");
                        }
                    }
                    trans.Commit();
                }

            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class DeleteAllViews : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F22CC78-A557-4819-AAF1-A678F6B22BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            //try
            //{
            //    string filename = @"T:\Transfer\lopez\Book1.xlsx";
            //    using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
            //    {
            //        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

            //        int column = 12;
            //        int number = Convert.ToInt32(sheet.Cells[2, column].Value);
            //        sheet.Cells[2, column].Value = (number + 1); ;
            //        package.Save();
            //    }
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Excel file not found", "");
            //}

            Form15 form = new Form15();
            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;

            try
            {
                Autodesk.Revit.DB.View viewTemplate = (from v in new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.View)).Cast<Autodesk.Revit.DB.View>()
                                                       where !v.IsTemplate && v.Name == "Home"
                                                       select v).First();
                uidoc.ActiveView = viewTemplate;
            }
            catch (Exception)
            {
                TaskDialog.Show("Warning", "'Home' view was not found in this project, name a view Home and try again");
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> collection = collector.OfClass(typeof(Autodesk.Revit.DB.View)).ToElements();

            using (Transaction t = new Transaction(doc, "Delete All Sheets and Views"))
            {
                t.Start();
                int x = 0;
                foreach (Element e in collection)
                {
                    try
                    {
                        Autodesk.Revit.DB.View view = e as Autodesk.Revit.DB.View;
                        doc.Delete(e.Id);
                        x += 1;
                    }
                    catch (Exception ex)
                    {
                    }
                }
                t.Commit();
                //doc.Regenerate();

                TaskDialog.Show("DeleteAllSheetsViews", "Views & Sheets Deleted: " + x.ToString());
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ReNumbering : IExternalCommand
    {

        private Parameter getParameterForReference(Autodesk.Revit.DB.Document doc, Reference r)
        {
            Element e = doc.GetElement(r);
            Parameter p = null;
            if (e is Autodesk.Revit.DB.Grid)
                p = e.LookupParameter("Name");
            else if (e is Room)
                p = e.LookupParameter("Number");
            else if (e is FamilyInstance)
                p = e.LookupParameter("Mark");
            else if (e is Viewport) // Viewport class is new to Revit 2013 API
                p = e.LookupParameter("Detail Number");
            else
            {
                TaskDialog.Show("Error", "Unsupported element");
                return null;
            }
            return p;
        }
        private void setParameterToValue(Parameter p, int i)
        {
            if (p.StorageType == StorageType.Integer)
                p.Set(i);
            else if (p.StorageType == StorageType.String)
                p.Set(i.ToString());
        }

        static AddInId appId = new AddInId(new Guid("F6FB1A7F-8410-6562-8420-06AE05C1239E"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            //try
            //{
            //    string filename = @"T:\Transfer\lopez\Book1.xlsx";
            //    using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
            //    {
            //        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

            //        int column = 10;
            //        int number = Convert.ToInt32(sheet.Cells[2, column].Value);
            //        sheet.Cells[2, column].Value = (number + 1); ;
            //        package.Save();
            //    }
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Excel file not found", "");
            //}

            IList<Reference> refList = new List<Reference>();

            TaskDialog.Show("Grids - family instances - Viewports", "Select elements in order to be renumbered and then press ESCAPE to finish (the sequence will start with the number of the first selected element)");

            try
            {
                while (true)
                    refList.Add(uidoc.Selection.PickObject(ObjectType.Element, "Select elements in order to be renumbered. ESC when finished."));
            }
            catch
            { }

            using (Transaction t = new Transaction(doc, "Renumber"))
            {
                t.Start();
                // need to avoid encountering the error "The name entered is already in use. Enter a unique name."
                // for example, if there is already a grid 2 we can't renumber some other grid to 2
                // therefore, first renumber everny element to a temporary name then to the real one
                int ctr = 1;
                int startValue = 0;
                foreach (Reference r in refList)
                {
                    Parameter p = getParameterForReference(doc, r);



                    // get the value of the first element to use as the start value for the renumbering in the next loop
                    if (ctr == 1)
                        startValue = Convert.ToInt16(p.AsString());

                    setParameterToValue(p, ctr + 12345); // hope this # is unused (could use Failure API to make this more robust
                    ctr++;
                }

                ctr = startValue;
                foreach (Reference r in refList)
                {
                    Parameter p = getParameterForReference(doc, r);
                    setParameterToValue(p, ctr);
                    ctr++;
                }
                t.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Wall_Angle_to : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F46AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            List<Element> ele = new List<Element>();
            List<Element> ele2 = new List<Element>();
            IList<Reference> refList = new List<Reference>();

            TaskDialog.Show("!", "Select a reference Grid to find orthogonal walls");

            Autodesk.Revit.DB.Grid levelBelow = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select Grid")) as Autodesk.Revit.DB.Grid;

            Autodesk.Revit.DB.Curve dircurve = levelBelow.Curve;
            Autodesk.Revit.DB.Line line = dircurve as Autodesk.Revit.DB.Line;
            XYZ dir = line.Direction;

            XYZ origin = line.Origin;
            XYZ viewdir = line.Direction;
            XYZ up = XYZ.BasisZ;
            XYZ right = up.CrossProduct(viewdir);

            foreach (Element wall in new FilteredElementCollector(doc).OfClass(typeof(Wall)))
            {
                LocationCurve lc = wall.Location as LocationCurve;
                Autodesk.Revit.DB.Transform curveTransform = lc.Curve.ComputeDerivatives(0.5, true);

                try
                {
                    XYZ origin2 = curveTransform.Origin;
                    XYZ viewdir2 = curveTransform.BasisX.Normalize();
                    XYZ viewdir2_back = curveTransform.BasisX.Normalize() * -1;

                    XYZ up2 = XYZ.BasisZ;
                    XYZ right2 = up.CrossProduct(viewdir2);
                    XYZ left2 = up.CrossProduct(viewdir2 * -1);

                    double y_onverted = Math.Round(-1 * viewdir2.X);

                    if (viewdir.IsAlmostEqualTo(right2/*, 0.3333333333*/))
                    {
                        ele.Add(wall);
                    }
                    if (viewdir.IsAlmostEqualTo(left2))
                    {
                        ele.Add(wall);
                    }
                    if (viewdir.IsAlmostEqualTo(viewdir2))
                    {
                        ele.Add(wall);
                    }
                    if (viewdir.IsAlmostEqualTo(viewdir2_back))
                    {
                        ele.Add(wall);
                    }


                }
                catch (Exception)
                {
                    return Autodesk.Revit.UI.Result.Cancelled;
                }
            }
            uidoc.Selection.SetElementIds(ele.Select(q => q.Id).ToList());
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class text_upper : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;



            FilteredElementCollector rooms1 = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(SpatialElement)).OfCategory(BuiltInCategory.OST_Rooms)
                                             ;
            ICollection<Element> room2 = rooms1.ToElements();

            IEnumerable<TextNote> textnoteslist = from elem in new FilteredElementCollector(doc)
                                               .OfClass(typeof(TextNote)).OfCategory(BuiltInCategory.OST_TextNotes)
                                                  let type = elem as TextNote
                                                  where type.Name != null
                                                  select type;

            IEnumerable<Autodesk.Revit.DB.View> view_list = from elem in new FilteredElementCollector(doc)
                                                .OfClass(typeof(Autodesk.Revit.DB.View))
                                                .OfCategory(BuiltInCategory.OST_Views)
                                                            let type = elem as Autodesk.Revit.DB.View
                                                            where type.Name != null
                                                            select type;

            IEnumerable<Autodesk.Revit.DB.Grid> Grid_list = from elem in new FilteredElementCollector(doc)
                                                .OfClass(typeof(Autodesk.Revit.DB.Grid))
                                                .OfCategory(BuiltInCategory.OST_Grids)
                                                            let type = elem as Autodesk.Revit.DB.Grid
                                                            where type.Name != null
                                                            select type;

            IEnumerable<FamilySymbol> familyList = from elem in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_TitleBlocks)
                                                   let type = elem as FamilySymbol
                                                   select type;

            FamilyInstance fi = new FilteredElementCollector(doc, doc.ActiveView.Id).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_TitleBlocks)
      .FirstOrDefault() as FamilyInstance;


            Form25 form = new Form25();
            form.ShowDialog();



            List<TextNote> textnotessearch = new List<TextNote>();
            foreach (var item in textnoteslist)
            {

                textnotessearch.Add(item);
            }

            //if (doc.ActiveView.ViewType == ViewType. ViewSheet)
            //{

            //}

            ViewType type_ = doc.ActiveView.ViewType;

            ViewSheet activeViewSheet = doc.ActiveView as ViewSheet;

            Autodesk.Revit.DB.View activeViewSheet_ = doc.ActiveView as Autodesk.Revit.DB.View;

            List<Viewport> viewports = new List<Viewport>();
            List<ElementId> views = new List<ElementId>();
            List<Autodesk.Revit.DB.Grid> g_list = new List<Autodesk.Revit.DB.Grid>();
            List<Autodesk.Revit.DB.Grid> tblock_list = new List<Autodesk.Revit.DB.Grid>();
            FamilySymbol FamilySymbol = null;
            IList<Element> ElementsOnSheet = new List<Element>();


            ICollection<ElementId> views_;

            try
            {
                views_ = activeViewSheet.GetAllPlacedViews();
            }
            catch (Exception)
            {
                MessageBox.Show("Warning", "You must be on a Revit sheet before running this tool");
                return Autodesk.Revit.UI.Result.Cancelled;
            }


            foreach (var item in views_)
            {

                views.Add(item);
            }
            foreach (Autodesk.Revit.DB.Grid item in Grid_list)
            {
                g_list.Add(item);
            }


            IList<Viewport> viewports__ = new FilteredElementCollector(doc).OfClass(typeof(Viewport)).Cast<Viewport>()
        .Where(q => q.SheetId == activeViewSheet.Id).ToList();
            foreach (var item in viewports__)
            {

                viewports.Add(item);
            }

            foreach (Element e in new FilteredElementCollector(doc).OwnedByView(/*sheet_.Id*/activeViewSheet.Id))
            {
                //if (e.Category != null && e.Category.Name == "Viewports")
                //{
                //    ViewPorts_.Add(e as Viewport);
                //}

                ElementsOnSheet.Add(e);
            }

            using (Transaction trans = new Transaction(doc, "Capital letters"))
            {
                trans.Start();

                if (form.checkBox1.Checked)
                {
                    foreach (var text_ in textnotessearch)
                    {
                        if (text_.OwnerViewId == activeViewSheet.Id)
                        {
                            string upper = text_.Text.ToUpper();
                            text_.Text = upper;
                        }

                    }
                }

                if (form.checkBox1.Checked)
                {
                    foreach (var text_ in textnotessearch)
                    {
                        foreach (var view_id in views)
                        {
                            if (text_.OwnerViewId == view_id)
                            {
                                string upper = text_.Text.ToUpper();
                                text_.Text = upper;
                            }
                        }
                    }
                }

                if (form.checkBox2.Checked)
                {
                    foreach (var item in room2)
                    {
                        string upper = item.Name.ToUpper();
                        item.Name = upper;
                    }
                }

                if (form.checkBox3.Checked)
                {
                    foreach (var view in view_list)
                    {
                        foreach (var viewid in views)
                        {
                            if (viewid == view.Id)
                            {
                                string upper = view.Name.ToUpper();
                                view.Name = upper;
                            }
                        }
                    }
                }

                if (form.checkBox4.Checked)
                {
                    foreach (Autodesk.Revit.DB.Grid item in g_list)
                    {
                        string upper = item.Name.ToUpper();
                        item.Name = upper;
                    }
                }

                if (form.checkBox6.Checked)
                {
                    string upper2 = fi.LookupParameter("Sheet Name").AsString().ToUpper();
                    Parameter param = fi.LookupParameter("Sheet Name") /*fi.get_Parameter("Sheet Name")*/;
                    param.Set(upper2);
                }

                //foreach (var item in viewports__)
                //{
                //    string upper = item.LookupParameter("View Name").ToString().ToUpper();

                //    p = item.LookupParameter("Title on Sheet");

                //    p.SetValueString("lfdkgjsdf");


                //    viewports.Add(item);
                //}

                trans.Commit();
            }


            //string appdataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            //string folderPath = Path.Combine(appdataFolder, @"Autodesk\Revit\Addins\2019\AJC_Commands\img\solar analisys guide.pdf");
            //Process.Start($"{folderPath}");

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class TotalLenght : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F92CC78-A127-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            double length = 0;

            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();

            //if (true)
            //{

            //}
            //TaskDialog.Show("Error", ids.ToList().Count ==);
            //return Autodesk.Revit.UI.Result.Cancelled;

            foreach (ElementId id in ids)
            {
                Element e = doc.GetElement(id);
                Parameter lengthParam = e.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                if (lengthParam == null)
                    continue;
                length += lengthParam.AsDouble();
            }
            string lengthWithUnits = UnitFormatUtils.Format(doc.GetUnits(), UnitType.UT_Length, length, false, false);
            TaskDialog.Show("Length", ids.Count + " elements = " + lengthWithUnits);

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class isolate : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            ElementId elementIdToIsolate = null;
            try
            {
                elementIdToIsolate = uidoc.Selection.PickObject(ObjectType.Element, "Select element or ESC to reset the view").ElementId;
            }
            catch
            {

            }


            OverrideGraphicSettings ogsFade = new OverrideGraphicSettings();
            ogsFade.SetSurfaceTransparency(80);
            ogsFade.SetSurfaceForegroundPatternVisible(false);
            ogsFade.SetSurfaceBackgroundPatternVisible(false);
            ogsFade.SetHalftone(true);
            

            OverrideGraphicSettings ogsIsolate = new OverrideGraphicSettings();
            ogsIsolate.SetSurfaceTransparency(0);
            ogsIsolate.SetSurfaceForegroundPatternVisible(true);
            ogsIsolate.SetSurfaceBackgroundPatternVisible(true);
            ogsIsolate.SetHalftone(false);

            using (Transaction t = new Transaction(doc, "Isolate with Fade"))
            {
                //t.Start();
                //foreach (Element e in new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType())
                //{
                //    if (e.Id != elementIdToIsolate || elementIdToIsolate != null)
                //        doc.ActiveView.SetElementOverrides(e.Id, ogsIsolate);
                //    else
                //        doc.ActiveView.SetElementOverrides(e.Id, ogsFade);
                //}
                //t.Commit();
                t.Start();
                doc.ActiveView.SetElementOverrides(elementIdToIsolate, ogsFade);
                
                t.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    public class Hide_ele : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            List<Element> ele_ = new List<Element>();

            foreach (var item in ids)
            {

                ele_.Add(doc.GetElement(item));
            }


            OverrideGraphicSettings ogsFade = new OverrideGraphicSettings();
            ogsFade.SetSurfaceTransparency(80);
            ogsFade.SetSurfaceForegroundPatternVisible(false);
            ogsFade.SetSurfaceBackgroundPatternVisible(false);
            ogsFade.SetHalftone(true);


            OverrideGraphicSettings ogsIsolate = new OverrideGraphicSettings();
            ogsIsolate.SetSurfaceTransparency(0);
            ogsIsolate.SetSurfaceForegroundPatternVisible(true);
            ogsIsolate.SetSurfaceBackgroundPatternVisible(true);
            ogsIsolate.SetHalftone(false);

            using (Transaction t = new Transaction(doc, "Isolate with Fade"))
            {
                //t.Start();
                //foreach (Element e in new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType())
                //{
                //    if (e.Id != elementIdToIsolate || elementIdToIsolate != null)
                //        doc.ActiveView.SetElementOverrides(e.Id, ogsIsolate);
                //    else
                //        doc.ActiveView.SetElementOverrides(e.Id, ogsFade);
                //}
                //t.Commit();
                t.Start();
                foreach (Element e in new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType())
                {
                   
                }

                //doc.ActiveView.HideElements(ele_.ToArray().id);

                t.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    public class isolate_category : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            Reference myRef2 = uidoc.Selection.PickObject(ObjectType.Element);
            Element e2 = doc.GetElement(myRef2.ElementId);

            GeometryObject geomObj2 = e2.GetGeometryObjectFromReference(myRef2);
            Wall wall_ = e2 as Wall;

            Category cat = e2.Category;
            var catname = cat.Name;

            OverrideGraphicSettings ogsFade = new OverrideGraphicSettings();
            ogsFade.SetSurfaceTransparency(80);
            ogsFade.SetSurfaceForegroundPatternVisible(false);
            ogsFade.SetSurfaceBackgroundPatternVisible(false);
            ogsFade.SetHalftone(true);

            OverrideGraphicSettings ogsIsolate = new OverrideGraphicSettings();
            ogsIsolate.SetSurfaceTransparency(0);
            ogsIsolate.SetSurfaceForegroundPatternVisible(true);
            ogsIsolate.SetSurfaceBackgroundPatternVisible(true);
            ogsIsolate.SetHalftone(false);

            using (Transaction t = new Transaction(doc, "Isolate with Fade"))
            {
                t.Start();
                foreach (Element e in new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType())
                {
                    if (e.Category != null)
                    {
                        if (e.Category.Name == cat.Name)
                            doc.ActiveView.SetElementOverrides(e.Id, ogsIsolate);
                        else
                            doc.ActiveView.SetElementOverrides(e.Id, ogsFade);
                    }

                }
                t.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Clean_view : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;



            OverrideGraphicSettings ogsFade = new OverrideGraphicSettings();
            ogsFade.SetSurfaceTransparency(85);
            ogsFade.SetSurfaceForegroundPatternVisible(false);
            ogsFade.SetSurfaceBackgroundPatternVisible(false);
            ogsFade.SetHalftone(true);

            OverrideGraphicSettings ogsIsolate = new OverrideGraphicSettings();
            ogsIsolate.SetSurfaceTransparency(0);
            ogsIsolate.SetSurfaceForegroundPatternVisible(true);
            ogsIsolate.SetSurfaceBackgroundPatternVisible(true);
            ogsIsolate.SetHalftone(false);

            using (Transaction t = new Transaction(doc, "Isolate with Fade"))
            {
                t.Start();
                foreach (Element e in new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType())
                {
                    doc.ActiveView.SetElementOverrides(e.Id, ogsIsolate);
                }
                t.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class DeleteLevel : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F92CC68-A127-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 5;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            //string comments = "DeleteLevel" + "_" + doc.Application.Username + "_" + doc.Title;
            //string filename = @"D:\Users\lopez\Desktop\Comments.txt";
            ////System.Diagnostics.Process.Start(filename);
            //StreamWriter writer = new StreamWriter(filename, true);
            ////writer.WriteLine( Environment.NewLine);
            //writer.WriteLine(DateTime.Now + " - " + comments);
            //writer.Close();

            MessageBox.Show("Please Select Level to be deleted.", "Steps 1", MessageBoxButtons.OK,
          MessageBoxIcon.Information);

            Level level = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select level")) as Level;

            MessageBox.Show("Please Select a new hosting level .", "Steps 2", MessageBoxButtons.OK,
          MessageBoxIcon.Information);

            Level levelBelow = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select level")) as Level;

            //Level levelBelow = new FilteredElementCollector(doc)
            //    .OfClass(typeof(Level))
            //    .Cast<Level>()
            //    .OrderBy(q => q.Elevation)
            //    .Where(q => q.Elevation <= level.Elevation)
            //    .FirstOrDefault();

            if (levelBelow == null)
            {
                TaskDialog.Show("Error", "No level below " + level.Elevation);
                return Autodesk.Revit.UI.Result.Succeeded;
            }

            List<string> paramsToAdjust = new List<string> { "Base Offset", "Sill Height", "Top Offset", "Height Offset From Level" };
            List<List<ElementId>> ids = new List<List<ElementId>>();
            GroupType gtype = null;
            List<Element> elements = new FilteredElementCollector(doc).WherePasses(new ElementLevelFilter(level.Id)).ToList();
            using (Transaction t = new Transaction(doc, "un group"))
            {
                t.Start();
                foreach (Element e in elements)
                {
                    Group gr_ = e as Group;
                    if (gr_ != null)
                    {
                        List<ElementId> ids_ = new List<ElementId>();
                        Group gr = e as Group;
                        gtype = gr.GroupType;
                        ICollection<ElementId> groups = gr.UngroupMembers();

                        foreach (var item in groups)
                        {
                            ids_.Add(item);
                        }

                        ids.Add(ids_);
                    }
                }
                t.Commit();
            }
            List<Element> elements2 = new FilteredElementCollector(doc).WherePasses(new ElementLevelFilter(level.Id)).ToList();
            using (Transaction t = new Transaction(doc, "Level Remap"))
            {
                t.Start();

                foreach (Element e in elements2)
                {
                    Parameter param = e.LookupParameter("Top Constraint");
                    if (param != null && param.ToString() != "Unconnected")
                    {
                        param.Set("Unconnected");
                    }
                    try
                    {
                        foreach (Parameter p in e.Parameters)
                        {
                            if (p.StorageType != StorageType.ElementId || p.IsReadOnly)
                                continue;
                            if (p.AsElementId() != level.Id)
                                continue;

                            double elevationDiff = level.Elevation - levelBelow.Elevation;

                            p.Set(levelBelow.Id);

                            foreach (string paramName in paramsToAdjust)
                            {
                                Parameter pToAdjust = e.LookupParameter(paramName);
                                if (pToAdjust != null && !pToAdjust.IsReadOnly)
                                    pToAdjust.Set(pToAdjust.AsDouble() + elevationDiff);
                            }
                        }
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("!", "RevitLookup", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        throw;
                    }
                }
                t.Commit();
                MessageBox.Show("Delete level.", "Steps 3", MessageBoxButtons.OK,
          MessageBoxIcon.Information);
            }
            using (Transaction t = new Transaction(doc, "Level Remap"))
            {
                t.Start("Group");
                foreach (var item in ids)
                {
                    Group grpNew = doc.Create.NewGroup(item);

                    // Access the name of the previous group type 
                    // and change the new group type to previous 
                    // group type to retain the previous group 
                    // configuration

                    FilteredElementCollector coll
                      = new FilteredElementCollector(doc)
                        .OfClass(typeof(GroupType));

                    IEnumerable<GroupType> grpTypes
                      = from GroupType g in coll
                        where g.Name == gtype.Name.ToString()
                        select g;

                    grpNew.GroupType = grpTypes.First<GroupType>();
                }


                t.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class remove_paint : IExternalCommand
    {


        public class MassSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element element)
            {
                if (element.Category.Name == "Walls")
                {
                    return true;
                }
                return false;
            }

            public bool AllowReference(Reference refer, XYZ point)
            {
                return false;
            }
        }

        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-6509-AAF8-A578F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 21;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }


            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            ReferenceArray ra = new ReferenceArray();
            ISelectionFilter selFilter = new MassSelectionFilter();

            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            List<Element> walls = new List<Element>();

            foreach (var item in ids)
            {

                walls.Add(doc.GetElement(item));
            }

            //IList<Element> refList = uidoc.Selection.PickElementsByRectangle(selFilter, "Select multiple faces") as IList<Element>;
            //Reference hasPickOne = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, selFilter);
            ICollection<Reference> refList = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Face, "Select ceilings to be reproduced in rhino geometry");

            //IList<Reference> refList = uidoc.Selection.PickObjects(ObjectType.Element, selFilter, "Pick elements to add to selection filter");


            using (Transaction trans = new Transaction(doc, "ViewDuplicate"))
            {
                trans.Start();

                foreach (var item_myRefWall in refList)
                {
                    Element e = doc.GetElement(item_myRefWall);


                    //Wall wall_ = e as Wall;
                    GeometryElement geometryElement = e.get_Geometry(new Options());
                    foreach (GeometryObject geometryObject in geometryElement)
                    {
                        if (geometryObject is Solid)
                        {
                            Solid solid = geometryObject as Solid;
                            foreach (Face face_ in solid.Faces)
                            {
                                doc.RemovePaint(e.Id, face_);
                            }
                        }
                    }
                }
                trans.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class make_line : IExternalCommand
    {
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            //Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateUnbound(ptb, XYZ.BasisZ /* XYZ.BasisZ*/);

            Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(pta,
             line.Evaluate(5, false), ptb);

            SketchPlane skplane = SketchPlane.Create(doc, pl);

            Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line2, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line2, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        static AddInId appId = new AddInId(new Guid("5F56AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            //ObjectSnapTypes snapTypes = ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections;
            //XYZ point = uidoc.Selection.PickPoint(snapTypes, "Select an end point or intersection");


            //  XYZ point_in_3d_1 = uidoc.Selection.PickPoint(snapTypes,
            //"Please pick a point on the plane"
            //+ " defined by the selected face");


            Reference myRef2 = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            Element e2 = doc.GetElement(myRef2.ElementId);

            GeometryObject geomObj2 = e2.GetGeometryObjectFromReference(myRef2);
            //			Point p = geomObj as Point;
            XYZ p2 = /*uidoc.Selection.PickPoint (ObjectSnapTypes.)  */ myRef2.GlobalPoint;
            //string pointString2 = p2 == null ? "<reference has no globalpoint>" : string.Format("{0}", p2);
            //TaskDialog.Show("Element Info", e2.Name + "   " + pointString2);

            Reference myRef1 = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            Element e1 = doc.GetElement(myRef1.ElementId);

            GeometryObject geomObj1 = e1.GetGeometryObjectFromReference(myRef1);
            //			Point p = geomObj as Point;
            XYZ p1 = myRef1.GlobalPoint;
            //string pointString1 = p1 == null ? "<reference has no globalpoint>" : string.Format("{0}", p1);
            //TaskDialog.Show("Element Info", e1.Name + "   " + pointString1);

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Make Line");

                ModelLine ml = Makeline(doc, p1, p2);

                //Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateUnbound(p2, XYZ.BasisZ);

                //Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(p1,
                // line.Evaluate(5, false), p2);

                SketchPlane sketch = ml.SketchPlane /*SketchPlane.Create(doc, pl)*/;
                doc.ActiveView.SketchPlane = sketch;
                doc.ActiveView.ShowActiveWorkPlane();

                try
                {

                }
                catch (Exception)
                {

                    //TaskDialog.Show("!", "the project does not contain the parameter (Audit Orthogonality)assigned to walls" +
                    //    " ");
                }

                tx.Commit();
            }




            //uidoc.Selection.SetElementIds(ele.Select(q => q.Id).ToList());
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class DeleteAllSheets : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F22CC78-A137-4819-AAF1-A678F6B22BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 11;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            Form15 form = new Form15();

            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;


            IEnumerable<ViewSheet> viewSheet = from elem in new FilteredElementCollector(doc)
                                                 .OfClass(typeof(ViewSheet))
                                                 .OfCategory(BuiltInCategory.OST_Sheets)
                                               let type = elem as ViewSheet
                                               where type.Name != null
                                               select type;

            IEnumerable<Autodesk.Revit.DB.Viewport> viewports = from elem in new FilteredElementCollector(doc)
                                                 .OfClass(typeof(Autodesk.Revit.DB.Viewport))
                                                 .OfCategory(BuiltInCategory.OST_Views)
                                                                let type = elem as Autodesk.Revit.DB.Viewport
                                                                where type.Name != null
                                                                select type;

            List<ViewSheet> viewSheet_list = new List<ViewSheet>();
            List<ViewSheet> viewports_ = new List<ViewSheet>();

            foreach (var item in viewSheet)
            {
                viewSheet_list.Add(item);
            }

            try
            {
                Autodesk.Revit.DB.View viewTemplate = (from v in new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.View)).Cast<Autodesk.Revit.DB.View>()
                                                       where !v.IsTemplate && v.Name == "Home"
                                                       select v).First();


                uidoc.ActiveView = viewTemplate;
            }
            catch (Exception)
            {
                TaskDialog.Show("Warning", "DISCLAIMER view was not found in this project");
                return Autodesk.Revit.UI.Result.Cancelled;
                //throw;
            }

            using (Transaction t = new Transaction(doc, "Delete All Sheets"))
            {
                t.Start();


                try
                {
                    for (int i = 0; i < viewports_.ToArray().Length; i++)
                    {
                        doc.Delete(viewports_.ToArray()[i].Id);
                    }
                }
                catch (Exception)
                {
                    TaskDialog.Show("Warning", "Active view must not be a sheet");
                    throw;
                }

                try
                {
                    for (int i = 0; i < viewSheet_list.ToArray().Length; i++)
                    {
                        doc.Delete(viewSheet_list.ToArray()[i].Id);
                    }
                }
                catch (Exception)
                {
                    TaskDialog.Show("Warning", "Active view must not be a sheet");
                    throw;
                }




                t.Commit();


            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }


    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Reading_from_rhino : IExternalCommand
    {

        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        public static bool IsZero(double a)
        {
            const double _eps = 1.0e-9;
            return _eps > Math.Abs(a);
        }

        public static bool IsEqual(double a, double b)
        {
            return IsZero(b - a);
        }

        public static int Compare(double a, double b)
        {
            return IsEqual(a, b) ? 0 : (a < b ? -1 : 1);
        }

        public static int Compare(XYZ p, XYZ q)
        {
            
            int diff = Compare(p.X, q.X);
            if (0 == diff)
            {
                diff = Compare(p.Y, q.Y);
                if (0 == diff)
                {
                    diff = Compare(p.Z, q.Z);
                }

            }

            return diff;
        }

        private static Wall CreateWall(FamilyInstance cube, Autodesk.Revit.DB.Curve curve, double height)
        {
            var doc = cube.Document;

            var wallTypeId = doc.GetDefaultElementTypeId(
              ElementTypeGroup.WallType);

            return Wall.Create(doc, curve.CreateReversed(),
              wallTypeId, cube.LevelId, height, 0, false,
              false);
        }
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 15;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }



            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }




            //string comments = "Create_floor_from_rhino" + "_" + doc.Application.Username + "_" + doc.Title;
            //string filename = @"D:\Users\lopez\Desktop\Comments.txt";
            //StreamWriter writer = new StreamWriter(filename, true);
            //writer.WriteLine(DateTime.Now + " - " + comments);
            //writer.Close();

            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            String info = "";
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            //List<FaceArray> face112 = new List<FaceArray>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);



            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }

            Point3d p1 = new Point3d(0, 0, 0);
            Rhino.Geometry.Point3d pt3d = new Point3d(10, 10, 0);

            File3dm m_modelfile = null;
            string m_name = null;
            string m_size = null;
            string m_created = null;
            string m_createdby = null;
            string m_edited = null;
            string m_editedby = null;
            string m_revision = null;
            string m_units = null;
            string m_notes = null;

            Object RhinoFile = filename2;
            if (RhinoFile is System.IO.FileInfo)
            {
                System.IO.FileInfo m_fileinfo = (System.IO.FileInfo)RhinoFile;
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }
            else if (RhinoFile is string)
            {
                System.IO.FileInfo m_fileinfo = new System.IO.FileInfo((string)RhinoFile);
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }

            File3dmLayerTable all_layers = m_modelfile.AllLayers;

            var objs_ = m_modelfile.Objects;

            //List<File3dmObject> m_objslist = new List<File3dmObject>();

            //foreach (Rhino.DocObjects.Layer lay in m_modelfile.AllLayers)
            //{
            //    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(lay.Name);
            //    foreach (File3dmObject obj in m_objs)
            //    {
            //        m_objslist.Add(obj);
            //    }
            //}

            List<Rhino.DocObjects.Layer> listadelayers = new List<Rhino.DocObjects.Layer>();

            List<Rhino.DocObjects.Layer> borrar = new List<Rhino.DocObjects.Layer>();

            List<List<Rhino.DocObjects.Layer>> multiplelistadelayers = new List<List<Rhino.DocObjects.Layer>>();

            string hola = "";

            MessageBox.Show("This tool will read 3D information only if the following Rhino Layers exist; Levels, Grids, Structure, Floor, Walls, Points", "!");

            foreach (var item in all_layers)
            {
                 

                //if (item.Name.Contains( "Level"))
                //{
                //    if (listadelayers.Count != 0)
                //    {
                //        multiplelistadelayers.Add(listadelayers);
                //        listadelayers.RemoveRange(0, listadelayers.Count);
                //    }
                    
                //    //listadelayers.Clear();
                //    hola = item.FullPath;
                   
                //    continue;
                //}
                
                if (item.FullPath.Contains(hola))
                {
                    listadelayers.Add(item);
                }
            }

            List<Rhino.Geometry.Brep> rh_breps = new List<Rhino.Geometry.Brep>();
            List<Rhino.Geometry.Curve> curves_frames = new List<Rhino.Geometry.Curve>();
            List<Autodesk.Revit.DB.Curve> revit_crv = new List<Autodesk.Revit.DB.Curve>();
            List<string> m_names = new List<string>();
            List<int> m_layerindeces = new List<int>();
            List<System.Drawing.Color> m_colors = new List<System.Drawing.Color>();
            List<string> m_guids = new List<string>();

            foreach (var Layer in listadelayers)
            {
                if (Layer.Name == "Grids")
                {
                    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);
                    foreach (File3dmObject obj in m_objs)
                    {
                        GeometryBase geo = obj.Geometry;
                        if (geo is Rhino.Geometry.Curve)
                        {
                            Rhino.Geometry.Curve crv_ = geo as Rhino.Geometry.Curve;
                            curves_frames.Add(crv_);

                            Point3d end = crv_.PointAtEnd;
                            Point3d start = crv_.PointAtStart;

                            double x_end = end.X / 304.8;
                            double y_end = end.Y / 304.8;
                            double z_end = end.Z / 304.8;

                            double x_start = start.X / 304.8;
                            double y_start = start.Y / 304.8;
                            double z_start = start.Z / 304.8;

                            XYZ pt_end = new XYZ(x_end, y_end, z_end);
                            XYZ pt_start = new XYZ(x_start, y_start, z_start);

                            using (Transaction t = new Transaction(doc, "Make Floor from Rhino"))
                            {
                                t.Start();

                                Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                Autodesk.Revit.DB.Grid lineGrid = Autodesk.Revit.DB.Grid.Create(doc, line);

                                revit_crv.Add(curve1);

                                t.Commit();
                            }

                            try
                            {
                                

                            }
                            catch (Exception)
                            {
                                //throw;
                            }
                        }
                    }
                }

                if (Layer.Name == "Levels")
                {
                    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);
                    foreach (File3dmObject obj in m_objs)
                    {
                        GeometryBase geo = obj.Geometry;
                        if (geo is Rhino.Geometry.Curve)
                        {
                            Rhino.Geometry.Curve crv_ = geo as Rhino.Geometry.Curve;
                            curves_frames.Add(crv_);

                            Point3d end = crv_.PointAtEnd;
                            Point3d start = crv_.PointAtStart;

                            double x_end = end.X / 304.8;
                            double y_end = end.Y / 304.8;
                            double z_end = end.Z / 304.8;

                            double x_start = start.X / 304.8;
                            double y_start = start.Y / 304.8;
                            double z_start = start.Z / 304.8;

                            XYZ pt_end = new XYZ(x_end, y_end, z_end);
                            XYZ pt_start = new XYZ(x_start, y_start, z_start);


                            using (Transaction t = new Transaction(doc, "Make ducts from Rhino"))
                            {
                                t.Start();

                                Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                Level level_ = Level.Create(doc, z_end);
                                

                                revit_crv.Add(curve1);

                                t.Commit();
                            }

                            try
                            {
                                

                            }
                            catch (Exception)
                            {
                                //throw;
                            }
                        }
                    }
                }
                
                if (Layer.Name == "Floors")
                {
                    List<List<Rhino.Geometry.Curve>> floor_breps_Curve = new List<List<Rhino.Geometry.Curve>>();
                    List<List<Autodesk.Revit.DB.Curve>> wall_breps_Curve = new List<List<Autodesk.Revit.DB.Curve>>();
                    List<Rhino.Geometry.Surface> srflist = new List<Rhino.Geometry.Surface>();

                    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);
                    foreach (var obj in m_objs)
                    {


                        GeometryBase brep = obj.Geometry;

                        Rhino.Geometry.Brep brep_ = brep as Rhino.Geometry.Brep;
                        Rhino.Geometry.Extrusion ext = brep as Rhino.Geometry.Extrusion;
                        
                       
                        if (brep_ is Rhino.Geometry.Brep)
                        {
                            foreach (BrepFace bf in brep_.Faces)
                            {
                               
                                List<Rhino.Geometry.NurbsSurface> m_nurbslist = new List<Rhino.Geometry.NurbsSurface>();

                                Vector3d _normal = bf.NormalAt(0.5, 0.5);

                                if (_normal.Z == 1.0)
                                {
                                   
                                   

                                    Rhino.Geometry.NurbsSurface m_nurbsurface = bf.ToNurbsSurface();
                                    m_nurbslist.Add(m_nurbsurface);

                                    List<Rhino.Geometry.Curve> rh_loops = new List<Rhino.Geometry.Curve>();

                                    if (bf.Loops.Count > 0)
                                    {
                                        foreach (BrepLoop bloop in bf.Loops)
                                        {
                                            List<Rhino.Geometry.Curve> Curve_ = new List<Rhino.Geometry.Curve>();

                                            List<Rhino.Geometry.Curve> m_trims = new List<Rhino.Geometry.Curve>();
                                            foreach (BrepTrim t in bloop.Trims)
                                            {
                                                if (t.TrimType == BrepTrimType.Boundary || t.TrimType == BrepTrimType.Mated) // ignore "seams"
                                                {
                                                    Rhino.Geometry.Curve m_edgecurve = t.Edge.DuplicateCurve();
                                                    Curve_.Add(m_edgecurve);
                                                }
                                            }
                                            floor_breps_Curve.Add(Curve_);
                                        }
                                    }
                                }

                            }
                        }
                    }

                    List<List<XYZ>> floor_listas_pts = new List<List<XYZ>>();
                    if (floor_breps_Curve.ToArray().Length > 0)
                    {
                        for (int i = 0; i < floor_breps_Curve.ToArray().Length; i++)
                        {
                            List<XYZ> lista_pt = new List<XYZ>();
                            foreach (var item in floor_breps_Curve[i])
                            {
                                Point3d pt = item.PointAtStart;
                                double x = pt.X / 304.8;
                                double y = pt.Y / 304.8;
                                double z = pt.Z / 304.8;

                                XYZ pt_ = new XYZ(x, y, z);
                                lista_pt.Add(pt_);
                            }
                            floor_listas_pts.Add(lista_pt);
                        }
                    }

                    if (floor_breps_Curve.ToArray().Length > 0)
                    {
                        using (Transaction t = new Transaction(doc, "Make Floor from Rhino"))
                        {
                            t.Start();

                            for (int i = 0; i < floor_listas_pts.ToArray().Length; i++)
                            {
                                CurveArray ca = new CurveArray();
                                XYZ prev = null;

                                foreach (var pt in floor_listas_pts[i])
                                {
                                    if (prev == null)
                                    {
                                        prev = floor_listas_pts[i].Last();
                                    }
                                    Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(prev, pt);
                                    ca.Append(line);
                                    prev = pt;
                                    try
                                    {
                                      
                                    }
                                    catch (Exception)
                                    {

                                        //throw;
                                    }
                                }
                                doc.Create.NewFloor(ca, false);
                            }
                            t.Commit();
                        }
                    }
                }

                if (Layer.Name == "Walls")
                {
                    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);
                    foreach (var obj in m_objs)
                    {
                        GeometryBase brep = obj.Geometry;
                        Rhino.Geometry.Brep brep_ = brep as Rhino.Geometry.Brep;
                        Rhino.Geometry.Extrusion ext = brep as Rhino.Geometry.Extrusion;
                        
                        if (brep_ is Rhino.Geometry.Brep)
                        {
                            foreach (BrepFace bf in brep_.Faces)
                            {
                                List<Rhino.Geometry.NurbsSurface> m_nurbslist = new List<Rhino.Geometry.NurbsSurface>();
                                Vector3d _normal = bf.NormalAt(0.5, 0.5);
                                if (_normal.Z == 0.0)
                                {
                                    Rhino.Geometry.NurbsSurface m_nurbsurface = bf.ToNurbsSurface();
                                    m_nurbslist.Add(m_nurbsurface);
                                    if (bf.Loops.Count > 0)
                                    {
                                        foreach (BrepLoop bloop in bf.Loops)
                                        {
                                            List<Autodesk.Revit.DB.Curve> Curve_ = new List<Autodesk.Revit.DB.Curve>();
                                            foreach (BrepTrim t in bloop.Trims)
                                            {
                                                if (t.TrimType == BrepTrimType.Boundary || t.TrimType == BrepTrimType.Mated) // ignore "seams"
                                                {
                                                    Rhino.Geometry.Curve m_edgecurve = t.Edge.DuplicateCurve();
                                                    Point3d end = m_edgecurve.PointAtEnd;
                                                    Point3d start = m_edgecurve.PointAtStart;

                                                    double x_end = end.X / 304.8;
                                                    double y_end = end.Y / 304.8;
                                                    double z_end = end.Z / 304.8;

                                                    double x_start = start.X / 304.8;
                                                    double y_start = start.Y / 304.8;
                                                    double z_start = start.Z / 304.8;

                                                    XYZ pt_end = new XYZ(x_end, y_end, z_end);
                                                    XYZ pt_start = new XYZ(x_start, y_start, z_start);

                                                    try
                                                    {
                                                        Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);

                                                        Curve_.Add(curve1);
                                                    }
                                                    catch (Exception)
                                                    {
                                                    }
                                                }
                                            }

                                            using (Transaction t = new Transaction(doc, "Make Floor from Rhino"))
                                            {
                                                t.Start();
                                                Autodesk.Revit.DB.Wall.Create (doc, Curve_, true);
                                                t.Commit();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }   
                }
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }  
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class writting_to_rhino : IExternalCommand
    {
        public static RhinoApplication RhinoApp(bool StartRhino) // bool IsVisible)
        {
            try
            {
                if (StartRhino == true)
                {
                    // CASE Rhino Connection Class
                    RhinoCon.RhinoApplication m_RhinoCOM = new RhinoCon.RhinoApplication("Rhino5x64.Application", true);


                    return m_RhinoCOM;
                }
                else { return null; }
            }
            catch { return null; }
        }

        public static List<List<Element>> GetMembersRecursive(Autodesk.Revit.DB.Document d, Group g, List<List<Element>> r, List<List<string>> strin_name, List<int> int_)
        {
            if (strin_name == null)
            {
                strin_name = new List<List<string>>();
            }
            if (r == null)
            {
                r = new List<List<Element>>();
            }
            if (int_ == null)
            {
                int_ = new List<int>();
            }
            List<string> lista_nombre_main = new List<string>();
            List<string> lista_nombre_buildings = new List<string>();
            List<string> lista_nombre_floor = new List<string>();
            List<Group> lista_de_buildings = new List<Group>();
            List<Group> lista_de_floor = new List<Group>();
            List<List<Element>> ele_list = new List<List<Element>>();
            List<Element> ceiling_groups1 = new List<Element>();

            List<Element> elems = g.GetMemberIds().Select(q => d.GetElement(q)).ToList();
            lista_nombre_main.Add(g.Name);
            lista_nombre_buildings.Add(g.Name);

            foreach (Element el in elems)
            {
                if (el.GetType() == typeof(Group))
                {
                    Group gp = el as Group;
                    lista_de_buildings.Add(gp);
                    lista_nombre_buildings.Add(el.Name);
                }
                if (el.GetType() == typeof(Ceiling))
                {
                    ceiling_groups1.Add(el);
                }
            }
            r.Add(ceiling_groups1);
            for (int i = 0; i < lista_de_buildings.ToArray().Length; i++)
            {
                Group gp2 = lista_de_buildings.ToArray()[i];

                List<Element> elems2 = gp2.GetMemberIds().Select(q => d.GetElement(q)).ToList();
                ele_list.Add(elems2);
            }

            for (int i = 0; i < ele_list.ToArray().Length; i++)
            {
                List<Element> lista1 = ele_list.ToArray()[i];
                List<Element> ceiling_groups = new List<Element>();
                foreach (var item in lista1)
                {
                    if (item.GetType() == typeof(Group))
                    {
                        Group gp4 = item as Group;
                        List<Element> elems3 = gp4.GetMemberIds().Select(q => d.GetElement(q)).ToList();
                        foreach (var item2 in elems3)
                        {
                            if (item2.GetType() == typeof(Group))
                            {
                                Group gp5 = item2 as Group;
                                List<Element> elems4 = gp5.GetMemberIds().Select(q => d.GetElement(q)).ToList();
                                foreach (var item3 in elems4)
                                {
                                    if (item3.GetType() == typeof(Ceiling))
                                    {
                                        ceiling_groups.Add(item3);
                                        lista_nombre_floor.Add(item.Name);
                                    }
                                }
                            }
                            if (item2.GetType() == typeof(Ceiling))
                            {
                                ceiling_groups.Add(item2);
                                lista_nombre_floor.Add(item.Name);
                            }
                        }

                        //lista_de_floor.Add(gp4);
                        //lista_nombre_floor.Add(item.Name);
                    }
                    if (item.GetType() == typeof(Ceiling))
                    {
                        ceiling_groups.Add(item);
                        lista_nombre_floor.Add(item.Name);

                    }
                }
                r.Add(ceiling_groups);
            }
            strin_name.Add(lista_nombre_main);
            strin_name.Add(lista_nombre_buildings);
            strin_name.Add(lista_nombre_floor);

            return r;

        }

        public static List<List<Face>> GetFaces(Autodesk.Revit.DB.Document doc_, List<List<Element>> list_elements, List<List<Face>> list_faces)
        {
            if (list_faces == null)
            {
                list_faces = new List<List<Face>>();
            }
            for (int i = 0; i < list_elements.ToArray().Length; i++)
            {
                List<Face> faces_list = new List<Face>();
                List<Element> ele = list_elements.ToArray()[i];

                foreach (var item in ele)
                {
                    Options op = new Options();
                    op.ComputeReferences = true;
                    foreach (var item2 in item.get_Geometry(op).Where(q => q is Solid).Cast<Solid>())
                    {
                        foreach (Face item3 in item2.Faces)
                        {
                            PlanarFace planarFace = item3 as PlanarFace;
                            XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));

                            if (normal.Z == 0 && normal.Y > -0.8 /*&& normal.X < 0*/)
                            {
                                Element e = doc_.GetElement(item3.Reference);
                                GeometryObject geoobj = e.GetGeometryObjectFromReference(item3.Reference);
                                Face face = geoobj as Face;
                                faces_list.Add(face);
                            }
                        }
                    }
                }
                if (faces_list.ToArray().Length > 0)
                {
                    list_faces.Add(faces_list);
                }

            }

            return list_faces;
        }

        public static List<List<XYZ>> GeTpoints(Autodesk.Revit.DB.Document doc_, List<List<XYZ>> xyz_faces, IList<CurveLoop> faceboundaries, List<List<Face>> list_faces)
        {
            if (list_faces == null)
            {
                list_faces = new List<List<Face>>();
            }

            for (int i = 0; i < list_faces.ToArray().Length; i++)
            {
                List<XYZ> puntos_ = new List<XYZ>();
                foreach (Face f in list_faces.ToArray()[i])
                {

                    faceboundaries = f.GetEdgesAsCurveLoops();//new trying to get the outline of the face instead of the edges
                    EdgeArrayArray edgeArrays = f.EdgeLoops;
                    foreach (CurveLoop edges in faceboundaries)
                    {
                        puntos_.Add(null);
                        foreach (Autodesk.Revit.DB.Curve edge in edges)
                        {
                            XYZ testPoint1 = edge.GetEndPoint(1);
                            double lenght = Math.Round(edge.ApproximateLength, 0);
                            double lenght_convert = UnitUtils.Convert(lenght, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);

                            double x = Math.Round(testPoint1.X, 0);
                            double y = Math.Round(testPoint1.Y, 0);
                            double z = Math.Round(testPoint1.Z, 0);

                            XYZ newpt = new XYZ(x, y, z);

                            if (!puntos_.Contains(testPoint1))
                            {
                                puntos_.Add(testPoint1);

                            }
                        }
                    }
                    int num = f.EdgeLoops.Size;

                }
                xyz_faces.Add(puntos_);
            }
            return xyz_faces;

        }

        static AddInId appId = new AddInId(new Guid("5F88CC09-A137-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 19;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            //string comments = "Rhino_access" + "_" + doc.Application.Username + "_" + doc.Title;
            //string filename = @"D:\Users\lopez\Desktop\Comments.txt";
            ////System.Diagnostics.Process.Start(filename);
            //StreamWriter writer = new StreamWriter(filename, true);
            ////writer.WriteLine( Environment.NewLine);
            //writer.WriteLine(DateTime.Now + " - " + comments);
            //writer.Close();

            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            String info = "";
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            //List<FaceArray> face112 = new List<FaceArray>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);

            Form7 form = new Form7();

            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //openDialog.Filter = "Rhino Files (*.3dm) |*.3dm)"; // TODO: Change to .csv
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }
            // this is use to open a new instance of rhino
            //RhinoCon.RhinoApplication m_RhinoCOM = RhinoApp(true);
            Point3d p1 = new Point3d(0, 0, 0);
            Rhino.Geometry.Point3d pt3d = new Point3d(10, 10, 0);
            //RhinoCon.RhinoApplication m_RhinoCOM = new RhinoCon.RhinoApplication("Rhino5x64.Application", true);
            //File3dm m_modelfile2 = m_RhinoCOM.RhinoScript.PointAdd(p1) /*as Rhino.FileIO.File3dm*/;
            File3dm m_modelfile = null;
            string m_name = null;
            string m_size = null;
            string m_created = null;
            string m_createdby = null;
            string m_edited = null;
            string m_editedby = null;
            string m_revision = null;
            string m_units = null;
            string m_notes = null;

            Object RhinoFile = filename2;
            if (RhinoFile is System.IO.FileInfo)
            {
                System.IO.FileInfo m_fileinfo = (System.IO.FileInfo)RhinoFile;
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }
            else if (RhinoFile is string)
            {
                System.IO.FileInfo m_fileinfo = new System.IO.FileInfo((string)RhinoFile);
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }

            Rhino.FileIO.File3dm df = m_modelfile as File3dm;

            RhinoCon.RhinoApplication m_RhinoCOM = new RhinoCon.RhinoApplication(filename2/*"Rhino5x64.Application"*/, true);

            

            //for (int j = 0; j <= 20; j++)

            //{

            //    m_modelfile.Objects.AddLine(new Rhino.Geometry.Line(j, 0, 5 - j, 5 + j, 0, j));

            //}
            //m_modelfile.Objects.AddPoint(p1);




            ICollection<Reference> my_faces = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Face, "Select ceilings to be reproduced in rhino geometry");



            List<Face> faces_picked = new List<Face>();
            List<string> name_of_roof = new List<string>();

            foreach (var item_myRefWall in my_faces)
            {

                Element e = doc.GetElement(item_myRefWall);
                GeometryObject geoobj = e.GetGeometryObjectFromReference(item_myRefWall);
                Face face = geoobj as Face;
                PlanarFace planarFace = face as PlanarFace;
                XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));


                name_of_roof.Add("roof");
                faces_picked.Add(face);


                if (item_myRefWall == my_faces.ToArray().Last())
                {
                    Faces_lists_excel.Add(faces_picked);

                    //names.Add(name_of_roof);
                }

            }

            Group grpExisting = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select an existing group")) as Group;



            GetMembersRecursive(doc, grpExisting, elemente_selected, names, numeros_);

            foreach (var item in elemente_selected)
            {
                foreach (var item2 in item)
                {
                    if (!form.listBox1.Items.Contains(item2.Name))
                    {
                        form.listBox1.Items.Add(item2.Name);
                    }
                }
            }

            //form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            List<List<Element>> elemente_selected_to_bedeleted = new List<List<Element>>();
            List<List<string>> nombres_familias = new List<List<string>>();

            foreach (var item in elemente_selected)
            {
                List<Element> ele_sel = new List<Element>();
                foreach (var item2 in item)
                {
                    foreach (var item3 in form.listBox2.Items)
                    {
                        if (item3.ToString() == item2.Name)
                        {
                            ele_sel.Add(item2);
                        }
                    }
                }
                elemente_selected_to_bedeleted.Add(ele_sel);
            }



            string name_of_group = names.ToArray()[0].ToArray()[0].ToString();

            names.ToArray()[1].Insert(1, "roof" + name_of_group);

            GetFaces(doc, /*elemente_selected_to_bedeleted*/elemente_selected, Faces_lists_excel);

            GeTpoints(doc, xyz_faces, faceboundaries, Faces_lists_excel);

            List<List<Point3d>> listas_pts = new List<List<Point3d>>();
            ;
            List<Rhino.DocObjects.Layer> child_layer = new List<Rhino.DocObjects.Layer>();


            IList<Rhino.DocObjects.Layer> EXISTING = m_modelfile.AllLayers;
            int count_ = 0;

            foreach (var item in EXISTING)
            {
                count_++;
            }

            Rhino.DocObjects.Layer Parentlayer = new Rhino.DocObjects.Layer();
            Parentlayer.Name = names.ToArray()[0].ToArray()[0].ToString();
            Parentlayer.Index = count_ + 1;
            m_modelfile.AllLayers.Add(Parentlayer);


            List<Rhino.DocObjects.Layer> created_layers = new List<Rhino.DocObjects.Layer>();

            foreach (var item in names.ToArray()[1])
            {
                if (Parentlayer.Name != item)
                {
                    Rhino.DocObjects.Layer Childlayer = new Rhino.DocObjects.Layer();
                    //Childlayer.ParentLayerId = Parentlayer.Id;
                    Childlayer.Name = item.ToString();
                    Childlayer.Color = System.Drawing.Color.Bisque;
                    Childlayer.Index = Parentlayer.Index;



                    m_modelfile.AllLayers.Add(Childlayer);

                    created_layers.Add(Childlayer);
                }

            }

            List<List<PlaneSurface>> psrf = new List<List<PlaneSurface>>();

            for (int i = 0; i < xyz_faces.ToArray().Length; i++)
            {
                List<PlaneSurface> srflist = new List<PlaneSurface>();
                List<Point3d> lista_de_puntos = new List<Point3d>();
                //lista_de_puntos.Clear();
                int num = 0;

                foreach (var item in xyz_faces.ToArray()[i])
                {

                    if (item != null)
                    {
                        num++;
                        double x = item.X * 304.8;
                        double y = item.Y * 304.8;
                        double z = item.Z * 304.8;

                        Point3d ptstart = new Point3d(x, y, z);
                        lista_de_puntos.Add(ptstart);
                    }
                    else
                    {
                        lista_de_puntos.Clear();
                        num = 0;

                    }
                    if (num == 3)
                    {
                        Rhino.Geometry.Point3d pt0 = lista_de_puntos.ToArray()[0];
                        Rhino.Geometry.Point3d pt1 = lista_de_puntos.ToArray()[1];
                        Rhino.Geometry.Point3d pt2 = lista_de_puntos.ToArray()[2];

                        Interval int1 = new Interval(0, pt0.DistanceTo(pt1));

                        var plane = new Rhino.Geometry.Plane(pt0, pt1, pt2);
                        var plane_surface = new PlaneSurface(plane, int1, new Interval(0, pt1.DistanceTo(pt2)));
                        //m_modelfile.Objects.AddSurface(plane_surface);
                        srflist.Add(plane_surface);
                        lista_de_puntos.Clear();
                        num = 0;
                    }

                }
                psrf.Add(srflist);
            }



            for (int i = 0; i < created_layers.ToArray().Length; i++)
            {
                count_++;
                foreach (var item in psrf.ToArray()[i])
                {

                    Rhino.DocObjects.ObjectAttributes myAtt = new Rhino.DocObjects.ObjectAttributes();
                    Rhino.DocObjects.Layer layer_ = m_modelfile.AllLayers.FindIndex(count_);

                    myAtt.LayerIndex = layer_.Index;
                    m_modelfile.Objects.AddSurface(item, myAtt);

                    //int layerIndex2 = thisDoc.Layers.FindByFullPath(child_layer.ToArray()[i].FullPath,false);
                }
            }
            File3dmWriteOptions options = new File3dmWriteOptions();
            options.SaveUserData = true;
            m_modelfile.Write(filename2, 0);

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class select_detailline : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            UIApplication uiapp = commandData.Application;
            //UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;



            Form3 form = new Form3();

            form.ShowDialog();
            
            DetailLine line = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select detail line")) as DetailLine;
            
           List< Autodesk.Revit.DB.Element> lines = new FilteredElementCollector(doc).OfClass(typeof(CurveElement)).OfCategory(BuiltInCategory.OST_Lines).Where(q => q is DetailLine).ToList(); ;

            List<DetailLine> selected = new List<DetailLine>();

            if (form.radioButton1.Checked == true)
            {
                foreach (var item in lines)
                {
                    DetailLine DL = item as DetailLine;
                    if (DL.LineStyle.Id == line.LineStyle.Id)
                    {
                        if (line.OwnerViewId == DL.OwnerViewId)
                        {
                            selected.Add(DL);
                        }
                    }
                }
            }


            if (form.radioButton2.Checked == true)
            {
                foreach (var item in lines)
                {
                    DetailLine DL = item as DetailLine;

                    if (DL.LineStyle.Id == line.LineStyle.Id)
                    {
                        selected.Add(DL);
                    }
                }

            }



            uidoc.Selection.SetElementIds(selected.Select(q => q.Id).ToList());


            return Autodesk.Revit.UI.Result.Cancelled;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class RoomElevations : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("07B86EFE-F18B-4354-AA9B-29F3E9C5F5AB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;


            //---------------------------------------- FILTERS ------------------------------------

            FilteredElementCollector levels = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_Levels);
            ICollection<Element> level1 = levels.ToElements();

            ViewFamilyType vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Section == x.ViewFamily);
            ViewFamilyType vftele = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Elevation == x.ViewFamily);

            FilteredElementCollector rooms1 = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(SpatialElement));
            ICollection<Element> room2 = rooms1.ToElements();

            //Autodesk.Revit.DB.View viewTemplate = (from v in new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.View)).Cast<Autodesk.Revit.DB.View>()
            //                                       where v.IsTemplate && v.Name == "Architectural Section"
            //                                       select v).First();

            IEnumerable<ViewFamilyType> viewFamilyTypes = from elem in new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType))
                                                          let type = elem as ViewFamilyType
                                                          where type.ViewFamily == ViewFamily.ThreeDimensional
                                                          select type;


            //Group grpExisting = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select an existing group")) as Group;
           
            //Element el = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select an existing group"));
            //Location loc = el.Location;
            //Autodesk.Revit.DB.LocationPoint positionPoint = loc as Autodesk.Revit.DB.LocationPoint;


            ViewFamilyType ceiling_plan_view = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.CeilingPlan == x.ViewFamily);
            ViewFamilyType floor_plan_view = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.FloorPlan == x.ViewFamily);
            string project_param = /*"Schedule Identifier"*/"Comments";

            using (Transaction t1 = new Transaction(doc))
            {
                List<Element> roomele = new List<Element>();
                List<Element> selected_room = new List<Element>();


                

                foreach (Element j in room2)
                {
                    try
                    {
                        ParameterMap parmap = j.ParametersMap;
                        Parameter par = parmap.get_Item(project_param);
                        string value = par.AsString();

                        if (value != null)
                        {
                            Element E = j;
                            roomele.Add(E);

                        }
                    }
                    catch (Exception)
                    {
                        TaskDialog.Show("Task Cancelled", "Project does not contain parameter " + project_param + "or may be no room has an identifier");
                        //TaskDialog.Show("Warning", "Task Cancelled");
                        return Autodesk.Revit.UI.Result.Cancelled;
                    }
                    
                   
                }

                Form_for_room_Schedule form = new Form_for_room_Schedule();



                foreach (var item in roomele)
                {
                    form.listBox1.Items.Add(item.Name);
                }
                

                if (roomele.Count == 0)
                {
                    TaskDialog.Show("Task Cancelled", "no room has an identifier");
                    
                    return Autodesk.Revit.UI.Result.Cancelled;
                }

                form.ShowDialog();


                if (form.DialogResult == DialogResult.Cancel)
                {
                    return Autodesk.Revit.UI.Result.Cancelled;
                }

                foreach (var item in roomele)
                {
                    try
                    {
                        for (int i = 0; i < form.listBox1.SelectedItems.Count; i++)
                        {
                            if (form.listBox1.SelectedItems[i].ToString() == item.Name)
                            {
                                selected_room.Add(item);
                            }
                        }

                        
                    }
                    catch (Exception)
                    {
                        TaskDialog.Show("Task Cancelled", "A room must be selected");
                        
                        return Autodesk.Revit.UI.Result.Cancelled;
                        throw;
                    }

                }


                t1.Start("Create views from selected view");

                if (form.checkBox1.Checked)
                {
                    foreach (Element item in selected_room)
                    {
                        ElementId id = item.LevelId;
                        try
                        {
                            ElementId level11 = doc.ActiveView.GenLevel.Id;
                        }
                        catch
                        {
                            TaskDialog.Show("Warning", "you must be in a floor plan");
                            goto final;
                        }

                        if (item.LevelId.IntegerValue == doc.ActiveView.GenLevel.Id.IntegerValue)
                        {

                            if (item.Location is LocationPoint)
                            {
                                LocationPoint lp = item.Location as LocationPoint;
                                double pz = lp.Point.X;
                                XYZ point = lp.Point;


                                try
                                {

                                    ElevationMarker marker = ElevationMarker.CreateElevationMarker(doc, vftele.Id, point, 100);
                                    for (int i = 0; i <  4; i++)
                                    {
                                        ViewSection elevation1 = marker.CreateElevation(doc, doc.ActiveView.Id, i);
                                        elevation1.Name = form.textBox1.Text + " "+  item.Name + " " + i.ToString() ;
                                    }
                                    //TaskDialog.Show("View/s created", "Interior Elevations were created for " + form.listBox1.SelectedItems.Count + " Rooms");
                                }
                                catch
                                {
                                   

                                    TaskDialog.Show("Element info", " Name must be unique ");
                                }
                            }
                        }
                    }
                }
                if (form.checkBox2.Checked)
                {
                    foreach (Element item in selected_room)
                    {
                        ElementId id = item.LevelId;
                        string name = item.Name;
                        ElementId level11 = doc.ActiveView.GenLevel.Id;

                        BoundingBoxXYZ room_box = item.get_BoundingBox(null);
                        XYZ mas_10 = new XYZ(4, 4, 4);
                        XYZ max_p = room_box.Max + mas_10;
                        XYZ min_p = room_box.Min - mas_10;

                        View3D view3D = View3D.CreateIsometric(doc, viewFamilyTypes.First().Id);

                        try
                        {
                            view3D.Name = form.textBox1.Text + " " + item.Name + item.Name;
                        }
                        catch (Exception)
                        {

                            TaskDialog.Show("Element info", " Name must be unique ");
                        }
                        
                        BoundingBoxXYZ boundingBoxXYZ = new BoundingBoxXYZ();
                        boundingBoxXYZ.Min = min_p;
                        boundingBoxXYZ.Max = max_p;
                        view3D.SetSectionBox(boundingBoxXYZ);
                        //TaskDialog.Show("View/s created", "Isometric view was created for " + form.listBox1.SelectedItems.Count + " Rooms");
                    }
                }

                if (form.checkBox3.Checked)
                {

                    foreach (Element item in selected_room)
                    {
                        ElementId id = item.LevelId;
                        string name = item.Name;


                        BoundingBoxXYZ room_box = item.get_BoundingBox(null);
                        XYZ mas_10 = new XYZ(4, 4, 4);
                        XYZ max_p = room_box.Max + mas_10;
                        XYZ min_p = room_box.Min - mas_10;

                        BoundingBoxXYZ boundingBoxXYZ = new BoundingBoxXYZ();
                        boundingBoxXYZ.Min = min_p;
                        boundingBoxXYZ.Max = max_p;

                        ElementId level11 = doc.ActiveView.GenLevel.Id;
                        ViewPlan floorView = ViewPlan.Create(doc, floor_plan_view.Id, id);

                        try
                        {
                            floorView.Name = form.textBox1.Text + " " + item.Name + item.Name;
                        }
                        catch (Exception)
                        {

                            TaskDialog.Show("Element info", " Name must be unique ");
                        }
                        
                        floorView.CropBox = boundingBoxXYZ;
                        floorView.CropBox.Enabled = true;
                        floorView.CropBoxActive = true;


                        //TaskDialog.Show("View/s created", "ViewPlan was created for " + form.listBox1.SelectedItems.Count + " Rooms");
                    }
                }

                if (form.checkBox4.Checked)
                {

                    foreach (Element item in selected_room)
                    {
                        ElementId id = item.LevelId;
                        string name = item.Name;


                        BoundingBoxXYZ room_box = item.get_BoundingBox(null);
                        XYZ mas_10 = new XYZ(4, 4, 4);
                        XYZ max_p = room_box.Max + mas_10;
                        XYZ min_p = room_box.Min - mas_10;



                        BoundingBoxXYZ boundingBoxXYZ = new BoundingBoxXYZ();
                        boundingBoxXYZ.Min = min_p;
                        boundingBoxXYZ.Max = max_p;

                        ElementId level11 = doc.ActiveView.GenLevel.Id;
                        ViewPlan ceiling_View = ViewPlan.Create(doc, ceiling_plan_view.Id, id);

                        try
                        {
                            ceiling_View.Name = form.textBox1.Text + " " + item.Name + item.Name;
                        }
                        catch (Exception)
                        {

                            TaskDialog.Show("Element info", " Name must be unique ");
                        }
                        
                        ceiling_View.CropBox = boundingBoxXYZ;
                        ceiling_View.CropBox.Enabled = true;
                        ceiling_View.CropBoxActive = true;
                        //TaskDialog.Show("View/s created", "Ceiling was created for " + form.listBox1.SelectedItems.Count + " Rooms");
                    }
                }
                t1.Commit();

            }
            final: return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Element_Elevations : IExternalCommand
    {
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {


            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }
        public XYZ InvCoord(XYZ MyCoord)
        {
            XYZ invcoord = new XYZ((Convert.ToDouble(MyCoord.X * -1)),
                (Convert.ToDouble(MyCoord.Y * -1)),
                (Convert.ToDouble(MyCoord.Z * -1)));
            return invcoord;
        }

        public void PickPoint(UIDocument uidoc)
        {
            ObjectSnapTypes snapTypes = ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections;
            XYZ point = uidoc.Selection.PickPoint(snapTypes, "Select an end point or intersection");

            string strCoords = "Selected point is " + point.ToString();

            TaskDialog.Show("Revit", strCoords);
        }

        public static bool IsZero(double a)
        {
            const double _eps = 1.0e-9;
            return _eps > Math.Abs(a);
        }
        public static bool IsEqual(double a, double b)
        {
            return IsZero(b - a);
        }
        public static int Compare(double a, double b)
        {
            return IsEqual(a, b) ? 0 : (a < b ? -1 : 1);
        }
        public static int Compare(XYZ p, XYZ q)
        {
            int diff = Compare(p.X, q.X);
            if (0 == diff)
            {
                diff = Compare(p.Y, q.Y);
                if (0 == diff)
                {
                    diff = Compare(p.Z, q.Z);
                }
            }
            return diff;
        }

        static AddInId appId = new AddInId(new Guid("07B86EFE-F18B-4354-AA9B-29F3E9C5F5AB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            ViewFamilyType floor_plan_view = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.FloorPlan == x.ViewFamily);

            ViewFamilyType vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Section == x.ViewFamily);
            ViewFamilyType vftele = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Elevation == x.ViewFamily);
            
            IEnumerable<ViewFamilyType> viewFamilyTypes = from elem in new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType))
                                                          let type = elem as ViewFamilyType
                                                          where type.ViewFamily == ViewFamily.ThreeDimensional
                                                          select type;

            List<ViewSection> views_ = new List<ViewSection>();

            Element el = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select an existing group"));


            Form_for_room_Schedule form = new Form_for_room_Schedule();

            form.listBox1.Items.Add(el.Name);
            
            form.ShowDialog();

            form.checkBox4.Enabled = false;

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            Location loc = el.Location;
            Autodesk.Revit.DB.LocationPoint lp = loc as Autodesk.Revit.DB.LocationPoint;
            var fi = el as FamilyInstance;
            var TRF = fi.GetTransform();

            var _X = TRF.BasisX;
            var _Z = TRF.BasisZ;

            var _Xtran = TRF.OfVector(TRF.BasisX);
            var objrot = TRF.BasisX.AngleOnPlaneTo(_Xtran, TRF.BasisZ) * (180 / Math.PI);
            XYZ DIR = fi.FacingOrientation;

            ElementId elementIdToIsolate = fi.Id;

            OverrideGraphicSettings ogsFade = new OverrideGraphicSettings();
            ogsFade.SetSurfaceTransparency(100);
            ogsFade.SetSurfaceForegroundPatternVisible(false);
            ogsFade.SetSurfaceBackgroundPatternVisible(false);
            ogsFade.SetHalftone(true);

            OverrideGraphicSettings ogsIsolate = new OverrideGraphicSettings();
            ogsIsolate.SetSurfaceTransparency(0);
            ogsIsolate.SetSurfaceForegroundPatternVisible(true);
            ogsIsolate.SetSurfaceBackgroundPatternVisible(true);
            ogsIsolate.SetHalftone(false);

            List<ElementId> visible = new List<ElementId>();

            ViewPlan floorView = null;

            BoundingBoxXYZ room_box2 = fi.get_BoundingBox(null);
            XYZ mas_10 = new XYZ(1, 1, 1);
            XYZ max_p = room_box2.Max + mas_10;
            XYZ min_p = room_box2.Min - mas_10;
            BoundingBoxXYZ boundingBoxXYZ = new BoundingBoxXYZ();
            boundingBoxXYZ.Min = min_p;
            boundingBoxXYZ.Max = max_p;
            View3D view3D = null;

            using (Transaction t1 = new Transaction(doc))
            {
                t1.Start("Create views from selected family");

                if (form.checkBox1.Checked)
                {
                    Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateUnbound(lp.Point, DIR);
                    XYZ vect1 = line2.Direction * (1100 / 304.8);
                    XYZ vect2 = vect1 + lp.Point;
                    ModelCurve mc = Makeline(doc, lp.Point, vect2);
                    Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateBound(lp.Point, vect2);

                    double angle1 = XYZ.BasisY.AngleTo(line3.Direction);
                    double angleDegrees = angle1 * 180 / Math.PI;
                    int number = lp.Point.X.CompareTo(line3.GetEndPoint(1).X);
                    double round = Math.Round(angleDegrees);
                    double rot_orig = lp.Rotation;
                    double rot_orig_deg = rot_orig * 180 / Math.PI;
                    if (lp.Point.X < line3.GetEndPoint(1).X)
                    {
                        angle1 = 2 * Math.PI - angle1;
                    }
                   


                    double angleDegreesCorrected = angle1 * 180 / Math.PI;
                    
                    Autodesk.Revit.DB.Line line_dir_back1 = Autodesk.Revit.DB.Line.CreateUnbound(lp.Point, InvCoord(DIR));
                    XYZ vect_dir_back1 = line_dir_back1.Direction * (1100 / 304.8);
                    XYZ vect_dir_back2 = vect_dir_back1 + lp.Point;
                    Autodesk.Revit.DB.Line line_dir_back2 = Autodesk.Revit.DB.Line.CreateBound(lp.Point, vect_dir_back2);

                    double angle2 = XYZ.BasisY.AngleTo(line_dir_back2.Direction);
                    double angleDegrees2 = angle2 * 180 / Math.PI;
                    double round2 = Math.Round(angleDegrees2);
                    double rot_orig2 = lp.Rotation;
                    if (lp.Point.X < line_dir_back2.GetEndPoint(1).X )
                    {
                        angle2 = 2 * Math.PI - angle2;
                    }

                    double angleDegreesCorrected2 = angle2 * 180 / Math.PI;
                    
                    Autodesk.Revit.DB.Transform curveTransform = line_dir_back2.ComputeDerivatives(0.5, true);
                    XYZ origin = curveTransform.Origin;
                    XYZ viewdir = curveTransform.BasisX.Normalize();
                    XYZ up = XYZ.BasisZ;
                    XYZ right = up.CrossProduct(viewdir);
                    
                    Autodesk.Revit.DB.Line line_dir_right1 = Autodesk.Revit.DB.Line.CreateUnbound(lp.Point, InvCoord(right));
                    XYZ vect_dir_right1 = line_dir_right1.Direction * (1100 / 304.8);
                    XYZ vect_dir_right2 = vect_dir_right1 + lp.Point;
                    ModelCurve mline_dir_right2 = Makeline(doc, lp.Point, vect_dir_right2);
                    Autodesk.Revit.DB.Line line_dir_right2 = Autodesk.Revit.DB.Line.CreateBound(lp.Point, vect_dir_right2);

                    double angle3 = XYZ.BasisY.AngleTo(line_dir_right2.Direction);
                    double angleDegrees3 = angle3 * 180 / Math.PI;
                    if (lp.Point.X < line_dir_right2.GetEndPoint(1).X)
                    {
                        angle3 = 2 * Math.PI - angle3;
                    }

                    double angleDegreesCorrected3 = angle3 * 180 / Math.PI;

                    Autodesk.Revit.DB.Line line_dir_left1 = Autodesk.Revit.DB.Line.CreateUnbound(lp.Point, right);
                    XYZ vect_dir_left1 = line_dir_left1.Direction * (1100 / 304.8);
                    XYZ vect_dir_left2 = vect_dir_left1 + lp.Point;
                    ModelCurve mline_dir_left2 = Makeline(doc, lp.Point, vect_dir_left2);
                    Autodesk.Revit.DB.Line mline_dir_left3 = Autodesk.Revit.DB.Line.CreateBound(lp.Point, vect_dir_left2);

                    double angle4 = XYZ.BasisY.AngleTo(mline_dir_left3.Direction);
                    double angleDegrees4 = angle4 * 180 / Math.PI;
                    if (lp.Point.X < mline_dir_left3.GetEndPoint(1).X)
                    {
                        angle4 = 2 * Math.PI - angle4;
                    }
                    double angleDegreesCorrected4 = angle4 * 180 / Math.PI;
                    
                    Autodesk.Revit.DB.Line axis = Autodesk.Revit.DB.Line.CreateBound(line3.GetEndPoint(1), new XYZ(line3.GetEndPoint(1).X, line3.GetEndPoint(1).Y, line3.GetEndPoint(1).Z + 10));
                    Autodesk.Revit.DB.Line axis2 = Autodesk.Revit.DB.Line.CreateBound(line_dir_back2.GetEndPoint(1), new XYZ(line_dir_back2.GetEndPoint(1).X, line_dir_back2.GetEndPoint(1).Y, line_dir_back2.GetEndPoint(1).Z + 10));
                    Autodesk.Revit.DB.Line axis3 = Autodesk.Revit.DB.Line.CreateBound(line_dir_right2.GetEndPoint(1), new XYZ(line_dir_right2.GetEndPoint(1).X, line_dir_right2.GetEndPoint(1).Y, line_dir_right2.GetEndPoint(1).Z + 10));
                    Autodesk.Revit.DB.Line axis4 = Autodesk.Revit.DB.Line.CreateBound(mline_dir_left3.GetEndPoint(1), new XYZ(mline_dir_left3.GetEndPoint(1).X, mline_dir_left3.GetEndPoint(1).Y, mline_dir_left3.GetEndPoint(1).Z + 10));
                    
                    ElevationMarker marker = ElevationMarker.CreateElevationMarker(doc, vftele.Id, line3.GetEndPoint(1), 100);
                    ViewSection elevation1 = marker.CreateElevation(doc, doc.ActiveView.Id, 1);
                    
                    try
                    {
                        elevation1.Name = el.Name + " Elevation 1" ;
                    }
                    catch (Exception)
                    {
                        
                        MessageBox.Show("Name or Number might be already in use!", "");
                    }

                    if (angleDegreesCorrected2 > 160 && angleDegreesCorrected2 < 200)
                    {
                        angle2 = angle2 / 2;
                        marker.Location.Rotate(axis, angle2);
                    }
                    marker.Location.Rotate(axis, angle2);

                    views_.Add(elevation1);

                    ElevationMarker marker2 = ElevationMarker.CreateElevationMarker(doc, vftele.Id, line_dir_back2.GetEndPoint(1), 100);
                    ViewSection elevation2 = marker2.CreateElevation(doc, doc.ActiveView.Id, 1);
                    //marker2.Location.Rotate(axis2, angle1);

                    try
                    {
                        elevation2.Name = el.Name + " Elevation 2";
                    }
                    catch (Exception)
                    {

                        MessageBox.Show("Name or Number might be already in use!", "");
                    }


                    if (angleDegreesCorrected > 160 && angleDegreesCorrected < 200)
                    {
                        angle1 = angle1 / 2;
                        marker2.Location.Rotate(axis2, angle1);
                    }
                    marker2.Location.Rotate(axis2, angle1);
                    views_.Add(elevation2);
                    //elevation2.CropBox = fi.get_BoundingBox(null);

                    ElevationMarker marker3 = ElevationMarker.CreateElevationMarker(doc, vftele.Id, line_dir_right2.GetEndPoint(1), 100);
                    ViewSection elevation3 = marker3.CreateElevation(doc, doc.ActiveView.Id, 1);
                    //marker3.Location.Rotate(axis3, angle4);


                    try
                    {
                        elevation3.Name = el.Name + " Elevation 3";
                    }
                    catch (Exception)
                    {

                        MessageBox.Show("Name or Number might be already in use!", "");
                    }


                    if (angleDegreesCorrected4 > 160 && angleDegreesCorrected4 < 200)
                    {
                        angle4 = angle4 / 2;
                        marker3.Location.Rotate(axis3, angle4);
                    }
                    marker3.Location.Rotate(axis3, angle4);
                    views_.Add(elevation3);
                    //elevation3.CropBox = fi.get_BoundingBox(null);

                    ElevationMarker marker4 = ElevationMarker.CreateElevationMarker(doc, vftele.Id, mline_dir_left3.GetEndPoint(1), 100);
                    ViewSection elevation4 = marker4.CreateElevation(doc, doc.ActiveView.Id, 1);
                    //marker4.Location.Rotate(axis4, angle3);

                    try
                    {
                        elevation4.Name = el.Name + " Elevation 4";
                    }
                    catch (Exception)
                    {

                        MessageBox.Show("Name or Number might be already in use!", "");
                    }

                    if (angleDegreesCorrected3 > 160 && angleDegreesCorrected3 < 200)
                    {
                        angle3 = angle3 / 2;
                        marker4.Location.Rotate(axis4, angle3);
                    }
                    marker4.Location.Rotate(axis4, angle3);
                    views_.Add(elevation4);
                    //elevation4.CropBox = fi.get_BoundingBox(null);



                    var _X2 = TRF.BasisX;
                    var _Z2 = TRF.BasisZ;
                    var _Xtran2 = TRF.OfVector(TRF.BasisX);
                    var objrot2 = TRF.BasisX.AngleOnPlaneTo(_Xtran, TRF.BasisZ) * (180 / Math.PI);


                    BoundingBoxXYZ room_box = fi.get_BoundingBox(null);
                    BoundingBoxXYZ bb = fi.get_BoundingBox(null);

                    XYZ minZ = bb.Min;
                    XYZ maxZ = bb.Max;

                    Autodesk.Revit.DB.Transform trf = bb.Transform;
                    
                    foreach (var item in views_)
                    {
                        foreach (Element e in new FilteredElementCollector(doc, item.Id).WhereElementIsNotElementType())
                        {
                            if (e.Id == elementIdToIsolate || elementIdToIsolate == null)
                            {
                                
                            }
                            else
                            {
                                if (e.CanBeHidden(item) )
                                {
                                    visible.Add(e.Id);
                                }


                            }
                        }
                    }
                    foreach (var item in views_)
                    {
                        item.HideElements(visible);
                    }
                    visible.Clear();
                }
                
                if (form.checkBox2.Checked)
                {
                    view3D = View3D.CreateIsometric(doc, viewFamilyTypes.First().Id);
                    
                    view3D.SetSectionBox(boundingBoxXYZ);

                    try
                    {
                        view3D.Name = el.Name  + " 3D View";
                    }
                    catch (Exception)
                    {

                        MessageBox.Show("Name or Number might be already in use!", "");
                    }

                   

                    foreach (Element e in new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType())
                    {
                        if (e.Id == elementIdToIsolate || elementIdToIsolate == null)
                        {
                            //floorView.SetElementOverrides(e.Id, ogsIsolate);
                            //view3D.SetElementOverrides(e.Id, ogsIsolate);

                        }
                        else
                        {
                            //floorView.SetElementOverrides(e.Id, ogsFade);
                            //view3D.SetElementOverrides(e.Id, ogsIsolate);

                            if (e.CanBeHidden(view3D) )
                            {
                                visible.Add(e.Id);
                            }

                        }
                    }

                    view3D.HideElements(visible);
                    visible.Clear();
                }


                if (form.checkBox3.Checked)
                {
                    try
                    {
                        ElementId id = fi.LevelId;
                        string name = fi.Name;

                        BoundingBoxXYZ room_box3 = fi.get_BoundingBox(null);
                        XYZ mas_103 = new XYZ(1, 1, 1);
                        XYZ max_p3 = room_box3.Max + mas_10;
                        XYZ min_p3 = room_box3.Min - mas_10;

                        BoundingBoxXYZ boundingBoxXYZ3 = new BoundingBoxXYZ();
                        boundingBoxXYZ3.Min = min_p;
                        boundingBoxXYZ3.Max = max_p;

                        //ElementId level113 = doc.ActiveView.GenLevel.Id;
                        floorView = ViewPlan.Create(doc, floor_plan_view.Id, id);

                        floorView.CropBox = boundingBoxXYZ;
                        floorView.CropBox.Enabled = true;


                        ViewCropRegionShapeManager regionManager = floorView.GetCropRegionShapeManager();

                        regionManager.BottomAnnotationCropOffset = 0.01;

                        regionManager.LeftAnnotationCropOffset = 0.01;

                        regionManager.RightAnnotationCropOffset = 0.01;

                        regionManager.TopAnnotationCropOffset = 0.01;
                        try
                        {
                            floorView.Name = el.Name + " PlanView";
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Name or Number might be already in use!", "");
                        }
                        floorView.CropBoxActive = true;

                        
                        foreach (Element e in new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType())
                        {
                            if (e.Id == elementIdToIsolate || elementIdToIsolate == null)
                            {
                               
                            }
                            else
                            {

                                if (e.CanBeHidden(floorView) )
                                {
                                    visible.Add(e.Id);
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Floor Plan can not be made, check family have a level!", "Failed");
                    }
                    floorView.HideElements(visible);
                }
                t1.Commit();
            }
        return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class CreateSchedule : IExternalCommand
    {

        static AddInId appId = new AddInId(new Guid("BFB59CE4-49D2-4C53-84D8-726E441220DD"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            UIApplication uiapp = commandData.Application;
            //---------------------------------------- FILTERS ------------------------------------


            ViewFamilyType vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Section == x.ViewFamily);
            ViewFamilyType vftele = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Elevation == x.ViewFamily);

            FilteredElementCollector viewTemplate = new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.View));
            ICollection<Element> VTcolector = viewTemplate.ToElements();
            
            FilteredElementCollector fec = new FilteredElementCollector(doc);
            fec.OfClass(typeof(ViewSection));
            var viewtem = fec.Cast<ViewSection>().Where<ViewSection>(vp => vp.IsTemplate);
            Form4 form = new Form4();
            ScheduleFieldId fieldid_1 = null;
            ScheduleField foundField = null;
            ScheduleDefinition definition = null;

            //------------------------Walls lists---------------------------------------------------
            List<Element> wallsEle = new List<Element>();
            List<string> _comment = new List<string>();
            List<string> _mark = new List<string>();
            List<ViewSchedule> lista_sch = new List<ViewSchedule>();
            List<ElementId> templaId = new List<ElementId>();

            string debug = null;

            string project_param = /*"Schedule Identifier"*/ "Comments";




            foreach (Element e in new FilteredElementCollector(doc).OfClass(typeof(Wall)))
            {
                try
                {
                    ParameterMap parmap = e.ParametersMap;
                    Parameter par = parmap.get_Item(project_param);
                    string value = par.AsString();

                    if (value != null)
                    {
                        _comment.Add(value);
                    }
                    
                }
                catch (Exception)
                {
                    TaskDialog.Show("!", "Project requires (Schedule Identifier) parameter");
                    return Autodesk.Revit.UI.Result.Cancelled;
                }
            }

           

            foreach (var item in _comment)
            {
                if (item != null)
                {
                    if (!form.comboBox1.Items.Contains(item))
                    {
                        form.comboBox1.Items.Add(item);
                    }
                }
            }

            foreach (var item in viewtem)
            {
                form.comboBox2.Items.Add(item.Name);
                templaId.Add(item.Id);
            }

            if (_comment.Count == 0 )
            {
                TaskDialog.Show("!", "No walls contain a identifier");
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            //int selectedindex = form.comboBox1.SelectedIndex;
            //string identifier = _comment.ElementAt(selectedindex).ToString();

            ElementId choose_template;
            try
            {
                int index_tBox = form.comboBox2.SelectedIndex;
                choose_template = templaId.ElementAt(index_tBox);
            }
            catch (Exception)
            {

                TaskDialog.Show("! ", "A template must be selected");
                TaskDialog.Show("! ", "Task Cancelled");
                return Autodesk.Revit.UI.Result.Cancelled;
            }
            

            if (_comment == null)
                form.Close();
            if (form.Equals(false))
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            foreach (Element e in new FilteredElementCollector(doc).OfClass(typeof(Wall)))
            {
                try
                {

                    ParameterMap parmap = e.ParametersMap;
                    Parameter par = parmap.get_Item(project_param);
                    string value = par.AsString();


                    if (form.comboBox1.SelectedItem.ToString() == value)
                    {
                        Element E = e;
                        wallsEle.Add(E);
                    }
                }
                catch (Exception)
                {
                    continue;
                }

            }

            if (form.checkBox1.Checked == true)
            {
                using (Transaction t = new Transaction(doc, "Create single-category"))
                {
                    t.Start();

                    int count = 0;
                    foreach (var item in wallsEle)
                    {


                        try
                        {
                            ViewSchedule vs_ = ViewSchedule.CreateSchedule(doc, new ElementId(BuiltInCategory.OST_Walls));

                            vs_.Name = item.Name + " " + count;
                            definition = vs_.Definition;
                        }
                        catch (Exception)
                        {
                            TaskDialog.Show("! ", "Task failed, check if view name is already in use");
                            TaskDialog.Show("! ", "Task Cancelled");
                            return Autodesk.Revit.UI.Result.Cancelled;
                        }



                        SchedulableField schedulableField = definition.GetSchedulableFields().FirstOrDefault<SchedulableField>(sf => sf.ParameterId == new ElementId(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS));
                        SchedulableField schedulableField_2 = definition.GetSchedulableFields().FirstOrDefault<SchedulableField>(sf => sf.ParameterId == new ElementId(BuiltInParameter.WALL_BASE_CONSTRAINT));
                        SchedulableField schedulableField_3 = definition.GetSchedulableFields().FirstOrDefault<SchedulableField>(sf => sf.ParameterId == new ElementId(BuiltInParameter.WALL_ATTR_WIDTH_PARAM));
                        SchedulableField schedulableField_4 = definition.GetSchedulableFields().FirstOrDefault<SchedulableField>(sf => sf.ParameterId == new ElementId(BuiltInParameter.ALL_MODEL_MARK));

                        if (schedulableField != null)
                        {

                            definition.AddField(schedulableField);
                            definition.AddField(schedulableField_2);
                            definition.AddField(schedulableField_3);
                            definition.AddField(schedulableField_4);
                        }

                        ElementId paramId = new ElementId(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);

                        foreach (ScheduleFieldId fieldId in definition.GetFieldOrder())
                        {
                            foundField = definition.GetField(fieldId);
                            if (foundField.ParameterId == paramId)
                            {
                                fieldid_1 = foundField.FieldId;
                            }
                        }

                        definition.AddFilter(new ScheduleFilter(fieldid_1, ScheduleFilterType.Equal, "IDENTIFIER"));

                        doc.Regenerate();

                        count++;
                    }
                    t.Commit();
                }

            }

            ViewSection vs = null;
            double lenght = 0;
            

            for (int i = 0; i < wallsEle.ToArray().Length; i++)
            {

                Wall wall = null;
                if (wallsEle.ToArray()[i] != null)

                    wall = wallsEle.ToArray()[i] as Wall;


                XYZ orientation1 = wall.Orientation;
                //data += orientation1.ToString() + "\n";


                BoundingBoxXYZ bb = wall.get_BoundingBox(null);
                double minZ = bb.Min.Z + 3.5;
                double maxZ = bb.Max.Z + 1;

                //double h = maxZ - minZ;
                //Level level = doc.ActiveView.GenLevel;
                //double top = 10 + level.Elevation;
                //double bottom = level.Elevation;

                LocationCurve lc = wall.Location as LocationCurve;
                Autodesk.Revit.DB.Line origWallLine = lc.Curve as Autodesk.Revit.DB.Line;


                CurtainGrid cgrid = wall.CurtainGrid;
                Options options = new Options();
                options.ComputeReferences = true;
                options.IncludeNonVisibleObjects = true;
                options.View = doc.ActiveView;

                GeometryElement geomElem = wall.get_Geometry(options);
                List<Autodesk.Revit.DB.Line> line_list = new List<Autodesk.Revit.DB.Line>();
                List<Autodesk.Revit.DB.Line> ver_line_list = new List<Autodesk.Revit.DB.Line>();


                try
                {
                    foreach (GeometryObject obj in geomElem)
                    {
                        Visibility vis = obj.Visibility;

                        string visString = vis.ToString();


                        Autodesk.Revit.DB.Line line_ = obj as Autodesk.Revit.DB.Line;
                        Solid solid = obj as Solid;

                        if (geomElem.ToArray().First() == obj)
                        {
                            lenght = line_.ApproximateLength;
                        }


                        XYZ dir = new XYZ(0, 0, 1);
                        if (line_ != null)
                        {
                            if (line_.ApproximateLength == lenght)
                            {
                                if (line_.Direction.Z == 1.0)
                                {
                                    line_list.Add(line_);
                                }
                            }
                        }

                        XYZ dir2 = new XYZ(0, 1, 0);
                        if (line_ != null)
                        {
                            if (line_.ApproximateLength == lenght)
                            {
                                if (line_.Direction.Z == 1.0)
                                {
                                    ver_line_list.Add(line_);
                                }
                            }
                        }
                    }
                }
                catch (Exception)
                {

                    TaskDialog.Show("! ", "Wall Selected must to be a curtain wall type");
                    return Autodesk.Revit.UI.Result.Cancelled;
                }

               





                Autodesk.Revit.DB.Curve offsetWallLine = origWallLine.CreateOffset(3, XYZ.BasisZ);
                //double unconnected_height = wall.LookupParameter("unconnected height");


                XYZ p = offsetWallLine.GetEndPoint(0);
                XYZ q = offsetWallLine.GetEndPoint(1);


                XYZ p2 = new XYZ(p.X, p.Y, maxZ + 1);
                XYZ q2 = new XYZ(q.X, q.Y, maxZ + 1);
                XYZ p2_ = new XYZ(p.X, p.Y, maxZ + 1);
                XYZ q2_ = new XYZ(q.X, q.Y, maxZ + 1);
                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(p2, q2);
                Autodesk.Revit.DB.Line line2_ = Autodesk.Revit.DB.Line.CreateBound(p2_, q2_);
                Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(p, p2);

                Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(p, p2);
                Autodesk.Revit.DB.Curve curve2 = Autodesk.Revit.DB.Line.CreateBound(q, q2);

                Autodesk.Revit.DB.Curve line3_ver = Autodesk.Revit.DB.Line.CreateBound(q, p);
                Autodesk.Revit.DB.Curve line4_ver = Autodesk.Revit.DB.Line.CreateBound(p2, q2);

                Autodesk.Revit.DB.Line ver_line1 = Autodesk.Revit.DB.Line.CreateBound(p, p2);


                double dist = p.DistanceTo(q);
                XYZ v = p - q;
                double halfLength = v.GetLength() / 2;
                double offset = 5; // offset by 3 feet.
                XYZ min = new XYZ(-halfLength, 0, -offset);
                XYZ max = new XYZ(halfLength, 10, offset);
                XYZ midpoint = q + 0.5 * v; // q get lower midpoint. 
                XYZ walldir = v.Normalize();
                XYZ up = XYZ.BasisZ;
                XYZ viewdir = walldir.CrossProduct(up);
                Autodesk.Revit.DB.Transform t = Autodesk.Revit.DB.Transform.Identity;
                t.Origin = midpoint;
                t.BasisX = walldir;
                t.BasisY = up;
                t.BasisZ = viewdir;
                BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
                sectionBox.Transform = t /*transform*/;
                sectionBox.Min = new XYZ(-halfLength , -0.5, 0);
                sectionBox.Max = new XYZ(halfLength + 0.5 , maxZ, 10);
                ViewFamilyType vft1 = vft;




                using (Transaction tx = new Transaction(doc))
                {
                    try
                    {
                        tx.Start("Create wall Section");
                        vs = ViewSection.CreateSection(doc, vft.Id, sectionBox);
                        vs.Name = form.textBox1.Text + " " + wallsEle.ToArray()[i].Name + " " + i;
                        vs.ViewTemplateId = choose_template;
                        vs.Scale = 100;
                        doc.Regenerate();



                        tx.Commit();
                        TaskDialog.Show("View created", "A section view was created");

                    }
                    catch
                    {

                        TaskDialog.Show("! ", "Task failed, check if view name is already in use");
                        TaskDialog.Show("! ", "Task Cancelled");
                        return Autodesk.Revit.UI.Result.Cancelled;
                    }



                }

               
                List<Autodesk.Revit.DB.DetailCurve> dCurve_list = new List<DetailCurve>();
                ReferenceArray refArray = new ReferenceArray();
                uidoc.ActiveView = vs;
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("create dimension by mullion");

                    DetailCurve dCurve1 = null;
                    
                    foreach (var Line in line_list)
                    {
                        if (!doc.IsFamilyDocument)
                        {
                            //Reference gridRef = null;
                            //gridRef = line.Reference;
                            //refArray.Append(gridRef);
                            dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, Line);
                            dCurve_list.Add(dCurve1);
                        }
                        else
                        {
                            //Reference gridRef = null;
                            //gridRef = line.Reference;
                            //refArray.Append(gridRef);
                            dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, Line);
                            dCurve_list.Add(dCurve1);
                        }
                    }
                    foreach (var curve_ in dCurve_list)
                    {
                        refArray.Append(curve_.GeometryCurve.Reference);
                    }


                    try
                    {
                        if (!doc.IsFamilyDocument)
                        {
                            doc.Create.NewDimension(
                              doc.ActiveView, line, refArray);
                        }
                        else
                        {
                            doc.FamilyCreate.NewDimension(
                              doc.ActiveView, line, refArray);
                        }
                    }
                    catch (Exception)
                    {

                        TaskDialog.Show("! ", "Task Cancelled");
                        goto finish;
                    }


                    finish: tx.Commit();
                }
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Overhall dimension");

                    DetailCurve dCurve1 = null;
                    DetailCurve dCurve2 = null;
                    //Reference gridRef1 = null;
                    //Reference gridRef2 = null;
                    ReferenceArray refArray2 = new ReferenceArray();

                    if (!doc.IsFamilyDocument)
                    {
                        dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, curve1);
                        dCurve2 = doc.Create.NewDetailCurve(doc.ActiveView, curve2);
                        //gridRef1 = curve1.Reference;
                    }
                    else
                    {
                        dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, curve1);
                        dCurve2 = doc.Create.NewDetailCurve(doc.ActiveView, curve2);
                        //gridRef2 = curve2.Reference;
                    }

                    refArray2.Append(/*curve1.Reference*/dCurve1.GeometryCurve.Reference /*gridRef1*/);
                    refArray2.Append(/*curve2.Reference*/dCurve2.GeometryCurve.Reference /*gridRef2*/);

                    if (!doc.IsFamilyDocument)
                    {
                        doc.Create.NewDimension(
                          doc.ActiveView, line2_, refArray2);
                    }
                    else
                    {
                        doc.FamilyCreate.NewDimension(
                          doc.ActiveView, line2_, refArray2);
                    }

                    tx.Commit();
                }
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("ver dim");

                    DetailCurve dCurve1 = null;
                    DetailCurve dCurve2 = null;
                    //Reference gridRef1 = null;
                    //Reference gridRef2 = null;
                    ReferenceArray refArray2 = new ReferenceArray();

                    if (!doc.IsFamilyDocument)
                    {
                        dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, line3_ver);
                        dCurve2 = doc.Create.NewDetailCurve(doc.ActiveView, line4_ver);
                        //gridRef1 = curve1.Reference;
                    }
                    else
                    {
                        dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, line3_ver);
                        dCurve2 = doc.Create.NewDetailCurve(doc.ActiveView, line4_ver);
                        //gridRef2 = curve2.Reference;
                    }

                    refArray2.Append(/*curve1.Reference*/dCurve1.GeometryCurve.Reference /*gridRef1*/);
                    refArray2.Append(/*curve2.Reference*/dCurve2.GeometryCurve.Reference /*gridRef2*/);

                    if (!doc.IsFamilyDocument)
                    {
                        doc.Create.NewDimension(
                          doc.ActiveView, ver_line1, refArray2);
                    }
                    else
                    {
                        doc.FamilyCreate.NewDimension(
                          doc.ActiveView, ver_line1, refArray2);
                    }
                    tx.Commit();
                }
            }
            //foreach (Element e in wallsEle)
            //{


            //    uidoc.ActiveView = vs;
            //    List<Autodesk.Revit.DB.DetailCurve> dCurve_list = new List<DetailCurve>();
            //    ReferenceArray refArray = new ReferenceArray();

            //    using (Transaction tx = new Transaction(doc))
            //    {
            //        tx.Start("create dimension by mullion");

            //        DetailCurve dCurve1 = null;

            //        foreach (var Line in line_list)
            //        {
            //            if (!doc.IsFamilyDocument)
            //            {
            //                //Reference gridRef = null;
            //                //gridRef = line.Reference;
            //                //refArray.Append(gridRef);
            //                dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, Line);
            //                dCurve_list.Add(dCurve1);
            //            }
            //            else
            //            {
            //                //Reference gridRef = null;
            //                //gridRef = line.Reference;
            //                //refArray.Append(gridRef);
            //                dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, Line);
            //                dCurve_list.Add(dCurve1);
            //            }
            //        }
            //        foreach (var curve_ in dCurve_list)
            //        {
            //            refArray.Append(curve_.GeometryCurve.Reference);
            //        }



            //        if (!doc.IsFamilyDocument)
            //        {
            //            doc.Create.NewDimension(
            //              doc.ActiveView, line, refArray);
            //        }
            //        else
            //        {
            //            doc.FamilyCreate.NewDimension(
            //              doc.ActiveView, line, refArray);
            //        }

            //        tx.Commit();
            //    }
            //    using (Transaction tx = new Transaction(doc))
            //    {
            //        tx.Start("Overhall dimension");

            //        DetailCurve dCurve1 = null;
            //        DetailCurve dCurve2 = null;
            //        //Reference gridRef1 = null;
            //        //Reference gridRef2 = null;
            //        ReferenceArray refArray2 = new ReferenceArray();

            //        if (!doc.IsFamilyDocument)
            //        {
            //            dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, curve1);
            //            dCurve2 = doc.Create.NewDetailCurve(doc.ActiveView, curve2);
            //            //gridRef1 = curve1.Reference;
            //        }
            //        else
            //        {
            //            dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, curve1);
            //            dCurve2 = doc.Create.NewDetailCurve(doc.ActiveView, curve2);
            //            //gridRef2 = curve2.Reference;
            //        }

            //        refArray2.Append(/*curve1.Reference*/dCurve1.GeometryCurve.Reference /*gridRef1*/);
            //        refArray2.Append(/*curve2.Reference*/dCurve2.GeometryCurve.Reference /*gridRef2*/);

            //        if (!doc.IsFamilyDocument)
            //        {
            //            doc.Create.NewDimension(
            //              doc.ActiveView, line2_, refArray2);
            //        }
            //        else
            //        {
            //            doc.FamilyCreate.NewDimension(
            //              doc.ActiveView, line2_, refArray2);
            //        }

            //        tx.Commit();
            //    }
            //    using (Transaction tx = new Transaction(doc))
            //    {
            //        tx.Start("ver dim");

            //        DetailCurve dCurve1 = null;
            //        DetailCurve dCurve2 = null;
            //        //Reference gridRef1 = null;
            //        //Reference gridRef2 = null;
            //        ReferenceArray refArray2 = new ReferenceArray();

            //        if (!doc.IsFamilyDocument)
            //        {
            //            dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, line3_ver);
            //            dCurve2 = doc.Create.NewDetailCurve(doc.ActiveView, line4_ver);
            //            //gridRef1 = curve1.Reference;
            //        }
            //        else
            //        {
            //            dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, line3_ver);
            //            dCurve2 = doc.Create.NewDetailCurve(doc.ActiveView, line4_ver);
            //            //gridRef2 = curve2.Reference;
            //        }

            //        refArray2.Append(/*curve1.Reference*/dCurve1.GeometryCurve.Reference /*gridRef1*/);
            //        refArray2.Append(/*curve2.Reference*/dCurve2.GeometryCurve.Reference /*gridRef2*/);

            //        if (!doc.IsFamilyDocument)
            //        {
            //            doc.Create.NewDimension(
            //              doc.ActiveView, ver_line1, refArray2);
            //        }
            //        else
            //        {
            //            doc.FamilyCreate.NewDimension(
            //              doc.ActiveView, ver_line1, refArray2);
            //        }
            //        tx.Commit();
            //    }
            //}
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Wall_Bounding_room : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F46AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            List<Element> ele = new List<Element>();

            foreach (Element e in new FilteredElementCollector(doc).OfClass(typeof(Wall)))
            {
                try
                {
                    ParameterMap parmap = e.ParametersMap;
                    Parameter par = parmap.get_Item("Room Bounding");
                    string value = par.AsString();

                    Parameter value2 = e.LookupParameter("Room Bounding");
                    string value3 = value2.Definition.Name;

                    var value_ = value2.AsValueString();

                  
                    if (value_ == "No")
                    {
                        ele.Add(e);
                    }
                }
                catch (Exception)
                {
                    //    TaskDialog.Show("!", "Project requires (Schedule Identifier) parameter");
                    return Autodesk.Revit.UI.Result.Cancelled;
                }
            }


            TaskDialog.Show("!", ele.Count.ToString() + " Wall were selected" );
            uidoc.Selection.SetElementIds(ele.Select(q => q.Id).ToList());

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class make_line_from_surface_normal : IExternalCommand
    {
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb, bool click)
        {

            Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = null;
                if (click)
                {
                    aSubBcrossz = aSubB.CrossProduct(XYZ.BasisX);
                   
                }

                if (click == false)
                {
                    aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                }

                
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }
            Autodesk.Revit.DB.Line line = null;

            //Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);
            if (click)
            {
                line = Autodesk.Revit.DB.Line.CreateUnbound(ptb, XYZ.BasisY);

                if (line2.Direction.X == 1 || line2.Direction.X == -1)
                {
                    line = Autodesk.Revit.DB.Line.CreateUnbound(ptb, XYZ.BasisY );
                }
              

                if (line2.Direction.Y == 1 || line2.Direction.Y == -1)
                {
                    line = Autodesk.Revit.DB.Line.CreateUnbound(ptb, XYZ.BasisX);
                }
                
            }
            else
            {
                line = Autodesk.Revit.DB.Line.CreateUnbound(ptb, XYZ.BasisZ);
            }


            Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(pta,
          line.Evaluate(5, false), ptb);
            
            SketchPlane skplane = SketchPlane.Create(doc, pl);
            
            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line2, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line2, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        static AddInId appId = new AddInId(new Guid("5F56AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            //ObjectSnapTypes snapTypes = ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections;
            //XYZ point = uidoc.Selection.PickPoint(snapTypes, "Select an end point or intersection");

            var pipeTypes = new FilteredElementCollector(doc).OfClass(typeof(SpotDimension)).OfType<SpotDimension>().ToList();

            Form1 form = new Form1();
            

            Reference r = uidoc.Selection.PickObject(ObjectType.Face, "Please pick a point on a " + "face for family instance insertion");

            Element e = doc.GetElement(r.ElementId);
            GeometryObject obj
              = e.GetGeometryObjectFromReference(r);

            XYZ p = r.GlobalPoint;
            XYZ v = null;

            form.ShowDialog();
            

            try
            {
                PlanarFace face = obj as PlanarFace;
                v = face.FaceNormal;
                if (v.IsZeroLength())
                {
                    v = face.FaceNormal.CrossProduct(XYZ.BasisX);
                }
            }
            catch (Exception)
            {
            }
            try
            {
                CylindricalFace face = obj as CylindricalFace;

               
               v = face.Axis.CrossProduct(XYZ.BasisZ);
                if (v.IsZeroLength())
                {
                    v = face.Axis.CrossProduct(XYZ.BasisX);
                }
            }
            catch (Exception)
            {
            }

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateUnbound(p, v);
           
            
            double movenumber;
            Double.TryParse(form.textBox1.Text, out movenumber);


            double rotnumber;
            Double.TryParse(form.textBox2.Text, out rotnumber);

            XYZ vect1 = line.Direction * (movenumber / 304.8);

            XYZ vect2 = vect1 + p;

            Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateUnbound(p, p.CrossProduct(line.Evaluate(5,false)));

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Make Line on face from input");

                ModelLine ml = Makeline(doc, p, vect2, form.radioButton1.Checked);

                //GeometryCreationUtilities.CreateLoftGeometry()
                
                XYZ cc = new XYZ(p.X, p.Y + 10, p.Z );
                Autodesk.Revit.DB.Line axis = Autodesk.Revit.DB.Line.CreateBound(p, cc);
                //ModelLine ml2 = Makeline(doc, p, cc);
                
                ElementTransformUtils.RotateElement(doc, ml.Id, axis, rotnumber);
                
                
                //XYZ bend = p2.Add(new XYZ(2, 2, 0));
                //XYZ end = p2.Add(new XYZ(3, 2, 0));
                //doc.Create.NewSpotElevation(doc.ActiveView, myRef2, p2, bend, end, p2, true);
                
                tx.Commit();
            }
            //uidoc.Selection.SetElementIds(ele.Select(q => q.Id).ToList());
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class make_pipe_by_line : IExternalCommand
    {
        public XYZ InvCoord(XYZ MyCoord)
        {
            XYZ invcoord = new XYZ((Convert.ToDouble(MyCoord.X * -1)),
                (Convert.ToDouble(MyCoord.Y * -1)),
                (Convert.ToDouble(MyCoord.Z * -1)));
            return invcoord;
        }
        public XYZ CrossProduct(XYZ v1, XYZ v2)
        {
            double x, y, z;
            x = v1.Y * v2.Z - v2.Y * v1.Z;
            y = (v1.X * v2.Z - v2.X * v1.Z) * -1;
            z = v1.X * v2.Y - v2.X * v1.Y;
            var rtnvector = new XYZ(x, y, z);
            return rtnvector;
        }

        private XYZ intersect(XYZ point, XYZ direction, Autodesk.Revit.DB.Curve curve)
        {
            Autodesk.Revit.DB.Line unbound = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(point.X, point.Y, curve.GetEndPoint(0).Z), direction);
            IntersectionResultArray ira = null;
            unbound.Intersect(curve, out ira);
            if (ira == null)
            {
                TaskDialog td = new TaskDialog("Error");
                td.MainInstruction = "no intersection";
                td.MainContent = point.ToString() + Environment.NewLine + direction.ToString();
                td.Show();

                return null;
            }
            IntersectionResult ir = ira.get_Item(0);
            return ir.XYZPoint;
        }
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        static AddInId appId = new AddInId(new Guid("5F56AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            List<Autodesk.Revit.DB.Curve> crvs = new List<Autodesk.Revit.DB.Curve>();

            MessageBox.Show("Chose one or more 3D lines to create revit pipes", "Select");


            ICollection<Reference> my_lines = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select lines");
            foreach (var item_myRefWall in my_lines)
            {
                Element ele = doc.GetElement(item_myRefWall);
                GeometryObject geoobj = ele.GetGeometryObjectFromReference(item_myRefWall);
                Face face = geoobj as Face;
                LocationCurve locationCurve2 = ele.Location as LocationCurve;

                crvs.Add(locationCurve2.Curve);
            }
            
            var mepSystemTypes = new FilteredElementCollector(doc).OfClass(typeof(PipingSystemType)).OfType<PipingSystemType>().ToList();
            
            var domesticHotWaterSystemType = mepSystemTypes.FirstOrDefault(st => st.SystemClassification ==MEPSystemClassification.DomesticHotWater);

            if (domesticHotWaterSystemType == null)
            {
                message = "Could not found Domestic Hot Water System Type";
                return Result.Failed;
            }
            
            var pipeTypes = new FilteredElementCollector(doc).OfClass(typeof(PipeType)).OfType<PipeType>().ToList();
            
            var firstPipeType = pipeTypes.FirstOrDefault();

            if (firstPipeType == null)
            {
                message = "Could not found Pipe Type";
                return Result.Failed;
            }
            
            FilteredElementCollector collector2 = new FilteredElementCollector(doc);
            collector2.OfClass(typeof(Level));
            Level lv = collector2.First() as Level;

            if (lv == null)
            {
                message = "Wrong Active View";
                return Result.Failed;
            }

            using (Transaction t = new Transaction(doc, "Create  pipe by line"))
            {
                t.Start();

                int num = 0;

                foreach (var crv in crvs)
                {

                    num++;

                    ////XYZ CROSSPT = CrossProduct(crv.GetEndPoint(0), crv.GetEndPoint(1));
                    //Autodesk.Revit.DB.Line xline = crv as Autodesk.Revit.DB.Line;
                    //XYZ Dir = xline.Direction;
                    //XYZ CROSSPT = CrossProduct(crv.GetEndPoint(0), Dir);
                    //ModelLine ml = Makeline(doc, crv.GetEndPoint(0), CROSSPT);


                    //XYZ norm = crv.GetEndPoint(0).CrossProduct(crv.GetEndPoint(1));
                    //if (norm.GetLength() == 0)
                    //{
                    //    XYZ aSubB = crv.GetEndPoint(0).Subtract(crv.GetEndPoint(1));
                    //    XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                    //    double crosslenght = aSubBcrossz.GetLength();
                    //    if (crosslenght == 0)
                    //    {
                    //        norm = XYZ.BasisY;
                    //    }
                    //    else
                    //    {
                    //        norm = XYZ.BasisZ;
                    //    }
                    //}

                    //Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, crv.GetEndPoint(1));


                   

                    //Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(crv.GetEndPoint(0), norm);

                  
                    //Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, crv.GetEndPoint(0));
                    //doc.ActiveView.ShowActiveWorkPlane();
                    //SketchPlane skplane = SketchPlane.Create(doc, plane);

                    //SketchPlane sketch = ml.SketchPlane 
                    //doc.ActiveView.SketchPlane = skplane;
                    //doc.ActiveView.ShowActiveWorkPlane();

                    //ModelLine ml = Makeline(doc, crv.GetEndPoint(0), norm);

                    //ModelLine ml2 = Makeline(doc, crv.GetEndPoint(0), crv.GetEndPoint(1));

                    //ElementTransformUtils.RotateElement(doc, ml2.Id, line, (90 / 308.4));


                    var pipe = Pipe.Create(doc, domesticHotWaterSystemType.Id, firstPipeType.Id,
                      lv.Id, crv.GetEndPoint(0), crv.GetEndPoint(1));



                }

                MessageBox.Show(num + " Revit pipes were created (not specific type were given to them)", "Select");

                t.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class make_duck_by_line : IExternalCommand
    {
        private XYZ intersect(XYZ point, XYZ direction, Autodesk.Revit.DB.Curve curve)
        {
            Autodesk.Revit.DB.Line unbound = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(point.X, point.Y, curve.GetEndPoint(0).Z), direction);
            IntersectionResultArray ira = null;
            unbound.Intersect(curve, out ira);
            if (ira == null)
            {
                TaskDialog td = new TaskDialog("Error");
                td.MainInstruction = "no intersection";
                td.MainContent = point.ToString() + Environment.NewLine + direction.ToString();
                td.Show();

                return null;
            }
            IntersectionResult ir = ira.get_Item(0);
            return ir.XYZPoint;
        }
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        static AddInId appId = new AddInId(new Guid("8F56AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            //Reference myRef = uidoc.Selection.PickObject(ObjectType.Element);
            //Element e = doc.GetElement(myRef);
            //LocationCurve locationCurve1 = e.Location as LocationCurve;
            //Autodesk.Revit.DB.Curve gridCurve = locationCurve1.Curve;
            //XYZ end1 = gridCurve.GetEndPoint(1);


            List<Autodesk.Revit.DB.Curve> crvs = new List<Autodesk.Revit.DB.Curve>();

            ICollection<Reference> my_lines = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select lines");
            foreach (var item_myRefWall in my_lines)
            {
                Element ele = doc.GetElement(item_myRefWall);
                GeometryObject geoobj = ele.GetGeometryObjectFromReference(item_myRefWall);
                Face face = geoobj as Face;
                LocationCurve locationCurve2 = ele.Location as LocationCurve;

                crvs.Add(locationCurve2.Curve);
            }


            var mepSystemTypes = new FilteredElementCollector(doc).OfClass(typeof(MEPSystemType))
            .OfType<MEPSystemType>().ToList();
            
            var ductTypes =
              new FilteredElementCollector(doc).OfClass(typeof(DuctType)).OfType<DuctType>().First();

           

            FilteredElementCollector collector2 = new FilteredElementCollector(doc);
            collector2.OfClass(typeof(Level));
            Level lv = collector2.First() as Level;

            if (lv == null)
            {
                message = "Wrong Active View";
                return Result.Failed;
            }

            var domesticHotWaterSystemType =
              mepSystemTypes.FirstOrDefault(
                st => st.SystemClassification ==
                  MEPSystemClassification.SupplyAir);

            if (domesticHotWaterSystemType == null)
            {
                message = "Could not found Domestic Hot Water System Type";
                return Result.Failed;
            }

            using (Transaction t = new Transaction(doc, "Create  pipe by line"))
            {
                t.Start();

                foreach (var crv in crvs)
                {
                    Duct.Create(doc, domesticHotWaterSystemType.Id, ductTypes.Id,
                      lv.Id, crv.GetEndPoint(0), crv.GetEndPoint(1));
                }
                

                t.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class line_point_plane : IExternalCommand
    {
        public List<TessellatedShapeBuilderResult> Build_Tessellate2(Face faces, Autodesk.Revit.DB.Document doc)
        {
            Autodesk.Revit.DB.Mesh mesh = faces.Triangulate();
            List<XYZ> vert = new List<XYZ>();

            foreach (XYZ ij in mesh.Vertices)
            {
                XYZ vertices = ij;
                vert.Add(vertices);
            }

            TessellatedShapeBuilder builder = new TessellatedShapeBuilder();

            builder.OpenConnectedFaceSet(false);

            //Filter for Title Blocks in active document
            FilteredElementCollector materials = new FilteredElementCollector(doc)
            .OfClass(typeof(Autodesk.Revit.DB.Material))
            .OfCategory(BuiltInCategory.OST_Materials);

            ElementId materialId = materials.First().Id;

            builder.AddFace(new TessellatedFace(vert, materialId));

            builder.CloseConnectedFaceSet();
            builder.Target = TessellatedShapeBuilderTarget.AnyGeometry;
            builder.Fallback = TessellatedShapeBuilderFallback.Mesh;

            builder.Build();

            TessellatedShapeBuilderResult result3 = builder.GetBuildResult();
            List<TessellatedShapeBuilderResult> res = new List<TessellatedShapeBuilderResult>();

            if (result3.Outcome.ToString() == "Sheet")
            {
                res.Add(result3);
            }

            return res;
        }
        static AddInId appId = new AddInId(new Guid("5F92CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            List<CurveLoop> cls = new List<CurveLoop>();

            MessageBox.Show("Select a 3D line as a base to crate a new working plane ", "Select step 1");

            Reference myRef = uidoc.Selection.PickObject(ObjectType.Element);
            Element e = doc.GetElement(myRef);

            MessageBox.Show("Select a point in space to give direction to the wroking plane", "Select step 2");

            Reference myRef2 = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            Element e2 = doc.GetElement(myRef2.ElementId);
            GeometryObject geomObj2 = e2.GetGeometryObjectFromReference(myRef2);
            XYZ p2 = myRef2.GlobalPoint;

            LocationCurve locationCurve1 = e.Location as LocationCurve;
            Autodesk.Revit.DB.Curve gridCurve = locationCurve1.Curve;

            

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Make plane by line and point");


                
                Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(gridCurve.GetEndPoint(0), p2
                , gridCurve.GetEndPoint(1));

                SketchPlane sketch = SketchPlane.Create(doc, pl);
                doc.ActiveView.SketchPlane = sketch;
                doc.ActiveView.ShowActiveWorkPlane();

                
                tx.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;
            
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Closest_point_2Lines : IExternalCommand
    {
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            //Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateUnbound(ptb, XYZ.BasisZ /* XYZ.BasisZ*/);

            Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(pta,
             line.Evaluate(5, false), ptb);

            SketchPlane skplane = SketchPlane.Create(doc, pl);

            Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line2, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line2, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        static AddInId appId = new AddInId(new Guid("5F92AA78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            List<CurveLoop> cls = new List<CurveLoop>();


            Reference myRef = uidoc.Selection.PickObject(ObjectType.Element);
            Element e = doc.GetElement(myRef);

            Reference myRef2 = uidoc.Selection.PickObject(ObjectType.Element);
            Element e2 = doc.GetElement(myRef2);

            //GeometryObject geomObj2 = e2.GetGeometryObjectFromReference(myRef2);
            //XYZ p2 = myRef2.GlobalPoint;

            LocationCurve locationCurve1 = e.Location as LocationCurve;
            Autodesk.Revit.DB.Curve line1 = locationCurve1.Curve;

            LocationCurve locationCurve2 = e2.Location as LocationCurve;
            Autodesk.Revit.DB.Curve line2 = locationCurve2.Curve;


            IntersectionResult intres = line2.Project(line1.GetEndPoint(0));
            IntersectionResult intres2 = line2.Project(line1.GetEndPoint(1));

            IList<ClosestPointsPairBetweenTwoCurves> closestPoints = new List<ClosestPointsPairBetweenTwoCurves>();

            line1.ComputeClosestPoints(line2, true, false, false, out closestPoints);
            XYZ closestPoint1 = closestPoints.FirstOrDefault().XYZPointOnFirstCurve;

            line2.ComputeClosestPoints(line1, true, false, false, out closestPoints);
            XYZ closestPoint2 = closestPoints.FirstOrDefault().XYZPointOnFirstCurve;


            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Make plane by line and point");


                ModelLine ml = Makeline(doc, closestPoint1, closestPoint2);

                //ModelLine ml1 = Makeline(doc, line1.GetEndPoint(0), intres.XYZPoint);
                //ModelLine ml2 = Makeline(doc, line1.GetEndPoint(1), intres2.XYZPoint);




                tx.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;

        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class loft : IExternalCommand
    {
        public List<TessellatedShapeBuilderResult> Build_Tessellate2(Face faces, Autodesk.Revit.DB.Document doc)
        {
            Autodesk.Revit.DB.Mesh mesh = faces.Triangulate();
            List<XYZ> vert = new List<XYZ>();

            foreach (XYZ ij in mesh.Vertices)
            {
                XYZ vertices = ij;
                vert.Add(vertices);
            }

            TessellatedShapeBuilder builder = new TessellatedShapeBuilder();

            builder.OpenConnectedFaceSet(false);

            //Filter for Title Blocks in active document
            FilteredElementCollector materials = new FilteredElementCollector(doc)
            .OfClass(typeof(Autodesk.Revit.DB.Material))
            .OfCategory(BuiltInCategory.OST_Materials);

            ElementId materialId = materials.First().Id;

            builder.AddFace(new TessellatedFace(vert, materialId));

            builder.CloseConnectedFaceSet();
            builder.Target = TessellatedShapeBuilderTarget.AnyGeometry;
            builder.Fallback = TessellatedShapeBuilderFallback.Mesh;

            builder.Build();

            TessellatedShapeBuilderResult result3 = builder.GetBuildResult();
            List<TessellatedShapeBuilderResult> res = new List<TessellatedShapeBuilderResult>();

            if (result3.Outcome.ToString() == "Sheet")
            {
                res.Add(result3);
            }

            return res;
        }
        static AddInId appId = new AddInId(new Guid("6F92CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            List<CurveLoop> cls = new List<CurveLoop>();
            

            Reference myRef = uidoc.Selection.PickObject(ObjectType.Element);
            Element e = doc.GetElement(myRef);


            Reference myRef3 = uidoc.Selection.PickObject(ObjectType.Element);
            Element e3 = doc.GetElement(myRef3);

            

            LocationCurve locationCurve1 = e.Location as LocationCurve;

            Autodesk.Revit.DB.Curve gridCurve = locationCurve1.Curve;

            LocationCurve locationCurve3 = e3.Location as LocationCurve;



            CurveLoop bottomLoop = new CurveLoop();
            CurveLoop topLoop = new CurveLoop();
            bottomLoop.Append(locationCurve1.Curve);
            topLoop.Append(locationCurve3.Curve);
            cls.Add(bottomLoop);
            cls.Add(topLoop);


            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Make loft");
                
                
                SolidOptions options = new SolidOptions(ElementId.InvalidElementId, ElementId.InvalidElementId);
                Solid mySolid = GeometryCreationUtilities.CreateLoftGeometry(cls, options);

            
                List<Solid> sol = new List<Solid>();

                sol.Add(mySolid);

                foreach (Face f in mySolid.Faces)
                {
                    //Build_Tessellate2(f, doc);
                    ElementId categoryId = new ElementId(BuiltInCategory.OST_GenericModel);
                    DirectShape ds = DirectShape.CreateElement(doc, categoryId);


                    ds.ApplicationId = System.Reflection.Assembly.GetExecutingAssembly().GetType().GUID.ToString();
                    ds.ApplicationDataId = Guid.NewGuid().ToString();


                    IList<GeometryObject> list = new List<GeometryObject>();
                    list.Add(mySolid);


                    // Create a direct shape.

                    DirectShape ds2 = DirectShape.CreateElement(doc,
                      new ElementId(BuiltInCategory.OST_GenericModel));

                    ds2.SetShape(list);
                    



                    //foreach (TessellatedShapeBuilderResult t1 in Build_Tessellate2(f, doc))
                    //{
                    //    ds.SetShape(t1.GetGeometricalObjects());

                    //    ds.Name = "Single_Surface";
                    //}
                }
                foreach (Solid s in sol)
                {

                }

                tx.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;

        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class rhino_points_to_revit_topo : IExternalCommand
    {
       
        static AddInId appId = new AddInId(new Guid("6F92CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 15;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }



            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            String info = "";
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);



            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }

            Point3d p1 = new Point3d(0, 0, 0);
            Rhino.Geometry.Point3d pt3d = new Point3d(10, 10, 0);

            File3dm m_modelfile = null;
            string m_name = null;
            string m_size = null;
            string m_created = null;
            string m_createdby = null;
            string m_edited = null;
            string m_editedby = null;
            string m_revision = null;
            string m_units = null;
            string m_notes = null;

            Object RhinoFile = filename2;
            if (RhinoFile is System.IO.FileInfo)
            {
                System.IO.FileInfo m_fileinfo = (System.IO.FileInfo)RhinoFile;
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }
            else if (RhinoFile is string)
            {
                System.IO.FileInfo m_fileinfo = new System.IO.FileInfo((string)RhinoFile);
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }

            File3dmLayerTable all_layers = m_modelfile.AllLayers;

            var objs_ = m_modelfile.Objects;
            List<Rhino.DocObjects.Layer> listadelayers = new List<Rhino.DocObjects.Layer>();
            List<Rhino.DocObjects.Layer> borrar = new List<Rhino.DocObjects.Layer>();
            List<List<Rhino.DocObjects.Layer>> multiplelistadelayers = new List<List<Rhino.DocObjects.Layer>>();
            string hola = "";

            MessageBox.Show("This tool will read 3D information only if the following Rhino Layers exist; Levels, Grids, Structure, Floor, Walls, Points", "!");

            foreach (var item in all_layers)
            {

                if (item.FullPath.Contains(hola))
                {
                    listadelayers.Add(item);
                }
            }
            List<Rhino.Geometry.Brep> rh_breps = new List<Rhino.Geometry.Brep>();
            List<Rhino.Geometry.Curve> curves_frames = new List<Rhino.Geometry.Curve>();
            List<Autodesk.Revit.DB.Curve> revit_crv = new List<Autodesk.Revit.DB.Curve>();
            List<string> m_names = new List<string>();
            List<int> m_layerindeces = new List<int>();
            List<System.Drawing.Color> m_colors = new List<System.Drawing.Color>();
            List<string> m_guids = new List<string>();

            foreach (var Layer in listadelayers)
            {
                if (Layer.Name == "Points")
                {
                    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);


                    List<Point3d> Points = new List<Point3d>();
                    List<XYZ> Revit_Points = new List<XYZ>();
                    List<XYZ> Revit_Points_notsameloc = new List<XYZ>();

                    foreach (var obj in m_objs)
                    {
                        GeometryBase brep = obj.Geometry;
                        Rhino.Geometry.Point pt = brep as Rhino.Geometry.Point;
                        Points.Add(pt.Location);
                    }
                    foreach (var item in Points)
                    {
                        double x = 0;
                        double y = 0;
                        double z = 0;

                        x = item.X / 304.8;
                        y = item.Y / 304.8;
                        z = item.Z / 304.8;

                        XYZ newpt = new XYZ(x, y, z);

                        Revit_Points.Add(newpt);
                    }

                    foreach (var item in Revit_Points)
                    {
                        Revit_Points_notsameloc.Add(item);
                    }

                    for (int i = 0; i < Revit_Points.ToArray().Length; i++)
                    {
                        for (int ij = 0; ij < Revit_Points.ToArray().Length; ij++)
                        {
                            if (Revit_Points.ToArray()[i].X == Revit_Points.ToArray()[ij].X &&
                                Revit_Points.ToArray()[i].Y == Revit_Points.ToArray()[ij].Y &&
                                Revit_Points.ToArray()[i].Z != Revit_Points.ToArray()[ij].Z)
                            {
                                Revit_Points_notsameloc.Remove(Revit_Points.ToArray()[i]);
                            }
                            if (i != ij && Revit_Points.ToArray()[i].X == Revit_Points.ToArray()[ij].X &&
                                Revit_Points.ToArray()[i].Y == Revit_Points.ToArray()[ij].Y &&
                                Revit_Points.ToArray()[i].Z == Revit_Points.ToArray()[ij].Z)
                            {
                                Revit_Points_notsameloc.Remove(Revit_Points.ToArray()[i]);
                            }
                        }
                    }
                    using (Transaction t = new Transaction(doc, "Make topograpgy from Rhino"))
                    {
                        t.Start();

                        TopographySurface ts = TopographySurface.Create(doc, Revit_Points_notsameloc);

                        t.Commit();
                    }
                }
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class rhino_lns_to_revit_lns : IExternalCommand
    {

        public static XYZ LocalToWorld(Autodesk.Revit.DB.Document document, XYZ coordinate, bool millimeters)
        {
            ProjectLocation oProjectLocation = document.ActiveProjectLocation;
            Autodesk.Revit.DB.Transform oTransform = oProjectLocation.GetTotalTransform().Inverse;

            XYZ oSurveyPoint = oTransform.OfPoint(coordinate);

            if (millimeters)
            {
                return new XYZ(oSurveyPoint.X, oSurveyPoint.Y, oSurveyPoint.Z) * 304.8;
            }

            return new XYZ(oSurveyPoint.X, oSurveyPoint.Y, oSurveyPoint.Z);
        }

        public static XYZ WorldToLocal(Autodesk.Revit.DB.Document document, XYZ coordinate, bool millimeters)
        {
            ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_ProjectBasePoint);
            FilteredElementCollector collector = new FilteredElementCollector(document);
            IList<Element> oProjectBasePoints = collector.WherePasses(filter).ToElements();

            Element oProjectBasePoint = null;

            foreach (Element bp in oProjectBasePoints)
            {
                oProjectBasePoint = bp;
                break;
            }

            double x = oProjectBasePoint.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble();
            double y = oProjectBasePoint.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble();
            double z = oProjectBasePoint.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble();
            double r = oProjectBasePoint.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM).AsDouble();

            XYZ result = new XYZ(
             coordinate.X * Math.Cos(r) - coordinate.Y * Math.Sin(r),
             coordinate.X * Math.Sin(r) + coordinate.Y * Math.Cos(r),
             0);

            if (millimeters)
            {
                return result * 304.8;
            }

            return result;
        }
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        public static bool IsZero(double a)
        {
            const double _eps = 1.0e-9;
            return _eps > Math.Abs(a);
        }

        public static bool IsEqual(double a, double b)
        {
            return IsZero(b - a);
        }

        public static int Compare(double a, double b)
        {
            return IsEqual(a, b) ? 0 : (a < b ? -1 : 1);
        }

        public static int Compare(XYZ p, XYZ q)
        {

            int diff = Compare(p.X, q.X);
            if (0 == diff)
            {
                diff = Compare(p.Y, q.Y);
                if (0 == diff)
                {
                    diff = Compare(p.Z, q.Z);
                }

            }

            return diff;
        }

        private static Wall CreateWall(FamilyInstance cube, Autodesk.Revit.DB.Curve curve, double height)
        {
            var doc = cube.Document;

            var wallTypeId = doc.GetDefaultElementTypeId(
              ElementTypeGroup.WallType);

            return Wall.Create(doc, curve.CreateReversed(),
              wallTypeId, cube.LevelId, height, 0, false,
              false);
        }
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 15;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }



            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            String info = "";
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            //List<FaceArray> face112 = new List<FaceArray>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);

            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }

            Point3d p1 = new Point3d(0, 0, 0);
            Rhino.Geometry.Point3d pt3d = new Point3d(10, 10, 0);

            File3dm m_modelfile = null;
            string m_name = null;
            string m_size = null;
            string m_created = null;
            string m_createdby = null;
            string m_edited = null;
            string m_editedby = null;
            string m_revision = null;
            string m_units = null;
            string m_notes = null;

            Object RhinoFile = filename2;
            if (RhinoFile is System.IO.FileInfo)
            {
                System.IO.FileInfo m_fileinfo = (System.IO.FileInfo)RhinoFile;
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }
            else if (RhinoFile is string)
            {
                System.IO.FileInfo m_fileinfo = new System.IO.FileInfo((string)RhinoFile);
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }

            File3dmLayerTable all_layers = m_modelfile.AllLayers;
            var objs_ = m_modelfile.Objects;
            List<Rhino.DocObjects.Layer> listadelayers = new List<Rhino.DocObjects.Layer>();
            List<Rhino.DocObjects.Layer> borrar = new List<Rhino.DocObjects.Layer>();
            List<List<Rhino.DocObjects.Layer>> multiplelistadelayers = new List<List<Rhino.DocObjects.Layer>>();
            string hola = "";

            MessageBox.Show("This tool will read 3D information only if the following Rhino Layers exist; Levels, Grids, Structure, Floor, Walls, Points", "!");

            foreach (var item in all_layers)
            {

                if (item.FullPath.Contains(hola))
                {
                    listadelayers.Add(item);
                }
            }

            List<Rhino.Geometry.Brep> rh_breps = new List<Rhino.Geometry.Brep>();
            List<XYZ> pts_ = new List<XYZ>();

            List<Autodesk.Revit.DB.Curve> revit_crv = new List<Autodesk.Revit.DB.Curve>();
            List<string> m_names = new List<string>();
            List<int> m_layerindeces = new List<int>();
            List<System.Drawing.Color> m_colors = new List<System.Drawing.Color>();
            List<string> m_guids = new List<string>();
            List<double> weith = new List<double>();



            foreach (var Layer in listadelayers)
            {


                if (Layer.Name == "Lines")
                {


                    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);

                    using (Transaction t = new Transaction(doc, "Lines from Rhino"))
                    {
                        t.Start();

                        foreach (File3dmObject obj in m_objs)
                        {
                            GeometryBase geo = obj.Geometry;
                            Type TYPE = geo.GetType();

                            if (TYPE.Name == "PolyCurve")
                            {
                                Rhino.Geometry.PolyCurve POLY = geo as Rhino.Geometry.PolyCurve;
                            }



                            if (TYPE.Name == "ArcCurve")
                            {
                                Rhino.Geometry.ArcCurve arc = geo as Rhino.Geometry.ArcCurve;

                                double x_end = arc.Arc.Plane.Normal.X / 304.8;
                                double y_end = arc.Arc.Plane.Normal.Y / 304.8;
                                double z_end = arc.Arc.Plane.Normal.Z / 304.8;

                                XYZ normal = new XYZ(x_end, y_end, z_end);

                                double x_ = arc.Arc.Plane.Origin.X / 304.8;
                                double y_ = arc.Arc.Plane.Origin.Y / 304.8;
                                double z_ = arc.Arc.Plane.Origin.Z / 304.8;
                                XYZ origin = new XYZ(x_, y_, z_);

                                double endAngle = 2 * Math.PI;        // this arc will be a circle

                                Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(normal, origin);

                                SketchPlane skplane = SketchPlane.Create(doc, pl);
                                doc.Create.NewModelCurve(Autodesk.Revit.DB.Arc.Create(pl, arc.Radius / 304.8, arc.Arc.StartAngle, arc.Arc.EndAngle), skplane);

                            }

                            if (TYPE.Name == "PolylineCurve")
                            {
                                Rhino.Geometry.PolylineCurve PolylineCurve = geo as Rhino.Geometry.PolylineCurve;
                                Rhino.Geometry.Polyline Polyline = PolylineCurve.ToPolyline();

                                for (int i = 0; i < Polyline.SegmentCount; i++)
                                {
                                    Point3d end = Polyline.SegmentAt(i).PointAt(0);
                                    Point3d start = Polyline.SegmentAt(i).PointAt(1);



                                    double x_end = end.X /*/ 304.8*/;
                                    double y_end = end.Y /*/ 304.8*/;
                                    double z_end = end.Z /*/ 304.8*/;

                                    double x_start = start.X /*/ 304.8*/;
                                    double y_start = start.Y /*/ 304.8*/ ;
                                    double z_start = start.Z /*/ 304.8*/;


                                    XYZ PointInSharedCoordinates_Meters_end = new XYZ(x_end, y_end, z_end);
                                    double x = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_end.X, DisplayUnitType.DUT_MILLIMETERS) /*/ 304.8*/;
                                    double y = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_end.Y, DisplayUnitType.DUT_MILLIMETERS)/* / 304.8*/;
                                    double z = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_end.Z, DisplayUnitType.DUT_MILLIMETERS) /*/ 304.8*/;
                                    XYZ pt_end = new XYZ(x, y, z);
                                    XYZ InternalPoint_end = doc.ActiveProjectLocation.GetTotalTransform().OfPoint(pt_end);



                                    XYZ PointInSharedCoordinates_Meters_start = new XYZ(x_start, y_start, z_start);
                                    double x_start2 = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_start.X, DisplayUnitType.DUT_MILLIMETERS) /*/ 304.8*/;
                                    double y_start2 = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_start.Y, DisplayUnitType.DUT_MILLIMETERS) /*/ 304.8*/;
                                    double z_start2 = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_start.Z, DisplayUnitType.DUT_MILLIMETERS) /*/ 304.8*/;
                                    XYZ pt_start = new XYZ(x_start2, y_start2, z_start2);
                                    XYZ InternalPoint_start = doc.ActiveProjectLocation.GetTotalTransform().OfPoint(pt_start);

                                    //XYZ pt_end = WorldToLocal(doc, new XYZ(x_end, y_end, z_end), true);
                                    //XYZ pt_start = WorldToLocal(doc, new XYZ(x_start, y_start, z_start), true);

                                    Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(InternalPoint_end, InternalPoint_start);
                                    Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(InternalPoint_end, InternalPoint_start);


                                    //Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);


                                    Makeline(doc, line.GetEndPoint(0), line.GetEndPoint(1));
                                    //revit_crv.Add(curve1);
                                }


                            }

                            if (TYPE.Name == "NurbsCurve")
                            {
                                Rhino.Geometry.NurbsCurve crv_ = geo as Rhino.Geometry.NurbsCurve;
                                Rhino.Geometry.Plane pl2;
                                crv_.TryGetPlane(out pl2);

                                double x_end = UnitUtils.ConvertToInternalUnits(pl2.Normal.X, DisplayUnitType.DUT_MILLIMETERS)   /*pl2.Normal.X / 304.8*/;
                                double y_end = UnitUtils.ConvertToInternalUnits(pl2.Normal.Y, DisplayUnitType.DUT_MILLIMETERS) /*pl2.Normal.Y / 304.8*/;
                                double z_end = UnitUtils.ConvertToInternalUnits(pl2.Normal.Z, DisplayUnitType.DUT_MILLIMETERS)  /*pl2.Normal.Z / 304.8*/;

                                XYZ normal = new XYZ(x_end, y_end, z_end);
                                XYZ InternalPoint_normal = doc.ActiveProjectLocation.GetTotalTransform().OfPoint(normal);

                                double x_2 = UnitUtils.ConvertToInternalUnits(pl2.Origin.X, DisplayUnitType.DUT_MILLIMETERS)  /*pl2.Origin.X / 304.8*/;
                                double y_2 = UnitUtils.ConvertToInternalUnits(pl2.Origin.Y, DisplayUnitType.DUT_MILLIMETERS)  /*pl2.Origin.Y / 304.8*/;
                                double z_2 = UnitUtils.ConvertToInternalUnits(pl2.Origin.Z, DisplayUnitType.DUT_MILLIMETERS)  /*pl2.Origin.Z / 304.8*/;
                                XYZ origin = new XYZ(x_2, y_2, z_2);
                                XYZ InternalPoint_otigin = doc.ActiveProjectLocation.GetTotalTransform().OfPoint(origin);

                                // this arc will be a circle

                                Autodesk.Revit.DB.Plane pl3 = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(InternalPoint_normal, InternalPoint_otigin);

                                var pts = crv_.Points;
                                foreach (var item in pts)
                                {
                                    double x_ = item.X;
                                    double y_ = item.Y /*/ 304.8*/;
                                    double z_ = item.Z /*/ 304.8*/;


                                    XYZ PointInSharedCoordinates_Meters_end = new XYZ(x_, y_, z_);
                                    double x = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_end.X, DisplayUnitType.DUT_MILLIMETERS) /*/ 304.8*/;
                                    double y = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_end.Y, DisplayUnitType.DUT_MILLIMETERS)/* / 304.8*/;
                                    double z = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_end.Z, DisplayUnitType.DUT_MILLIMETERS) /*/ 304.8*/;
                                    XYZ pt_end = new XYZ(x, y, z);
                                    XYZ InternalPoint_end = doc.ActiveProjectLocation.GetTotalTransform().OfPoint(pt_end);

                                    XYZ pt = new XYZ(x_, y_, z_);
                                    weith.Add(item.Weight);
                                    pts_.Add(/*pt*/ InternalPoint_end);
                                }

                                Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pts_.First(), pts_.Last());

                                XYZ norm = pts_.First().CrossProduct(curve1.Evaluate(5, false));
                                if (norm.GetLength() == 0)
                                {
                                    XYZ aSubB = pts_.First().Subtract(pts_.Last());
                                    XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                                    double crosslenght = aSubBcrossz.GetLength();
                                    if (crosslenght == 0)
                                    {
                                        norm = XYZ.BasisY;
                                    }
                                    else
                                    {
                                        norm = XYZ.BasisZ;
                                    }
                                }
                                //Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, pts_.First());
                                Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(pts_.First(), pts_.ElementAt(2), pts_.Last());
                                SketchPlane skplane = SketchPlane.Create(doc, pl);

                                Autodesk.Revit.DB.Curve nspline = Autodesk.Revit.DB.NurbSpline.CreateCurve(pts_, weith);
                                doc.Create.NewModelCurve(nspline, skplane);

                                pts_.Clear();
                                weith.Clear();
                            }

                            if (TYPE.Name == "Curve")
                            {
                                Rhino.Geometry.Curve crv_ = geo as Rhino.Geometry.Curve;


                                Point3d end = crv_.PointAtEnd;
                                Point3d start = crv_.PointAtStart;

                                double x_end = end.X / 304.8;
                                double y_end = end.Y / 304.8;
                                double z_end = end.Z / 304.8;

                                double x_start = start.X / 304.8;
                                double y_start = start.Y / 304.8;
                                double z_start = start.Z / 304.8;

                                XYZ pt_end = new XYZ(x_end, y_end, z_end);
                                XYZ pt_start = new XYZ(x_start, y_start, z_start);

                                Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);

                                Makeline(doc, line.GetEndPoint(0), line.GetEndPoint(1));
                                revit_crv.Add(curve1);
                            }
                        }
                        t.Commit();
                    }
                }
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class rhino_pt_to_revit_clash : IExternalCommand
    {

        public static XYZ LocalToWorld(Autodesk.Revit.DB.Document document, XYZ coordinate, bool millimeters)
        {
            ProjectLocation oProjectLocation = document.ActiveProjectLocation;
            Autodesk.Revit.DB.Transform oTransform = oProjectLocation.GetTotalTransform().Inverse;

            XYZ oSurveyPoint = oTransform.OfPoint(coordinate);

            if (millimeters)
            {
                return new XYZ(oSurveyPoint.X, oSurveyPoint.Y, oSurveyPoint.Z) * 304.8;
            }

            return new XYZ(oSurveyPoint.X, oSurveyPoint.Y, oSurveyPoint.Z);
        }

        public static XYZ WorldToLocal(Autodesk.Revit.DB.Document document, XYZ coordinate, bool millimeters)
        {
            ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_ProjectBasePoint);
            FilteredElementCollector collector = new FilteredElementCollector(document);
            IList<Element> oProjectBasePoints = collector.WherePasses(filter).ToElements();

            Element oProjectBasePoint = null;

            foreach (Element bp in oProjectBasePoints)
            {
                oProjectBasePoint = bp;
                break;
            }

            double x = oProjectBasePoint.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble();
            double y = oProjectBasePoint.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble();
            double z = oProjectBasePoint.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble();
            double r = oProjectBasePoint.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM).AsDouble();

            XYZ result = new XYZ(
             coordinate.X * Math.Cos(r) - coordinate.Y * Math.Sin(r),
             coordinate.X * Math.Sin(r) + coordinate.Y * Math.Cos(r),
             0);

            if (millimeters)
            {
                return result * 304.8;
            }

            return result;
        }
       
        
        static AddInId appId = new AddInId(new Guid("C9EBBF36-C2CC-4018-B511-6F17EF5AD6B6"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            //try
            //{
            //    string filename = @"T:\Transfer\lopez\Book1.xlsx";
            //    using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
            //    {
            //        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

            //        int column = 15;
            //        int number = Convert.ToInt32(sheet.Cells[2, column].Value);
            //        sheet.Cells[2, column].Value = (number + 1); ;
            //        package.Save();
            //    }
            //}
            
            //catch (Exception)
            //{
            //    MessageBox.Show("Excel file not found", "");
            //}

            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            String info = "";
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            //List<FaceArray> face112 = new List<FaceArray>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);

            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }

            Point3d p1 = new Point3d(0, 0, 0);
            Rhino.Geometry.Point3d pt3d = new Point3d(10, 10, 0);

            File3dm m_modelfile = null;
            string m_name = null;
            string m_size = null;
            string m_created = null;
            string m_createdby = null;
            string m_edited = null;
            string m_editedby = null;
            string m_revision = null;
            string m_units = null;
            string m_notes = null;

            Object RhinoFile = filename2;
            if (RhinoFile is System.IO.FileInfo)
            {
                System.IO.FileInfo m_fileinfo = (System.IO.FileInfo)RhinoFile;
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }
            else if (RhinoFile is string)
            {
                System.IO.FileInfo m_fileinfo = new System.IO.FileInfo((string)RhinoFile);
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }

            File3dmLayerTable all_layers = m_modelfile.AllLayers;
            var objs_ = m_modelfile.Objects;
            List<Rhino.DocObjects.Layer> listadelayers = new List<Rhino.DocObjects.Layer>();
            List<Rhino.DocObjects.Layer> borrar = new List<Rhino.DocObjects.Layer>();
            List<List<Rhino.DocObjects.Layer>> multiplelistadelayers = new List<List<Rhino.DocObjects.Layer>>();
            string hola = "";

            MessageBox.Show("This tool will read 3D information only if the following Rhino Layers exist; Levels, Grids, Structure, Floor, Walls, Points", "!");

            foreach (var item in all_layers)
            {

                if (item.FullPath.Contains(hola))
                {
                    listadelayers.Add(item);
                }
            }

            List<Rhino.Geometry.Brep> rh_breps = new List<Rhino.Geometry.Brep>();
            List<XYZ> pts_ = new List<XYZ>();

            List<Autodesk.Revit.DB.Curve> revit_crv = new List<Autodesk.Revit.DB.Curve>();
            List<string> m_names = new List<string>();
            List<int> m_layerindeces = new List<int>();
            List<System.Drawing.Color> m_colors = new List<System.Drawing.Color>();
            List<string> m_guids = new List<string>();
            List<double> weith = new List<double>();



            foreach (var Layer in listadelayers)
            
                if (Layer.Name == "Points")
                {
                    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);

                    foreach (File3dmObject obj in m_objs)
                    {
                        GeometryBase geo = obj.Geometry;
                        Type TYPE = geo.GetType();

                        if (TYPE.Name == "Point")
                        {
                            Rhino.Geometry.Point PolylineCurve = geo as Rhino.Geometry.Point;

                            double x_end = PolylineCurve.Location.X;
                            double y_end = PolylineCurve.Location.Y;
                            double z_end = PolylineCurve.Location.Z;

                            XYZ PointInSharedCoordinates_Meters_end = new XYZ(x_end, y_end, z_end);
                            double x = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_end.X, DisplayUnitType.DUT_MILLIMETERS);
                            double y = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_end.Y, DisplayUnitType.DUT_MILLIMETERS);
                            double z = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_end.Z, DisplayUnitType.DUT_MILLIMETERS);
                            XYZ pt_end = new XYZ(x, y, z);
                            XYZ InternalPoint_end = doc.ActiveProjectLocation.GetTotalTransform().OfPoint(pt_end);

                            FilteredElementCollector col = new FilteredElementCollector(doc);
                            col.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_GenericModel);

                            FamilySymbol symbol = col.FirstElement() as FamilySymbol;

                            foreach (var item in col)
                            {
                                if (item.Name == "ES - Clash point - R19")
                                {
                                    symbol = item as FamilySymbol;
                                }
                            }

                                                     
                            using (Transaction t = new Transaction(doc, "Penetration point"))
                            {
                                t.Start();
                                
                                FamilyInstance fi2 = doc.Create.NewFamilyInstance(InternalPoint_end, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                
                                t.Commit();
                                
                            }
                        }
                    }
                }
            
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class excel_to_revit_clash : IExternalCommand
    {

        public static XYZ LocalToWorld(Autodesk.Revit.DB.Document document, XYZ coordinate, bool millimeters)
        {
            ProjectLocation oProjectLocation = document.ActiveProjectLocation;
            Autodesk.Revit.DB.Transform oTransform = oProjectLocation.GetTotalTransform().Inverse;

            XYZ oSurveyPoint = oTransform.OfPoint(coordinate);

            if (millimeters)
            {
                return new XYZ(oSurveyPoint.X, oSurveyPoint.Y, oSurveyPoint.Z) * 304.8;
            }

            return new XYZ(oSurveyPoint.X, oSurveyPoint.Y, oSurveyPoint.Z);
        }

        public static XYZ WorldToLocal(Autodesk.Revit.DB.Document document, XYZ coordinate, bool millimeters)
        {
            ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_ProjectBasePoint);
            FilteredElementCollector collector = new FilteredElementCollector(document);
            IList<Element> oProjectBasePoints = collector.WherePasses(filter).ToElements();

            Element oProjectBasePoint = null;

            foreach (Element bp in oProjectBasePoints)
            {
                oProjectBasePoint = bp;
                break;
            }

            double x = oProjectBasePoint.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble();
            double y = oProjectBasePoint.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble();
            double z = oProjectBasePoint.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble();
            double r = oProjectBasePoint.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM).AsDouble();

            XYZ result = new XYZ(
             coordinate.X * Math.Cos(r) - coordinate.Y * Math.Sin(r),
             coordinate.X * Math.Sin(r) + coordinate.Y * Math.Cos(r),
             0);

            if (millimeters)
            {
                return result * 304.8;
            }

            return result;
        }


        static AddInId appId = new AddInId(new Guid("C7EBBF36-C2CC-4018-B521-6F13EF5AD6B6"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            //string filename2 = Path.Combine(Path.GetTempPath(), "Book1.xlsx");
            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            //openDialog.Filter = "Excel Files (*.xlsx) |*.xslx)";

            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }
            else
            {
                return Result.Cancelled;
            }


            double x_ = 0;
            double y_ = 0;
            double z_ = 0;


            using (ExcelPackage package = new ExcelPackage(new FileInfo(filename2)))
            {

                ExcelWorksheet sheet = package.Workbook.Worksheets[1];


                for (int row = 1; row < 9999; row++)
                {
                    using (Transaction t = new Transaction(doc, "Penetration point"))
                    {
                        t.Start();

                        string row4 = null;
                        string row5 = null;
                        string row6 = null;


                        for (int col = 1; col < 9999; col++)
                        {
                            var thisValue = sheet.Cells[row, col].Value;

                            if (thisValue == null && col == 1)
                            {
                                return Autodesk.Revit.UI.Result.Succeeded;
                            }

                            if (thisValue != null)
                            {
                                if (col == 1)
                                {
                                    x_ = (double)thisValue;
                                }
                                if (col == 2)
                                {
                                    y_ = (double)thisValue;
                                }
                                if (col == 3)
                                {
                                    z_ = (double)thisValue;
                                }
                                if (col == 4)
                                {
                                    row4 = thisValue as string;
                                }
                                if (col == 5)
                                {
                                    row5 = thisValue as string;
                                }
                                if (col == 6)
                                {
                                    row6 = thisValue as string;
                                }

                            }
                            else
                            {



                                XYZ pt_ = new XYZ(x_, y_, z_);

                                double x_end = pt_.X;
                                double y_end = pt_.Y;
                                double z_end = pt_.Z;

                                XYZ PointInSharedCoordinates_Meters_end = new XYZ(x_end, y_end, z_end);
                                double x = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_end.X, DisplayUnitType.DUT_MILLIMETERS);
                                double y = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_end.Y, DisplayUnitType.DUT_MILLIMETERS);
                                double z = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_end.Z, DisplayUnitType.DUT_MILLIMETERS);
                                XYZ pt_end = new XYZ(x, y, z);
                                XYZ InternalPoint_end = doc.ActiveProjectLocation.GetTotalTransform().OfPoint(pt_end);

                                FilteredElementCollector col_ = new FilteredElementCollector(doc);
                                col_.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_GenericModel);

                                FamilySymbol symbol = col_.FirstElement() as FamilySymbol;

                                foreach (var item in col_)
                                {
                                    if (item.Name == "ES - Clash point - R19")
                                    {
                                        symbol = item as FamilySymbol;
                                    }
                                }

                                FamilyInstance fi2 = doc.Create.NewFamilyInstance(InternalPoint_end, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                try
                                {
                                    string upper1 = row4;
                                    Parameter param = fi2.LookupParameter("Mark");
                                    param.Set(upper1);

                                }
                                catch (Exception)
                                {

                                    
                                }

                                try
                                {
                                    string upper2 = row6;
                                    Parameter param2 = fi2.LookupParameter("Comments");
                                    param2.Set(upper2);
                                }
                                catch (Exception)
                                {

                                    
                                }

                                try
                                {
                                    string upper3 = row4;
                                    Parameter param3 = fi2.LookupParameter("Name");
                                    param3.Set(upper3);
                                }
                                catch (Exception)
                                {

                                    
                                }

                                try
                                {
                                    double upper4 = x_;
                                    Parameter param4 = fi2.LookupParameter("x-value");
                                    param4.Set(upper4);
                                }
                                catch (Exception)
                                {

                                    
                                }

                                try
                                {
                                    double upper5 = x_;
                                    Parameter param5 = fi2.LookupParameter("y-value");
                                    param5.Set(upper5);
                                }
                                catch (Exception)
                                {

                                   
                                }

                                try
                                {
                                    double upper6 = z_;
                                    Parameter param6 = fi2.LookupParameter("z-value");
                                    param6.Set(upper6);
                                }
                                catch (Exception)
                                {

                                }
                                

                                
                                break;
                                row++;
                            }

                        }

                        t.Commit();
                    }

                    
                }
            }

           

            //foreach (var Layer in listadelayers)

            //    if (Layer.Name == "Points")
            //    {
            //        File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);

            //        foreach (File3dmObject obj in m_objs)
            //        {
            //            GeometryBase geo = obj.Geometry;
            //            Type TYPE = geo.GetType();

            //            if (TYPE.Name == "Point")
            //            {
            //                Rhino.Geometry.Point PolylineCurve = geo as Rhino.Geometry.Point;

            //                double x_end = PolylineCurve.Location.X;
            //                double y_end = PolylineCurve.Location.Y;
            //                double z_end = PolylineCurve.Location.Z;

            //                XYZ PointInSharedCoordinates_Meters_end = new XYZ(x_end, y_end, z_end);
            //                double x = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_end.X, DisplayUnitType.DUT_MILLIMETERS);
            //                double y = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_end.Y, DisplayUnitType.DUT_MILLIMETERS);
            //                double z = UnitUtils.ConvertToInternalUnits(PointInSharedCoordinates_Meters_end.Z, DisplayUnitType.DUT_MILLIMETERS);
            //                XYZ pt_end = new XYZ(x, y, z);
            //                XYZ InternalPoint_end = doc.ActiveProjectLocation.GetTotalTransform().OfPoint(pt_end);

            //                FilteredElementCollector col = new FilteredElementCollector(doc);
            //                col.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_GenericModel);

            //                FamilySymbol symbol = col.FirstElement() as FamilySymbol;

            //                foreach (var item in col)
            //                {
            //                    if (item.Name == "ES - Clash point - R19")
            //                    {
            //                        symbol = item as FamilySymbol;
            //                    }
            //                }


            //                using (Transaction t = new Transaction(doc, "Penetration point"))
            //                {
            //                    t.Start();

            //                    FamilyInstance fi2 = doc.Create.NewFamilyInstance(InternalPoint_end, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

            //                    t.Commit();

            //                }
            //            }
            //        }
            //    }

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Create_flex_ducts_from_line : IExternalCommand
    {


        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;


            List<Autodesk.Revit.DB.XYZ> crvs = new List<Autodesk.Revit.DB.XYZ>();
            Autodesk.Revit.DB.NurbSpline line = null;

            ICollection<Reference> my_lines = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select lines");
            foreach (var item_myRefWall in my_lines)
            {
                Element ele = doc.GetElement(item_myRefWall);
                GeometryObject geoobj = ele.GetGeometryObjectFromReference(item_myRefWall);
                Face face = geoobj as Face;
                LocationCurve locationCurve2 = ele.Location as LocationCurve;
                


                if (locationCurve2.Curve.GetType().Name == "NurbSpline")
                {
                    line = locationCurve2.Curve as Autodesk.Revit.DB.NurbSpline;
                    
                }
            }

            var ductTypes =new FilteredElementCollector(doc).OfClass(typeof(FlexDuctType)).OfType<FlexDuctType>().First();

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("flex"
                  + "flex");
                FlexDuct flexDuct = doc.Create.NewFlexDuct(line.CtrlPoints, ductTypes);

                tx.Commit();
            }

            

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class rhino_PTS_to_revitPTS : IExternalCommand
    {


        public static XYZ TransformPoint(XYZ point, Autodesk.Revit.DB.Transform transform)
        {
            double x = point.X;
            double y = point.Y;
            double z = point.Z;

            //transform basis of the old coordinate system in the new coordinate // system
            XYZ b0 = transform.get_Basis(0);
            XYZ b1 = transform.get_Basis(1);
            XYZ b2 = transform.get_Basis(2);
            XYZ origin = transform.Origin;

            //transform the origin of the old coordinate system in the new 
            //coordinate system
            double xTemp = x * b0.X + y * b1.X + z * b2.X + origin.X;
            double yTemp = x * b0.Y + y * b1.Y + z * b2.Y + origin.Y;
            double zTemp = x * b0.Z + y * b1.Z + z * b2.Z + origin.Z;

            return new XYZ(xTemp, yTemp, zTemp);
        }

        private static Wall CreateWall(FamilyInstance cube, Autodesk.Revit.DB.Curve curve, double height)
        {
            var doc = cube.Document;

            var wallTypeId = doc.GetDefaultElementTypeId(
              ElementTypeGroup.WallType);

            return Wall.Create(doc, curve.CreateReversed(),
              wallTypeId, cube.LevelId, height, 0, false,
              false);
        }
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 15;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }

            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            String info = "";
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            //List<FaceArray> face112 = new List<FaceArray>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);

            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }

            Point3d p1 = new Point3d(0, 0, 0);
            Rhino.Geometry.Point3d pt3d = new Point3d(10, 10, 0);

            File3dm m_modelfile = null;
            string m_name = null;
            string m_size = null;
            string m_created = null;
            string m_createdby = null;
            string m_edited = null;
            string m_editedby = null;
            string m_revision = null;
            string m_units = null;
            string m_notes = null;

            Object RhinoFile = filename2;
            if (RhinoFile is System.IO.FileInfo)
            {
                System.IO.FileInfo m_fileinfo = (System.IO.FileInfo)RhinoFile;
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }
            else if (RhinoFile is string)
            {
                System.IO.FileInfo m_fileinfo = new System.IO.FileInfo((string)RhinoFile);
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }

            File3dmLayerTable all_layers = m_modelfile.AllLayers;
            var objs_ = m_modelfile.Objects;

            List<Rhino.DocObjects.Layer> listadelayers = new List<Rhino.DocObjects.Layer>();
            List<Rhino.DocObjects.Layer> borrar = new List<Rhino.DocObjects.Layer>();
            List<List<Rhino.DocObjects.Layer>> multiplelistadelayers = new List<List<Rhino.DocObjects.Layer>>();
            string hola = "";
            MessageBox.Show("This tool will read 3D information only if the following Rhino Layers exist; Levels, Grids, Structure, Floor, Walls, Points", "!");

            foreach (var item in all_layers)
            {
                if (item.FullPath.Contains(hola))
                {
                    listadelayers.Add(item);
                }
            }

            List<Rhino.Geometry.Brep> rh_breps = new List<Rhino.Geometry.Brep>();
            List<Rhino.Geometry.Curve> curves_frames = new List<Rhino.Geometry.Curve>();
            List<Autodesk.Revit.DB.Curve> revit_crv = new List<Autodesk.Revit.DB.Curve>();
            List<string> m_names = new List<string>();
            List<int> m_layerindeces = new List<int>();
            List<System.Drawing.Color> m_colors = new List<System.Drawing.Color>();
            List<string> m_guids = new List<string>();

            List<Element> FamilySymbols_ = new List<Element>();
            
            FilteredElementCollector collector= new FilteredElementCollector(doc);

            collector.OfCategory(BuiltInCategory.OST_GenericModel);collector.OfClass(typeof(FamilySymbol));

            foreach (var item in collector)
            {
                
                FamilySymbols_.Add(item);
               
            }


            //foreach (var item in FamilySymbols_)
            //{

            //    FamilyInstance fInstance = item as FamilyInstance;

            //    FamilySymbol FType = fInstance.Symbol;

            //    Family Fam = FType.Family;
            //    if (Fam.Name == "ES - Clash point - R19")
            //    {
            //        symbol = FType;
            //    }
            //}
            

            foreach (var Layer in listadelayers)
            {
                FamilySymbol beamSymbol = null;

                if (Layer.Name == "Points")
                {

                    using (Transaction t = new Transaction(doc, "Make structural frame from Rhino"))
                    {
                        t.Start();
                        File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);
                        foreach (File3dmObject obj in m_objs)
                        {
                            GeometryBase geo = obj.Geometry;
                            if (geo is Rhino.Geometry.Point)
                            {
                                Rhino.Geometry.Point point = (Rhino.Geometry.Point)obj.Geometry;
                                Point3d pt = point.Location;

                                
                              

                                double x_end = UnitUtils.Convert(pt.X, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS); /*/ 304.8*/;
                                double y_end = UnitUtils.Convert(pt.Y, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS); /*/ 304.8*/;

                                double z_end = UnitUtils.Convert(pt.Z, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS); /*/ 304.8*/;


                                XYZ Rpt = new XYZ(x_end, y_end, z_end);

                                FilteredElementCollector col = new FilteredElementCollector(doc);
                                col.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_GenericModel);

                                FamilySymbol symbol = col.FirstElement() as FamilySymbol;

                                foreach (var item in col)
                                {
                                    if (item.Name == "ES - Clash point - R19")
                                    {
                                        symbol = item as FamilySymbol;
                                    }
                                }

                                

                                FamilyInstance fi = doc.Create.NewFamilyInstance(Rpt, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                //LocationPoint location = fi.Location as LocationPoint;
                                //XYZ origin = location.Point;

                                XYZ ti_1 = TransformPoint(Rpt, fi.GetTransform());

                                Autodesk.Revit.DB.Transform trf = fi.GetTransform();
                                trf = trf.Inverse;

                                XYZ tPoint = trf.OfPoint(ti_1);

                                FamilyInstance fi2 = doc.Create.NewFamilyInstance(tPoint, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);


                            }
                        }
                        t.Commit();
                    }
                }

                //if (Layer.Name == "Structure")
                //{
                //    Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().OrderBy(q => q.Elevation).First();
                //    List<FamilyInstance> famsimbol = new List<FamilyInstance>();

                //    foreach (var item in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>())
                //    {
                //        if (item.IsActive == true && item.Family.FamilyCategory.Name == "Structural Framing")
                //        {
                //            beamSymbol = item;
                //        }
                //    }

                //    using (Transaction t = new Transaction(doc, "Make structural frame from Rhino"))
                //    {
                //        t.Start();
                //        File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);
                //        foreach (File3dmObject obj in m_objs)
                //        {
                //            GeometryBase geo = obj.Geometry;
                //            if (geo is Rhino.Geometry.Curve)
                //            {
                //                Rhino.Geometry.Curve crv_ = geo as Rhino.Geometry.Curve;
                //                curves_frames.Add(crv_);

                //                Point3d end = crv_.PointAtEnd;
                //                Point3d start = crv_.PointAtStart;

                //                double x_end = end.X / 304.8;
                //                double y_end = end.Y / 304.8;
                //                double z_end = end.Z / 304.8;

                //                double x_start = start.X / 304.8;
                //                double y_start = start.Y / 304.8;
                //                double z_start = start.Z / 304.8;

                //                XYZ pt_end = new XYZ(x_end, y_end, z_end);
                //                XYZ pt_start = new XYZ(x_start, y_start, z_start);

                //                Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                //                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);

                //                foreach (FamilyInstance fi_ in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_StructuralFraming).Cast<FamilyInstance>())
                //                {
                //                    famsimbol.Add(fi_);
                //                }

                //                FamilyInstance fi = doc.Create.NewFamilyInstance(curve1, beamSymbol, level, Autodesk.Revit.DB.Structure.StructuralType.Beam);

                //                //try
                //                //{
                //                //    using (Transaction t = new Transaction(doc, "Make structural frame from Rhino"))
                //                //    {
                //                //        t.Start();

                //                //        Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                //                //        Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);


                //                //        foreach (FamilyInstance fi_ in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_StructuralFraming).Cast<FamilyInstance>())
                //                //        {
                //                //            famsimbol.Add(fi_);
                //                //        }

                //                //        FamilyInstance fi = doc.Create.NewFamilyInstance(curve1, beamSymbol, level, Autodesk.Revit.DB.Structure.StructuralType.Beam);

                //                //        t.Commit();
                //                //    }

                //                //}
                //                //catch (Exception)
                //                //{
                //                //    //throw;
                //                //}
                //            }
                //        }

                //        t.Commit();
                //    }
                //}
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class duct_elbo : IExternalCommand
    {

        public static bool IsZero(double a)
        {
            const double _eps = 1.0e-9;
            return _eps > Math.Abs(a);
        }

        public static bool IsEqual(double a, double b)
        {
            return IsZero(b - a);
        }

        public static int Compare(double a, double b)
        {
            return IsEqual(a, b) ? 0 : (a < b ? -1 : 1);
        }

        public static int Compare(XYZ p, XYZ q)
        {

            int diff = Compare(p.X, q.X);
            if (0 == diff)
            {
                diff = Compare(p.Y, q.Y);
                if (0 == diff)
                {
                    diff = Compare(p.Z, q.Z);
                }

            }

            return diff;
        }

        static FilteredElementCollector GetConnectorElements(Autodesk.Revit.DB.Document doc,bool include_wires)
        {
            // what categories of family instances
            // are we interested in?

             BuiltInCategory[] bics = new BuiltInCategory[] {
                //BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting,
                //BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting,
                //BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_LightingDevices,
                BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_MechanicalEquipment,
             //BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_SpecialityEquipment,
                BuiltInCategory.OST_Sprinklers,
                //BuiltInCategory.OST_Wire,
            };

            IList<ElementFilter> a = new List<ElementFilter>(bics.Count());

            foreach (BuiltInCategory bic in bics)
            {
                a.Add(new ElementCategoryFilter(bic));
            }

            LogicalOrFilter categoryFilter = new LogicalOrFilter(a);

            LogicalAndFilter familyInstanceFilter = new LogicalAndFilter(categoryFilter, new ElementClassFilter(
                  typeof(FamilyInstance)));

            IList<ElementFilter> b = new List<ElementFilter>(6);

            b.Add(new ElementClassFilter(typeof(CableTray)));
            b.Add(new ElementClassFilter(typeof(Conduit)));
            b.Add(new ElementClassFilter(typeof(Duct)));
            b.Add(new ElementClassFilter(typeof(Pipe)));

            if (include_wires)
            {
                b.Add(new ElementClassFilter(typeof(Wire)));
            }

            b.Add(familyInstanceFilter);

            LogicalOrFilter classFilter
              = new LogicalOrFilter(b);

            FilteredElementCollector collector
              = new FilteredElementCollector(doc);

            collector.WherePasses(classFilter);

            return collector;
        }

        static ConnectorSet GetConnectors(Element e)
        {
            ConnectorSet connectors = null;

            if (e is FamilyInstance)
            {
                MEPModel m = ((FamilyInstance)e).MEPModel;
               
                if (null != m && null != m.ConnectorManager)
                {

                    connectors = m.ConnectorManager.Connectors;
                }
            }
            else if (e is Wire)
            {
                connectors = ((Wire)e) .ConnectorManager.Connectors;
            }
            else
            {
                Debug.Assert(
                  e.GetType().IsSubclassOf(typeof(MEPCurve)),
                  "expected all candidate connector provider "
                  + "elements to be either family instances or "
                  + "derived from MEPCurve");

                if (e is MEPCurve)
                {
                    connectors = ((MEPCurve)e).ConnectorManager.Connectors;
                }
            }

           
            return connectors;
        }

        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {


            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;


            List<Autodesk.Revit.DB.XYZ> crvs = new List<Autodesk.Revit.DB.XYZ>();
            Autodesk.Revit.DB.NurbSpline line = null;

            //ICollection<Reference> my_lines = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select lines");
            //foreach (var item_myRefWall in my_lines)
            //{
            //    Element ele = doc.GetElement(item_myRefWall);
            //    GeometryObject geoobj = ele.GetGeometryObjectFromReference(item_myRefWall);
            //    Face face = geoobj as Face;
            //    LocationCurve locationCurve2 = ele.Location as LocationCurve;
            
            //    if (locationCurve2.Curve.GetType().Name == "NurbSpline")
            //    {
            //        line = locationCurve2.Curve as Autodesk.Revit.DB.NurbSpline;

            //    }
            //}

            Reference myRef2 = uidoc.Selection.PickObject(ObjectType.Element);
            Element e2 = doc.GetElement(myRef2.ElementId);
            ConnectorSet connector = GetConnectors(e2);

            List<XYZ> pts = new List<XYZ>();

            foreach (Connector item in connector)
            {
                pts.Add(item.Origin);
            }
            
            Reference myRef1 = uidoc.Selection.PickObject(ObjectType.Element);
            Element e1 = doc.GetElement(myRef1.ElementId);
            ConnectorSet connector2 = GetConnectors(e1);

            foreach (Connector item in connector2)
            {
                pts.Add(item.Origin);
            }

            List<double> distances = new List<double>();
            
            foreach (var item in pts)
            {

                foreach (var item2 in pts)
                {

                    int hola = Compare(item, item2);

                    if (Compare(item, item2) != 0) 
                    {
                        distances.Add(item.DistanceTo(item2));
                    }
                }
            }

            distances.Sort();

            foreach (Connector item in connector)
            {
                foreach (Connector item2 in connector2)
                {
                    if (item.Origin.DistanceTo(item2.Origin) == distances.First())
                    {
                        //var ductTypes = new FilteredElementCollector(doc)
                        //.OfClass(typeof(MechanicalFitting)).OfType<MechanicalFitting>().First();
                        using (Transaction tx = new Transaction(doc))
                        {
                            tx.Start("flex"
                              + "flex");
                            FamilyInstance flexDuct = doc.Create.NewTransitionFitting(item, item2);
                            tx.Commit();
                        }

                        continue;
                    }
                }
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Pipe_slab_intersetion : IExternalCommand
    {
        private XYZ intersect(XYZ point, XYZ direction, Autodesk.Revit.DB.Curve curve)
        {
            Autodesk.Revit.DB.Line unbound = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(point.X, point.Y, curve.GetEndPoint(0).Z), direction);
            IntersectionResultArray ira = null;
            unbound.Intersect(curve, out ira);
            if (ira == null)
            {
                TaskDialog td = new TaskDialog("Error");
                td.MainInstruction = "no intersection";
                td.MainContent = point.ToString() + Environment.NewLine + direction.ToString();
                td.Show();

                return null;
            }
            IntersectionResult ir = ira.get_Item(0);
            return ir.XYZPoint;
        }
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        public static XYZ LinePlaneIntersection(Autodesk.Revit.DB.Line line, Autodesk.Revit.DB.Plane plane, out double lineParameter)
        {
            XYZ planePoint = plane.Origin;
            XYZ planeNormal = plane.Normal;
            XYZ linePoint = line.GetEndPoint(0);

            XYZ lineDirection = (line.GetEndPoint(1)
              - linePoint).Normalize();

            // Is the line parallel to the plane, i.e.,
            // perpendicular to the plane normal?

            if (IsZero(planeNormal.DotProduct(lineDirection)))
            {
                lineParameter = double.NaN;
                return null;
            }

            lineParameter = (planeNormal.DotProduct(planePoint)
              - planeNormal.DotProduct(linePoint))
                / planeNormal.DotProduct(lineDirection);

            return linePoint + lineParameter * lineDirection;
        }
        public static bool IsZero(double a)
        {
            const double _eps = 1.0e-9;
            return _eps > Math.Abs(a);
        }

        static AddInId appId = new AddInId(new Guid("7AF2027D-0C24-49DB-A975-99E2C8CC1B1C"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            MessageBox.Show("Select a floor face", "clashing surface");

            Reference myRef1 = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            XYZ p1 = myRef1.GlobalPoint;
            Reference myRef2 = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            XYZ p2 = myRef2.GlobalPoint;
            Reference myRef3 = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            XYZ p3 = myRef3.GlobalPoint;

            //Reference my_faces_ = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Face, "Select ceilings to be reproduced in rhino geometry");


            List<Face> faces_picked = new List<Face>();
            List<string> name_of_roof = new List<string>();

            //Element e = doc.GetElement(my_faces_);
            //GeometryObject geoobj2 = e.GetGeometryObjectFromReference(my_faces_);
            //Face face = geoobj2 as Face;
            //PlanarFace planarFace = face as PlanarFace;


            //XYZ p1 = null;
            //XYZ p2 = null;
            //XYZ p3 = null;

            Autodesk.Revit.DB.Plane plane = null;

            plane = Autodesk.Revit.DB.Plane.CreateByThreePoints(p1, p2, p3/*edges.get_Item(0).AsCurve().GetEndPoint(0), edges.get_Item(1).AsCurve().GetEndPoint(0), edges.get_Item(2).AsCurve().GetEndPoint(0)*/);

            //if (face != null)
            //{
            //    EdgeArrayArray edgeArrays = face.EdgeLoops;



            //    foreach (EdgeArray edges in edgeArrays)
            //    {
            //        //for (int i = 0; i < edges.Size; i++)
            //        //{
            //        //    edges.get_Item()
            //        //}


            //        plane = Autodesk.Revit.DB.Plane.CreateByThreePoints(p1,p2,p3/*edges.get_Item(0).AsCurve().GetEndPoint(0), edges.get_Item(1).AsCurve().GetEndPoint(0), edges.get_Item(2).AsCurve().GetEndPoint(0)*/);

            //        //foreach (Edge edge in edges)
            //        //{
            //        //    edge_crvs.Add(edge.AsCurve());

            //        //    Autodesk.Revit.DB.Line line = locationCurve2.Curve as Autodesk.Revit.DB.Line;
                      

            //        //    Autodesk.Revit.DB.Transform curveTransform = edge.AsCurve().ComputeDerivatives(0.5, true);

            //        //    try
            //        //    {
                          
            //        //    }
            //        //    catch (Exception)
            //        //    {
            //        //        return Autodesk.Revit.UI.Result.Cancelled;
            //        //    }

            //        //}


            //    }

            //}

           



            //Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByThreePoints(p1, p2, p3);


            List<Autodesk.Revit.DB.Curve> crvs = new List<Autodesk.Revit.DB.Curve>();

            MessageBox.Show("select one or more pipes", "Clashing element");

            ICollection<Reference> my_lines = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select pipes");
            foreach (var item_myRefWall in my_lines)
            {
                Element ele = doc.GetElement(item_myRefWall);
                

                GeometryObject geoobj = ele.GetGeometryObjectFromReference(item_myRefWall);
               
                
                LocationCurve locationCurve2 = ele.Location as LocationCurve;

                crvs.Add(locationCurve2.Curve);


            }


            FilteredElementCollector col = new FilteredElementCollector(doc);
            col.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_GenericModel);

            FamilySymbol symbol = col.FirstElement() as FamilySymbol;

            foreach (var item in col)
            {
                if (item.Name == "ES - Clash point - R19")
                {
                    symbol = item as FamilySymbol;
                }
            }

            //IEnumerable viewFamilyTypes = from elem in new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType))
            //                              let type = elem as ViewFamilyType
            //                              where type.ViewFamily == ViewFamily.ThreeDimensional
            //                              select type;

            using (Transaction t = new Transaction(doc, "Penetration point"))
            {
                t.Start();

                int num = 0;

                foreach (var crv in crvs)
                {
                    num++;
                    XYZ intpt = LinePlaneIntersection(crv as Autodesk.Revit.DB.Line, plane, out double lineParameter);

                    FamilyInstance fi2 = doc.Create.NewFamilyInstance(intpt, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                    
                    
                }

                MessageBox.Show(num + " Penetration references were place where the element intersect", "Results");

                t.Commit();
                return Autodesk.Revit.UI.Result.Succeeded;
            }
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Closestpt_togrids : IExternalCommand
    {
        private XYZ intersect(XYZ point, XYZ direction, Autodesk.Revit.DB.Curve curve)
        {
            Autodesk.Revit.DB.Line unbound = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(point.X, point.Y, curve.GetEndPoint(0).Z), direction);
            IntersectionResultArray ira = null;
            unbound.Intersect(curve, out ira);
            if (ira == null)
            {
                TaskDialog td = new TaskDialog("Error");
                td.MainInstruction = "no intersection";
                td.MainContent = point.ToString() + Environment.NewLine + direction.ToString();
                td.Show();

                return null;
            }
            IntersectionResult ir = ira.get_Item(0);
            return ir.XYZPoint;
        }
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        static AddInId appId = new AddInId(new Guid("5F56AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            List<Autodesk.Revit.DB.Curve> crvs = new List<Autodesk.Revit.DB.Curve>();
            List<Double> crvssort = new List<Double>();


            FilteredElementCollector colGrids = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_Grids).OfClass(typeof(Autodesk.Revit.DB.Grid));


            foreach (var grid   in colGrids)
            {
                Autodesk.Revit.DB.Grid gcurve = grid as Autodesk.Revit.DB.Grid;
                
                Autodesk.Revit.DB.Curve line1 = gcurve.Curve;

                crvs.Add(line1);
                crvssort.Add(line1.Length);
            }
            crvssort.Sort();


            Reference pt_ = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select Peno point");
            Element peno = doc.GetElement(pt_);
            LocationPoint locationpt = peno.Location as LocationPoint;
            XYZ loc_ = locationpt.Point;



            using (Transaction t = new Transaction(doc, "Closes point from peno to grid"))
            {
                t.Start();

                foreach (var line in crvs)
                {
                    if (crvssort.ToArray()[0] == line.Length || crvssort.ToArray()[1] == line.Length)
                    {

                        XYZ pointH = line.Project(loc_).XYZPoint;


                        foreach (var crv in crvs)
                        {
                            ModelLine ml = Makeline(doc, pointH, loc_);
                        }
                    }
                }

                t.Commit();
            }

            


            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Pipe_lines: IExternalCommand
    {
       

        public static XYZ LocalToWorld(Document document, XYZ coordinate, bool millimeters)
        {
            ProjectLocation oProjectLocation = document.ActiveProjectLocation;
            Autodesk.Revit.DB.Transform oTransform = oProjectLocation.GetTotalTransform().Inverse;

            XYZ oSurveyPoint = oTransform.OfPoint(coordinate);

            if (millimeters)
            {
                return new XYZ(oSurveyPoint.X, oSurveyPoint.Y, oSurveyPoint.Z) * 304.8;
            }

            return new XYZ(oSurveyPoint.X, oSurveyPoint.Y, oSurveyPoint.Z);
        }

        private XYZ intersect(XYZ point, XYZ direction, Autodesk.Revit.DB.Curve curve)
        {
            Autodesk.Revit.DB.Line unbound = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(point.X, point.Y, curve.GetEndPoint(0).Z), direction);
            IntersectionResultArray ira = null;
            unbound.Intersect(curve, out ira);
            if (ira == null)
            {
                TaskDialog td = new TaskDialog("Error");
                td.MainInstruction = "no intersection";
                td.MainContent = point.ToString() + Environment.NewLine + direction.ToString();
                td.Show();

                return null;
            }
            IntersectionResult ir = ira.get_Item(0);
            return ir.XYZPoint;
        }
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }
        
        static AddInId appId = new AddInId(new Guid("7A8665BD-55D4-46CE-8CA7-791F66CDC87B"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            List<List<Autodesk.Revit.DB.Curve>> crvs = new List<List<Autodesk.Revit.DB.Curve>>();

            List<Autodesk.Revit.DB.LocationCurve> mid_crvs = new List<Autodesk.Revit.DB.LocationCurve>();

            List<Autodesk.Revit.DB.Curve> edge_crvs = new List<Autodesk.Revit.DB.Curve>();

            List<Autodesk.Revit.DB.Element> elect = new List<Autodesk.Revit.DB.Element>();




            linesfrompipes form = new linesfrompipes();

            MessageBox.Show("Chose the justification lines you want to extract from the pipe", "Select");

            form.ShowDialog();

            

            List<DetailLine> selected = new List<DetailLine>();

            MessageBox.Show("Select one or more pipes", "Select");

            ICollection<Reference> my_lines = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select pipes");
            foreach (var item_myRefWall in my_lines)
            {
                Element ele = doc.GetElement(item_myRefWall);

                Pipe pp = ele as Pipe;
               


                 Options opt = new Options();
                opt.ComputeReferences = true;
                opt.IncludeNonVisibleObjects = true;
                opt.View = doc.ActiveView;

                GeometryObject geoobj = ele.GetGeometryObjectFromReference(item_myRefWall);
                Face face = geoobj as Face;


                mid_crvs.Add(ele.Location as LocationCurve);


                LocationCurve locationCurve2 = ele.Location as LocationCurve;
                Autodesk.Revit.DB.Line line = locationCurve2.Curve as Autodesk.Revit.DB.Line;
                

                foreach (var geoObj in ele.get_Geometry(opt))
                {
                    //Autodesk.Revit.DB.Curve cv = geoObj as Autodesk.Revit.DB.Curve;
                    Solid solid = geoObj as Solid;
                    Autodesk.Revit.DB.Point PT = geoObj as Autodesk.Revit.DB.Point;


                    if (null != solid)
                    {
                        List<Autodesk.Revit.DB.Curve> parallel_crvs = new List<Autodesk.Revit.DB.Curve>();

                        if (line.Direction.Z < -0.9 || line.Direction.Normalize().Z > 0.9)
                        {

                            parallel_crvs.Add(locationCurve2.Curve);
                            crvs.Add(parallel_crvs);
                            continue;
                        }

                        foreach (Face fC in solid.Faces)
                        {


                            if (fC != null)
                            {
                                EdgeArrayArray edgeArrays = fC.EdgeLoops;
                                foreach (EdgeArray edges in edgeArrays)
                                {

                                    foreach (Edge edge in edges)
                                    {
                                        edge_crvs.Add(edge.AsCurve());

                                       
                                        XYZ dir = line.Direction;

                                       

                                        XYZ origin = line.Origin;
                                        XYZ viewdir = line.Direction;
                                        XYZ up = XYZ.BasisZ;
                                        XYZ right = up.CrossProduct(viewdir);

                                        //using (Transaction t = new Transaction(doc, "Create  pipe by line"))
                                        //{
                                        //    t.Start();

                                        //    ModelLine ml = Makeline(doc, locationCurve2.Curve.Evaluate(0.5, true), right);


                                        //    t.Commit();
                                            
                                        //}
                                        

                                        Autodesk.Revit.DB.Transform curveTransform = edge.AsCurve().ComputeDerivatives(0.5, true);

                                        try
                                        {
                                            XYZ origin2 = curveTransform.Origin;
                                            XYZ viewdir2 = curveTransform.BasisX.Normalize();
                                            XYZ viewdir2_back = curveTransform.BasisX.Normalize() * -1;

                                            XYZ up2 = XYZ.BasisZ;
                                            XYZ right2 = up.CrossProduct(viewdir2);
                                            XYZ left2 = up.CrossProduct(viewdir2 * -1);

                                            double y_onverted = Math.Round(-1 * viewdir2.X);

                                            //if (viewdir.IsAlmostEqualTo(right2/*, 0.3333333333*/))
                                            //{
                                            //    parallel_crvs.Add(crv);
                                            //}
                                            //if (viewdir.IsAlmostEqualTo(left2))
                                            //{
                                            //    parallel_crvs.Add(crv);
                                            //}
                                            if (viewdir.IsAlmostEqualTo(viewdir2))
                                            {
                                                parallel_crvs.Add(edge.AsCurve());
                                            }
                                            if (viewdir.IsAlmostEqualTo(viewdir2_back))
                                            {
                                                parallel_crvs.Add(edge.AsCurve());
                                            }
                                        }
                                        catch (Exception)
                                        {
                                            return Autodesk.Revit.UI.Result.Cancelled;
                                        }

                                    }


                                }

                            }
                        }
                        if (parallel_crvs.Count > 0)
                        {
                            crvs.Add(parallel_crvs);
                        }
                        else
                        {
                            elect.Add(ele);

                        }

                    }
                }
            }



            Finish: using (Transaction t = new Transaction(doc, "Create  pipe by line"))
            {
                t.Start();


                if (form.radioButton1.Checked == true)
                {

                    for (int i = 0; i < crvs.ToArray().Length; i++)
                    {
                        if (crvs.ToArray()[i].Count() == 1)
                        {
                            ModelLine ml = Makeline(doc, crvs.ToArray()[i][0].GetEndPoint(0), crvs.ToArray()[i][0].GetEndPoint(1));
                        }
                        else
                        {
                            ModelLine ml = Makeline(doc, crvs.ToArray()[i][1].GetEndPoint(0), crvs.ToArray()[i][1].GetEndPoint(1));
                        }
                       
                    }
                }
                if (form.radioButton2.Checked == true)
                {
                    foreach (var item in mid_crvs)
                    {
                        ModelLine ml = Makeline(doc, item.Curve.GetEndPoint(0), item.Curve.GetEndPoint(1)/* intpt*/);
                    }
                }
                if (form.radioButton3.Checked == true)
                {
                    for (int i = 0; i < crvs.ToArray().Length; i++)
                    {
                        if (crvs.ToArray()[i].Count() == 1)
                        {
                            ModelLine ml = Makeline(doc, crvs.ToArray()[i][0].GetEndPoint(0), crvs.ToArray()[i][0].GetEndPoint(1));
                        }
                        else
                        {
                            ModelLine ml2 = Makeline(doc, crvs.ToArray()[i][5].GetEndPoint(0), crvs.ToArray()[i][5].GetEndPoint(1));
                        }
                        
                    }
                }

                t.Commit();
                
               
            }
            uidoc.Selection.SetElementIds(elect.Select(q => q.Id).ToList());
            return Autodesk.Revit.UI.Result.Succeeded;
        }
        
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Pipe_lines_excel : IExternalCommand
    {
        public XYZ CrossProduct(XYZ v1, XYZ v2)
        {
            double x, y, z;
            x = v1.Y * v2.Z - v2.Y * v1.Z;
            y = (v1.X * v2.Z - v2.X * v1.Z) * -1;
            z = v1.X * v2.Y - v2.X * v1.Y;
            var rtnvector = new XYZ(x, y, z);
            return rtnvector;
        }

        public void WriteTextFile(List< List<string>>  pt_lines)
        {
            string filename = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename = openDialog.FileName;
            }

            string tempPath = Path.GetTempPath();
            string filename2 = Path.Combine(tempPath, filename);

            if (File.Exists(filename2))
            {
                File.Delete(filename2);
            }
            using (StreamWriter writer = new StreamWriter(filename2))
            {
                foreach (var item in pt_lines)
                {
                    foreach (var item2 in item)
                    {
                        writer.WriteLine(item2 /*+ Environment.NewLine*/);
                    }
                    
                }

            }
        }

        public static XYZ LocalToWorld(Document document, XYZ coordinate, bool millimeters)
        {
            ProjectLocation oProjectLocation = document.ActiveProjectLocation;
            Autodesk.Revit.DB.Transform oTransform = oProjectLocation.GetTotalTransform().Inverse;

            XYZ oSurveyPoint = oTransform.OfPoint(coordinate);

            if (millimeters)
            {
                return new XYZ(oSurveyPoint.X, oSurveyPoint.Y, oSurveyPoint.Z) * 304.8;
            }

            return new XYZ(oSurveyPoint.X, oSurveyPoint.Y, oSurveyPoint.Z);
        }

        private XYZ intersect(XYZ point, XYZ direction, Autodesk.Revit.DB.Curve curve)
        {
            Autodesk.Revit.DB.Line unbound = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(point.X, point.Y, curve.GetEndPoint(0).Z), direction);
            IntersectionResultArray ira = null;
            unbound.Intersect(curve, out ira);
            if (ira == null)
            {
                TaskDialog td = new TaskDialog("Error");
                td.MainInstruction = "no intersection";
                td.MainContent = point.ToString() + Environment.NewLine + direction.ToString();
                td.Show();

                return null;
            }
            IntersectionResult ir = ira.get_Item(0);
            return ir.XYZPoint;
        }

        public XYZ InvCoord(XYZ MyCoord)
        {
            XYZ invcoord = new XYZ((Convert.ToDouble(MyCoord.X * -1)),
                (Convert.ToDouble(MyCoord.Y * -1)),
                (Convert.ToDouble(MyCoord.Z * -1)));
            return invcoord;
        }

        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {


            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }



        static AddInId appId = new AddInId(new Guid("7A8665BD-55D4-46CE-8CA7-791F66CDC87B"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            List<List<Autodesk.Revit.DB.Curve>> crvs = new List<List<Autodesk.Revit.DB.Curve>>();

            List<Autodesk.Revit.DB.LocationCurve> mid_crvs = new List<Autodesk.Revit.DB.LocationCurve>();
           
            List<Element> ele_pipes = new List<Element>();

            linesfrompipes form = new linesfrompipes();

            List<Autodesk.Revit.DB.Element> elect = new List<Autodesk.Revit.DB.Element>();

            //MessageBox.Show("Chose the justification lines you want to extract from the pipe", "Select");

            form.ShowDialog();
            

            MessageBox.Show("Select one or more pipes", "Select");

            ICollection<Reference> my_lines = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element , "Select pipes");

            int count = 0;
            foreach (var item in my_lines)
            {

                Element ele = doc.GetElement(item);
                if (ele.Category.Name == "Pipes")
                {

                    ele_pipes.Add(ele);
                    count++;
                }
                else
                {
                    MessageBox.Show("Within the selection there an object which is not a pipe","Error");
                    return Autodesk.Revit.UI.Result.Cancelled;
                }
            }

            TaskDialog.Show("info", "Numeber of Elements selected = " + count.ToString());

            if (ele_pipes.Count == 0)
            {
                TaskDialog.Show("Failed", "the selection does not contains pipes elements");
                return Autodesk.Revit.UI.Result.Cancelled;
            }
            
            List<List<string>> lista_pts = new List<List<string>>();


            foreach (var item_myRefWall in my_lines)
            {
                Element ele = doc.GetElement(item_myRefWall);
                Pipe type = ele as Pipe;
               var dia = type.LookupParameter("Size").AsString();
                

                List<string> pts = new List<string>();
                try
                {
                    Parameter params_ = type.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                    var name = params_.AsValueString();
                    try
                    {
                        name = name.Replace("_", " ");
                    }
                    catch (Exception)
                    {
                        
                    }
                    try
                    {
                        name = name.Replace("-", " ");

                    }
                    catch (Exception)
                    {
                        
                    }
                    try
                    {
                        pts.Add("model " + "\"" + name.Replace("_"," ") + " "+ dia +  "\"" );
                        pts.Add("string super {");
                        pts.Add("data_3d {" );
                    }
                    catch (Exception)
                    {
                    }
                    LocationCurve locCrv = ele.Location as LocationCurve;
                }
                catch (Exception)
                {
                    TaskDialog.Show("Failed", "The tool could not foundt pipe type");
                    return Autodesk.Revit.UI.Result.Cancelled;
                }

                Options opt = new Options();
                opt.ComputeReferences = true;
                opt.IncludeNonVisibleObjects = true;
                opt.View = doc.ActiveView;
                GeometryObject geoobj = ele.GetGeometryObjectFromReference(item_myRefWall);
                Face face = geoobj as Face;
                mid_crvs.Add(ele.Location as LocationCurve);
                LocationCurve locationCurve2 = ele.Location as LocationCurve;

                
                Autodesk.Revit.DB.Line Orig_cent_line = locationCurve2.Curve as Autodesk.Revit.DB.Line;

                
                XYZ pntCenter = Orig_cent_line.Evaluate(0.5, true);
                XYZ normal = Orig_cent_line.Direction.Normalize();
                XYZ dir2 = new XYZ(0, 0, 1);
                XYZ cross = normal.CrossProduct(dir2);
                XYZ pntEnd = pntCenter + cross.Multiply(10);


                try
                {
                    Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(pntCenter, pntEnd);


                    XYZ cross_p = CrossProduct(Orig_cent_line.Direction, line2.Direction);

                    Autodesk.Revit.DB.Line Perp_to_cent_line = Autodesk.Revit.DB.Line.CreateUnbound(line2.Evaluate(0, true), cross_p);

                    InvCoord(cross_p);

                    Autodesk.Revit.DB.View currentView = uidoc.ActiveView;
                    View3D view3D = currentView as View3D;

                   

                    ElementClassFilter filter_1 = new ElementClassFilter(typeof(Pipe));
                    ReferenceIntersector refIntersector_1 = new ReferenceIntersector(filter_1, FindReferenceTarget.Face, view3D);
                    ReferenceWithContext referenceWithContext_1 = refIntersector_1.FindNearest(Orig_cent_line.Evaluate(0, true), InvCoord(cross_p));
                    Reference reference_1 = referenceWithContext_1.GetReference();
                    XYZ intersection_1 = reference_1.GlobalPoint;

                    ReferenceIntersector refIntersector_neg_1 = new ReferenceIntersector(filter_1, FindReferenceTarget.Face, view3D);
                    ReferenceWithContext referenceWithContext_neg_1 = refIntersector_neg_1.FindNearest(Orig_cent_line.Evaluate(1, true), InvCoord(cross_p));
                    Reference reference_neg_1 = referenceWithContext_neg_1.GetReference();
                    XYZ intersection_neg_1 = reference_neg_1.GlobalPoint;

                    ReferenceIntersector refIntersector_2 = new ReferenceIntersector(filter_1, FindReferenceTarget.Face, view3D);
                    ReferenceWithContext referenceWithContext_2 = refIntersector_2.FindNearest(Orig_cent_line.Evaluate(0, true), cross_p);
                    Reference reference_2 = referenceWithContext_2.GetReference();
                    XYZ intersection_2 = reference_2.GlobalPoint;

                    ReferenceIntersector refIntersector_neg_2 = new ReferenceIntersector(filter_1, FindReferenceTarget.Face, view3D);
                    ReferenceWithContext referenceWithContext_neg_2 = refIntersector_neg_2.FindNearest(Orig_cent_line.Evaluate(1, true), cross_p);
                    Reference reference_neg_2 = referenceWithContext_neg_2.GetReference();
                    XYZ intersection_neg_2 = reference_neg_2.GlobalPoint;

                    if (form.radioButton1.Checked == true)
                    {
                        XYZ newpt0 = LocalToWorld(doc, new XYZ(intersection_1.X, intersection_1.Y, intersection_1.Z), true);
                        double x_ = UnitUtils.Convert(newpt0.X, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        double y_ = UnitUtils.Convert(newpt0.Y, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        double z_ = UnitUtils.Convert(newpt0.Z, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        XYZ newpt00 = new XYZ(Math.Round(x_, 6), Math.Round(y_, 6), Math.Round(z_, 6));
                        pts.Add(newpt00.X.ToString() + " " + newpt00.Y.ToString() + " " + newpt00.Z.ToString());

                        XYZ newpt1 = LocalToWorld(doc, new XYZ(intersection_neg_1.X, intersection_neg_1.Y, intersection_neg_1.Z), true);
                        double x_1 = UnitUtils.Convert(newpt1.X, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        double y_1 = UnitUtils.Convert(newpt1.Y, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        double z_1 = UnitUtils.Convert(newpt1.Z, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        XYZ newpt11 = new XYZ(Math.Round(x_1, 6), Math.Round(y_1, 6), Math.Round(z_1, 6));
                        pts.Add(newpt11.X.ToString() + " " + newpt11.Y.ToString() + " " + newpt11.Z.ToString());
                        //using (Transaction t = new Transaction(doc, "Create  pipe by line"))
                        //{
                        //    t.Start();
                        //    ModelLine ml1 = Makeline(doc, Perp_to_cent_line.Origin, Perp_to_cent_line.Evaluate(5, false));
                        //    ModelLine Line_perp_neg = Makeline(doc, line2.Evaluate(0, true), line2.Evaluate(1, true));
                        //    ModelLine obvert = Makeline(doc, intersection_1, intersection_neg_1);
                        //    t.Commit();
                        //}
                        type.LookupParameter("Size").AsDouble();
                        double diameter = UnitUtils.Convert((type.Diameter * 304.8), DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);

                        
                        pts.Add("}");
                        pts.Add("pipe_value {");
                        pts.Add("diameter " + diameter);
                        pts.Add("}");
                        pts.Add("justify top");
                        pts.Add("}");
                    }
                    if (form.radioButton2.Checked == true)
                    {
                        double diameter = UnitUtils.Convert((type.Diameter * 304.8), DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        
                        try
                        {
                            XYZ newpt0 = LocalToWorld(doc, new XYZ(Orig_cent_line.GetEndPoint(0).X, Orig_cent_line.GetEndPoint(0).Y, Orig_cent_line.GetEndPoint(0).Z), true);
                            double x_ = UnitUtils.Convert(newpt0.X, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                            double y_ = UnitUtils.Convert(newpt0.Y, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                            double z_ = UnitUtils.Convert(newpt0.Z, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                            XYZ newpt00 = new XYZ(Math.Round(x_, 6), Math.Round(y_, 6), Math.Round(z_, 6));
                            pts.Add(newpt00.X.ToString() + " " + newpt00.Y.ToString() + " " + newpt00.Z.ToString());

                            XYZ newpt1 = LocalToWorld(doc, new XYZ(Orig_cent_line.GetEndPoint(1).X, Orig_cent_line.GetEndPoint(1).Y, Orig_cent_line.GetEndPoint(1).Z), true);
                            double x_1 = UnitUtils.Convert(newpt1.X, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                            double y_1 = UnitUtils.Convert(newpt1.Y, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                            double z_1 = UnitUtils.Convert(newpt1.Z, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                            XYZ newpt11 = new XYZ(Math.Round(x_1, 6), Math.Round(y_1, 6), Math.Round(z_1, 6));
                            pts.Add(newpt11.X.ToString() + " " + newpt11.Y.ToString() + " " + newpt11.Z.ToString());
                        }
                        catch (Exception)
                        {
                        }
                        pts.Add("}");
                        pts.Add("pipe_value {");
                        pts.Add("diameter " + diameter);
                        pts.Add("}");
                        pts.Add("justify centre");
                        pts.Add("}");
                    }
                    if (form.radioButton3.Checked == true)
                    {
                        XYZ newpt0 = LocalToWorld(doc, new XYZ(intersection_2.X, intersection_2.Y, intersection_2.Z), true);
                        double x_ = UnitUtils.Convert(newpt0.X, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        double y_ = UnitUtils.Convert(newpt0.Y, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        double z_ = UnitUtils.Convert(newpt0.Z, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        XYZ newpt00 = new XYZ(Math.Round(x_, 6), Math.Round(y_, 6), Math.Round(z_, 6));
                        pts.Add(newpt00.X.ToString() + " " + newpt00.Y.ToString() + " " + newpt00.Z.ToString());

                        XYZ newpt1 = LocalToWorld(doc, new XYZ(intersection_neg_2.X, intersection_neg_2.Y, intersection_neg_2.Z), true);
                        double x_1 = UnitUtils.Convert(newpt1.X, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        double y_1 = UnitUtils.Convert(newpt1.Y, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        double z_1 = UnitUtils.Convert(newpt1.Z, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        XYZ newpt11 = new XYZ(Math.Round(x_1, 6), Math.Round(y_1, 6), Math.Round(z_1, 6));

                        pts.Add(newpt11.X.ToString() + " " + newpt11.Y.ToString() + " " + newpt11.Z.ToString());
                        //using (Transaction t = new Transaction(doc, "Create  pipe by line"))
                        //{
                        //    t.Start();
                        //    ModelLine ml1 = Makeline(doc, Perp_to_cent_line.Origin, Perp_to_cent_line.Evaluate(5, false));
                        //    ModelLine Line_perp_neg = Makeline(doc, line2.Evaluate(0, true), line2.Evaluate(1, true));
                        //    ModelLine invert = Makeline(doc, intersection_2, intersection_neg_2);
                        //    t.Commit();
                        //}
                        double diameter = UnitUtils.Convert((type.Diameter * 304.8), DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        
                        pts.Add("}");
                        pts.Add("pipe_value {");
                        pts.Add("diameter " + diameter);
                        pts.Add("}");
                        pts.Add("justify bottom");
                        pts.Add("}");
                    }
                }
                catch (Exception)
                {
                    double diameter = UnitUtils.Convert((type.Diameter * 304.8), DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);

                    try
                    {
                        XYZ newpt0 = LocalToWorld(doc, new XYZ(Orig_cent_line.GetEndPoint(0).X, Orig_cent_line.GetEndPoint(0).Y, Orig_cent_line.GetEndPoint(0).Z), true);
                        double x_ = UnitUtils.Convert(newpt0.X, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        double y_ = UnitUtils.Convert(newpt0.Y, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        double z_ = UnitUtils.Convert(newpt0.Z, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        XYZ newpt00 = new XYZ(Math.Round(x_, 6), Math.Round(y_, 6), Math.Round(z_, 6));
                        pts.Add(newpt00.X.ToString() + " " + newpt00.Y.ToString() + " " + newpt00.Z.ToString());

                        XYZ newpt1 = LocalToWorld(doc, new XYZ(Orig_cent_line.GetEndPoint(1).X, Orig_cent_line.GetEndPoint(1).Y, Orig_cent_line.GetEndPoint(1).Z), true);
                        double x_1 = UnitUtils.Convert(newpt1.X, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        double y_1 = UnitUtils.Convert(newpt1.Y, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        double z_1 = UnitUtils.Convert(newpt1.Z, DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_METERS);
                        XYZ newpt11 = new XYZ(Math.Round(x_1, 6), Math.Round(y_1, 6), Math.Round(z_1, 6));
                        pts.Add(newpt11.X.ToString() + " " + newpt11.Y.ToString() + " " + newpt11.Z.ToString());
                    }
                    catch (Exception)
                    {
                    }
                    pts.Add("}");
                    pts.Add("pipe_value {");
                    pts.Add("diameter " + diameter);
                    pts.Add("}");
                    pts.Add("justify centre");
                    pts.Add("}");

                }


                lista_pts.Add(pts);
            }

            TaskDialog.Show("Info", "Select a TEXT file in your system");
            WriteTextFile(lista_pts);

            if (elect.ToArray().Length != 0)
            {
                TaskDialog.Show("Info", "The tool selected pipes that could not be pocessed, total = "+ elect.ToArray().Length.ToString());
                uidoc.Selection.SetElementIds(elect.Select(q => q.Id).ToList());
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Isolate_Pipe : IExternalCommand
    {
        
        static AddInId appId = new AddInId(new Guid("5F56AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            List<List<Autodesk.Revit.DB.Curve>> crvs = new List<List<Autodesk.Revit.DB.Curve>>();
            List<Autodesk.Revit.DB.LocationCurve> mid_crvs = new List<Autodesk.Revit.DB.LocationCurve>();
            List<DetailLine> selected = new List<DetailLine>();
            List<Element> ele_pipes = new List<Element>();
            List<Pipe> all_pipes = new List<Pipe>();
            List<string> names = new List<string>();
            
            ICollection<Reference> my_lines = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select pipes");

            Pipe typepipe = null;
            foreach (var item_myRefWall in my_lines)
            {
                Element element = doc.GetElement(item_myRefWall);

                if (element.GetType().Name == "Pipe")
                {
                    typepipe = element as Pipe;
                    Parameter params_ = typepipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                    var name = params_.AsValueString();
                    names.Add(name);
                }
                else
                {
                    TaskDialog.Show("Errror", "there are objects in the selection which are not pipes, please try again");
                    return Autodesk.Revit.UI.Result.Cancelled;
                }
            }
            IEnumerable<Pipe> pipes = from elem in new FilteredElementCollector(doc).OfClass(typeof(Pipe)).OfCategory(BuiltInCategory.OST_PipeCurves)
                                      let type = elem as Pipe where type.Name != null select type;

            //FilteredElementCollector systemTypeColl = new FilteredElementCollector(doc);
            //systemTypeColl.OfClass(typeof(PipingSystemType));
            //PipingSystemType systemType = systemTypeColl.FirstElement() as PipingSystemType;
            //List<Element> elements = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).Where(a => a.LookupParameter("System Type").AsString() == system).ToList();
            

            FilteredElementCollector allElementsInView = new FilteredElementCollector(doc, doc.ActiveView.Id);
            IList elementsInView = (IList)allElementsInView.ToElements();

            
            List<Element> elelist = new List<Element>();


            foreach (var pp in pipes)
            {
                Parameter params_2 = pp.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                var name2 = params_2.AsValueString();

                foreach (var selectedname in names)
                {
                    if (selectedname == name2)
                    {
                        elelist.Add(pp);
                    }
                }

                
            }

        
            uidoc.Selection.SetElementIds(elelist.Select(q => q.Id).ToList());
            
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class rhino_lns_to_csv : IExternalCommand
    {


        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        public static bool IsZero(double a)
        {
            const double _eps = 1.0e-9;
            return _eps > Math.Abs(a);
        }

        public static bool IsEqual(double a, double b)
        {
            return IsZero(b - a);
        }

        public static int Compare(double a, double b)
        {
            return IsEqual(a, b) ? 0 : (a < b ? -1 : 1);
        }

        public static int Compare(XYZ p, XYZ q)
        {

            int diff = Compare(p.X, q.X);
            if (0 == diff)
            {
                diff = Compare(p.Y, q.Y);
                if (0 == diff)
                {
                    diff = Compare(p.Z, q.Z);
                }

            }

            return diff;
        }

        private static Wall CreateWall(FamilyInstance cube, Autodesk.Revit.DB.Curve curve, double height)
        {
            var doc = cube.Document;

            var wallTypeId = doc.GetDefaultElementTypeId(
              ElementTypeGroup.WallType);

            return Wall.Create(doc, curve.CreateReversed(),
              wallTypeId, cube.LevelId, height, 0, false,
              false);
        }
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 15;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }



            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            String info = "";
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            //List<FaceArray> face112 = new List<FaceArray>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);

            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }

            Point3d p1 = new Point3d(0, 0, 0);
            Rhino.Geometry.Point3d pt3d = new Point3d(10, 10, 0);

            File3dm m_modelfile = null;
            string m_name = null;
            string m_size = null;
            string m_created = null;
            string m_createdby = null;
            string m_edited = null;
            string m_editedby = null;
            string m_revision = null;
            string m_units = null;
            string m_notes = null;

            Object RhinoFile = filename2;
            if (RhinoFile is System.IO.FileInfo)
            {
                System.IO.FileInfo m_fileinfo = (System.IO.FileInfo)RhinoFile;
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }
            else if (RhinoFile is string)
            {
                System.IO.FileInfo m_fileinfo = new System.IO.FileInfo((string)RhinoFile);
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }

            File3dmLayerTable all_layers = m_modelfile.AllLayers;
            var objs_ = m_modelfile.Objects;
            List<Rhino.DocObjects.Layer> listadelayers = new List<Rhino.DocObjects.Layer>();
            List<Rhino.DocObjects.Layer> borrar = new List<Rhino.DocObjects.Layer>();
            List<List<Rhino.DocObjects.Layer>> multiplelistadelayers = new List<List<Rhino.DocObjects.Layer>>();
            string hola = "";

            MessageBox.Show("This tool will read 3D information only if the following Rhino Layers exist; Levels, Grids, Structure, Floor, Walls, Points", "!");

            foreach (var item in all_layers)
            {

                if (item.FullPath.Contains(hola))
                {
                    listadelayers.Add(item);
                }
            }

            List<Rhino.Geometry.Brep> rh_breps = new List<Rhino.Geometry.Brep>();
            List<XYZ> pts_ = new List<XYZ>();

            List<Autodesk.Revit.DB.Curve> revit_crv = new List<Autodesk.Revit.DB.Curve>();
            List<string> m_names = new List<string>();
            List<int> m_layerindeces = new List<int>();
            List<System.Drawing.Color> m_colors = new List<System.Drawing.Color>();
            List<string> m_guids = new List<string>();
            List<double> weith = new List<double>();



            foreach (var Layer in listadelayers)
            {


                if (Layer.Name == "Lines")
                {


                    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);

                    using (Transaction t = new Transaction(doc, "Lines from Rhino"))
                    {
                        t.Start();

                        foreach (File3dmObject obj in m_objs)
                        {
                            GeometryBase geo = obj.Geometry;
                            Type TYPE = geo.GetType();

                            if (TYPE.Name == "PolyCurve")
                            {
                                Rhino.Geometry.PolyCurve POLY = geo as Rhino.Geometry.PolyCurve;
                            }



                            if (TYPE.Name == "ArcCurve")
                            {
                                Rhino.Geometry.ArcCurve arc = geo as Rhino.Geometry.ArcCurve;

                                double x_end = arc.Arc.Plane.Normal.X / 304.8;
                                double y_end = arc.Arc.Plane.Normal.Y / 304.8;
                                double z_end = arc.Arc.Plane.Normal.Z / 304.8;

                                XYZ normal = new XYZ(x_end, y_end, z_end);

                                double x_ = arc.Arc.Plane.Origin.X / 304.8;
                                double y_ = arc.Arc.Plane.Origin.Y / 304.8;
                                double z_ = arc.Arc.Plane.Origin.Z / 304.8;
                                XYZ origin = new XYZ(x_, y_, z_);

                                double endAngle = 2 * Math.PI;        // this arc will be a circle

                                Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(normal, origin);

                                SketchPlane skplane = SketchPlane.Create(doc, pl);
                                doc.Create.NewModelCurve(Autodesk.Revit.DB.Arc.Create(pl, arc.Radius / 304.8, arc.Arc.StartAngle, arc.Arc.EndAngle), skplane);

                            }

                            if (TYPE.Name == "PolylineCurve")
                            {
                                Rhino.Geometry.PolylineCurve PolylineCurve = geo as Rhino.Geometry.PolylineCurve;
                                Rhino.Geometry.Polyline Polyline = PolylineCurve.ToPolyline();

                                for (int i = 0; i < Polyline.SegmentCount; i++)
                                {
                                    Point3d end = Polyline.SegmentAt(i).PointAt(0);
                                    Point3d start = Polyline.SegmentAt(i).PointAt(1);


                                    double x_end = end.X / 304.8;
                                    double y_end = end.Y / 304.8;
                                    double z_end = end.Z / 304.8;

                                    double x_start = start.X / 304.8;
                                    double y_start = start.Y / 304.8;
                                    double z_start = start.Z / 304.8;

                                    XYZ pt_end = new XYZ(x_end, y_end, z_end);
                                    XYZ pt_start = new XYZ(x_start, y_start, z_start);

                                    Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                    Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);

                                    Makeline(doc, line.GetEndPoint(0), line.GetEndPoint(1));
                                    revit_crv.Add(curve1);
                                }


                            }

                            if (TYPE.Name == "NurbsCurve")
                            {
                                Rhino.Geometry.NurbsCurve crv_ = geo as Rhino.Geometry.NurbsCurve;
                                Rhino.Geometry.Plane pl2;
                                crv_.TryGetPlane(out pl2);

                                double x_end = pl2.Normal.X / 304.8;
                                double y_end = pl2.Normal.Y / 304.8;
                                double z_end = pl2.Normal.Z / 304.8;

                                XYZ normal = new XYZ(x_end, y_end, z_end);

                                double x_2 = pl2.Origin.X / 304.8;
                                double y_2 = pl2.Origin.Y / 304.8;
                                double z_2 = pl2.Origin.Z / 304.8;
                                XYZ origin = new XYZ(x_2, y_2, z_2);

                                // this arc will be a circle

                                Autodesk.Revit.DB.Plane pl3 = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(normal, origin);

                                var pts = crv_.Points;
                                foreach (var item in pts)
                                {
                                    double x_ = item.X / 304.8;
                                    double y_ = item.Y / 304.8;
                                    double z_ = item.Z / 304.8;

                                    XYZ pt = new XYZ(x_, y_, z_);
                                    weith.Add(item.Weight);
                                    pts_.Add(pt);
                                }

                                Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pts_.First(), pts_.Last());

                                XYZ norm = pts_.First().CrossProduct(curve1.Evaluate(5, false));
                                if (norm.GetLength() == 0)
                                {
                                    XYZ aSubB = pts_.First().Subtract(pts_.Last());
                                    XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                                    double crosslenght = aSubBcrossz.GetLength();
                                    if (crosslenght == 0)
                                    {
                                        norm = XYZ.BasisY;
                                    }
                                    else
                                    {
                                        norm = XYZ.BasisZ;
                                    }
                                }
                                //Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, pts_.First());
                                Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(pts_.First(), pts_.ElementAt(2), pts_.Last());
                                SketchPlane skplane = SketchPlane.Create(doc, pl);

                                Autodesk.Revit.DB.Curve nspline = Autodesk.Revit.DB.NurbSpline.CreateCurve(pts_, weith);
                                doc.Create.NewModelCurve(nspline, skplane);

                                pts_.Clear();
                                weith.Clear();
                            }

                            if (TYPE.Name == "Curve")
                            {
                                Rhino.Geometry.Curve crv_ = geo as Rhino.Geometry.Curve;


                                Point3d end = crv_.PointAtEnd;
                                Point3d start = crv_.PointAtStart;

                                double x_end = end.X / 304.8;
                                double y_end = end.Y / 304.8;
                                double z_end = end.Z / 304.8;

                                double x_start = start.X / 304.8;
                                double y_start = start.Y / 304.8;
                                double z_start = start.Z / 304.8;

                                XYZ pt_end = new XYZ(x_end, y_end, z_end);
                                XYZ pt_start = new XYZ(x_start, y_start, z_start);

                                Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);

                                Makeline(doc, line.GetEndPoint(0), line.GetEndPoint(1));
                                revit_crv.Add(curve1);
                            }
                        }
                        t.Commit();
                    }
                }
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Wall_Elevation : IExternalCommand
    {
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {


            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }



        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            ViewFamilyType vftele = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Elevation == x.ViewFamily);



            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();

            Wall wall = null;

            foreach (ElementId id in ids)
            {
                Element e = doc.GetElement(id);

                wall = e as Wall;

                LocationCurve lc = wall.Location as LocationCurve;

                Autodesk.Revit.DB.Line line = lc.Curve as Autodesk.Revit.DB.Line;

                if (null == line)
                {
                    message = "Unable to retrieve wall location line.";

                    return Result.Failed;
                }

                //XYZ p = line.GetEndPoint(0);
                //XYZ q = line.GetEndPoint(1);
                //XYZ v = q - p;

                //BoundingBoxXYZ bb = wall.get_BoundingBox(null);
                //double minZ = bb.Min.Z;
                //double maxZ = bb.Max.Z;

                //double w = v.GetLength();
                //double h = maxZ - minZ;
                //double d = wall.WallType.Width;
                //double offset = 0.1 * w;

                //XYZ min = new XYZ(-w, minZ - offset, -offset);
                //XYZ max = new XYZ(w, maxZ + offset, 0);

                //XYZ midpoint = p + 0.5 * v;
                //XYZ walldir = v.Normalize();
                //XYZ up = XYZ.BasisZ;
                //XYZ viewdir = walldir.CrossProduct(up);

                //Autodesk.Revit.DB.Transform t = Autodesk.Revit.DB.Transform.Identity;
                //t.Origin = midpoint;
                //t.BasisX = walldir;
                //t.BasisY = up;
                //t.BasisZ = viewdir;

                //BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
                //sectionBox.Transform = t;
                //sectionBox.Min = min;
                //sectionBox.Max = max;

                XYZ pntCenter = line.Evaluate(0.5, true);
                XYZ normal = line.Direction.Normalize();
                XYZ dir = new XYZ(0, 0, 1);
                XYZ cross = normal.CrossProduct(dir * -1);
                XYZ pntEnd = pntCenter + cross.Multiply(2);

                XYZ poscross = normal.CrossProduct(dir);
                XYZ pospntEnd = pntCenter + poscross.Multiply(2);


                XYZ vect1 = line.Direction * (-1100 / 304.8);
                vect1 = vect1.Negate();
                XYZ vect2 = vect1 + line.Evaluate(0.5, true);
                Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateBound(line.Evaluate(0.5, true), vect2);


                Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(pntCenter, pospntEnd);

                double angle4 = XYZ.BasisY.AngleTo(line2.Direction);
                double angleDegrees4 = angle4 * 180 / Math.PI;
                if (pntCenter.X < line2.GetEndPoint(1).X)
                {
                    angle4 = 2 * Math.PI - angle4;
                }
                double angleDegreesCorrected4 = angle4 * 180 / Math.PI;

                Autodesk.Revit.DB.Line axis = Autodesk.Revit.DB.Line.CreateUnbound(pntEnd, XYZ.BasisZ);

                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Create Wall Section View");



                    ElevationMarker marker = ElevationMarker.CreateElevationMarker(doc, vftele.Id, pntEnd/*line.Evaluate(0.5,true)*/, 100);
                    ViewSection elevation1 = marker.CreateElevation(doc, doc.ActiveView.Id, 1);



                    if (angleDegreesCorrected4 > 160 && angleDegreesCorrected4 < 200)
                    {
                        angle4 = angle4 / 2;

                        marker.Location.Rotate(axis, angle4);
                    }

                    marker.Location.Rotate(axis, angle4);
                    //ModelCurve mline_dir_right2 = Makeline(doc, line.Evaluate(0.5, true), pntEnd);
                    //elevation1.CropBox = e.get_BoundingBox(null);

                    tx.Commit();
                }
            }








            return Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Rot_Wal_Angle_To_Grid : IExternalCommand
    {
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {


            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-5609-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            ViewFamilyType vftele = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Elevation == x.ViewFamily);

            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();

            List<Wall> Wall_Sel = new List<Wall>();
            foreach (ElementId id in ids)
            {
                Element e = doc.GetElement(id);

                Wall_Sel.Add(e as Wall);
            }

            List<Element> ele = new List<Element>();
            List<Element> ele2 = new List<Element>();
            IList<Reference> refList = new List<Reference>();

            TaskDialog.Show("!", "Select a reference Grid to find orthogonal walls");
            Autodesk.Revit.DB.Grid levelBelow = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select Grid")) as Autodesk.Revit.DB.Grid;

            Autodesk.Revit.DB.Curve dircurve = levelBelow.Curve;
            Autodesk.Revit.DB.Line Grid_line = dircurve as Autodesk.Revit.DB.Line;
            //XYZ dir = Grid_line.Direction;


            XYZ grid_dir = Grid_line.Direction;
            XYZ up = XYZ.BasisZ;
            XYZ right = up.CrossProduct(grid_dir);
            List<Element> To_Be_Rotated = new List<Element>();

            foreach (var wall in Wall_Sel)
            {
                LocationCurve lc = wall.Location as LocationCurve;
                Autodesk.Revit.DB.Transform curveTransform = lc.Curve.ComputeDerivatives(0.5, true);

                

                foreach (Wall id in Wall_Sel)
                {

                    try
                    {
                        XYZ origin2 = curveTransform.Origin;
                        XYZ viewdir2 = curveTransform.BasisX.Normalize();
                        XYZ viewdir2_back = curveTransform.BasisX.Normalize() * -1;

                        XYZ up2 = XYZ.BasisZ;
                        XYZ right2 = up.CrossProduct(viewdir2);
                        XYZ left2 = up.CrossProduct(viewdir2 * -1);

                        double y_onverted = Math.Round(-1 * viewdir2.X);

                        if (grid_dir.IsAlmostEqualTo(right2/*, 0.3333333333*/) || grid_dir.IsAlmostEqualTo(left2) || grid_dir.IsAlmostEqualTo(viewdir2) || grid_dir.IsAlmostEqualTo(viewdir2_back))
                        {
                            ele.Add(wall);
                        }
                        else
                        {
                            To_Be_Rotated.Add(wall);
                        }
                    }
                    catch (Exception)
                    {
                        return Autodesk.Revit.UI.Result.Cancelled;
                    }
                }

                

            }

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Create Wall Section View");

                foreach (var wall2 in To_Be_Rotated)
                {
                    LocationCurve lc2 = wall2.Location as LocationCurve;

                    Autodesk.Revit.DB.Line rot_line2 = lc2.Curve as Autodesk.Revit.DB.Line;

                    double angle4 = grid_dir.AngleTo(rot_line2.Direction);
                    double angleDegrees4 = angle4 * 180 / Math.PI;
                    if (grid_dir.X < rot_line2.GetEndPoint(1).X)
                    {
                        angle4 = 2 * Math.PI - angle4;
                    }
                    else
                    {
                        angle4 = angle4 * -1;
                    }
                    double angleDegreesCorrected4 = angle4 * 180 / Math.PI;

                    Autodesk.Revit.DB.Line axis = Autodesk.Revit.DB.Line.CreateUnbound(lc2.Curve.Evaluate(0.5, true), XYZ.BasisZ);

                    //if (angleDegreesCorrected4 > 160 && angleDegreesCorrected4 < 200)
                    //{
                    //    angle4 = angle4 / 2;

                    //    wall2.Location.Rotate(axis, angleDegreesCorrected4);
                    //}
                    wall2.Location.Rotate(axis, angle4);
                    //Makeline(doc, axis.Evaluate(0, false), XYZ.BasisZ);


                }
                //ModelCurve mline_dir_right2 = Makeline(doc, line.Evaluate(0.5, true), pntEnd);
                //elevation1.CropBox = e.get_BoundingBox(null);

                tx.Commit();
            }
            return Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Set_Annotation_Crop : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("6C22CC72-A167-4819-AAF1-A178F6B44BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
           
            // Get selected elements from current document.;
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

            string s = "You Picked:" + "\n";

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Set Annotation Crop size");

                foreach (ElementId selectedid in selectedIds)
                {
                    Autodesk.Revit.DB.View e = doc.GetElement(selectedid) as Autodesk.Revit.DB.View;

                    s += " View id = " + e.Id + "\n";

                    s += " View name = " + e.Name + "\n";

                    s += "\n";

                    ViewCropRegionShapeManager regionManager = e.GetCropRegionShapeManager();

                    regionManager.BottomAnnotationCropOffset = 0.01;

                    regionManager.LeftAnnotationCropOffset = 0.01;

                    regionManager.RightAnnotationCropOffset = 0.01;

                    regionManager.TopAnnotationCropOffset = 0.01;
                }
                {
                    TaskDialog.Show("Basic Element Info", s);
                }
                tx.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Register_ : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("6C22CC72-A167-4819-AAF1-A178F6B44BAB"));
        static public Autodesk.Revit.ApplicationServices.Application m_app;

        public System.Timers.Timer timer;
        public int h, m, s;
        public void CountdownTimer_Load(object sender, EventArgs e)
        {
            timer = new System.Timers.Timer();
            timer.Interval = 100;
            timer.Elapsed += OnTimeEvent;
        }
        public void OnTimeEvent(object sender, ElapsedEventArgs e)
        {
            s += 1;
            if (s == 60)
            {
                s = 0;
                m += 1;
            }
            if (m == 60)
            {
                m = 0;
                h += 1;
            }

        }

        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            //m_app.DocumentSynchronizingWithCentral += new EventHandler<DocumentSynchronizingWithCentralEventArgs>(myCommand.myDocumentSaving);
            //return Autodesk.Revit.UI.Result.Succeeded;
            //doc.DocumentSaving += new EventHandler<DocumentSavingEventArgs>(myCommand.myDocumentSaving);


            if (doc.Title == "RHR_BUILDING_A22")
            {
               
                double reset = 10000000;
                SyncListUpdater SyncListUpdater_ = new SyncListUpdater();

                string user = doc.Application.Username;

                string s = "People syncing:" + "\n";

            beggining:
                try
                {
                    string Sync_Manager = @"T:\Lopez\Sync_Manager.xlsx";
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
                    {
                        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);
                        var Time_ = DateTime.Now;
                        //---------------------------------------------------------
                        for (int row = 1; row < 20; row++)
                        {
                            if (sheet.Cells[row, 1].Value == null)
                            {
                                break;
                            }
                            if (sheet.Cells[row, 1].Value != null)
                            {
                                var Value1 = sheet.Cells[row, 1].Value;
                                var Value2 = sheet.Cells[row, 2].Value;
                                //s += Value1 + " + " + Value2.ToString() + "\n";
                                SyncListUpdater_.listBox1.Items.Add(Value1 + " + " + Value2.ToString() + "\n");
                            }

                        }

                        SyncListUpdater_.Show();

                    //TaskDialog.Show("Current Sync List ", s);
                    //MessageBox.Show("Current Sync List ", "");

                    nonvalue:
                        //---------------------------------------------------------
                        try
                        {
                            if (sheet.Cells[1, 1].Value == null)
                            {
                                sheet.Cells[1, 1].Value = Time_.ToString();
                                sheet.Cells[1, 2].Value = user;
                                package.Save();
                                goto finish_;
                            }
                        }
                        catch (Exception)
                        {
                            goto nonvalue;
                        }

                        //---------------------------------------------------------
                        if (sheet.Cells[1, 2].Value.ToString() != "Alex synced")
                        {
                            for (int row = 1; row < 9999; row++)
                            {
                                var thisValue = sheet.Cells[row, 1].Value;

                                if (thisValue == null)
                                {
                                    sheet.Cells[row, 1].Value = Time_.ToString();
                                    sheet.Cells[row, 2].Value = user;
                                    package.Save();
                                    goto finder;
                                }
                                else
                                {


                                }
                            }
                        }
                        //---------------------------------------------------------

                        //---------------------------------------------------------
                    }
                }
                catch (Exception)
                {
                    //MessageBox.Show("Excel file not found", "");
                    //return;
                    goto beggining;

                }
            finder:
                for (int i = 0; i < reset; i++)
                {
                    if (i == 9999999)
                    {
                        try
                        {
                            SyncListUpdater_.listBox1.Items.Clear();
                            string Sync_Manager = @"T:\Lopez\Sync_Manager.xlsx";
                            using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
                            {
                                ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                                for (int row = 1; row < 20; row++)
                                {
                                    if (sheet.Cells[row, 1].Value == null)
                                    {
                                        break;
                                    }
                                    if (sheet.Cells[row, 1].Value != null)
                                    {
                                        var Value1 = sheet.Cells[row, 1].Value;
                                        var Value2 = sheet.Cells[row, 2].Value;

                                        SyncListUpdater_.listBox1.Items.Add(Value1 + " + " + Value2.ToString());
                                    }

                                }

                                if (sheet.Cells[1, 1].Value != null && sheet.Cells[1, 2].Value.ToString() != user)
                                {
                                    SyncListUpdater_.listBox1.Refresh();
                                    SyncListUpdater_.listBox1.Update();
                                    goto finder;
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }
                        catch (Exception)
                        {

                            goto finder;

                        }
                    }
                }

            //System.Timers.Timer timer1 = new System.Timers.Timer
            //{
            //    Interval = 2000
            //};
            //timer1.Enabled = true;
            //timer1.Tick += new System.EventHandler(OnTimerEvent);

            //CountdownTimer timer = new CountdownTimer();
            //timer.Show();
            //OnTimeEvent();



            finish_:;

                SyncListUpdater_.Close();
                //timer_();

                TransactWithCentralOptions transact = new TransactWithCentralOptions();
                SynchronizeWithCentralOptions synch = new SynchronizeWithCentralOptions();
                //synch.Comment = "Autosaved by the API at " + DateTime.Now;
                RelinquishOptions relinquishOptions = new RelinquishOptions(true);
                relinquishOptions.CheckedOutElements = true;
                synch.SetRelinquishOptions(relinquishOptions);

                //uiApp.Application.WriteJournalComment("AutoSave To Central", true);
                doc.SynchronizeWithCentral(transact, synch);


                try
                {
                    string Sync_Manager = @"T:\Lopez\Sync_Manager.xlsx";
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
                    {
                        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);
                        if (sheet.Cells[1, 2].Value != null)
                        {
                            if (sheet.Cells[1, 2].Value.ToString() == user)
                            {
                                sheet.DeleteRow(1, 1);
                                package.Save();
                            }
                        }

                    }
                }
                catch (Exception)
                {
                    //MessageBox.Show("Excel file not found", "");
                    //return;
                }



                return Autodesk.Revit.UI.Result.Succeeded;
            }
            else
            {
                TaskDialog.Show(doc.PathName.ToString(), "This tool is active only for =" + "RHR_BUILDING_A22");
                return Autodesk.Revit.UI.Result.Cancelled;
            }
            
        }
    }
    //public static class myCommand
    //{
    //    static DateTime lastSaveTime;
    //    public static void myDocumentSaving(object sender, DocumentSynchronizingWithCentralEventArgs args)
    //    {
           
    //        double reset = 10000000;
    //        SyncListUpdater SyncListUpdater_ = new SyncListUpdater();

    //        string user = args.Document.Application.Username;

    //        string s = "People syncing:" + "\n";
    //        try
    //        {
    //            string Sync_Manager = @"T:\Lopez\Sync_Manager.xlsx";
    //            using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
    //            {
    //                ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);
    //                var Time_ = DateTime.Now;
    //                //---------------------------------------------------------
    //                for (int row = 1; row < 20; row++)
    //                {
    //                    if (sheet.Cells[row, 1].Value == null)
    //                    {
    //                        break;
    //                    }
    //                    if (sheet.Cells[row, 1].Value != null)
    //                    {
    //                        var Value1 = sheet.Cells[row, 1].Value;
    //                        var Value2 = sheet.Cells[row, 2].Value;
    //                        //s += Value1 + " + " + Value2.ToString() + "\n";
    //                        SyncListUpdater_.listBox1.Items.Add(Value1 + " + " + Value2.ToString() + "\n");
    //                    }

    //                }

    //                SyncListUpdater_.Show();
                        
    //                //TaskDialog.Show("Current Sync List ", s);
    //                //MessageBox.Show("Current Sync List ", "");


    //                //---------------------------------------------------------
    //                if (sheet.Cells[1, 1].Value == null)
    //                {
    //                    sheet.Cells[1, 1].Value = Time_.ToString();
    //                    sheet.Cells[1, 2].Value = user;
    //                    package.Save();
    //                    goto finish_;
    //                }
    //                //---------------------------------------------------------
    //                if (sheet.Cells[1, 2].Value.ToString() != "Alex synced")
    //                {
    //                    for (int row = 1; row < 9999; row++)
    //                    {
    //                        var thisValue = sheet.Cells[row, 1].Value;

    //                        if (thisValue == null)
    //                        {
    //                            sheet.Cells[row, 1].Value = Time_.ToString();
    //                            sheet.Cells[row, 2].Value = user;
    //                            package.Save();
    //                            goto finder;
    //                        }
    //                        else
    //                        {
                              
                               
    //                        }
    //                    }
    //                }
    //                //---------------------------------------------------------
                    
    //            //---------------------------------------------------------
    //            }
    //        }
    //        catch (Exception)
    //        {
    //            MessageBox.Show("Excel file not found", "");
    //            //return;
    //            args.Cancel();
    //        }
    //        finder:
    //        for (int i = 0; i < reset; i++)
    //        {
    //            if (i == 9999999)
    //            {
    //                try
    //                {
    //                    SyncListUpdater_.listBox1.Items.Clear();
    //                    string Sync_Manager = @"T:\Lopez\Sync_Manager.xlsx";
    //                    using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
    //                    {
    //                        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

    //                        for (int row = 1; row < 20; row++)
    //                        {
    //                            if (sheet.Cells[row, 1].Value == null)
    //                            {
    //                                break;
    //                            }
    //                            if (sheet.Cells[row, 1].Value != null)
    //                            {
    //                                var Value1 = sheet.Cells[row, 1].Value;
    //                                var Value2 = sheet.Cells[row, 2].Value;

    //                                SyncListUpdater_.listBox1.Items.Add(Value1 + " + " + Value2.ToString());
    //                            }

    //                        }

    //                        if (sheet.Cells[1, 1].Value != null && sheet.Cells[1, 2].Value.ToString() != user)
    //                        {
    //                            SyncListUpdater_.listBox1.Refresh();
    //                            SyncListUpdater_.listBox1.Update();
    //                            goto finder;
    //                        }
    //                        else
    //                        {
    //                            break;
    //                        }
    //                    }
    //                }
    //                catch (Exception)
    //                {

    //                    goto finder;

    //                }
    //            }
    //        }
    //        //CountdownTimer timer = new CountdownTimer();
    //        //timer.Show();
            
    //        finish_:;
    //        SyncListUpdater_.Close();
    //        //timer_();


         
    //    }

       
    //    public static void myDocumentSaved(object sender, DocumentSynchronizedWithCentralEventArgs args)
    //    {

    //        string user = args.Document.Application.Username;
    //        SyncListUpdater SyncListUpdater_ = new SyncListUpdater();
    //        try
    //        {
    //            string Sync_Manager = @"T:\Lopez\Sync_Manager.xlsx";
    //            using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
    //            {
    //                ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);
    //                if (sheet.Cells[1, 2].Value.ToString() == user)
    //                {
    //                    sheet.DeleteRow(1, 1);
    //                    package.Save();
    //                }
    //            }
    //        }
    //        catch (Exception)
    //        {
    //            //MessageBox.Show("Excel file not found", "");
    //            //return;
    //        }
    //    }
       
    //    public static void idleUpdate(object sender, IdlingEventArgs e)
    //    {
    //        // set an initial value for the last saved time
    //        if (lastSaveTime == DateTime.MinValue)
    //            lastSaveTime = DateTime.Now;
    //        DateTime now = DateTime.Now;

    //        TimeSpan elapsedTime = now.Subtract(lastSaveTime);
    //        double minutes = elapsedTime.Minutes;
    //        UIApplication uiApp = sender as UIApplication;
    //        uiApp.Application.WriteJournalComment("Idle check. Elapsed time = " + minutes, true);
    //        if (minutes < 1)
    //            return;

    //        Document doc = uiApp.ActiveUIDocument.Document;
    //        if (!doc.IsWorkshared)
    //            return;

    //        TransactWithCentralOptions transact = new TransactWithCentralOptions();
    //        SynchronizeWithCentralOptions synch = new SynchronizeWithCentralOptions();
    //        synch.Comment = "Autosaved by the API at " + DateTime.Now;
    //        RelinquishOptions relinquishOptions = new RelinquishOptions(true);
    //        relinquishOptions.CheckedOutElements = true;
    //        synch.SetRelinquishOptions(relinquishOptions);

    //        uiApp.Application.WriteJournalComment("AutoSave To Central", true);
    //        doc.SynchronizeWithCentral(transact, synch);
    //        lastSaveTime = DateTime.Now;
            
    //    }
        
    //}

    class ribbonUI : IExternalApplication
    {
        
        public static FailureDefinitionId failureDefinitionId = new FailureDefinitionId(new Guid("E7BC1F65-781D-48E8-AF37-1136B62913F5"));
       
        public Autodesk.Revit.UI.Result OnStartup(UIControlledApplication application)
        {
            string appdataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folderPath = Path.Combine(appdataFolder, @"Autodesk\Revit\Addins\2022\ES_commands\img");
            string dll = Assembly.GetExecutingAssembly().Location;
            string myRibbon_1 = "Alex Tools";



            //application.Idling += new EventHandler<IdlingEventArgs>(myCommand.idleUpdate);
            //return Result.Succeeded;

            
            //application.ControlledApplication.DocumentSynchronizingWithCentral += new EventHandler<DocumentSynchronizingWithCentralEventArgs>(myCommand.myDocumentSaving);
            //application.ControlledApplication.DocumentSynchronizedWithCentral += new EventHandler<DocumentSynchronizedWithCentralEventArgs>(myCommand.myDocumentSaved);
            


            application.CreateRibbonTab(myRibbon_1);
            RibbonPanel panel_1_a = application.CreateRibbonPanel(myRibbon_1, "Views / Sheets Tools");
            RibbonPanel panel_2_a = application.CreateRibbonPanel(myRibbon_1, "Working tools");
            RibbonPanel panel_3_a = application.CreateRibbonPanel(myRibbon_1, "Analisys");
            //RibbonPanel panel_4_a = application.CreateRibbonPanel(myRibbon_1, "Legend views");
            //RibbonPanel panel_5_a = application.CreateRibbonPanel(myRibbon_1, "Version");
            RibbonPanel panel_6_a = application.CreateRibbonPanel(myRibbon_1, "Rhino Tools");
            RibbonPanel panel_7_a = application.CreateRibbonPanel(myRibbon_1, "Lines / planes creation");
            PushButtonData b1 = new PushButtonData("ButtonNameA", "Create Sheet/s views", dll, "BoostYourBIM.PlaceView_CrateSheet");
            b1.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "createsheet_32.png"), UriKind.Absolute));
            PushButtonData b2 = new PushButtonData("Create Multiple Sheets", "Create Multiple Empty Sheets", dll, "BoostYourBIM.CreatMultipleSheet");
            b2.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "createsheet_32.png"), UriKind.Absolute));
            PushButtonData b3 = new PushButtonData("Duplicate one Sheet", "Duplicate one Sheet", dll, "BoostYourBIM.Duplicate_0ne_sheet");
            b3.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "storeselect.png"), UriKind.Absolute));
            PushButtonData b4 = new PushButtonData("Copy Schedule", "Schedule/legend placement", dll, "BoostYourBIM.copy_schedule");
            b4.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "Form.png"), UriKind.Absolute));
            PushButtonData b5 = new PushButtonData("Create Views", "Create Views", dll, "BoostYourBIM.Element_Elevations");
            b5.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "roomSchedule_32.png"), UriKind.Absolute));
            PushButtonData b6 = new PushButtonData("Room Elevations", "Room Elevations", dll, "BoostYourBIM.RoomElevations");
            b6.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "roomSchedule_32.png"), UriKind.Absolute));
            SplitButtonData sb1 = new SplitButtonData("Sheet creating option", "Options to create Sheets sets");
            SplitButton sb = panel_1_a.AddItem(sb1) as SplitButton;
            sb.IsSynchronizedWithCurrentItem = false;
            sb.ItemText = "hola";
            sb.AddPushButton(b1);
            sb.AddPushButton(b2);
            sb.AddPushButton(b3);
            sb.AddPushButton(b4);
            sb.AddPushButton(b5);
            sb.AddPushButton(b6);
            PushButton aa_1 = (PushButton)panel_1_a.AddItem(new PushButtonData("Register_", "Register_", dll, "BoostYourBIM.Register_"));
            aa_1.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "Erase.png"), UriKind.Absolute));
          
            PushButton a_1 = (PushButton)panel_1_a.AddItem(new PushButtonData("Delete All Views", "Delete All Views", dll, "BoostYourBIM.DeleteAllViews"));
            a_1.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "Erase.png"), UriKind.Absolute));
            a_1.ToolTip = "This tool will delete all views and sheets in the project living only one (the view named Home)";
            a_1.LongDescription = "...";
            PushButton a_2_1 = (PushButton)panel_2_a.AddItem(new PushButtonData("ReNumbering", "ReNumbering", dll, "BoostYourBIM.ReNumbering"));
            a_2_1.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rename.png"), UriKind.Absolute));
            a_2_1.ToolTip = "Renumbers a secuence of revit elements (Viewports, Doors/room number, Grids) giving that the parameter does not contain a text character";
            a_2_1.LongDescription = "...";
            PushButton a_2 = (PushButton)panel_2_a.AddItem(new PushButtonData("Delete Level ", "Delete Level ", dll, "BoostYourBIM.DeleteLevel"));
            a_2.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "deletelevel_32.png"), UriKind.Absolute));
            a_2.ToolTip = "Reallocated hosted element from one level to another so elements are not lose when level is deleted";
            PushButton a_3 = (PushButton)panel_2_a.AddItem(new PushButtonData("Remove paint", "Remove paint", dll, "BoostYourBIM.remove_paint"));
            a_3.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "remove_paint.png"), UriKind.Absolute));
            a_3.ToolTip = "Removes paint from selected set of walls";

            PushButton a_3_2 = (PushButton)panel_3_a.AddItem(new PushButtonData("Wall Elevation", "Wall Elevation", dll, "BoostYourBIM.Wall_Elevation"));
            a_3_2.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "iconfinder_Angle_131818.png"), UriKind.Absolute));
            a_3_2.ToolTip = "...";

            PushButton a_3_3 = (PushButton)panel_3_a.AddItem(new PushButtonData("Rot_Wal_Angle_To_Grid", "Rot_Wal_Angle_To_Grid", dll, "BoostYourBIM.Rot_Wal_Angle_To_Grid"));
            a_3_3.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "iconfinder_Angle_131818.png"), UriKind.Absolute));
            a_3_3.ToolTip = "..."; 
            PushButton a_3_4 = (PushButton)panel_3_a.AddItem(new PushButtonData("Set_Annotation_Crop", "Set_Annotation_Crop", dll, "BoostYourBIM.Set_Annotation_Crop"));
            a_3_4.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "iconfinder_Angle_131818.png"), UriKind.Absolute));
            a_3_4.ToolTip = "...";


            PushButton a_12 = (PushButton)panel_2_a.AddItem(new PushButtonData("Text to Uppercase", "Text to Uppercase", dll, "BoostYourBIM.text_upper"));
            a_12.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "upper_.png"), UriKind.Absolute));
            a_12.ToolTip = "";
            PushButton a_3_1 = (PushButton)panel_3_a.AddItem(new PushButtonData("Wall Angle", "Wall Angle", dll, "BoostYourBIM.Wall_Angle_to"));
            a_3_1.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "iconfinder_Angle_131818.png"), UriKind.Absolute));
            a_3_1.ToolTip = "...";
            PushButton a_8 = (PushButton)panel_2_a.AddItem(new PushButtonData("Detail Line select", "Detail Line select", dll, "BoostYourBIM.select_detailline"));
            a_8.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "Selection.png"), UriKind.Absolute));
            a_8.ToolTip = "After selecting a Detail line the tool will select all intances of the type in the view or the project";
            PushButton a_13 = (PushButton)panel_2_a.AddItem(new PushButtonData("Total Lenght", "Total Lenght", dll, "BoostYourBIM.TotalLenght"));
            a_13.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "TotalLenght.png"), UriKind.Absolute));
            a_13.ToolTip = "Retrives the total lenght of line base elements ";
            PushButton a_21 = (PushButton)panel_2_a.AddItem(new PushButtonData("Transparensy", "Transparency", dll, "BoostYourBIM.isolate"));
            a_21.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "solid icon.png"), UriKind.Absolute));
            a_21.ToolTip = "";
            PushButton a_22 = (PushButton)panel_2_a.AddItem(new PushButtonData("Isolate category", "Isolate category", dll, "BoostYourBIM.isolate_category"));
            a_22.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "solid icon.png"), UriKind.Absolute));
            a_22.ToolTip = "";
            PushButton a_23 = (PushButton)panel_2_a.AddItem(new PushButtonData("Clean view", "Clean view", dll, "BoostYourBIM.Clean_view"));
            a_23.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "solid icon.png"), UriKind.Absolute));
            a_23.ToolTip = "";
            PushButtonData D_1 = new PushButtonData("Line", "Line", dll, "BoostYourBIM.make_line");
            D_1.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));
            PushButton a_16 = (PushButton)panel_3_a.AddItem(new PushButtonData("Select NB wall", "non bounding wall", dll, "BoostYourBIM.Wall_Bounding_room"));
            a_16.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "nonroombounding.png"), UriKind.Absolute));
            a_16.ToolTip = "...";
            PushButtonData D_2 = new PushButtonData("Plane", "Plane", dll, "BoostYourBIM.line_point_plane");
            D_2.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));
            PushButtonData D_3 = new PushButtonData("Line distance", "Line from surface", dll, "BoostYourBIM.make_line_from_surface_normal");
            D_3.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));
            PushButtonData D_5 = new PushButtonData("Loft geometry", "Loft geometry", dll, "BoostYourBIM.loft");
            D_5.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));
            PushButtonData D_6 = new PushButtonData("Duct from line", "Duct geometry", dll, "BoostYourBIM.make_duck_by_line");
            D_6.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));
            PushButtonData D_7 = new PushButtonData("Create flex ducts", "Create flex ducts", dll, "BoostYourBIM.Create_flex_ducts_from_line");
            D_7.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));
            PushButtonData D_8 = new PushButtonData("Create ducts fitting", "Create ducts fitting", dll, "BoostYourBIM.duct_elbo");
            D_8.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));
            PushButtonData D_9 = new PushButtonData("Closest_point", "Closest_point", dll, "BoostYourBIM.Closest_point_2Lines");
            D_9.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));
            PushButton D_11 = (PushButton)panel_7_a.AddItem(new PushButtonData("Pipe 3D Lines", "Pipe 3D Lines", dll, "BoostYourBIM.Pipe_lines"));
            D_11.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "pipelinedgw.png"), UriKind.Absolute));
            D_11.ToolTip = "...";
            PushButton D_12 = (PushButton)panel_7_a.AddItem(new PushButtonData("PENO Point", "PENO Point", dll, "BoostYourBIM.Pipe_slab_intersetion"));
            D_12.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "PENOPT.png"), UriKind.Absolute));
            D_12.ToolTip = "...";
            PushButton D_13 = (PushButton)panel_7_a.AddItem(new PushButtonData("Closest grid to Peno", "Closest grid to Peno", dll, "BoostYourBIM.Closestpt_togrids"));
            D_13.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "closestgrid.png"), UriKind.Absolute));
            D_13.ToolTip = "...";
            PushButton D_14 = (PushButton)panel_7_a.AddItem(new PushButtonData("Pipe from lines", "Pipe from lines", dll, "BoostYourBIM.make_pipe_by_line"));
            D_14.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "Makepipefromline.png"), UriKind.Absolute));
            D_14.ToolTip = "...";
            PushButton D_15 = (PushButton)panel_7_a.AddItem(new PushButtonData("Plane by line & point", "Plane by line & point", dll, "BoostYourBIM.line_point_plane"));
            D_15.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "Planebylinepoint.png"), UriKind.Absolute));
            D_15.ToolTip = "...";
            PushButton D_16 = (PushButton)panel_7_a.AddItem(new PushButtonData(" Pipe lines to 12D ", " Pipe lines to 12D ", dll, "BoostYourBIM.Pipe_lines_excel"));
            D_16.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "pipeline12d.png"), UriKind.Absolute));
            D_16.ToolTip = "...";
           PushButton D_17 = (PushButton)panel_7_a.AddItem(new PushButtonData(" Isolate Pipe type ", " Isolate Pipe type ", dll, "BoostYourBIM.Isolate_Pipe"));
           D_17.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "selectpipe.png"), UriKind.Absolute));
           D_17.ToolTip = "...";
            SplitButtonData sb4 = new SplitButtonData("Sheet creating option", "Options to create Sheets sets");
            SplitButton sb_4 = panel_7_a.AddItem(sb4) as SplitButton;
            sb_4.IsSynchronizedWithCurrentItem = false;
            sb_4.ItemText = "hola";
            sb_4.IsSynchronizedWithCurrentItem = false;
            sb_4.AddPushButton(D_1);
            sb_4.AddPushButton(D_2);
            sb_4.AddPushButton(D_3);
            sb_4.AddPushButton(D_5);
            sb_4.AddPushButton(D_6);
            sb_4.AddPushButton(D_7);
            sb_4.AddPushButton(D_8);
            sb_4.AddPushButton(D_9);
            PushButton a_18 = (PushButton)panel_3_a.AddItem(new PushButtonData("Floor from topo", "Floor from topo", dll, "BoostYourBIM.revision_on_project"));
            a_18.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "topo.png"), UriKind.Absolute));
            a_18.ToolTip = "...";
            PushButtonData C_11 = new PushButtonData("Family Geometry Search", "Family GeoGeometry Search", dll, "BoostYourBIM.Rhino_access_faces");
            C_11.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));
            PushButtonData C_12 = new PushButtonData("Rhino Object to Revit", "Rhino Object to Revit", dll, "BoostYourBIM.Reading_from_rhino");
            C_12.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32_copy.png"), UriKind.Absolute));
            PushButtonData C_13 = new PushButtonData("Group Geometry Exporter", "Group Geometry Exporter", dll, "BoostYourBIM.Rhino_access");
            C_13.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));
            PushButtonData C_14 = new PushButtonData("Rhino pnts to Revit topo", "Rhino pnts to Revit topo", dll, "BoostYourBIM.rhino_points_to_revit_topo");
            C_14.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32_copy.png"), UriKind.Absolute));
            PushButtonData C_15 = new PushButtonData("Rhino lns to Revit lines", "Rhino lns to Revit lines", dll, "BoostYourBIM.rhino_lns_to_revit_lns");
            C_15.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));
            PushButtonData C_16 = new PushButtonData("Rhino lns to Frame", "Rhino lns to Frame", dll, "BoostYourBIM.rhino_PTS_to_revitPTS");
            C_16.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));
            PushButtonData C_17 = new PushButtonData("Create_solid_rooms", "Create_solid_rooms", dll, "BoostYourBIM.Create_solid_rooms");
            C_17.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));
            PushButtonData C_18 = new PushButtonData("Import Rhino clash points", "Import Rhino clash points", dll, "BoostYourBIM.rhino_pt_to_revit_clash");
            C_18.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));
            PushButtonData C_19 = new PushButtonData("Import Excel clash points", "Import Excel clash points", dll, "BoostYourBIM.excel_to_revit_clash");
            C_19.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));
              SplitButtonData sb3 = new SplitButtonData("Sheet creating option", "Options to create Sheets sets");
            SplitButton sb_3 = panel_6_a.AddItem(sb3) as SplitButton;
            sb_3.IsSynchronizedWithCurrentItem = false;
            sb_3.ItemText = "hola";
            sb_3.IsSynchronizedWithCurrentItem = false;
            sb_3.AddPushButton(C_11);
            sb_3.AddPushButton(C_12);
            sb_3.AddPushButton(C_13);
            sb_3.AddPushButton(C_14);
            sb_3.AddPushButton(C_15);
            sb_3.AddPushButton(C_16);
            sb_3.AddPushButton(C_17);
            sb_3.AddPushButton(C_18);
            sb_3.AddPushButton(C_19);
            PushButtonData b_8 = new PushButtonData("Door Views", "Door Views", dll, "BoostYourBIM.Door_Section");
            b_8.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "doorSchedule.png"), UriKind.Absolute));
            b_8.ToolTip = "a section view is created of elements containing an identifier in the parameter field (Schedule Identifier)";
            PushButtonData b_9 = new PushButtonData("Room Views", "Room Views", dll, "BoostYourBIM.RoomElevations");
            b_9.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "roomSchedule_32.png"), UriKind.Absolute));
            b_9.ToolTip = "Section, floor/ceiling plans, Int elevation and Isometric views can be created of one room containing an identifier in the parameter field (Schedule Identifier)";
            PushButtonData b_10 = new PushButtonData("Wall Views", "Wall Views", dll, "BoostYourBIM.CreateSchedule");
            b_10.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "cutingwall_32.png"), UriKind.Absolute));
            b_10.ToolTip = "a section view is created of elements containing an identifier in the parameter field (Schedule Identifier)";
            PushButtonData b_11 = new PushButtonData("Window Views", "Window Views", dll, "BoostYourBIM.WindowSection");
            b_11.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "windowSchedule_32.png"), UriKind.Absolute));
            b_11.ToolTip = "a section view is created of elements containing an identifier in the parameter field (Schedule Identifier)";

            //PushButtonData b5 = new PushButtonData("Delete all Sheets", "Delete all Sheets", dll, "BoostYourBIM.DeleteAllSheets");
            //b5.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "Erase.png"), UriKind.Absolute));

            //adWin.ComponentManager.UIElementActivated += new EventHandler<adWin.UIElementActivatedEventArgs>(ComponentManager_UIElementActivated);

            try
            {
                foreach (Autodesk.Windows.RibbonTab tab in Autodesk.Windows.ComponentManager.Ribbon.Tabs)
                {
                    if (tab.Title == "Insert")
                    {
                        tab.IsVisible = false;
                    }
                }

                adWin.RibbonControl ribbon = adWin.ComponentManager.Ribbon;

                //ImageSource imgbg = new BitmapImage(new Uri(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                //    "gradient.png"), UriKind.Relative));



                ImageBrush picBrush = new ImageBrush();
                //picBrush.ImageSource = imgbg;
                picBrush.AlignmentX = AlignmentX.Left;
                picBrush.AlignmentY = AlignmentY.Top;
                picBrush.Stretch = Stretch.None;
                picBrush.TileMode = TileMode.FlipXY;

               

                LinearGradientBrush gradientBrush = new LinearGradientBrush();

                gradientBrush.StartPoint = new System.Windows.Point(0, 0);

                gradientBrush.EndPoint = new System.Windows.Point(0, 1);

                gradientBrush.GradientStops.Add(new GradientStop(Colors.White, 0.0));

                gradientBrush.GradientStops.Add(new GradientStop(Colors.Orange, 0.95));



                ribbon.FontFamily = new System.Windows.Media.FontFamily("Bauhaus 93");
                ribbon.Opacity = 70;
                ribbon.FontSize = 10;
                //ribbon.Background = picBrush;


                foreach (adWin.RibbonTab tab in ribbon.Tabs)
                {

                    string name = tab.AutomationName;

                    if (name == "Insert")
                    {
                        foreach (adWin.RibbonPanel panel in tab.Panels)
                        {

                            adWin.RibbonItemCollection items = panel.Source.Items;
                            foreach (var item in items)
                            {
                                string name_ = item.Id;
                                if (name_ == "ID_FILE_IMPORT")
                                {
                                    item.IsEnabled = false;
                                }
                            }


                        }
                    }
                    foreach (adWin.RibbonPanel panel in tab.Panels)
                    {
                        string name1 = panel.AutomationName;
                        panel.CustomPanelTitleBarBackground = gradientBrush;
                        panel.CustomPanelBackground = picBrush;
                    }
                }
                RibbonItemEventArgs jk = new RibbonItemEventArgs();
                List<RibbonPanel> ribbons = jk.Application.GetRibbonPanels();

            }
            catch (Exception ex)
            {
                winform.MessageBox.Show(
                  ex.StackTrace + "\r\n" + ex.InnerException,
                  "Error", winform.MessageBoxButtons.OK);

                return Result.Failed;
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }

        public Autodesk.Revit.UI.Result OnShutdown(UIControlledApplication application)
        {
            return Autodesk.Revit.UI.Result.Succeeded;
        }

    }
}

