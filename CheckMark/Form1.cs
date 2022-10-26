using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures;
//using Tekla.Structures.Model.UI;
//using Tekla.Structures.Datatype;
using Tekla.Structures.Drawing;
using Tekla.Structures.Model.Operations;
//using Tekla.Structures.Model;
using tsm = Tekla.Structures.Model;

namespace CheckMark
{
    public partial class Form1 : Form
    {
        tsm.Model _model = new tsm.Model();        
        DrawingHandler _drawinghandler = new DrawingHandler();

        List<string> _allmarksMK = new List<string>();
        List<string> _allmarksID = new List<string>();
        List<string> _planmarksMK = new List<string>();
        List<string> _planmarksID = new List<string>();
        List<string> missingmarksMK = new List<string>();
        List<string> missingmarksID = new List<string>();
        ArrayList ObjectsToSelect = new ArrayList();
        ViewBase view = null;
        Tekla.Structures.Model.Operations.Operation.ProgressBar progress = new Tekla.Structures.Model.Operations.Operation.ProgressBar();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                tsm.ModelInfo Info = _model.GetInfo();
                string modelPath = Info.ModelPath;
                List<Task> _tasks = new List<Task>();

                if (!_model.GetConnectionStatus() || !_drawinghandler.GetConnectionStatus())
                    return;

                var picker = _drawinghandler.GetPicker();

                var prompt = "Pick View";
                DrawingObject p = null;
                picker.PickObject(prompt, out p, out view);


                bool displayResult = progress.Display(100, "MissingMark", "Program is running", "cancel..", " ");

                var types = new[] { typeof(Part), typeof(Mark) };

                var drawingObjects = view.GetObjects(types);

                while (drawingObjects.MoveNext())
                {
                    if (progress.Canceled())
                        break;

                    if (drawingObjects.Current is Part part)
                    {
                        _tasks.Add(Task.Factory.StartNew(() =>
                        {
                            CheckPart(part, _allmarksMK, _allmarksID);
                        }));
                        continue;
                    }

                    if (drawingObjects.Current is Mark mark)
                    {
                        _tasks.Add(Task.Factory.StartNew(() =>
                        {
                            var relatedObjects = mark.GetRelatedObjects(new[] { typeof(Part) });

                            while (relatedObjects.MoveNext())
                            {
                                if (relatedObjects.Current is Part partFromMark && !mark.Hideable.IsHidden && mark.Attributes.Content.Count > 0)
                                    CheckPart(partFromMark, _planmarksMK, _planmarksID);
                            }
                        }));
                    }
                }

                Task.WaitAll(_tasks.ToArray());

                //Ищу каких марок не хватает на виде (Это работает при условии, что allmarks более полный чем planmarks)
                missingmarksMK = differenceInLists(missingmarksMK, _allmarksMK, _planmarksMK);
                missingmarksID = differenceInLists(missingmarksID, _allmarksID, _planmarksID);

                progress.Close();

                string fileName = $@"{modelPath}\attributes\Missing marks.dsf";
                CreateFilter(fileName, missingmarksID);
                ObjectsToSelect = SelectObjects(missingmarksID);
                CreateReport(modelPath, ObjectsToSelect);


            }
            catch (Exception ex)
            {
                MessageBox.Show(new Form { TopMost = true }, ex.ToString()/*"Something goes wrong :("*/);
                this.Close();
            }                   

        }

        public void CheckPart(Part part, List<string> marks, List<string> marksId)
        {
            var partModel = _model.SelectModelObject(part.ModelIdentifier) as tsm.Part;

            if (partModel == null)
                return;

            var assembly = partModel.GetAssembly();
            var assemblyType = assembly.GetAssemblyType();

            if (assemblyType == tsm.Assembly.AssemblyTypeEnum.STEEL_ASSEMBLY)
            {
                var checkPositions = CheckPositions(assembly, partModel);

                if (checkPositions)
                {
                    marks.Add(partModel.GetPartMark());
                    marksId.Add("id" + partModel.Identifier.GUID);
                    progress.SetProgress(partModel.GetPartMark(), 70);
                }
            }
        }

        public bool CheckPositions(tsm.Assembly assembly, tsm.Part partModel)
        {
            var partposition = "";
            partModel.GetReportProperty("PART_POS", ref partposition);
            var assposition = "";
            assembly.GetReportProperty("ASSEMBLY_POS", ref assposition);

            return partposition == assposition;
        }

        public string displayMembers(List<string> allmarks)
        {
            return string.Join(Environment.NewLine, allmarks.ToArray());
        }

        public void CreateFilter(string fileName, List<string> IDlist)
        {
            int IDlistLength = IDlist.Count;
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }
            using (StreamWriter sw = File.CreateText(fileName))
            {
                sw.WriteLine("TITLE_OBJECT_GROUP");
                sw.WriteLine("{");
                sw.WriteLine("    Version= 1.05 ");
                sw.WriteLine($"    Count= {IDlistLength} ");
                foreach (string s in missingmarksID)
                {
                    sw.WriteLine("    SECTION_OBJECT_GROUP");
                    sw.WriteLine("    {");
                    sw.WriteLine("        0 ");
                    sw.WriteLine("        1 ");
                    sw.WriteLine("        co_object ");
                    sw.WriteLine("        proGUID ");
                    sw.WriteLine("        albl_Guid ");
                    sw.WriteLine("        != ");
                    sw.WriteLine("        albl_DoesNotEqual ");
                    sw.WriteLine(s.ToUpper());
                    sw.WriteLine("        0 ");
                    sw.WriteLine("        && ");
                    sw.WriteLine("        }");
                }
                sw.WriteLine("    }");
            }
        }

        public void CreateReport(string modelPath,ArrayList ObjectsToSelect)
        {
            string subPath = $@"{modelPath}\Reports";

            bool exists = System.IO.Directory.Exists(subPath);

            if (!exists)
                System.IO.Directory.CreateDirectory(subPath);

            Tekla.Structures.Model.UI.ModelObjectSelector ModelSelector = new Tekla.Structures.Model.UI.ModelObjectSelector();
            ModelSelector.Select(ObjectsToSelect);
            Operation.CreateReportFromSelected("FF_MISSING_LIST", "FF_MISSING_LIST.txt", "MyTitle", "", "");
            textBox1.Text = File.ReadAllText($@"{modelPath}\Reports\FF_MISSING_LIST.txt");
        }

        public ArrayList SelectObjects(List<string> IDlist)
        {
            ArrayList ObjectsToSelect = new ArrayList();
            foreach (var item in IDlist)
            {
                string my_GUID = item.ToUpper();
                tsm.ModelObject myPart = _model.SelectModelObject(_model.GetIdentifierByGUID(my_GUID));
                if (myPart == null)
                    MessageBox.Show($"No objects");
                else
                    ObjectsToSelect.Add(myPart);
            }
            return ObjectsToSelect;
        }

        public List<string> differenceInLists(List<string> box, List<string> main, List<string> secondary)
        {
            box = main;
            foreach (var item in secondary)
            {
                box.Remove(item);
            }
            return box;
        }

        private void button1_Click(object sender, EventArgs e)
        {            

        }

        private void button2_Click(object sender, EventArgs e)
        {           
            Tekla.Structures.Model.UI.ModelObjectSelector ModelSelector = new Tekla.Structures.Model.UI.ModelObjectSelector();
            ModelSelector.Select(ObjectsToSelect);
            Operation.CreateReportFromSelected("FF_MISSING_LIST_GUID", "FF_MISSING_LIST_GUID.rpt", "MyTitle", "", "");
            Operation.DisplayReport("FF_MISSING_LIST_GUID.rpt");
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            TeklaStructures.Connect();

            var macrobuilder1 = new Tekla.Structures.MacroBuilder();
            
            macrobuilder1.ValueChange("main_frame", "gr_sel_all", "0");
            macrobuilder1.ValueChange("main_frame", "gr_sel_drawing_part", "1");
            macrobuilder1.Callback("acmd_display_gr_select_filter_dialog", "", "main_frame");
            macrobuilder1.ValueChange("diaSelDrawingObjectGroupDialogInstance", "get_menu", "Missing marks");
            macrobuilder1.PushButton("dia_pa_apply", "diaSelDrawingObjectGroupDialogInstance");

            macrobuilder1.Run();
        }
    }
}
