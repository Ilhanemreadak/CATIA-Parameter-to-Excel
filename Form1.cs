using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using INFITF;
using KnowledgewareTypeLib;
using MECMOD;
using PARTITF;
using ProductStructureTypeLib;
using Microsoft.Office.Interop.Excel;
using Parameter = KnowledgewareTypeLib.Parameter;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GetParametersToExcel
{
    public partial class Form1 : Form
    {
        INFITF.Application myCATIA;
        List<string> parameterList = new List<string>();
        HashSet<string> processedParts = new HashSet<string>();
        Dictionary<string, int> partCount = new Dictionary<string, int>();
        int paramCol = 3;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            myCATIA = (INFITF.Application)Marshal.GetActiveObject("CATIA.Application");
            Console.WriteLine("Checked. Connected");
        }

        private async void button1_Click_1(object sender, EventArgs e)
        {
            await Task.Run(() =>
            {
                Document oDoc = myCATIA.ActiveDocument;
                ProductDocument productDocument = (ProductDocument)oDoc;

                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                object Missing = Type.Missing;
                Workbook workbook = excel.Workbooks.Add(Missing);
                Worksheet sheet = (Worksheet)workbook.Sheets[1];

                partCount.Clear();
                processedParts.Clear();

                Invoke(new System.Action(() => treeView1.Nodes.Clear()));

                AddParametertoExcel(sheet);

                TreeNode rootNode = new TreeNode() { Text = productDocument.Product.get_Name() };
                Invoke(new System.Action(() => treeView1.Nodes.Add(rootNode)));

                GetParameters(productDocument.Product, rootNode, sheet);

                UpdatePartCounts(sheet);

                Invoke(new System.Action(() =>
                {
                    treeView1.ExpandAll();
                    excel.Visible = true;
                }));

                Console.WriteLine("Finish!");
            });
        }

        private void AddParametertoExcel(Worksheet sheet)
        {
            int row = 1;

            ((Range)sheet.Cells[row, 1]).Value2 = "Parça Adı";
            ((Range)sheet.Cells[row, 2]).Value2 = "Adet";
            row++;

            foreach (var paramName in parameterList)
            {
                ((Range)sheet.Cells[1, paramCol]).Value2 = paramName;
                paramCol++;
            }
        }

        public void GetParameters(Product oInstance, TreeNode treeNode, Worksheet sheet)
        {
            Products oInstances = oInstance.Products;
            int numberOfInstance = oInstances.Count;
            Product currentPart;

            KnowledgewareTypeLib.Parameter parameter;
            int numberOfParameters;

            int row = 2;

            if (treeNode == null) { return; }

            if (numberOfInstance == 0)
            {
                return;
            }
            else if (numberOfInstance > 0)
            {
                for (int i = 1; i <= numberOfInstance; i++)
                {
                    currentPart = oInstances.Item(i);
                    string partNumber = currentPart.get_PartNumber();

                    if (processedParts.Contains(partNumber))
                    {
                        partCount[partNumber]++;
                        continue;
                    }

                    processedParts.Add(partNumber);
                    partCount[partNumber] = 1;

                    TreeNode iNode = new TreeNode() { Text = currentPart.get_Name() };

                    Invoke(new System.Action(() => treeNode.Nodes.Add(iNode)));

                    numberOfParameters = currentPart.Parameters.Count;

                    for (int j = 1; j <= numberOfParameters; j++)
                    {
                        parameter = currentPart.Parameters.Item(j);

                        if (parameterList.Contains(parameter.get_Name().Split('\\').Last()))
                        {
                            TreeNode jNode = new TreeNode() { Text = parameter.get_Name() + " = " + parameter.ValueAsString() };

                            ((Range)sheet.Cells[row, 1]).Value2 = partNumber;
                            ((Range)sheet.Cells[row, 2]).Value2 = partCount[partNumber];
                            int index = parameterList.IndexOf(parameter.get_Name().Split('\\').Last());
                            ((Range)sheet.Cells[row, (index + 3)]).Value2 = parameter.ValueAsString();


                            Invoke(new System.Action(() => iNode.Nodes.Add(jNode)));
                        }
                    }
                    row++;
                    GetParameters(currentPart, iNode, sheet);
                }
            }
        }

        private void UpdatePartCounts(Worksheet sheet)
        {
            int lastRow = sheet.Cells[sheet.Rows.Count, 1].End[XlDirection.xlUp].Row;

            for (int row = 2; row <= lastRow; row++)
            {
                string partNumber = ((Range)sheet.Cells[row, 1]).Value2.ToString();

                if (processedParts.Contains(partNumber))
                {
                    ((Range)sheet.Cells[row, 2]).Value2 = partCount[partNumber];
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string parameterName = textBox1.Text;

            if (!string.IsNullOrEmpty(parameterName))
            {
                parameterList.Add(parameterName);

                listBox1.Items.Add(parameterName);

                textBox1.Clear();
            }
        }
    }
}
