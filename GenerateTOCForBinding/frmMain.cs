using Microsoft.Office.Tools;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;

namespace GenerateTOCForBinding
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
            
            button2_Click(this, EventArgs.Empty);
            button1_Click(this, EventArgs.Empty);


            //Dispose();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try {
            //////Microsoft.Office.Tools.Excel.Workbook theWorkbook = Microsoft.Office.Tools.Excel.ex .ExcelObj.Workbooks.Open(
            //////     this.txtXmlPath.Text, 0, true, 5,
            //////      "", "", true, Microsoft.Office.Tools.Excel.xl.xlWindows, "\t", false, false,
            //////      0, true);

            //////Excel.Sheets sheets = theWorkbook.Worksheets;
            //////Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);

            //////int i=1;
            //////string[] strArray;
            //////System.Array myvalues;

            //////do {
            //////    Excel.Range range = worksheet.get_Range("A" + i.ToString(), "O" + i.ToString());
            //////    myvalues = (System.Array)range.Cells.Value;
            //////    strArray = ConvertToStringArray(myvalues);
            //////} while (strArray[0].Length > 0);
            
            
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.txtXlsPath.Text + @";Extended Properties=""Excel 8.0;HDR=YES;""";

            DbProviderFactory factory = DbProviderFactories.GetFactory("System.Data.OleDb");

            using (DbConnection connection = factory.CreateConnection())
            {
                connection.ConnectionString = connectionString;

                ////using (DbCommand command = connection.CreateCommand())
                ////{
                ////    command.CommandText = "INSERT INTO [Cities$]
                ////     (ID, City, State) VALUES(4,\"Tampa\",\"Florida\")";

                connection.Open();

                ////    command.ExecuteNonQuery();
                ////}
                
                System.Data.Common.DbCommand cmd = connection.CreateCommand();
                cmd.CommandText = "SELECT * FROM [volumes$]";
                System.Data.Common.DbDataReader drdr = cmd.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(drdr);
                drdr.Dispose();
                cmd.Dispose();

                XmlDocument docXml = new XmlDocument();
                XmlNode nodeRoot = docXml.AppendChild(docXml.CreateNode(XmlNodeType.Element, "root", docXml.NamespaceURI));
                XmlNode nodeVolumes = nodeRoot.AppendChild(docXml.CreateNode(XmlNodeType.Element, "volumes", docXml.NamespaceURI));
                XmlNode nodeVolume;

                XmlNode nodeColumns = nodeRoot.AppendChild(docXml.CreateNode(XmlNodeType.Element, "columns", docXml.NamespaceURI));
                foreach (DataColumn dc in dt.Columns) {
                    XmlNode nodeColumn = nodeColumns.AppendChild(docXml.CreateNode(XmlNodeType.Element, "column", docXml.NamespaceURI));
                    XmlAttribute attFldName = nodeColumn.Attributes.Append((XmlAttribute)docXml.CreateNode(XmlNodeType.Attribute, "FldName", docXml.NamespaceURI));
                    attFldName.InnerText = dc.ColumnName;
                }

                foreach(DataRow dr in dt.Rows) {

                    nodeVolume = nodeVolumes.SelectSingleNode("./volume[@volume_code = '" + dr["Volume"].ToString() + "']");
                    if(nodeVolume == null) {
                        nodeVolume = nodeVolumes.AppendChild(docXml.CreateNode(XmlNodeType.Element, "volume", docXml.NamespaceURI));
                        XmlAttribute attVolumeCode = nodeVolume.Attributes.Append((XmlAttribute) docXml.CreateNode(XmlNodeType.Attribute, "volume_code", docXml.NamespaceURI));
                        attVolumeCode.InnerText = dr["volume"].ToString();
                    }

                    XmlNode nodeIssue = nodeVolume.AppendChild(docXml.CreateNode(XmlNodeType.Element, "issue", docXml.NamespaceURI));

                    foreach (DataColumn dc in dt.Columns)
                    {
                        if(string.Compare(dc.ColumnName, "volume", true) != 0) {
                            XmlNode nodeFld;
                            nodeFld = nodeIssue.AppendChild(docXml.CreateNode(XmlNodeType.Element, "Fld", docXml.NamespaceURI));
                            XmlAttribute attFldName = nodeFld.Attributes.Append( (XmlAttribute)(docXml.CreateNode(XmlNodeType.Attribute, "Name", docXml.NamespaceURI)));
                            attFldName.InnerText = dc.ColumnName;//.Replace(" ", "_");
                            nodeFld.InnerText = dr[dc].ToString();
                        }
                    }
                }

                string strXmlFilePath = this.txtXlsPath.Text.Substring(0, this.txtXlsPath.Text.LastIndexOf(".xls")) + ".xml";

                docXml.Save(strXmlFilePath);
                this.txtXmlPath.Text = strXmlFilePath;
            }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try {
            //Create a new XslTransform object.
            XslTransform xslt = new XslTransform();

            //Load the stylesheet.
            xslt.Load(@"C:\dev\jDev\fop\Defenders_TOC\XSLTFile2.xslt");

            //Create a new XPathDocument and load the XML data to be transformed.
            XPathDocument mydata = new XPathDocument(this.txtXmlPath.Text);

            //Create an XmlTextWriter which outputs to the console.
            XmlWriter writer = new XmlTextWriter(this.txtXmlPath.Text.Substring(0, this.txtXmlPath.Text.LastIndexOf(".xml")) + ".html", Encoding.UTF8);

            //Transform the data and send the output to the console.
            xslt.Transform(mydata, null, writer, null);

            writer.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
