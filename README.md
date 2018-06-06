using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Client;


namespace FirstWordAddIn
{
    public partial class ProposalTextPartsTaskPane : UserControl
    {
        private Dictionary<string, string> textPartIdsToContent { get; set; }
        private string linkToWebsite = "http://sharepoint2016/vorlagen-und-dokumente/";

        public ProposalTextPartsTaskPane()
        {
            InitializeComponent();
        }

        private void ProposalTextPartsTaskPane_Load(object sender, EventArgs e)
        {

        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            ClientContext remoteContext = new ClientContext(linkToWebsite);
            var remoteSite = remoteContext.Web;
            remoteContext.Load(remoteSite);
            remoteContext.ExecuteQuery();

            Microsoft.SharePoint.Client.List listTextParts =
              remoteSite.Lists.GetByTitle("Textbausteine");
            var query = new CamlQuery()
            { ViewXml = "<View></View>" };

            var textPartItems = listTextParts.GetItems(query);
            remoteContext.Load(textPartItems);
            remoteContext.ExecuteQuery();

            List<string> categories = new List<string>();
            foreach (var textPartItem in textPartItems)
            {
                if (!categories.Contains(textPartItem["Kategorie"].ToString()))
                    categories.Add(textPartItem["Kategorie"].ToString());
            }
            cbTextPartCategory.DataSource = categories;
        }

        private void cbTextPartCategory_SelectedIndexChanged(object sender, EventArgs e)
        {
            ClientContext remoteContext = new ClientContext(linkToWebsite);
            var remoteSite = remoteContext.Web;
            remoteContext.Load(remoteSite);
            remoteContext.ExecuteQuery();

            Microsoft.SharePoint.Client.List listTextParts = remoteSite.Lists.GetByTitle("Textbausteine");
            var query = new CamlQuery()
            {
                ViewXml = @"<View><Query><Where><Eq><FieldRef Name='Kategorie' /><Value Type='Choice'>" + cbTextPartCategory.Text +
                "</Value></Eq></Where></Query></View>"
            };

            var textPartItems = listTextParts.GetItems(query);
            remoteContext.Load(textPartItems);
            remoteContext.ExecuteQuery();

            textPartIdsToContent = new Dictionary<string, string>();
            dgTextParts.Rows.Clear();
            foreach (var textPartItem in textPartItems)
            {
                dgTextParts.Rows.Add(textPartItem["ID"].ToString(), textPartItem["Title"].ToString());
                textPartIdsToContent.Add(textPartItem["ID"].ToString(), textPartItem["Inhalt"].ToString());
            }
        }

        private void btnAddTextPart_Click(object sender, EventArgs e)
        {
            if (dgTextParts.SelectedRows.Count > 0)
            {
                Microsoft.Office.Interop.Word.Selection currentSelection =
                  Globals.ThisAddIn.Application.Selection;
                currentSelection.ParagraphFormat.LineSpacingRule =
                  WdLineSpacing.wdLineSpaceSingle;
                currentSelection.TypeText(textPartIdsToContent.FirstOrDefault(
                  t => t.Value == dgTextParts.SelectedRows[0].Cells[0].Value.ToString()).Value); currentSelection.TypeParagraph();
            }
        }
    }
}
