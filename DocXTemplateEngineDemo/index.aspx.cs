using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.UI;
using TemplateEngine.Docx;

namespace DocXTemplateEngineDemo
{
    public partial class index : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        protected void downloadFile_btn_OnClick(object sender, EventArgs e)
        {
            var targetPath = "D:/demo/";
            string fileName = $"{DateTime.Now:yyyyMMddHHmmss}.docx";
            File.Copy(Server.MapPath("~/Template/template.docx"), targetPath + fileName);

            var valuesToFill = new Content();
            valuesToFill.Fields.Add(new FieldContent("name", "王小明"));
            valuesToFill.Fields.Add(new FieldContent("special_1", "■"));
            valuesToFill.Fields.Add(new FieldContent("special_2", "□"));

            var tableContent = new TableContent("row");
            foreach (var f in FakeData())
                tableContent.AddRow(
                    new FieldContent("r_name", f.Name),
                    new FieldContent("r_chinese", f.Chinese),
                    new FieldContent("r_english", f.English),
                    new FieldContent("r_math", f.Math)
                );
            valuesToFill.Tables.Add(tableContent);
            var avg_chinese = FakeData().Average(x => Convert.ToDecimal(x.Chinese));
            var avg_english = FakeData().Average(x => Convert.ToDecimal(x.English));
            var avg_math = FakeData().Average(x => Convert.ToDecimal(x.Math));
            valuesToFill.Fields.Add(new FieldContent("avg_chinese", avg_chinese.ToString("0.0")));
            valuesToFill.Fields.Add(new FieldContent("avg_english", avg_english.ToString("0.0")));
            valuesToFill.Fields.Add(new FieldContent("avg_math", avg_math.ToString("0.0")));

            using (var outputDocument = new TemplateProcessor(targetPath + fileName).SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            }
        }

        public List<ScoreSheet> FakeData()
        {
            return new List<ScoreSheet>
            {
                new ScoreSheet
                {
                    Name = "王XX",
                    Chinese = "100",
                    English = "70",
                    Math = "90"
                },
                new ScoreSheet
                {
                    Name = "鄭XX",
                    Chinese = "90",
                    English = "60",
                    Math = "80"
                },
                new ScoreSheet
                {
                    Name = "陳XX",
                    Chinese = "80",
                    English = "50",
                    Math = "60"
                },
                new ScoreSheet
                {
                    Name = "張XX",
                    Chinese = "70",
                    English = "80",
                    Math = "70"
                }
            };
        }

        public class ScoreSheet
        {
            public string Name { get; set; }
            public string Chinese { get; set; }
            public string English { get; set; }
            public string Math { get; set; }
        }
    }
}