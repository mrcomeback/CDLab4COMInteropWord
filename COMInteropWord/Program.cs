using Microsoft.Office.Interop.Word;

namespace SimpleMarkupToWord
{
    class Program
    {
        static void Main(string[] args)
        {
            var wordApp = new Application { Visible = false };
            var doc = wordApp.Documents.Add();

            var lines = File.ReadAllLines("input.txt");

            foreach (var rawLine in lines)
            {
                bool isPlainText = rawLine.StartsWith("###");
                string content = isPlainText
                    ? rawLine.Substring(3).TrimStart()
                    : rawLine.Trim();

                var paragraph = doc.Content.Paragraphs.Add();
                paragraph.Range.Text = content;

                if (isPlainText)
                {
                    paragraph.Range.set_Style(WdBuiltinStyle.wdStyleNormal);
                }
                else
                {
                    paragraph.Range.set_Style(WdBuiltinStyle.wdStyleHeading1);
                }

                foreach (Microsoft.Office.Interop.Word.Range wordRange in paragraph.Range.Words)
                {
                    string w = wordRange.Text.Trim();
                    if (w.IndexOf('v', StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        wordRange.Font.Superscript = 1;
                    }
                }

                paragraph.Range.InsertParagraphAfter();
            }

            string outputPath = Path.Combine(
                Directory.GetCurrentDirectory(),
                "output.docx");
            doc.SaveAs2(outputPath);
            doc.Close();
            wordApp.Quit();

            Console.WriteLine($"Готово! Документ збережено {outputPath}");
        }
    }
}