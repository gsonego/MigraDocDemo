using System.Diagnostics;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Pdf;

namespace MigraDocDemo.Program
{
    class Program
    {
        static void Main()
        {
            //SimpleDemo();

            CompleteDemo();
        }

        static void CompleteDemo()
        {
            // Create a MigraDoc document
            Document document = Documents.CreateDocument();

            //string ddl = MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToString(document);
            MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToFile(document, "MigraDoc.mdddl");

            PdfDocumentRenderer renderer = new PdfDocumentRenderer(true, PdfFontEmbedding.Always);
            renderer.Document = document;

            renderer.RenderDocument();

            // Save the document...
            const string filename = "HelloMigraDoc.pdf";
            renderer.PdfDocument.Save(filename);

            // ...and start a viewer.
            Process.Start(filename);
        }

        private static void SimpleDemo()
        {
            // tudo comeca com um documento
            var document = new Document();

            // no documento nao se adiciona páginas e sim secoes
            Section section = document.AddSection();

            // adionamos um novo paragrafo à secao criada
            section.AddParagraph("Hello, World!");

            // adiciona paragrafo vazio
            section.AddParagraph();

            // cria um objeto paragrafo para modificacao
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Font.Color = Color.FromCmyk(100, 30, 20, 50);
            paragraph.AddFormattedText("Hello, World!", TextFormat.Underline);

            // também pode-se criar um objeto de texto formatado para ser manipulado
            FormattedText ft = paragraph.AddFormattedText("Small text", TextFormat.Bold);
            ft.Font.Size = 6;

            // podemos renderizar o documento em PDF ou RTF
            var pdfRenderer = new PdfDocumentRenderer(false, PdfFontEmbedding.Always);

            // basta passar o documento criado anteriormente ao renderizador
            pdfRenderer.Document = document;

            // solicita a renderizacao
            pdfRenderer.RenderDocument();

            // salva o arquivo gerado em disco
            const string filename = "HelloWorld.pdf";
            pdfRenderer.PdfDocument.Save(filename);

            // solicita a abertura do arquivo
            Process.Start(filename);
        }
    }
}
