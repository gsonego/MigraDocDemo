using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;

namespace MigraDocDemo.Program
{
    public class Cover
    {
        /// <summary>
        /// Defines the cover page.
        /// </summary>
        public static void DefineCover(Document document)
        {
            Section section = document.AddSection();

            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.SpaceAfter = "3cm";

            Image image = section.AddImage("../../images/Logo.png");
            image.Width = "10cm";

            paragraph = section.AddParagraph("A sample document that demonstrates the\ncapabilities of MigraDoc");
            paragraph.Format.Font.Size = 16;
            paragraph.Format.Font.Color = Colors.DarkRed;
            paragraph.Format.SpaceBefore = "8cm";
            paragraph.Format.SpaceAfter = "3cm";

            paragraph = section.AddParagraph("Rendering date: ");
            paragraph.AddDateField();
        }
    }
}
