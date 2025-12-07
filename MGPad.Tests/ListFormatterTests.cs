using System.Windows.Controls;
using System.Windows.Documents;
using Xunit;
using Xunit.StaFact;

namespace MGPad.Tests;

public class ListFormatterTests
{
    [StaFact]
    public void RenumberLists_SeparatesNumberingByIndent()
    {
        RichTextBox editor = new()
        {
            Document = new FlowDocument()
        };

        TextRange range = new(editor.Document.ContentStart, editor.Document.ContentEnd)
        {
            Text = "1. Parent\n    5. Child\n    9. Second child\n3. Next parent\n"
        };

        ListFormatter.RenumberLists(editor);

        Assert.Equal(
            "1. Parent\r\n    1. Child\r\n    2. Second child\r\n2. Next parent\r\n",
            range.Text);
    }
}
