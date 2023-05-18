namespace XlsxMerge.Extensions;

public static class TextBoxExtension
{
    public static Binding BindingText(this TextBox textBox, object dataSource, string dataMember)
    {
        return textBox.DataBindings.Add(nameof(textBox.Text), dataSource, dataMember);
    }
}
