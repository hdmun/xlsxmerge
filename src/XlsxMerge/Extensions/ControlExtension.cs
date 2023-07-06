namespace XlsxMerge.Extensions;

public static class ControlExtension
{
    public static Binding BindingText(this TextBox textBox, object dataSource, string dataMember)
    {
        return textBox.DataBindings.Add(nameof(textBox.Text), dataSource, dataMember);
    }

    public static Binding BindingText(this Label label, object dataSource, string dataMember)
    {
        return label.DataBindings.Add(nameof(label.Text), dataSource, dataMember);
    }

    public static Binding BindingVisible(this Control control, object dataSource, string dataMember)
    {
        return control.DataBindings.Add(nameof(control.Visible), dataSource, dataMember);
    }

    public static Binding BindingEnabled(this Control control, object dataSource, string dataMember)
    {
        return control.DataBindings.Add(nameof(control.Enabled), dataSource, dataMember);
    }
}
