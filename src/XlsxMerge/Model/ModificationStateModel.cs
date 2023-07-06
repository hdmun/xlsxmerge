namespace XlsxMerge.Model;

public class ModificationStateModel
{
    public readonly string Name;
    public readonly Color Color;

    public ModificationStateModel(string name, Color color)
    {
        Name = name;
        Color = color;
    }
}
