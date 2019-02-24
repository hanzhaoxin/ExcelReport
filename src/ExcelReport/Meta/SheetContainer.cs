namespace ExcelReport.Meta
{
    public class SheetContainer : Named
    {
        public Container<Parameter> Parameters { get; } = new Container<Parameter>();

        public Container<Repeater> Repeaters { get; } = new Container<Repeater>();
    }
}