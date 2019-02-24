namespace ExcelReport.Accumulations
{
    public class Accumulation
    {
        public int Value { get; private set; }

        public Accumulation()
        {
            Value = 0;
        }

        public void Add(int increment)
        {
            Value += increment;
        }
    }
}