namespace Spreadsheet_Project
{
    public class DoubleType : IValue
    {
        private readonly double doubleValue;

        public dynamic TypeValue => doubleValue;

        public DoubleType(double doubleValue)
        {
            this.doubleValue = doubleValue;
        }
    }
}
