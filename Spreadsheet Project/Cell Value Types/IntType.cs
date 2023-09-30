namespace Spreadsheet_Project
{
    public class IntType : IValue
    {
        private readonly int intValue;

        public dynamic TypeValue => intValue;

        public IntType(int intValue)
        {
            this.intValue = intValue;
        }
    }
}
