namespace Spreadsheet_Project
{
    public class StringType : IValue
    {
        private readonly string stringValue;

        public dynamic TypeValue => stringValue;

        public StringType(string stringValue)
        {
            this.stringValue = stringValue;
        }
    }
}
