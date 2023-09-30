namespace Spreadsheet_Project
{
    public class GlobalCopyCutVariable
    {
        private static readonly GlobalCopyCutVariable instance = new();

        private (
            (IValue Content, string Formula) Contents,
            bool cut,
            int sheetReferenceIndex
        ) cutCopyData;

        private GlobalCopyCutVariable() { }

        public static GlobalCopyCutVariable GetInstance()
        {
            return instance;
        }

        public (
            (IValue Content, string Formula) Contents,
            bool cut,
            int sheetReferenceIndex
        ) GetValue()
        {
            return cutCopyData;
        }

        public void SetValue(
            (
                (IValue Content, string Formula) Contents,
                bool cut,
                int sheetReferenceIndex
            ) cutCopyData
        )
        {
            this.cutCopyData = cutCopyData;
        }
    }

}
