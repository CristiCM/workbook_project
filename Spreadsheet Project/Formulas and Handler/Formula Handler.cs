namespace Spreadsheet_Project.Formulas
{
    public class FormulaHandler
    {
        readonly Formulas formulas;
        readonly IEnumerable<
            Func<(int, int), (bool Match, string formulaResult)>
        > allAvailableFormulas;

        public FormulaHandler(Sheet sheetReference)
        {
            this.formulas = new Formulas(sheetReference);

            allAvailableFormulas = formulas.GetAllAvailableFormulas();
        }

        /// <summary>
        /// Takes the contents of a indicated cell based on a currentPosition and tries to find the matching formula.
        /// In the case of a match it returns the formula result. Otherwise it returns null.
        /// </summary>
        public string? TryFormulas((int X, int Y) currentPosition)
        {
            foreach (var formula in allAvailableFormulas)
            {
                var (formulaMatch, formulaResult) = formula.Invoke(currentPosition);

                if (formulaMatch)
                {
                    return formulaResult;
                }
            }

            return null;
        }
    }
}
