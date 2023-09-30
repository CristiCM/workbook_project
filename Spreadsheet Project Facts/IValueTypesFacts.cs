using Spreadsheet_Project;

namespace Spreadsheet_Project_Facts
{
    public class IValueTypesFacts
    {
        [Fact]
        public void Type_Int_ShouldParseStringAndCreateProperObject()
        {
            var typeObject = new IntType(1234);

            Assert.Equal(1234, typeObject.TypeValue);
        }

        [Fact]
        public void Type_Double_ShouldParseStringAndCreateProperObject()
        {
            var typeObject = new DoubleType(1.79769313486231570E+308);

            Assert.Equal(1.79769313486231570E+308, typeObject.TypeValue);
        }

        [Fact]
        public void Type_String_ShouldParseAndCreateProperObject()
        {
            var typeObject = new StringType("testing");

            Assert.Equal("testing", typeObject.TypeValue);
        }
    }
}