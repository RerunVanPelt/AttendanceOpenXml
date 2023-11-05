using JetBrains.Annotations;
using Xunit;

namespace OpenXmlEngine.Tests;

[TestSubject(typeof(WorksheetProcessor))]
public class WorksheetProcessorTest
{
    [Fact]
    public void SetCellReferenceShouldProvideRightReferences()
    {
        // Arange 
        var expected = "J";

        // Actual
        var actual = WorksheetProcessor.SetCellReference(10);

        // Assert
        Assert.Equal(expected, actual);
    }


    [Theory]
    [InlineData(3, 1, "C1")]
    [InlineData(10, 1, "J1")]
    [InlineData(27, 1, "AA1")]
    [InlineData(28, 1, "AB1")]
    [InlineData(64, 1, "BL1")]
    [InlineData(720, 1, "AAR1")]
    [InlineData(780, 1, "ACZ1")]
    [InlineData(1449, 2, "BCS2")]
    public void SetCellReferenceShouldProvideRightSetOfReferences(uint cellColumn, uint cellRow, string expected)
    {
        // Arange 

        // Actual
        var actual = WorksheetProcessor.SetCellReference(cellColumn, cellRow);

        // Assert
        Assert.Equal(expected, actual);
    }
}