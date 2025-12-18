using System.Globalization;
using System.Linq;
using Xunit;
using Xunit.StaFact;

namespace MGPad.Tests;

public class MainWindowTests
{
    [StaFact]
    public void ExtractNumbersFromSelection_PreservesFractionDigitCounts()
    {
        string selection = "1.23 apples 4.5 bananas 6";

        var numbers = MainWindow.ExtractNumbersFromSelection(selection);

        Assert.Collection(
            numbers,
            first =>
            {
                Assert.Equal(1.23m, first.Value);
                Assert.Equal(2, first.FractionDigits);
            },
            second =>
            {
                Assert.Equal(4.5m, second.Value);
                Assert.Equal(1, second.FractionDigits);
            },
            third =>
            {
                Assert.Equal(6m, third.Value);
                Assert.Equal(0, third.FractionDigits);
            });
    }

    [StaFact]
    public void SumSelection_FormatsUsingMaxFractionDigits()
    {
        string selection = "1.1\n2.25 3";

        var numbers = MainWindow.ExtractNumbersFromSelection(selection);

        decimal sum = numbers.Sum(number => number.Value);
        int maxFractionDigits = numbers.Max(number => number.FractionDigits);

        string formatted = sum.ToString($"F{maxFractionDigits}", CultureInfo.InvariantCulture);

        Assert.Equal(3, numbers.Count);
        Assert.Equal("6.35", formatted);
    }

    [StaFact]
    public void SumSelection_PreservesTrailingZeros()
    {
        string selection = "1.50 2.5";

        var numbers = MainWindow.ExtractNumbersFromSelection(selection);

        decimal sum = numbers.Sum(number => number.Value);
        int maxFractionDigits = numbers.Max(number => number.FractionDigits);

        string formatted = sum.ToString($"F{maxFractionDigits}", CultureInfo.InvariantCulture);

        Assert.Equal("4.00", formatted);
    }
}
