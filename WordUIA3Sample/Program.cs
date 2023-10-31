using System.Diagnostics;
using UIAutomationClient;

// This is an example of getting visible paragraphs in Word
// using UIAutomationClient COM API (UIA3) in C#.
//
// Here is a sample output:
//
// Paragraph @0-15: Hello, World!
// + CommentedRect[0]: X = 195, Y = 246, Width = 15, Height = 26
// Paragraph @15-16:
// Paragraph @16-463: Lorem ipsum dolor sit amet, consectetur adipiscing elit, ...
// + CommentedRect[0]: X = 691, Y = 299, Width = 96, Height = 26
// + CommentedRect[1]: X = 149, Y = 326, Width = 42, Height = 26

// Get the word process.
var process = Process.GetProcessesByName("WINWORD").FirstOrDefault();
if (process == null)
{
    Console.Error.WriteLine("Failed to get the Word process. Ensure Word is running.");
    return;
}

// CUIAutomation is also available on Windows 7 or prior.
// CUIAutomation8 allows you to configure the timeout.
var uia = new CUIAutomation8
{
    ConnectionTimeout = 1000,
    TransactionTimeout = (uint)TimeSpan.FromSeconds(1).TotalMilliseconds,
};

// Get the document element of the Word window.
var windowElement = uia.ElementFromHandle(process.MainWindowHandle);
var documentCondition = uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_DocumentControlTypeId);
var documentElement = windowElement.FindFirst(TreeScope.TreeScope_Descendants, documentCondition);
if (documentElement == null)
{
    Console.Error.WriteLine("Failed to get the document element. Ensure a file is opened in Word.");
    return;
}

// Get the text pattern of the document element.
var textPattern = documentElement.GetCurrentPattern(UIA_PatternIds.UIA_TextPatternId) as IUIAutomationTextPattern;
if (textPattern == null)
{
    Console.Error.WriteLine("Failed to get the text pattern.");
    return;
}

// Get ranges with comments in all the visible ranges.
var ranges = textPattern.GetVisibleRanges();
for (int i = 0, n = ranges.Length; i < n; ++i)
{
    // Get the range for the enclosing paragraph.
    // Visible ranges seem to be split uniquely,
    // different from paragraphs or any other text unit.
    // Please note paragraphs here (UI paragraphs) can be split differently
    // from paragraphs recognized by Word (Word paragraphs,
    // i.e., Microsoft.Office.Interop.Word.Paragraph).
    // A Word paragraph may contain multiple UI paragraphs, or vice versa.
    var range = ranges.GetElement(i);
    range.ExpandToEnclosingUnit(TextUnit.TextUnit_Paragraph);

    // Get the position of the paragraph relative to the beginning of the document.
    var startIndex = range.CompareEndpoints(TextPatternRangeEndpoint.TextPatternRangeEndpoint_Start, textPattern.DocumentRange, TextPatternRangeEndpoint.TextPatternRangeEndpoint_Start);
    var endIndex = range.CompareEndpoints(TextPatternRangeEndpoint.TextPatternRangeEndpoint_End, textPattern.DocumentRange, TextPatternRangeEndpoint.TextPatternRangeEndpoint_Start);

    // Get the text of the enclosing paragraph.
    // There seems an upper limit (approx. 65000 characters)
    // on the number of texts to be retrieved.
    var paragraphText = range.GetText(-1);

    Console.WriteLine("Paragraph @{0}-{1}: {2}", startIndex, endIndex, paragraphText);

    // Check if the range has comments.
    // Otherwise, FindAttribute() below throws InvalidOperationException
    // when there is no attribute found.
    var annotationTypes = range.GetAttributeValue(UIA_TextAttributeIds.UIA_AnnotationTypesAttributeId) as int[];
    if (!(annotationTypes?.Contains(UIA_AnnotationTypes.AnnotationType_Comment) ?? false))
    {
        // Check the next range.
        continue;
    }

    // Get the range for the first comment.
    var commentRange = range.FindAttribute(
        UIA_TextAttributeIds.UIA_AnnotationTypesAttributeId,
        UIA_AnnotationTypes.AnnotationType_Comment,
        0 /* backward: false (= forward) */);

    // Locate the rectangles of the comment range.
    // GetBoundingRectangles() returns an Array of double values
    // indicating X, Y, Width, Height of each rectangle.
    // Please note a range can have multiple rectangles
    // when the range wraps at the end of the line.
    var rectangleValues = commentRange.GetBoundingRectangles().OfType<double>().ToArray();
    for (int j = 0, m = rectangleValues.Length / 4; j < m; ++j)
    {
        Console.WriteLine("+ CommentedRect[{0}]: X={1}, Y={2}, Width={3}, Height={4}",
            j,
            rectangleValues[j * 4],
            rectangleValues[j * 4 + 1],
            rectangleValues[j * 4 + 2],
            rectangleValues[j * 4 + 3]);
    }
}
