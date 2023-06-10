namespace MailMerge;

/// <summary>
/// <para>This folder implements a very small fragment of OfficeOpenXml, sufficient for
/// programmatically merging MSWord docx documents.</para>
///
/// <para>The full OoXml spec for WordProcessingML is ECMA-376
/// https://www.ecma-international.org/publications/standards/Ecma-376.htm .
/// It is very long.
/// We only implement a fraction of it, namely the ability to do some basic merge field transforms.</para> 
///
/// <para>There is a stunningly helpful, explanatory, mini-reference at http://officeopenxml.com/ ,
/// created by Daniel Dick, http://officeopenxml.com/aboutThisSite.php</para>
/// </summary>
/// 
/// <remarks>
/// When adding new capabilities, put them in this folder.
/// </remarks>
public static partial class KnownWordProcessingMLTransformationsReadMe
{
    /// <summary>Transformations so far implemented</summary>
    /// 
    static string[] AllKnown = new[]
    {
        nameof(TransformSimpleMergeFields.MergeSimpleMergeFields),
        nameof(TransformComplexMergeFields.MergeComplexMergeFields),
        nameof(TransformDateFields.MergeDateFields),
    };

}