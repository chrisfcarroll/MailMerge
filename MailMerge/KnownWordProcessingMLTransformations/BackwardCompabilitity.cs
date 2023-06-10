using System;
using System.Collections.Generic;
using System.Xml;
using Microsoft.Extensions.Logging;

namespace MailMerge
{
    /// <remarks>This class refactored to class-per-transformation</remarks>
    public static class KnownWordProcessingMLTransforms
    {
        /// <summary>Renamed to <seealso cref="TransformComplexMergeFields.MergeComplexMergeFields"/></summary>
        [Obsolete("Renamed to TransformComplexMergeFields.MergeComplexMergeFields")]
        public static void ComplexMergeFields(this XmlDocument mainDocumentPart,
                                              Dictionary<string, string> fieldValues, 
                                              ILogger log)
            => TransformComplexMergeFields.MergeComplexMergeFields(mainDocumentPart, fieldValues, log);

        /// <summary>Renamed to <seealso cref="TransformSimpleMergeFields.MergeSimpleMergeFields"/></summary>
        [Obsolete("Renamed to TransformSimpleMergeFields.MergeSimpleMergeFields")]
        public static void SimpleMergeFields(this XmlDocument mainDocumentPart,
                                                       Dictionary<string, string> fieldValues, ILogger log)
            => TransformSimpleMergeFields.MergeSimpleMergeFields(mainDocumentPart, fieldValues, log);

        /// <summary>Renamed to <seealso cref="TransformDateFields.MergeDateFields(System.Xml.XmlDocument,Microsoft.Extensions.Logging.ILogger,string,string[])"/></summary>
        [Obsolete("Renamed to TransformDateFields.MergeDateFields")]
        public static void MergeDate(this XmlDocument mainDocumentPart, 
                                               ILogger logger, 
                                               string formattedFixedDate,
                                               string[] datesToReplace = default)
            => TransformDateFields.MergeDateFields(mainDocumentPart, logger, formattedFixedDate, datesToReplace);
        
        /// <summary>Renamed to <seealso cref="TransformDateFields.MergeDateFields(System.Xml.XmlDocument,Microsoft.Extensions.Logging.ILogger,System.Nullable{System.DateTime},string,string[])"/></summary>
        [Obsolete("Renamed to TransformDateFields.MergeDateFields")]
        public static void MergeDate(this XmlDocument mainDocumentPart, ILogger logger, DateTime? date, string formattedFixedDate = null, string[] datesToReplace=null)
            => TransformDateFields.MergeDateFields(mainDocumentPart, logger, date, formattedFixedDate, datesToReplace);
        
        /// <summary>Renamed to <seealso cref="TransformDateFields.DefaultDatesToReplace"/></summary>
        [Obsolete("Renamed to TransformDateFields.DefaultDatesToReplace")]
        public static readonly string[] DefaultDatesToReplace = TransformDateFields.DefaultDatesToReplace;
        
        /// <summary>Renamed to <seealso cref="XmlNodeTextExtensions.NewLineSeparators"/></summary>
        [Obsolete("Renamed to XmlNodeTextExtensions.NewLineSeparators")]
        public static readonly string[] NewLineSeparators =  XmlNodeTextExtensions.NewLineSeparators;
    }
}
