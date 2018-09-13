// ----------------------------------------------------------------------
// <copyright file="SampleRefreshableDocumentGenerator.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Client
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using WordDocumentGenerator.Library;

    /// <summary>
    /// Sample refreshable document generator for Test_Template - 1.docx template
    /// </summary>
    public class SampleRefreshableDocumentGenerator : SampleDocumentGenerator
    {
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="SampleRefreshableDocumentGenerator"/> class.
        /// </summary>
        /// <param name="generationInfo">The generation info.</param>
        public SampleRefreshableDocumentGenerator(DocumentGenerationInfo generationInfo)
            : base(generationInfo)
        {

        }

        #endregion

        #region Overridden methods

        /// <summary>
        /// Containers the placeholder found.
        /// </summary>
        /// <param name="placeholderTag">The placeholder tag.</param>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        protected override void ContainerPlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (openXmlElementDataContext == null || openXmlElementDataContext.Element == null || openXmlElementDataContext.DataContext == null)
            {
                return;
            }
                        
            string tagPlaceHolderValue = string.Empty;
            string tagGuidPart = string.Empty;
            GetTagValue(openXmlElementDataContext.Element as SdtElement, out tagPlaceHolderValue, out tagGuidPart);

            string tagValue = string.Empty;
            string content = string.Empty;

            switch (tagPlaceHolderValue)
            {
                case PlaceholderContainerA:
                    SdtElement parentContainer = openXmlElementDataContext.Element as SdtElement;
                    // Sets the parentContainer from CustomXmlPart if refresh else saves the parentContainer markup to CustomXmlPart 
                    this.GetParentContainer(ref parentContainer, tagPlaceHolderValue);
                    base.ContainerPlaceholderFound(placeholderTag, new OpenXmlElementDataContext() { Element = parentContainer, DataContext = openXmlElementDataContext.DataContext });
                    break;
            }
        }

        #endregion
    }
}
