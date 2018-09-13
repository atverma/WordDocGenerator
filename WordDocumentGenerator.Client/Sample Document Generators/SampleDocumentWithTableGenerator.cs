// ----------------------------------------------------------------------
// <copyright file="SampleDocumentWithTableGenerator.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Client
{
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Wordprocessing;
    using WordDocumentGenerator.Library;

    /// <summary>
    /// Sample refreshable document generator for Test_Template - 2.docx(has table) template
    /// </summary>
    public class SampleDocumentWithTableGenerator : SampleRefreshableDocumentGenerator
    {
        // Content Control Tags - Table tags are different. Other Tags are same so reusing the base class's code
        protected const string VendorDetailRow = "VendorDetailRow";
        protected const string VendorId = "VendorId";
        protected const string VendorName = "VendorName";

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="SampleDocumentWithTableGenerator"/> class.
        /// </summary>
        /// <param name="generationInfo">The generation info.</param>
        public SampleDocumentWithTableGenerator(DocumentGenerationInfo generationInfo)
            : base(generationInfo)
        {
        }

        #endregion

        #region Overridden methods

        /// <summary>
        /// Gets the place holder tag to type collection.
        /// </summary>
        /// <returns></returns>
        protected override Dictionary<string, PlaceHolderType> GetPlaceHolderTagToTypeCollection()
        {
            Dictionary<string, PlaceHolderType> placeHolderTagToTypeCollection = base.GetPlaceHolderTagToTypeCollection();

            // Handle recursive placeholders            
            placeHolderTagToTypeCollection.Add(VendorDetailRow, PlaceHolderType.Recursive);

            // Handle non recursive placeholders
            placeHolderTagToTypeCollection.Add(VendorId, PlaceHolderType.NonRecursive);
            placeHolderTagToTypeCollection.Add(VendorName, PlaceHolderType.NonRecursive);

            return placeHolderTagToTypeCollection;
        }

        /// <summary>
        /// Non recursive placeholder found.
        /// </summary>
        /// <param name="placeholderTag">The placeholder tag.</param>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        protected override void NonRecursivePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext)
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

            // Reuse base class code and handle only tags unavailable in base class
            bool bubblePlaceHolder = true;

            switch (tagPlaceHolderValue)
            {
                case VendorId:
                    bubblePlaceHolder = false;
                    tagValue = ((openXmlElementDataContext.DataContext) as Vendor).Id.ToString();
                    content = tagValue;
                    break;
                case VendorName:
                    bubblePlaceHolder = false;
                    tagValue = ((openXmlElementDataContext.DataContext) as Vendor).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as Vendor).Name;
                    break;
            }

            if (bubblePlaceHolder)
            {
                // Use base class code as tags are already defined in base class.
                base.NonRecursivePlaceholderFound(placeholderTag, openXmlElementDataContext);
            }
            else
            {
                // Set the tag for the content control
                if (!string.IsNullOrEmpty(tagValue))
                {
                    this.SetTagValue(openXmlElementDataContext.Element as SdtElement, GetFullTagValue(tagPlaceHolderValue, tagValue));
                }

                // Set the content for the content control
                this.SetContentOfContentControl(openXmlElementDataContext.Element as SdtElement, content);
            }
        }

        /// <summary>
        /// Recursive placeholder found.
        /// </summary>
        /// <param name="placeholderTag">The placeholder tag.</param>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        protected override void RecursivePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (openXmlElementDataContext == null || openXmlElementDataContext.Element == null || openXmlElementDataContext.DataContext == null)
            {
                return;
            }

            string tagPlaceHolderValue = string.Empty;
            string tagGuidPart = string.Empty;
            GetTagValue(openXmlElementDataContext.Element as SdtElement, out tagPlaceHolderValue, out tagGuidPart);

            // Reuse base class code and handle only tags unavailable in base class
            bool bubblePlaceHolder = true;

            switch (tagPlaceHolderValue)
            {
                case VendorDetailRow:
                    bubblePlaceHolder = false;

                    foreach (Vendor testB in ((openXmlElementDataContext.DataContext) as Order).vendors)
                    {
                        SdtElement clonedElement = this.CloneElementAndSetContentInPlaceholders(new OpenXmlElementDataContext() { Element = openXmlElementDataContext.Element, DataContext = testB });
                    }

                    openXmlElementDataContext.Element.Remove();
                    break;
            }

            if (bubblePlaceHolder)
            {
                // Use base class code as tags are already defined in base class.
                base.RecursivePlaceholderFound(placeholderTag, openXmlElementDataContext);
            }
        }

        #endregion
    }
}
