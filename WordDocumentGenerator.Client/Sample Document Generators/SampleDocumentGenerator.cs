// ----------------------------------------------------------------------
// <copyright file="SampleDocumentGenerator.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Client
{
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Wordprocessing;
    using WordDocumentGenerator.Library;

    /// <summary>
    /// Sample non-refreshable document generator for Test_Template - 1.docx template
    /// </summary>
    public class SampleDocumentGenerator : DocumentGenerator
    {
        // Content Control Tags
        protected const string PlaceholderIgnoreA = "PlaceholderIgnoreA";
        protected const string PlaceholderIgnoreB = "PlaceholderIgnoreB";
        
        protected const string PlaceholderContainerA = "PlaceholderContainerA";
        
        protected const string PlaceholderRecursiveA = "PlaceholderRecursiveA";
        protected const string PlaceholderRecursiveB = "PlaceholderRecursiveB";
        
        protected const string PlaceholderNonRecursiveA = "PlaceholderNonRecursiveA";
        protected const string PlaceholderNonRecursiveB = "PlaceholderNonRecursiveB";
        protected const string PlaceholderNonRecursiveC = "PlaceholderNonRecursiveC";
        protected const string PlaceholderNonRecursiveD = "PlaceholderNonRecursiveD";

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="SampleDocumentGenerator"/> class.
        /// </summary>
        /// <param name="generationInfo">The generation info.</param>
        public SampleDocumentGenerator(DocumentGenerationInfo generationInfo)
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
            Dictionary<string, PlaceHolderType> placeHolderTagToTypeCollection = new Dictionary<string, PlaceHolderType>();

            // Handle ignore placeholders
            placeHolderTagToTypeCollection.Add(PlaceholderIgnoreA, PlaceHolderType.Ignore);
            placeHolderTagToTypeCollection.Add(PlaceholderIgnoreB, PlaceHolderType.Ignore);

            // Handle container placeholders            
            placeHolderTagToTypeCollection.Add(PlaceholderContainerA, PlaceHolderType.Container);

            // Handle recursive placeholders            
            placeHolderTagToTypeCollection.Add(PlaceholderRecursiveA, PlaceHolderType.Recursive);
            placeHolderTagToTypeCollection.Add(PlaceholderRecursiveB, PlaceHolderType.Recursive);

            // Handle non recursive placeholders
            placeHolderTagToTypeCollection.Add(PlaceholderNonRecursiveA, PlaceHolderType.NonRecursive);
            placeHolderTagToTypeCollection.Add(PlaceholderNonRecursiveB, PlaceHolderType.NonRecursive);
            placeHolderTagToTypeCollection.Add(PlaceholderNonRecursiveC, PlaceHolderType.NonRecursive);
            placeHolderTagToTypeCollection.Add(PlaceholderNonRecursiveD, PlaceHolderType.NonRecursive);

            return placeHolderTagToTypeCollection;
        }

        /// <summary>
        /// Ignore placeholder found.
        /// </summary>
        /// <param name="placeholderTag">The placeholder tag.</param>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        protected override void IgnorePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext)
        {
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

            switch (tagPlaceHolderValue)
            {
                case PlaceholderNonRecursiveA:
                    tagValue = ((openXmlElementDataContext.DataContext) as Vendor).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as Vendor).Name;
                    break;
                case PlaceholderNonRecursiveB:
                    tagValue = ((openXmlElementDataContext.DataContext) as Item).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as Item).Name;
                    break;
                case PlaceholderNonRecursiveC:
                    tagValue = ((openXmlElementDataContext.DataContext) as Order).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as Order).Name;
                    break;
                case PlaceholderNonRecursiveD:
                    tagValue = ((openXmlElementDataContext.DataContext) as Order).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as Order).Description;
                    break;
            }

            // Set the tag for the content control
            if (!string.IsNullOrEmpty(tagValue))
            {
                this.SetTagValue(openXmlElementDataContext.Element as SdtElement, GetFullTagValue(tagPlaceHolderValue, tagValue));
            }

            // Set text without data binding
            this.SetContentOfContentControl(openXmlElementDataContext.Element as SdtElement, content);
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

            switch (tagPlaceHolderValue)
            {
                case PlaceholderRecursiveA:

                    foreach (Vendor testB in ((openXmlElementDataContext.DataContext) as Order).vendors)
                    {
                        SdtElement clonedElement = this.CloneElementAndSetContentInPlaceholders(new OpenXmlElementDataContext() { Element = openXmlElementDataContext.Element, DataContext = testB });
                    }

                    openXmlElementDataContext.Element.Remove();

                    break;
                case PlaceholderRecursiveB:

                    foreach (Item testC in ((openXmlElementDataContext.DataContext) as Order).items)
                    {
                        SdtElement clonedElement = this.CloneElementAndSetContentInPlaceholders(new OpenXmlElementDataContext() { Element = openXmlElementDataContext.Element, DataContext = testC });
                    }

                    openXmlElementDataContext.Element.Remove();
                    break;
            }
        }

        /// <summary>
        /// Container placeholder found.
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
                    // As this sample is non-refreshable hence we don't call GetRecursiveTemplateElementForContainer method( Sets the parentContainer from CustomXmlPart if refresh else saves the parentContainer markup to CustomXmlPart)
                    tagValue = (openXmlElementDataContext.DataContext as Order).Id.ToString();

                    if (!string.IsNullOrEmpty(tagValue))
                    {
                        this.SetTagValue(openXmlElementDataContext.Element as SdtElement, GetFullTagValue(tagPlaceHolderValue, tagValue));
                    }

                    foreach (var v in openXmlElementDataContext.Element.Elements())
                    {
                        this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = v, DataContext = openXmlElementDataContext.DataContext });
                    }

                    break;
            }
        }

        #endregion        
    }    
}
