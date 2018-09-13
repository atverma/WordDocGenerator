// ----------------------------------------------------------------------
// <copyright file="SampleDocumentWithTableGeneratorUsingDatabinding.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Client
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using WordDocumentGenerator.Library;

    /// <summary>
    /// Sample refreshable document generator for Test_Template - 2.docx(has table) template using data bound content controls
    /// </summary>
    public class SampleDocumentWithTableGeneratorUsingDatabinding : SampleDocumentWithTableGenerator
    {
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="SampleDocumentWithTableGeneratorUsingDatabinding"/> class.
        /// </summary>
        /// <param name="generationInfo">The generation info.</param>
        public SampleDocumentWithTableGeneratorUsingDatabinding(DocumentGenerationInfo generationInfo)
            : base(generationInfo)
        {
        }

        #endregion

        #region Overridden methods

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

            // Index is used to build XPath for controls that are bound to collection item
            int index = -1;

            // XPath to be used for data binding
            string xPath = string.Empty;

            switch (tagPlaceHolderValue)
            {
                case PlaceholderNonRecursiveA:
                    tagValue = ((openXmlElementDataContext.DataContext) as Vendor).Id.ToString();
                    index = (this.GetDataContext() as Order).vendors.IndexOf((openXmlElementDataContext.DataContext) as Vendor);
                    xPath = "/ns0:DocumentRootNode[1]/ns0:DataBoundControlsDataStore[1]/ns0:Data[1]/Order[1]/vendors[1]/Vendor[" + (index + 1).ToString() + "]/Name[1]";
                    break;
                case PlaceholderNonRecursiveB:
                    tagValue = ((openXmlElementDataContext.DataContext) as Item).Id.ToString();
                    index = (this.GetDataContext() as Order).items.IndexOf((openXmlElementDataContext.DataContext) as Item);
                    xPath = "/ns0:DocumentRootNode[1]/ns0:DataBoundControlsDataStore[1]/ns0:Data[1]/Order[1]/items[1]/Item[" + (index + 1).ToString() + "]/Name[1]";
                    break;
                case PlaceholderNonRecursiveC:
                    tagValue = (this.GetDataContext() as Order).Id.ToString();
                    xPath = "/ns0:DocumentRootNode[1]/ns0:DataBoundControlsDataStore[1]/ns0:Data[1]/Order[1]/Name[1]";
                    break;
                case PlaceholderNonRecursiveD:
                    tagValue = (this.GetDataContext() as Order).Id.ToString();
                    xPath = "/ns0:DocumentRootNode[1]/ns0:DataBoundControlsDataStore[1]/ns0:Data[1]/Order[1]/Description[1]";
                    break;
                case VendorId:
                    tagValue = ((openXmlElementDataContext.DataContext) as Vendor).Id.ToString();
                    index = (this.GetDataContext() as Order).vendors.IndexOf((openXmlElementDataContext.DataContext) as Vendor);
                    xPath = "/ns0:DocumentRootNode[1]/ns0:DataBoundControlsDataStore[1]/ns0:Data[1]/Order[1]/vendors[1]/Vendor[" + (index + 1).ToString() + "]/Id[1]";
                    break;
                case VendorName:
                    tagValue = ((openXmlElementDataContext.DataContext) as Vendor).Id.ToString();
                    index = (this.GetDataContext() as Order).vendors.IndexOf((openXmlElementDataContext.DataContext) as Vendor);
                    xPath = "/ns0:DocumentRootNode[1]/ns0:DataBoundControlsDataStore[1]/ns0:Data[1]/Order[1]/vendors[1]/Vendor[" + (index + 1).ToString() + "]/Name[1]";
                    break;
            }

            // Set the tag for the content control
            if (!string.IsNullOrEmpty(tagValue))
            {
                this.SetTagValue(openXmlElementDataContext.Element as SdtElement, GetFullTagValue(tagPlaceHolderValue, tagValue));
            }

            // Set the data binding for content control
            if (!string.IsNullOrEmpty(xPath))
            {
                this.SetDataBinding(xPath, (openXmlElementDataContext.Element) as SdtElement);
            }
        }

        #endregion
    }
}
