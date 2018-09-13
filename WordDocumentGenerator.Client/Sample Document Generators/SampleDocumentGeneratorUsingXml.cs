// ----------------------------------------------------------------------
// <copyright file="SampleDocumentGeneratorUsingXml.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Client
{
    using System.Collections.Generic;
    using System.Xml;
    using DocumentFormat.OpenXml.Wordprocessing;
    using WordDocumentGenerator.Library;

    /// <summary>
    /// Sample generic non-refreshable document generator using Xml data for Test_Template - 1.docx template
    /// </summary>
    public class SampleDocumentGeneratorUsingXml : DocumentGenerator
    {
        private Dictionary<string, ContentControlXmlMetadata> placeHolderNameToContentControlXmlMetadataCollection;
                
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="SampleDocumentGenerator"/> class.
        /// </summary>
        /// <param name="generationInfo">The generation info.</param>
        /// <param name="placeHolderNameToContentControlXmlMetadataCollection">The place holder name to content control XML metadata collection.</param>
        public SampleDocumentGeneratorUsingXml(DocumentGenerationInfo generationInfo, Dictionary<string, ContentControlXmlMetadata> placeHolderNameToContentControlXmlMetadataCollection)
            : base(generationInfo)
        {
            this.placeHolderNameToContentControlXmlMetadataCollection = placeHolderNameToContentControlXmlMetadataCollection;            
        }

        #endregion

        #region Overridden methods

        /// <summary>
        /// Gets the place holder tag to type collection.
        /// </summary>
        /// <returns></returns>
        protected override Dictionary<string, PlaceHolderType> GetPlaceHolderTagToTypeCollection()
        {
            Dictionary<string, PlaceHolderType> dict = new Dictionary<string, PlaceHolderType>();
            
            foreach (string key in this.placeHolderNameToContentControlXmlMetadataCollection.Keys)
            {
                dict.Add(key, placeHolderNameToContentControlXmlMetadataCollection[key].Type);
            }
            
            return dict;
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

            if (this.placeHolderNameToContentControlXmlMetadataCollection.ContainsKey(tagPlaceHolderValue))
            {                
                tagValue = this.GetNodeText(openXmlElementDataContext.DataContext, this.placeHolderNameToContentControlXmlMetadataCollection[tagPlaceHolderValue].ControlTagXPath);
                content = this.GetNodeText(openXmlElementDataContext.DataContext, this.placeHolderNameToContentControlXmlMetadataCollection[tagPlaceHolderValue].ControlValueXPath);                
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

            if (this.placeHolderNameToContentControlXmlMetadataCollection.ContainsKey(tagPlaceHolderValue))
            {
                XmlNode node = GetNode(openXmlElementDataContext.DataContext, this.placeHolderNameToContentControlXmlMetadataCollection[tagPlaceHolderValue].ControlValueXPath);

                foreach (XmlNode childNode in node.ChildNodes)
                {
                    SdtElement clonedElement = this.CloneElementAndSetContentInPlaceholders(new OpenXmlElementDataContext() { Element = openXmlElementDataContext.Element, DataContext = childNode});
                }

                openXmlElementDataContext.Element.Remove();                
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

            if (this.placeHolderNameToContentControlXmlMetadataCollection.ContainsKey(tagPlaceHolderValue))
            {                
                tagValue = GetNodeText(openXmlElementDataContext.DataContext, this.placeHolderNameToContentControlXmlMetadataCollection[tagPlaceHolderValue].ControlTagXPath);

                if (!string.IsNullOrEmpty(tagValue))
                {
                    this.SetTagValue(openXmlElementDataContext.Element as SdtElement, GetFullTagValue(tagPlaceHolderValue, tagValue));
                }

                foreach (var v in openXmlElementDataContext.Element.Elements())
                {
                    this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = v, DataContext = openXmlElementDataContext.DataContext });
                }
            }
        }

        #endregion

        #region Private Methods        

        /// <summary>
        /// Gets the node.
        /// </summary>
        /// <param name="node">The node.</param>
        /// <param name="xPath">The x path.</param>
        /// <returns></returns>
        private XmlNode GetNode(object node, string xPath)
        {
            XmlNode childNode = null;

            if (node as XmlNode != null)
            {
                XmlNamespaceManager mgr = new XmlNamespaceManager(new NameTable());
                mgr.AddNamespace("ns0", DocumentGenerationInfo.NamespaceUri);

                childNode = (node as XmlNode).SelectSingleNode(xPath, mgr);
            }

            return childNode;
        }

        /// <summary>
        /// Gets the node text.
        /// </summary>
        /// <param name="node">The node.</param>
        /// <param name="xPath">The x path.</param>
        /// <returns></returns>
        private string GetNodeText(object node, string xPath)
        {
            string text = string.Empty;
            XmlNode childNode = GetNode(node, xPath);

            if (childNode != null)
            {
                text = childNode.InnerText;
            } 

            return text;
        }

        #endregion
    }
}