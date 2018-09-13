// ----------------------------------------------------------------------
// <copyright file="SampleDocumentGenerator.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Client
{
    using System.Collections.Generic;
    using System.Xml;
    using System.Xml.XPath;
    using DocumentFormat.OpenXml.Wordprocessing;
    using WordDocumentGenerator.Library;
    using System.Linq;

    /// <summary>
    /// Sample non-refreshable generic document generator for Test_Template - 1.docx & Test_Template - 2.docx templates using XML
    /// </summary>
    public class SampleGenericDocumentGeneratorUsingXml : DocumentGenerator
    {
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="SampleGenericDocumentGeneratorUsingXml"/> class.
        /// </summary>
        /// <param name="generationInfo">The generation info.</param>
        public SampleGenericDocumentGeneratorUsingXml(DocumentGenerationInfo generationInfo)
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
            Dictionary<string, PlaceHolderType> dict = new Dictionary<string, PlaceHolderType>();

            XmlElement dataContext = this.GetDataContext() as XmlElement;

            if (dataContext != null)
            {
                foreach (XmlNode elem in dataContext.OwnerDocument.GetElementsByTagName("contentControl"))
                {
                    XmlAttribute attrType = elem.Attributes["type"];
                    XmlAttribute attrTag = elem.Attributes["tag"];

                    if (!dict.ContainsKey(attrTag.Value))
                    {
                        dict.Add(attrTag.Value, (PlaceHolderType)(int.Parse(attrType.Value)));
                    }
                }
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
            string refTagValue = string.Empty;
            string refControlValue = string.Empty;

            XPathNavigator navigator = Parse(openXmlElementDataContext, tagPlaceHolderValue, ref refTagValue, ref refControlValue);

            if (navigator != null)
            {
                tagValue = navigator.GetAttribute(refTagValue, (openXmlElementDataContext.DataContext as XmlNode).NamespaceURI);
                content = navigator.GetAttribute(refControlValue, (openXmlElementDataContext.DataContext as XmlNode).NamespaceURI);
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

            string refTagValue = string.Empty;
            string refControlValue = string.Empty;

            XPathNavigator navigator = Parse(openXmlElementDataContext, tagPlaceHolderValue, ref refTagValue, ref refControlValue);

            if (navigator != null)
            {
                XPathNodeIterator nodeIterator = navigator.SelectDescendants("field", (openXmlElementDataContext.DataContext as XmlNode).NamespaceURI, false);

                while (nodeIterator.MoveNext())
                {
                    // Get the Ancestors
                    XPathNodeIterator xIterator = nodeIterator.Current.SelectAncestors("field", (openXmlElementDataContext.DataContext as XmlNode).NamespaceURI, false);

                    // Only the first Ancestor of Node
                    if (xIterator.MoveNext())
                    {
                        // Get the attribute of the first Ancestor
                        string attr = xIterator.Current.GetAttribute("contentControlTagREFS", (openXmlElementDataContext.DataContext as XmlNode).NamespaceURI);

                        if (!string.IsNullOrEmpty(attr))
                        {
                            // If Ancestor attribute contains the current place holder then only clone element
                            if ((new List<string>(attr.Split(' '))).Contains(tagPlaceHolderValue))
                            {
                                XmlDocument e = new XmlDocument();
                                e.LoadXml(nodeIterator.Current.OuterXml);
                                SdtElement clonedElement = this.CloneElementAndSetContentInPlaceholders(new OpenXmlElementDataContext() { Element = openXmlElementDataContext.Element, DataContext = e.DocumentElement });
                            }
                        }
                    }
                }
            }

            openXmlElementDataContext.Element.Remove();
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

            string refTagValue = string.Empty;
            string refControlValue = string.Empty;

            XPathNavigator navigator = Parse(openXmlElementDataContext, tagPlaceHolderValue, ref refTagValue, ref refControlValue);

            if (navigator != null)
            {
                tagValue = navigator.GetAttribute(refTagValue, (openXmlElementDataContext.DataContext as XmlNode).NamespaceURI);
            }

            if (!string.IsNullOrEmpty(tagValue))
            {
                this.SetTagValue(openXmlElementDataContext.Element as SdtElement, GetFullTagValue(tagPlaceHolderValue, tagValue));
            }

            foreach (var v in openXmlElementDataContext.Element.Elements())
            {
                this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = v, DataContext = openXmlElementDataContext.DataContext });
            }
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Parses the specified open XML element data context.
        /// </summary>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        /// <param name="tagPlaceHolderValue">The tag place holder value.</param>
        /// <param name="refTagValue">The ref tag value.</param>
        /// <param name="refControlValue">The ref control value.</param>
        /// <returns></returns>
        private XPathNavigator Parse(OpenXmlElementDataContext openXmlElementDataContext, string tagPlaceHolderValue, ref string refTagValue, ref string refControlValue)
        {
            XPathNavigator contentControlsNavigator = (this.GetDataContext() as XmlNode).CreateNavigator();
            XPathNodeIterator contentControlsIterator = contentControlsNavigator.SelectDescendants("contentControl", (openXmlElementDataContext.DataContext as XmlNode).NamespaceURI, true);

            while (contentControlsIterator.MoveNext())
            {
                string contentControlTagAttr = contentControlsIterator.Current.GetAttribute("tag", (openXmlElementDataContext.DataContext as XmlNode).NamespaceURI);
                refControlValue = contentControlsIterator.Current.GetAttribute("refControlValue", (openXmlElementDataContext.DataContext as XmlNode).NamespaceURI);
                refTagValue = contentControlsIterator.Current.GetAttribute("refTagValue", (openXmlElementDataContext.DataContext as XmlNode).NamespaceURI);

                if ((!string.IsNullOrEmpty(contentControlTagAttr)) && contentControlTagAttr == tagPlaceHolderValue)
                {
                    XPathNavigator navigator = (openXmlElementDataContext.DataContext as XmlNode).CreateNavigator();
                    XPathNodeIterator iterator = navigator.SelectDescendants("field", (openXmlElementDataContext.DataContext as XmlNode).NamespaceURI, true);

                    while (iterator.MoveNext())
                    {
                        string attr = iterator.Current.GetAttribute("contentControlTagREFS", (openXmlElementDataContext.DataContext as XmlNode).NamespaceURI);

                        if (!string.IsNullOrEmpty(attr))
                        {
                            if ((new List<string>(attr.Split(' '))).Contains(contentControlTagAttr))
                            {
                                return iterator.Current;
                            }
                        }
                    }

                    break;
                }
            }

            return null;
        }

        #endregion
    }

}
