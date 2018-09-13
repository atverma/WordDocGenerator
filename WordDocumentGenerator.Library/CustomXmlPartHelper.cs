// ----------------------------------------------------------------------
// <copyright file="CustomXmlPartHelper.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Library
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Xml;
    using System.Xml.Linq;
    using DocumentFormat.OpenXml.Packaging;

    /// <summary>
    /// Helper class for Word CustomXml part operations
    /// </summary>
    public class CustomXmlPartHelper
    {
        #region Members

        public readonly CustomXmlPartCore customXmlPartCore = null;        

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomXmlPartHelper"/> class.
        /// </summary>
        /// <param name="documentNamespace">The namespace URI.</param>
        public CustomXmlPartHelper(string namespaceUri)
        {
            this.customXmlPartCore = new CustomXmlPartCore(namespaceUri);
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Sets the type of the element from name to value collection for.
        /// </summary>
        /// <param name="mainDocumentPart">The main document part.</param>
        /// <param name="rootElementName">Name of the root element.</param>
        /// <param name="childElementName">Name of the child element.</param>
        /// <param name="nameToValueCollection">The name to value collection.</param>
        /// <param name="forNodeType">Type of for node.</param>
        public void SetElementFromNameToValueCollectionForType(MainDocumentPart mainDocumentPart, string rootElementName, string childElementName, Dictionary<string, string> nameToValueCollection, NodeType forNodeType)
        {
            if (mainDocumentPart == null)
            {
                throw new ArgumentNullException("mainDocumentPart");
            }

            if (string.IsNullOrEmpty(rootElementName))
            {
                throw new ArgumentNullException("rootElementName");
            }

            if (string.IsNullOrEmpty(childElementName))
            {
                throw new ArgumentNullException("childElementName");
            }

            if (nameToValueCollection == null)
            {
                throw new ArgumentNullException("nameToValueCollection");
            }

            XName rootElementXName = XName.Get(rootElementName, this.customXmlPartCore.namespaceUri);
            XName childElementXName = XName.Get(childElementName, this.customXmlPartCore.namespaceUri);
            XElement rootElement = new XElement(rootElementXName);
            XElement childElement = null;
            CustomXmlPart customXmlPart = this.customXmlPartCore.GetCustomXmlPart(mainDocumentPart);

            if (customXmlPart != null)
            {
                // Root element shall never be null if Custom Xml part is present
                rootElement = this.customXmlPartCore.GetFirstElementFromCustomXmlPart(customXmlPart, rootElementName);

                childElement = (from e in rootElement.Descendants(childElementXName)
                                select e).FirstOrDefault();

                if (childElement != null)
                {
                    foreach (KeyValuePair<string, string> idToValue in nameToValueCollection)
                    {
                        if (forNodeType == NodeType.Attribute)
                        {
                            AddOrUpdateAttribute(childElement, idToValue.Key, idToValue.Value);
                        }
                        else if (forNodeType == NodeType.Element)
                        {
                            AddOrUpdateChildElement(childElement, idToValue.Key, idToValue.Value);
                        }
                    }

                    this.customXmlPartCore.WriteElementToCustomXmlPart(customXmlPart, rootElement);
                }
                else
                {
                    childElement = GetElementFromNameToValueCollectionForType(nameToValueCollection, childElementXName, forNodeType);
                    rootElement.Add(childElement);
                }
            }
            else
            {
                customXmlPart = this.customXmlPartCore.AddCustomXmlPart(mainDocumentPart, rootElementName);
                childElement = GetElementFromNameToValueCollectionForType(nameToValueCollection, childElementXName, forNodeType);
                rootElement.Add(childElement);
            }

            this.customXmlPartCore.WriteElementToCustomXmlPart(customXmlPart, rootElement);
        }

        /// <summary>
        /// Gets the type of the name to value collection from element for.
        /// </summary>
        /// <param name="mainDocumentPart">The main document part.</param>
        /// <param name="elementName">Name of the element.</param>
        /// <param name="forNodeType">Type of for node.</param>
        /// <returns></returns>
        public Dictionary<string, string> GetNameToValueCollectionFromElementForType(MainDocumentPart mainDocumentPart, string elementName, NodeType forNodeType)
        {
            Dictionary<string, string> nameToValueCollection = new Dictionary<string, string>();
            CustomXmlPart customXmlPart = this.customXmlPartCore.GetCustomXmlPart(mainDocumentPart);

            if (customXmlPart != null)
            {
                XElement element = this.customXmlPartCore.GetFirstElementFromCustomXmlPart(customXmlPart, elementName);

                if (element != null)
                {
                    if (forNodeType == NodeType.Element)
                    {
                        foreach (XElement elem in element.Elements())
                        {
                            nameToValueCollection.Add(elem.Name.LocalName, elem.Nodes().Where(node => node.NodeType == XmlNodeType.Element).FirstOrDefault().ToString());
                        }
                    }
                    else if (forNodeType == NodeType.Attribute)
                    {
                        foreach (XAttribute attr in element.Attributes())
                        {
                            nameToValueCollection.Add(attr.Name.LocalName, attr.Value);
                        }
                    }
                }
            }

            return nameToValueCollection;
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Gets the type of the element from name to value collection for.
        /// </summary>
        /// <param name="nameToValueCollection">The name to value collection.</param>
        /// <param name="elementXName">Name of the element X.</param>
        /// <param name="nodeType">Type of the node.</param>
        /// <returns></returns>
        private XElement GetElementFromNameToValueCollectionForType(Dictionary<string, string> nameToValueCollection, XName elementXName, NodeType nodeType)
        {
            XElement element = new XElement(elementXName);

            foreach (KeyValuePair<string, string> idToValue in nameToValueCollection)
            {
                if (nodeType == NodeType.Element)
                {
                    AddOrUpdateChildElement(element, idToValue.Key, idToValue.Value);
                }
                else if (nodeType == NodeType.Attribute)
                {
                    AddOrUpdateAttribute(element, idToValue.Key, idToValue.Value);
                }
            }

            return element;
        }

        /// <summary>
        /// Adds the or update attribute.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="attributeName">Name of the attribute.</param>
        /// <param name="attributeValue">The attribute value.</param>
        private void AddOrUpdateAttribute(XElement element, string attributeName, string attributeValue)
        {
            XAttribute attrToUpdate = element.Attributes().Where(attr => attr.Name.LocalName.Equals(attributeName)).FirstOrDefault();

            if (attrToUpdate != null)
            {
                attrToUpdate.Value = attributeValue;
            }
            else
            {
                XAttribute attr = new XAttribute(attributeName, attributeValue);
                element.Add(attr);
            }
        }

        /// <summary>
        /// Adds the or update child element.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="childElementName">Name of the child element.</param>
        /// <param name="childElementValue">The child element value.</param>
        private void AddOrUpdateChildElement(XElement element, string childElementName, string childElementValue)
        {
            XElement childElement = element.Elements().Where(elem => elem.Name.LocalName.Equals(childElementName)).FirstOrDefault();
            XElement newChildElement = new XElement(XName.Get(childElementName, this.customXmlPartCore.namespaceUri));
            newChildElement.Add(XElement.Parse(childElementValue));

            if (childElement != null)
            {
                childElement.ReplaceWith(newChildElement);
            }
            else
            {
                element.Add(newChildElement);
            }
        }

        #endregion
    }
}
