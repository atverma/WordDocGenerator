// ----------------------------------------------------------------------
// <copyright file="CustomXmlPartCore.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Library
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Xml;
    using System.Xml.Linq;
    using DocumentFormat.OpenXml.CustomXmlDataProperties;
    using DocumentFormat.OpenXml.Packaging;

    /// <summary>
    /// Helper class for Word CustomXml part operations
    /// </summary>
    public class CustomXmlPartCore
    {
        #region Members

        /// <summary>
        /// Namespace Uri
        /// </summary>
        public readonly string namespaceUri = string.Empty;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomXmlPartCore"/> class.
        /// </summary>
        /// <param name="namespaceUri">The namespace URI.</param>
        public CustomXmlPartCore(string namespaceUri)
        {
            this.namespaceUri = namespaceUri;
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Adds the custom XML part.
        /// </summary>
        /// <param name="mainDocumentPart">The main part.</param>
        /// <param name="rootElementName">Name of the root element.</param>
        /// <returns>
        /// Returns CustomXmlPart
        /// </returns>
        public CustomXmlPart AddCustomXmlPart(MainDocumentPart mainDocumentPart, string rootElementName)
        {
            if (mainDocumentPart == null)
            {
                throw new ArgumentNullException("mainDocumentPart");
            }

            if (string.IsNullOrEmpty(rootElementName))
            {
                throw new ArgumentNullException("rootElementName");
            }

            XName rootElementXName = XName.Get(rootElementName, this.namespaceUri);
            XElement rootElement = new XElement(rootElementXName);
            CustomXmlPart customXmlPart = mainDocumentPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            CustomXmlPropertiesPart customXmlPropertiesPart = customXmlPart.AddNewPart<CustomXmlPropertiesPart>();
            GenerateCustomXmlPropertiesPartContent(customXmlPropertiesPart);
            WriteElementToCustomXmlPart(customXmlPart, rootElement);

            return customXmlPart;
        }

        /// <summary>
        /// Removes the custom XML part.
        /// </summary>
        /// <param name="mainDocumentPart">The main part.</param>
        /// <param name="customXmlPart">The custom XML part.</param>
        public void RemoveCustomXmlPart(MainDocumentPart mainDocumentPart, CustomXmlPart customXmlPart)
        {
            if (mainDocumentPart == null)
            {
                throw new ArgumentNullException("mainDocumentPart");
            }

            if (customXmlPart != null)
            {
                RemoveCustomXmlParts(mainDocumentPart, new List<CustomXmlPart>(new CustomXmlPart[] { customXmlPart }));
            }
        }

        /// <summary>
        /// Removes the custom XML parts.
        /// </summary>
        /// <param name="mainDocumentPart">The main document part.</param>
        /// <param name="customXmlParts">The custom XML parts.</param>
        public void RemoveCustomXmlParts(OpenXmlPartContainer mainDocumentPart, IList<CustomXmlPart> customXmlParts)
        {
            if (mainDocumentPart == null)
            {
                throw new ArgumentNullException("mainDocumentPart");
            }

            if (customXmlParts != null)
            {
                mainDocumentPart.DeleteParts<CustomXmlPart>(customXmlParts);
            }
        }

        /// <summary>
        /// Gets the custom XML part.
        /// </summary>
        /// <param name="mainDocumentPart">The main document part.</param>
        /// <returns></returns>
        public CustomXmlPart GetCustomXmlPart(MainDocumentPart mainDocumentPart)
        {
            if (mainDocumentPart == null)
            {
                throw new ArgumentNullException("mainDocumentPart");
            }

            CustomXmlPart result = null;

            foreach (CustomXmlPart part in mainDocumentPart.CustomXmlParts)
            {
                using (XmlTextReader reader = new XmlTextReader(part.GetStream(FileMode.Open, FileAccess.Read)))
                {
                    reader.MoveToContent();
                    bool exists = reader.NamespaceURI.Equals(this.namespaceUri);

                    if (exists)
                    {
                        result = part;
                        break;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Gets the store item id.
        /// </summary>
        /// <param name="mainDocumentPart">The main document part.</param>
        /// <returns></returns>
        public string GetStoreItemId(MainDocumentPart mainDocumentPart)
        {
            if (mainDocumentPart == null)
            {
                throw new ArgumentNullException("mainDocumentPart");
            }

            CustomXmlPart customXmlPart = GetCustomXmlPart(mainDocumentPart);
            CustomXmlPropertiesPart customXmlPropertiesPart = customXmlPart.CustomXmlPropertiesPart;
            return customXmlPropertiesPart.DataStoreItem.ItemId.ToString();
        }

        /// <summary>
        /// Gets the first element from custom XML part.
        /// </summary>
        /// <param name="customXmlPart">The custom XML part.</param>
        /// <param name="elementName">Name of the element.</param>
        /// <returns></returns>
        public XElement GetFirstElementFromCustomXmlPart(CustomXmlPart customXmlPart, string elementName)
        {
            if (customXmlPart == null)
            {
                throw new ArgumentNullException("customXmlPart");
            }

            if (string.IsNullOrEmpty(elementName))
            {
                throw new ArgumentNullException("elementName");
            }

            XDocument customPartDoc = null;

            using (XmlReader reader = XmlReader.Create(customXmlPart.GetStream(FileMode.Open, FileAccess.Read)))
            {
                customPartDoc = XDocument.Load(reader);
            }

            XElement element = null;

            if (customPartDoc != null)
            {
                XName elementXName = XName.Get(elementName, this.namespaceUri);
                element = (from e in customPartDoc.Descendants(elementXName)
                                    select e).FirstOrDefault();
            }

            return element;
        }

        /// <summary>
        /// Writes the element to custom XML part.
        /// </summary>
        /// <param name="customXmlPart">The custom XML part.</param>
        /// <param name="rootElement">The root element.</param>
        public void WriteElementToCustomXmlPart(CustomXmlPart customXmlPart, XElement rootElement)
        {
            if (customXmlPart == null)
            {
                throw new ArgumentNullException("customXmlPart");
            }

            if (rootElement == null)
            {
                throw new ArgumentNullException("rootElement");
            }

            using (XmlWriter writer = XmlWriter.Create(customXmlPart.GetStream(FileMode.Create, FileAccess.Write)))
            {
                rootElement.WriteTo(writer);
                writer.Flush();
            }
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Generates the content of the custom XML properties part.
        /// </summary>
        /// <param name="customXmlPropertiesPart">The custom XML properties part1.</param>
        private void GenerateCustomXmlPropertiesPartContent(CustomXmlPropertiesPart customXmlPropertiesPart)
        {
            DataStoreItem dataStoreItem = new DataStoreItem() { ItemId ="{" + Guid.NewGuid().ToString() + "}" };
            dataStoreItem.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");
            SchemaReferences schemaReferences = new SchemaReferences();
            SchemaReference schemaReference = new SchemaReference() { Uri = namespaceUri };
            schemaReferences.Append(schemaReference);
            dataStoreItem.Append(schemaReferences);
            customXmlPropertiesPart.DataStoreItem = dataStoreItem;
        }

        #endregion
    }
}