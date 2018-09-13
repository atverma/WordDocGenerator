// ----------------------------------------------------------------------
// <copyright file="DocumentGenerator.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Library
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Xml;
    using System.Xml.Serialization;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// Base class for document generation
    /// </summary>
    public abstract class DocumentGenerator
    {
        #region Constants

        /// <summary>
        /// Root Node of CustomXML Part
        /// </summary>
        protected const string DocumentRootNode = "DocumentRootNode";

        /// <summary>
        /// Document Node
        /// </summary>
        protected const string DocumentNode = "Document";

        /// <summary>
        /// Document Container PlaceHolders Node
        /// </summary>
        protected const string DocumentContainerPlaceHoldersNode = "DocumentContainerPlaceHolders";

        /// <summary>
        /// Data bound controls data store Node
        /// </summary>
        protected const string DataBoundControlsDataStoreNode = "DataBoundControlsDataStore";

        /// <summary>
        /// Data node in Data bound controls data store
        /// </summary>
        protected const string DataNode = "Data";

        /// <summary>
        /// Document Type Attribute
        /// </summary>
        protected const string DocumentTypeNodeName = "DocumentType";

        /// <summary>
        /// Document Version Attribute
        /// </summary>
        protected const string DocumentVersionNodeName = "Version";

        #endregion

        #region Members

        /// <summary>
        /// Instance of Document generation info
        /// </summary>
        private DocumentGenerationInfo generationInfo;

        /// <summary>
        /// Instance of CustomXml Part Helper
        /// </summary>
        private readonly CustomXmlPartHelper customXmlPartHelper = new CustomXmlPartHelper(DocumentGenerationInfo.NamespaceUri);

        /// <summary>
        /// Instance of Open Xml Helper
        /// </summary>
        private readonly OpenXmlHelper openXmlHelper = new OpenXmlHelper(DocumentGenerationInfo.NamespaceUri);

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentGenerator"/> class.
        /// </summary>
        /// <param name="generationInfo">The generation info.</param>
        public DocumentGenerator(DocumentGenerationInfo generationInfo)
        {
            this.generationInfo = generationInfo;
        }

        #endregion

        #region Protected Methods

        /// <summary>
        /// Gets the place holder tag to type collection.
        /// </summary>
        /// <returns></returns>
        protected abstract Dictionary<string, PlaceHolderType> GetPlaceHolderTagToTypeCollection();

        /// <summary>
        /// Ignore placeholder found.
        /// </summary>
        /// <param name="placeholderTag">The placeholder tag.</param>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        protected abstract void IgnorePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext);

        /// <summary>
        /// Non recursive placeholder found.
        /// </summary>
        /// <param name="placeholderTag">The placeholder tag.</param>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        protected abstract void NonRecursivePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext);

        /// <summary>
        /// Recursive placeholder found.
        /// </summary>
        /// <param name="placeholderTag">The placeholder tag.</param>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        protected abstract void RecursivePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext);

        /// <summary>
        /// Container placeholder found.
        /// </summary>
        /// <param name="placeholderTag">The placeholder tag.</param>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        protected abstract void ContainerPlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext);

        /// <summary>
        /// Gets the serialized data context.
        /// </summary>
        /// <returns></returns>
        protected virtual string SerializeDataContextToXml()
        {
            StringBuilder sb = new StringBuilder();

            if (generationInfo != null && generationInfo.DataContext != null)
            {
                XmlSerializer serializer = new XmlSerializer(generationInfo.DataContext.GetType());
                XmlWriterSettings writerSettings = new XmlWriterSettings();
                writerSettings.OmitXmlDeclaration = true;

                using (XmlWriter writer = XmlWriter.Create(sb, writerSettings))
                {
                    serializer.Serialize(writer, generationInfo.DataContext);
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// Gets the parent container.
        /// </summary>
        /// <param name="parentContainer">The parent container.</param>
        /// <param name="placeHolder">The place holder.</param>
        /// <returns></returns>
        protected bool GetParentContainer(ref SdtElement parentContainer, string placeHolder)
        {
            bool isRefresh = false;
            MainDocumentPart mainDocumentPart = parentContainer.Ancestors<Document>().First().MainDocumentPart;
            KeyValuePair<string, string> nameToValue = this.customXmlPartHelper.GetNameToValueCollectionFromElementForType(mainDocumentPart, DocumentContainerPlaceHoldersNode, NodeType.Element).Where(f => f.Key.Equals(placeHolder)).FirstOrDefault();

            isRefresh = !string.IsNullOrEmpty(nameToValue.Value);

            if (isRefresh)
            {
                SdtElement parentElementFromCustomXmlPart = new SdtBlock(nameToValue.Value);
                parentContainer.Parent.ReplaceChild(parentElementFromCustomXmlPart, parentContainer);
                parentContainer = parentElementFromCustomXmlPart;
            }
            else
            {
                Dictionary<string, string> nameToValueCollection = new Dictionary<string, string>();
                nameToValueCollection.Add(placeHolder, parentContainer.OuterXml);
                this.customXmlPartHelper.SetElementFromNameToValueCollectionForType(mainDocumentPart, DocumentRootNode, DocumentContainerPlaceHoldersNode, nameToValueCollection, NodeType.Element);
            }

            return isRefresh;
        }

        /// <summary>
        /// Gets the tag value.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="templateTagPart">The template tag part.</param>
        /// <param name="tagGuidPart">The tag GUID part.</param>
        /// <returns></returns>
        protected string GetTagValue(SdtElement element, out string templateTagPart, out string tagGuidPart)
        {
            templateTagPart = string.Empty;
            tagGuidPart = string.Empty;
            Tag tag = openXmlHelper.GetTag(element);

            string fullTag = (tag == null || (tag.Val.HasValue == false)) ? string.Empty : tag.Val.Value;

            if (!string.IsNullOrEmpty(fullTag))
            {
                string[] tagParts = fullTag.Split(':');

                if (tagParts.Length == 2)
                {
                    templateTagPart = tagParts[0];
                    tagGuidPart = tagParts[1];
                }
                else if (tagParts.Length == 1)
                {
                    templateTagPart = tagParts[0];
                }
            }

            return fullTag;
        }

        /// <summary>
        /// Gets the full tag value.
        /// </summary>
        /// <param name="templateTagPart">The template tag part.</param>
        /// <param name="tagGuidPart">The tag GUID part.</param>
        /// <returns></returns>
        protected string GetFullTagValue(string templateTagPart, string tagGuidPart)
        {
            return templateTagPart + ":" + tagGuidPart;
        }

        /// <summary>
        /// Saves the data content to data bound controls data store.
        /// </summary>
        /// <param name="mainDocumentPart">The main document part.</param>
        protected void SaveDataToDataBoundControlsDataStore(MainDocumentPart mainDocumentPart)
        {
            string dataContextAsXml = this.SerializeDataContextToXml();
            Dictionary<string, string> nameToValueCollection = new Dictionary<string, string>();
            nameToValueCollection.Add(DataNode, dataContextAsXml);
            this.customXmlPartHelper.SetElementFromNameToValueCollectionForType(mainDocumentPart, DocumentRootNode, DataBoundControlsDataStoreNode, nameToValueCollection, NodeType.Element);
        }

        /// <summary>
        /// Sets the data binding.
        /// </summary>
        /// <param name="xPath">The x path.</param>
        /// <param name="element">The element.</param>
        protected void SetDataBinding(string xPath, SdtElement element)
        {
            element.SdtProperties.RemoveAllChildren<DataBinding>();
            DataBinding dataBinding = new DataBinding() { XPath = xPath, StoreItemId = new StringValue(this.customXmlPartHelper.customXmlPartCore.GetStoreItemId(element.Ancestors<Document>().First().MainDocumentPart)) };
            element.SdtProperties.Append(dataBinding);
        }

        /// <summary>
        /// Gets the data context.
        /// </summary>
        /// <returns></returns>
        protected object GetDataContext()
        {
            return generationInfo != null ? this.generationInfo.DataContext : null;
        }

        /// <summary>
        /// Sets the tag value.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="fullTagValue">The full tag value.</param>
        protected void SetTagValue(SdtElement element, string fullTagValue)
        {
            // Set the tag for the content control
            if (!string.IsNullOrEmpty(fullTagValue))
            {
                this.openXmlHelper.SetTagValue(element, fullTagValue);
            }
        }

        /// <summary>
        /// Sets the content of content control.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="content">The content.</param>
        protected void SetContentOfContentControl(SdtElement element, string content)
        {
            // Set text without data binding
            this.openXmlHelper.SetContentOfContentControl(element, content);
        }

        /// <summary>
        /// Sets the content in placeholders.
        /// </summary>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        protected void SetContentInPlaceholders(OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (IsContentControl(openXmlElementDataContext))
            {
                string templateTagPart = string.Empty;
                string tagGuidPart = string.Empty;
                SdtElement element = openXmlElementDataContext.Element as SdtElement;
                GetTagValue(element, out templateTagPart, out tagGuidPart);

                if (this.generationInfo.PlaceHolderTagToTypeCollection.ContainsKey(templateTagPart))
                {
                    this.OnPlaceHolderFound(openXmlElementDataContext);
                }
                else
                {
                    this.PopulateOtherOpenXmlElements(openXmlElementDataContext);
                }
            }
            else
            {
                this.PopulateOtherOpenXmlElements(openXmlElementDataContext);
            }
        }

        /// <summary>
        /// Clones the element and set content in placeholders.
        /// </summary>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        /// <returns></returns>
        protected SdtElement CloneElementAndSetContentInPlaceholders(OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (openXmlElementDataContext == null)
            {
                throw new ArgumentNullException("openXmlElementDataContext");
            }

            if (openXmlElementDataContext.Element == null)
            {
                throw new ArgumentNullException("openXmlElementDataContext.element");
            }

            SdtElement clonedSdtElement = null;

            if (openXmlElementDataContext.Element.Parent != null && openXmlElementDataContext.Element.Parent is Paragraph)
            {
                Paragraph clonedPara = openXmlElementDataContext.Element.Parent.InsertBeforeSelf(openXmlElementDataContext.Element.Parent.CloneNode(true) as Paragraph);
                clonedSdtElement = clonedPara.Descendants<SdtElement>().First();
            }
            else
            {
                clonedSdtElement = openXmlElementDataContext.Element.InsertBeforeSelf(openXmlElementDataContext.Element.CloneNode(true) as SdtElement);
            }

            foreach (var v in clonedSdtElement.Elements())
            {
                this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = v, DataContext = openXmlElementDataContext.DataContext });
            }

            return clonedSdtElement;
        }

        /// <summary>
        /// Sets the document properties.
        /// </summary>
        /// <param name="mainDocumentPart">The main document part.</param>
        /// <param name="docProperties">The doc properties.</param>
        protected void SetDocumentProperties(MainDocumentPart mainDocumentPart, DocumentMetadata docProperties)
        {
            if (mainDocumentPart == null)
            {
                throw new ArgumentNullException("mainDocumentPart");
            }

            if (docProperties == null)
            {
                throw new ArgumentNullException("docProperties");
            }

            Dictionary<string, string> idtoValues = new Dictionary<string, string>();
            idtoValues.Add(DocumentTypeNodeName, string.IsNullOrEmpty(docProperties.DocumentType) ? string.Empty : docProperties.DocumentType);
            idtoValues.Add(DocumentVersionNodeName, string.IsNullOrEmpty(docProperties.DocumentVersion) ? string.Empty : docProperties.DocumentVersion);
            this.customXmlPartHelper.SetElementFromNameToValueCollectionForType(mainDocumentPart, DocumentRootNode, DocumentNode, idtoValues, NodeType.Attribute);
        }

        /// <summary>
        /// Determines whether [is template tag equal] [the specified element].
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="placeholderName">Name of the placeholder.</param>
        /// <returns>
        ///   <c>true</c> if [is template tag equal] [the specified element]; otherwise, <c>false</c>.
        /// </returns>
        protected bool IsTemplateTagEqual(SdtElement element, string placeholderName)
        {
            if (element == null)
            {
                throw new ArgumentNullException("element");
            }

            if (placeholderName == null)
            {
                throw new ArgumentNullException("placeholderName");
            }

            string templateTagPart = string.Empty;
            string tagGuidPart = string.Empty;
            GetTagValue(element, out templateTagPart, out tagGuidPart);
            return placeholderName.Equals(templateTagPart);
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Generates the document.
        /// </summary>
        /// <returns></returns>
        public byte[] GenerateDocument()
        {
            if (this.generationInfo == null)
            {
                throw new ArgumentNullException("generationInfo");
            }

            if (this.generationInfo.TemplateData == null)
            {
                throw new ArgumentNullException("templateData");
            }

            this.generationInfo.PlaceHolderTagToTypeCollection = this.GetPlaceHolderTagToTypeCollection();

            if (this.generationInfo.PlaceHolderTagToTypeCollection == null)
            {
                throw new ArgumentNullException("PlaceHolderTagToTypeCollection");
            }

            return SetContentInPlaceholders();
        }

        #endregion

        #region Private Methods        

        /// <summary>
        /// Sets the content in placeholders.
        /// </summary>
        /// <returns></returns>
        private byte[] SetContentInPlaceholders()
        {
            byte[] output = null;

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(this.generationInfo.TemplateData, 0, this.generationInfo.TemplateData.Length);

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(ms, true))
                {
                    wordDocument.ChangeDocumentType(WordprocessingDocumentType.Document);
                    MainDocumentPart mainDocumentPart = wordDocument.MainDocumentPart;
                    Document document = mainDocumentPart.Document;

                    if (this.generationInfo.Metadata != null)
                    {
                        SetDocumentProperties(mainDocumentPart, this.generationInfo.Metadata);
                    }

                    if (this.generationInfo.IsDataBoundControls)
                    {
                        SaveDataToDataBoundControlsDataStore(mainDocumentPart);
                    }

                    foreach (HeaderPart part in mainDocumentPart.HeaderParts)
                    {
                        this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = part.Header, DataContext = this.generationInfo.DataContext });
                        part.Header.Save();
                    }

                    foreach (FooterPart part in mainDocumentPart.FooterParts)
                    {
                        this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = part.Footer, DataContext = this.generationInfo.DataContext });
                        part.Footer.Save();
                    }

                    this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = document, DataContext = this.generationInfo.DataContext });

                    this.openXmlHelper.EnsureUniqueContentControlIdsForMainDocumentPart(mainDocumentPart);

                    document.Save();
                }

                ms.Position = 0;
                output = new byte[ms.Length];
                ms.Read(output, 0, output.Length);
            }

            return output;
        }

        /// <summary>
        /// Populates the other open XML elements.
        /// </summary>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        private void PopulateOtherOpenXmlElements(OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (openXmlElementDataContext.Element is OpenXmlCompositeElement && openXmlElementDataContext.Element.HasChildren)
            {
                List<OpenXmlElement> elements = openXmlElementDataContext.Element.Elements().ToList();

                foreach (var element in elements)
                {
                    if (element is OpenXmlCompositeElement)
                    {
                        this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = element, DataContext = openXmlElementDataContext.DataContext });
                    }
                }
            }
        }

        /// <summary>
        /// Determines whether [is content control] [the specified open XML element data context].
        /// </summary>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        /// <returns>
        ///   <c>true</c> if [is content control] [the specified open XML element data context]; otherwise, <c>false</c>.
        /// </returns>
        private bool IsContentControl(OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (openXmlElementDataContext == null || openXmlElementDataContext.Element == null)
            {
                return false;
            }

            return openXmlElementDataContext.Element is SdtBlock || openXmlElementDataContext.Element is SdtRun || openXmlElementDataContext.Element is SdtRow || openXmlElementDataContext.Element is SdtCell;
        }

        /// <summary>
        /// Called when [place holder found].
        /// </summary>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        private void OnPlaceHolderFound(OpenXmlElementDataContext openXmlElementDataContext)
        {
            string templateTagPart = string.Empty;
            string tagGuidPart = string.Empty;
            SdtElement element = openXmlElementDataContext.Element as SdtElement;
            GetTagValue(element, out templateTagPart, out tagGuidPart);

            if (this.generationInfo.PlaceHolderTagToTypeCollection.ContainsKey(templateTagPart))
            {
                switch (this.generationInfo.PlaceHolderTagToTypeCollection[templateTagPart])
                {
                    case PlaceHolderType.None:
                        break;
                    case PlaceHolderType.NonRecursive:
                        this.NonRecursivePlaceholderFound(templateTagPart, openXmlElementDataContext);
                        break;
                    case PlaceHolderType.Recursive:
                        this.RecursivePlaceholderFound(templateTagPart, openXmlElementDataContext);
                        break;
                    case PlaceHolderType.Ignore:
                        this.IgnorePlaceholderFound(templateTagPart, openXmlElementDataContext);
                        break;
                    case PlaceHolderType.Container:
                        this.ContainerPlaceholderFound(templateTagPart, openXmlElementDataContext);
                        break;
                }
            }
        }

        #endregion
    }
}
