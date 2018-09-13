// ----------------------------------------------------------------------
// <copyright file="DocumentGenerationInfo.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Library
{
    using System.Collections.Generic;

    public class DocumentGenerationInfo
    {
        /// <summary>
        /// Namespace Uri for CustomXML part
        /// </summary>
        public const string NamespaceUri = "http://schemas.WordDocumentGenerator.com/DocumentGeneration";

        private DocumentMetadata metadata;
        private byte[] templateData;
        private object dataContext;
        private Dictionary<string, PlaceHolderType> placeHolderTagToTypeCollection;
        private bool isDataBoundControls;

        /// <summary>
        /// Gets or sets the place holder tag to type collection.
        /// </summary>
        /// <value>
        /// The place holder tag to type collection.
        /// </value>
        public Dictionary<string, PlaceHolderType> PlaceHolderTagToTypeCollection
        {
            get { return placeHolderTagToTypeCollection; }
            set { placeHolderTagToTypeCollection = value; }
        }

        /// <summary>
        /// Gets or sets the metadata.
        /// </summary>
        /// <value>
        /// The metadata.
        /// </value>
        public DocumentMetadata Metadata
        {
            get { return metadata; }
            set { metadata = value; }
        }

        /// <summary>
        /// Gets or sets the template data.
        /// </summary>
        /// <value>
        /// The template data.
        /// </value>
        public byte[] TemplateData
        {
            get { return templateData; }
            set { templateData = value; }
        }

        /// <summary>
        /// Gets or sets the data context.
        /// </summary>
        /// <value>
        /// The data context.
        /// </value>
        public object DataContext
        {
            get { return dataContext; }
            set { dataContext = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is data bound controls.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is data bound controls; otherwise, <c>false</c>.
        /// </value>
        public bool IsDataBoundControls
        {
            get { return isDataBoundControls; }
            set { isDataBoundControls = value; }
        }
    }
}
