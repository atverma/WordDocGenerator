// ----------------------------------------------------------------------
// <copyright file="DocumentMetadata.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Library
{
    /// <summary>
    /// Defines the metadata for a Word document
    /// </summary>
    public class DocumentMetadata
    {
        #region Members

        /// <summary>
        /// Gets or sets the type of the document.
        /// </summary>
        /// <value>
        /// The type of the document.
        /// </value>
        public string DocumentType { get; set; }

        /// <summary>
        /// Gets or sets the document version.
        /// </summary>
        /// <value>
        /// The document version.
        /// </value>
        public string DocumentVersion { get; set; }

        #endregion
    }
}
