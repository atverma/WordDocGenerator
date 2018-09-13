// ----------------------------------------------------------------------
// <copyright file="OpenXmlElementDataContext.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Library
{
    using DocumentFormat.OpenXml;

    /// <summary>
    /// OpenXml element and data context
    /// </summary>
    public class OpenXmlElementDataContext
    {
        private OpenXmlElement element;
        private object dataContext;

        /// <summary>
        /// Gets or sets the element.
        /// </summary>
        /// <value>
        /// The element.
        /// </value>
        public OpenXmlElement Element
        {
            get { return element; }
            set { element = value; }
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
    }
}
