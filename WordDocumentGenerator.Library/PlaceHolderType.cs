// ----------------------------------------------------------------------
// <copyright file="PlaceHolderType.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Library
{
    /// <summary>
    /// Defines the type of the PlaceHolder
    /// </summary>
    public enum PlaceHolderType
    {
        None = 0,
        Recursive = 1,
        NonRecursive = 2,
        Ignore = 3,
        Container = 4
    }
}
