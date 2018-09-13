// ----------------------------------------------------------------------
// <copyright file="TestData.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Client
{
    using System;
    using System.Collections.Generic;

    public class Order
    {
        public List<Vendor> vendors = null;
        public List<Item> items = null;

        public Guid Id = Guid.Empty;
        public string Name = string.Empty;
        public string Description = string.Empty;
    }

    public class Vendor
    {
        public Guid Id = Guid.Empty;
        public string Name = string.Empty;        
    }

    public class Item
    {
        public Guid Id = Guid.Empty;
        public string Name = string.Empty;
    }
}
