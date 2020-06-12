using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OPCLibrary;



namespace OPCLibraryTest
{
    [TestClass]
    public class OPCServerTest
    {
        [TestMethod]
        public void TestConstructor()
        {
            OPCServer server = new OPCServer();
            Assert.AreNotEqual(null, server);
            Assert.AreEqual("default", server.Name);
        }
    }
}
