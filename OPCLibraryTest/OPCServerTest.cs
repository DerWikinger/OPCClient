using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OPCLibrary;
using OPCLibrary.DZ.Opc.Integration.Internal;



namespace OPCLibraryTest
{
    [TestClass]
    public class OPCServerTest
    {
        [TestMethod]
        public void TestOPCServer()
        {
            OPCServer server = new OPCServer();
            Assert.AreNotEqual(null, server);
            server.ServerName = "Name";
            Assert.AreEqual("Name", server.ServerName);
            server.ServerDescription = "Description";
            Assert.AreEqual("Description", server.ServerDescription);
            Guid g = new Guid("00000000-0000-0000-0000-000000000000");
            server.Guid = g;
            Assert.AreEqual(g, server.Guid);
            server.HostName = "localhost";
            Assert.AreEqual("localhost", server.HostName);
            server.ProgId = "ProgID";
            Assert.AreEqual("ProgID", server.ProgId);

            List<OPCServer> servers = OPCServer.BrowseServers("127.0.0.1");
            Assert.AreNotEqual(null, servers);
            Assert.AreNotEqual(0, servers.Count);
        }

        [TestMethod]
        public void TestInterop()
        {
            string message = OPCLibrary.DZ.Opc.Integration.Internal.Interop.GetSystemMessage(0);
            Assert.AreNotEqual(null, message);
            Assert.AreNotEqual("", message);
        }

        [TestMethod]
        public void TestServerInfo()
        {
            OPCLibrary.DZ.Opc.Integration.Internal.ServerInfo sInfo = new OPCLibrary.DZ.Opc.Integration.Internal.ServerInfo();
            Assert.AreNotEqual(null, sInfo);
            COSERVERINFO info = sInfo.Allocate("localhost", new System.Net.NetworkCredential());
            Assert.AreNotEqual(null, info); 
        }
    }
}
