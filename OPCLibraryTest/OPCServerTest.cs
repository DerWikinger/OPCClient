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

            string hostname = "127.0.0.1";
            List<OPCServer> servers = OPCServer.BrowseServers(hostname);
            Assert.AreNotEqual(null, servers);
            Assert.AreNotEqual(0, servers.Count);

            string progID = "Insat";
            server = OPCServer.FindServerByProgID(progID, hostname);
            Assert.AreNotEqual(null, server);

            server.Connect();
            List<OPCItem> arr = server.GetItems();
            Assert.AreNotEqual(null, arr);
            Assert.AreEqual(27, arr.Count); // 27 items are defiened in OPC Server config
            server.Disconnect();
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
            ServerInfo sInfo = new ServerInfo();
            Assert.AreNotEqual(null, sInfo);
            COSERVERINFO info = sInfo.Allocate("localhost", new System.Net.NetworkCredential());
            Assert.AreNotEqual(null, info); 
        }
    }
}
