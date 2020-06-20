using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OPCLibrary;
using OPCLibrary.DZ.Opc.Integration.Internal;
using System.Runtime.InteropServices;



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
            List<OPCItem> arr = server.GetItems(true);
            Assert.AreNotEqual(null, arr);
            Assert.AreEqual(27, arr.Count); // 27 items are defiened in OPC Server config
            List<OPCItem> items = new List<OPCItem>(arr.FindAll(i => i.ItemType == OPCItemType.LEAF));
            server.SyncRead(items);
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

        [TestMethod]
        public void TestConverter()
        {
            Assert.AreEqual("Bad", OPCLibrary.Converter.GetQualityString(0x00));
            Assert.AreEqual("Config Error", OPCLibrary.Converter.GetQualityString(0x04));
            Assert.AreEqual("Not Connected", OPCLibrary.Converter.GetQualityString(0x08));
            Assert.AreEqual("Device Failure", OPCLibrary.Converter.GetQualityString(0x0C));
            Assert.AreEqual("Sensor Failure", OPCLibrary.Converter.GetQualityString(0x10));
            Assert.AreEqual("Last Known", OPCLibrary.Converter.GetQualityString(0x14));
            Assert.AreEqual("Comm Failure", OPCLibrary.Converter.GetQualityString(0x18));
            Assert.AreEqual("Out of Service", OPCLibrary.Converter.GetQualityString(0x1C));
            Assert.AreEqual("Initializing", OPCLibrary.Converter.GetQualityString(0x20));
            Assert.AreEqual("Uncertain", OPCLibrary.Converter.GetQualityString(0x40));
            Assert.AreEqual("Last Usable", OPCLibrary.Converter.GetQualityString(0x44));
            Assert.AreEqual("Sensor Calibration", OPCLibrary.Converter.GetQualityString(0x50));
            Assert.AreEqual("EGU Exceeded", OPCLibrary.Converter.GetQualityString(0x54));
            Assert.AreEqual("Sub Normal", OPCLibrary.Converter.GetQualityString(0x58));
            Assert.AreEqual("Good", OPCLibrary.Converter.GetQualityString(0xC0));
            Assert.AreEqual("Local Override", OPCLibrary.Converter.GetQualityString(0xD8));
            Assert.AreEqual("Unknown", OPCLibrary.Converter.GetQualityString(0xFF));

            Assert.AreEqual("VT_EMPTY", OPCLibrary.Converter.GetVTString(0));
        }

    }
}
