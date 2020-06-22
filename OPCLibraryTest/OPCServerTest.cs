using System;
using System.Threading;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OPCLibrary;
using OPCLibrary.Internal;
using System.Runtime.InteropServices;
using System.Linq;



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
            server.Disconnect();
        }

        [TestMethod]
        public void TestInterop()
        {
            string message = OPCLibrary.Internal.Interop.GetSystemMessage(0);
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

        [TestMethod]
        public void TestOPCGroupSyncRead()
        {
            string progID = "Insat";
            int errorCount = 0;
            OPCServer server = OPCServer.FindServerByProgID(progID, "localhost");
            server.Connect();
            List<OPCItem> items = server.GetItems();
            var opcBranches = from item in items
                              where item.ItemType == OPCItemType.BRANCH
                              let leaves = from elem in items
                                           where elem.ItemType == OPCItemType.LEAF
                                           where elem.Parent != null && elem.Parent.ItemName == item.ItemName
                                           select elem
                              where leaves.Count() > 0
                              where !item.ItemName.Contains("МВ110")
                              select item;
            foreach(OPCItem branch in opcBranches)
            {
                string groupName = branch.ItemID;                
                var opcItems = from item in items
                               where item.ItemType == OPCItemType.LEAF
                               where item.Parent != null
                               where item.Parent.ItemName == branch.ItemName
                               select item;
                List<OPCItem> list = new List<OPCItem>(opcItems);
                uint udateRate = 100;
                OPCGroup group = new OPCGroup(name: groupName, items: list)
                {
                    UpdateRate = udateRate
                };
                Assert.AreEqual(groupName, group.Name);
                Assert.AreSame(list, group.Items);
                Assert.AreEqual(udateRate, group.UpdateRate);
                group.RegInServer(server);
                try
                {
                    group.SyncRead();
                }
                catch(ServerException ex)
                {
                    string msg;
                    //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                    server.Server.GetErrorString(ex.HResult, 2, out msg);
                    //Показываем сообщение ошибки
                    Console.Out.WriteLine(msg, "Ошибка");
                }
                List<string> results1 = new List<string>(group.Items.Select(i => i.TimeStamp ?? "1"));
                System.Threading.Thread.Sleep(1000);
                try
                {
                    group.SyncRead();
                }
                catch (ServerException ex)
                {
                    string msg;
                    //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                    server.Server.GetErrorString(ex.HResult, 2, out msg);
                    //Показываем сообщение ошибки
                    Console.Out.WriteLine(msg, "Ошибка");
                }
                List<string> results2 = new List<string>(group.Items.Select(i => i.TimeStamp ?? "2"));
                Assert.AreEqual(results1.Count, results2.Count);
                Assert.AreNotEqual(results1[1], results2[1]);
                group.RemoveFromServer();
                try
                {
                    group.SyncRead();
                }
                catch (ServerException ex)
                {
                    errorCount++;
                    string msg;
                    //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                    server.Server.GetErrorString(ex.HResult, 2, out msg);
                    //Показываем сообщение ошибки
                    Console.Out.WriteLine(msg, "Ошибка");
                }
            }
            Assert.AreEqual(errorCount, opcBranches.Count());
            server.Disconnect();      
        }

        [TestMethod]
        public void TestOPCGroupASyncRead()
        {
            string progID = "Insat";
            int errorCount = 0;
            OPCServer server = OPCServer.FindServerByProgID(progID, "localhost");
            server.Connect();
            List<OPCItem> items = server.GetItems();
            var opcBranches = from item in items
                              where item.ItemType == OPCItemType.BRANCH
                              let leaves = from elem in items
                                           where elem.ItemType == OPCItemType.LEAF
                                           where elem.Parent != null && elem.Parent.ItemName == item.ItemName
                                           select elem
                              where leaves.Count() > 0
                              where !item.ItemName.Contains("МВ110")
                              select item;
            foreach (OPCItem branch in opcBranches)
            {
                string groupName = branch.ItemID;
                var opcItems = from item in items
                               where item.ItemType == OPCItemType.LEAF
                               where item.Parent != null
                               where item.Parent.ItemName == branch.ItemName
                               select item;
                List<OPCItem> list = new List<OPCItem>(opcItems);
                uint udateRate = 100;
                OPCGroup group = new OPCGroup(name: groupName, items: list)
                {
                    UpdateRate = udateRate
                };
                Assert.AreEqual(groupName, group.Name);
                Assert.AreSame(list, group.Items);
                Assert.AreEqual(udateRate, group.UpdateRate);
                group.RegInServer(server);
                group.AdviseOnServer(server, group.Items.ToArray()[1]);
                
                try
                {
                    group.ASyncRead();
                }
                catch (ServerException ex)
                {
                    string msg;
                    //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                    server.Server.GetErrorString(ex.HResult, 2, out msg);
                    //Показываем сообщение ошибки
                    Console.Out.WriteLine(msg, "Ошибка");
                }
                List<string> results1 = new List<string>(group.Items.Select(i => i.TimeStamp ?? "1"));
                System.Threading.Thread.Sleep(1000);
/*                try
                {
                    group.SyncRead();
                }
                catch (ServerException ex)
                {
                    string msg;
                    //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                    server.Server.GetErrorString(ex.HResult, 2, out msg);
                    //Показываем сообщение ошибки
                    Console.Out.WriteLine(msg, "Ошибка");
                }
*/                List<string> results2 = new List<string>(group.Items.Select(i => i.TimeStamp ?? "2"));
                Assert.AreEqual(results1.Count, results2.Count);
                Assert.AreNotEqual(results1[1], results2[1]);
                group.RemoveFromServer();
/*                try
                {
                    group.SyncRead();
                }
                catch (ServerException ex)
                {
                    errorCount++;
                    string msg;
                    //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT 
                    server.Server.GetErrorString(ex.HResult, 2, out msg);
                    //Показываем сообщение ошибки
                    Console.Out.WriteLine(msg, "Ошибка");
                }*/
            }
            //Assert.AreEqual(errorCount, opcBranches.Count());
            server.Disconnect();
        }
    }
}
