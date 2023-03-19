using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics; // Для процессов
using System.Runtime.InteropServices; // пространство имен для маршала

namespace Netlist_to_Exel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WindowHeight = Console.LargestWindowHeight/2;
            Console.WindowWidth = Console.LargestWindowWidth/2;
            Console.BufferHeight = 2000;         

            Console.ForegroundColor = ConsoleColor.Green;
            Excel.Application myExcel;

            Excel._Worksheet mySheet;

            Pin onePin = new Pin();
            Cable oneCable = new Cable();
            myExcel = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            mySheet = (Excel._Worksheet)myExcel.ActiveSheet;
            if (mySheet != null)
                Console.WriteLine(mySheet.Name);
            else
                Console.WriteLine("Ups!!");
            
            string[] lines = System.IO.File.ReadAllLines("F:\\Ринат\\Коробки соединительные\\ТИС-БУ САРД\\Netlist_1.txt");
 
            oneCable.readConductors(mySheet);
            oneCable.readNetList("F:\\Ринат\\Коробки соединительные\\ТИС-БУ САРД\\Netlist_1.txt");
            oneCable.initialization();
            Console.ReadLine();
        }
        private static void SetNet(Excel._Worksheet s_Sheet, Pin s_Pin, string s_net)
        {
            Pin p1 = new Pin();
            Pin p2 = new Pin();
            Console.WriteLine("Ищем элемент {0,10}  |{1,2} ....", s_Pin.connector, s_Pin.pinNumber, s_net);
            for (int j = 2; j < 285; j++)
            {
                if (s_Sheet.Cells[j, 3].value == "")
                {
                    p1.connector = s_Sheet.Cells[j, 4].value;
                    p1.pinNumber = s_Sheet.Cells[j, 5].Value;
                    p2.connector = s_Sheet.Cells[j, 8].value;
                    p2.pinNumber = s_Sheet.Cells[j, 9].Value;

                    if (s_Pin == p1 && s_Pin != p2)
                    {
                        s_Sheet.Cells[j, 3].value = s_net;
                        Console.WriteLine("{0,10}  |{1,10} | присквоена цепь {2,5}", p1.connector, p1.pinNumber, s_net);
                        SetNet(s_Sheet, p2, s_net);
                    }
                    if (s_Pin == p2 && s_Pin != p1)
                    {
                        s_Sheet.Cells[j, 3].value = s_net;
                        Console.WriteLine("{0,10}  |{1,10} | присквоена цепь {2,5}", p2.connector, p2.pinNumber, s_net);
                        SetNet(s_Sheet, p1, s_net);
                    }
                }
            }
        }
    }

    public class Pin // контакт
    {
        public string connector;
        public string pinNumber;
        //public string net;
        public Pin()
        {
            connector = "";
            pinNumber= "";
        }
        public Pin(string connectorName, string pinNo)
        {
            connector = connectorName;
            pinNumber = pinNo;
            if (connectorName.Length > 2)
            {
                if (connectorName.Remove(2) == "XA")
                {                    
                    connector = "A1";
                    pinNumber = connectorName;
                }
                if (connectorName.Remove(2) == "XB")
                {
                    connector = "A2";
                    pinNumber = connectorName;
                }
            }
            if (connectorName == "S7")
            {
                if (pinNo.Remove(1) == "A" || pinNo.Remove(1) == "B")
                {
                    connector = "S7.1";
                }
                if (pinNo.Remove(1) == "C" || pinNo.Remove(1) == "D")
                {
                    connector = "S7.2";
                }
            }
            //net = "";
        }
        public bool Equals (Pin s_Pin)
        {
            return (s_Pin.connector == connector && s_Pin.pinNumber == pinNumber);
        }
        public void Connect(string connectorName, string pinNo)
        {
            connector = connectorName;
            pinNumber = pinNo;
        }
    }

    
    public class Conductor //жила
    {
        public Pin pIn;
        public Pin pOut;
        public Pin pInCsreen;
        public Pin pOutScreen;
        public string net;
        public string number;
        public Conductor()
        {
            pIn = new Pin();
            pOut = new Pin();
            pInCsreen = new Pin();
            pOutScreen = new Pin();
        }
        public Conductor(Pin pinIn, Pin pinOut, string net)
        {
            pIn = pinIn;
            pOut = pinOut;
        }
        public void WriteToConsole()
        {
            string[] splitStringIN = new string[1];
            string[] splitStringOUT =new string[1];
            string sConnIn, sConOut;
            sConnIn = pIn.connector;
            sConOut = pOut.connector;

            if (pIn.connector.Length > 10)
            {
                splitStringIN = pIn.connector.Split(' ');
                splitStringOUT = new string[splitStringIN.Length];
                sConnIn = splitStringIN[0]+ " " + splitStringIN[1];
            }
            if (pOut.connector.Length > 10)
            {
                splitStringOUT = pOut.connector.Split(' ');
                splitStringIN = new string[splitStringOUT.Length];
                sConOut = splitStringOUT[0] +" "+ splitStringOUT[1];
            }
            if (splitStringIN.Length > 1)
            {
                Console.WriteLine(new string('-',72));
            }
            Console.WriteLine("{0,5} |{1,10} |{2,12} |{3,10} |{4,12} |{5,10} |{6,8} |{7,8} |{8,8} |{9,8} |",
                    number, net,
                    sConnIn, pIn.pinNumber,
                    sConOut, pOut.pinNumber,
                    pInCsreen.connector, pInCsreen.pinNumber,
                    pOutScreen.connector, pOutScreen.pinNumber);

            if (splitStringIN.Length > 1)
            {
                for (int i = 3; i < splitStringIN.Length; i++)
                {
                    Console.WriteLine("{0,5} |{1,10} |{2,12} |{3,10} |{4,12} |{5,10} |{6,8} |{7,8} |{8,8} |{9,8} |",
                            "", "",
                            splitStringIN[i], "",
                            splitStringOUT[i], "",
                            "","","","");
                }     
                Console.WriteLine(new string('-', 72));
            }
        }
    }
    public class Wire // провод
    {
        bool screen;
        int position;
        int index;
        List<Conductor> conductorList;
        public Wire(Conductor s_Ware)
        {
            conductorList = new List<Conductor>();
            conductorList.Add(s_Ware);
        }
        public Wire(List<Conductor> s_condList, string s_ID)
        {
            conductorList = new List<Conductor>();
            conductorList = s_condList.FindAll(N => N.number == s_ID);
        }
        public void Add(Conductor s_Conductor)
        {
            conductorList.Add(s_Conductor);
        }
        public void Add(List<Conductor> s_condList, string s_ID)
        {
            conductorList = s_condList.FindAll(N => N.number == s_ID);
        }
        public void WriteAtConsole()
        {
            if (conductorList != null)
            {                
                foreach (Conductor C in conductorList)
                {
                    C.WriteToConsole();
                }
                Console.WriteLine(new string('-', 72));
            }
        }
        public bool Equals(Wire s_Wire)
        {
            return (
                s_Wire.conductorList.Equals(conductorList) &&
                s_Wire.index ==  index &&
                s_Wire.position == position &&
                s_Wire.screen == screen
                );
        }
    }
    public class GroupWires // группа проводов с объедененным экраном
    {
        List<Wire> cableList;
        Pin pinScreenIN;
        Pin pinScreenOUT;
    }
    
    public class Net // электрическая цепь 
    {
        public string name;
        public List<Pin> pinList;
        public Net(string s_name)
        {
            name = s_name;
            pinList = new List<Pin>();
        }
    }
    public class Cable // класс 
    {
        List<GroupWires> GroupWiresList; // группы проводов
        List<Wire> wiresList; // провода
        List<Conductor> conductorList; // проводники
        List<Net> netList; //
        List<string> ConnectorList;
        List<Conductor> screenConductors;
        public void initialization()
        {

            SetNets();
            GenerateWaries();
            this.ScrensToConductors();
            writeConductorsToConsole();             
            Console.WriteLine("END");
        }
        private void AddConnector(string s_Connector)
        {
            if (ConnectorList != null)
            {
                if (ConnectorList.Contains(s_Connector) == false)
                    ConnectorList.Add(s_Connector);
            }
        }
        private List<Pin> SplitStringToPins (string s_String)
        {
            List<Pin> myPins = new List<Pin>();
            bool bConn = true;
            string[] myStrings = s_String.Split(' ');
            List<string> localConnectors = new List<string>();
            List<string> localNO = new List<string>();            
            for (int i = myStrings.Length-1; i>=0; i--) 
            {
                string oneStr = myStrings[i];
                // если есть пробелы удаляем их------------------------------------------------
                while (oneStr.Contains(' '))
                {
                    int z = oneStr.IndexOf(' ');
                    oneStr.Remove(i, 1);                    
                }
                //------------------------------------------------------------------------------

                if (oneStr[oneStr.Length - 1] == ',') // если есть запятая
                {
                    oneStr = oneStr.Remove(myStrings[i].Length - 1); // удалем запятую
                }
                // строка точно без запятой
                if (bConn == true)
                {
                    if (ConnectorList.Contains(oneStr)) // если есть совпадение
                        localConnectors.Add(oneStr); // добавляем коннектор
                    else // нет совпадения
                    {
                        bConn = false;
                        continue; // завершаем
                    }
                }
                else // false
                {
                    localNO.Add(oneStr); // добавляем номер
                    if (i > 0)
                    {
                        string nextString = myStrings[i - 1];
                        if (nextString[nextString.Length - 1] != ',') // строка без запятой
                        {
                            bConn = true;
                            //добавляем накопленное коллекцию
                            foreach (string Con in localConnectors)
                            {
                                foreach (string No in localNO)
                                {
                                    myPins.Add(new Pin(Con, No));
                                }
                            }
                            localConnectors.Clear();
                            localNO.Clear();
                        }
                    }
                }                
            }
            return myPins;
        }

        public void readNetList(string fileRoot) //"F:\\Ринат\\Коробки соединительные\\ТИС-БУ САРД\\Netlist_1.txt"
        {
            string[] lines = System.IO.File.ReadAllLines(fileRoot);
            string net = "";
            string element = "";
            string contact = "";
            Net oneNet;
            Pin onePin;
            netList = new List<Net>();
            oneNet = new Net("");
            for (int i = 0; i < lines.Count(); i++)
            {
                if (lines[i].Contains("(net") == true)
                {
                    if (oneNet.name != "")
                        netList.Add(oneNet);
                    net = lines[i].Split('"')[1];
                    oneNet = new Net(net);
                }                                
                if (lines[i].Contains("(node") == true)
                {
                    element = lines[i].Split('"')[1];
                    contact = lines[i].Split('"')[3];
                    onePin = new Pin(element, contact);
                    oneNet.pinList.Add(onePin);
                }
            }
        }
        public void writeNetList ()
        {
            if (netList != null)
            {
                foreach (Net s_Net in netList)
                { 
                   Console.WriteLine(s_Net.name);
                    foreach (Pin p in s_Net.pinList)
                    {
                        Console.WriteLine("{0,10}  |{1,10} ",p.connector, p.pinNumber);
                    }
                
                }
                
            }
        }
        public void readConductors(Excel._Worksheet s_Sheet)
        {
            Conductor one_Conductor;
            conductorList = new List<Conductor>();
            ConnectorList = new List<string>();
            screenConductors = new List<Conductor>();
            Pin p1, p2;
            for (int i = 2; i < 285; i++)
            {
                one_Conductor = new Conductor();
                one_Conductor.number = (string)s_Sheet.Cells[i, 2].value.ToString();
                p1 = new Pin((string)s_Sheet.Cells[i, 4].value, (string)s_Sheet.Cells[i, 5].value);
                p2 = new Pin((string)s_Sheet.Cells[i, 8].value, (string)s_Sheet.Cells[i, 9].value);
                one_Conductor.pIn = p1;
                one_Conductor.pOut = p2;
                AddConnector(p1.connector);
                AddConnector(p2.connector);
                if (one_Conductor.pIn.connector.Length > 10 || one_Conductor.pOut.connector.Length > 10)
                {
                    screenConductors.Add(one_Conductor);
                }
                else
                conductorList.Add(one_Conductor);                
            }

        }
        public void writeConductorsToConsole()
        {
            if (conductorList != null)
            {
                foreach (Conductor cond in conductorList)
                {
                    cond.WriteToConsole();                        
                }
            }
        }
        private void SetNet(Pin s_Pin, string s_net)
        {
            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.Red;
            Pin p1, p2;
            //Console.WriteLine("Ищем элемент  {0,7}  |{1,4} | net {2,7}|", s_Pin.connector, s_Pin.pinNumber, s_net);
            Console.ForegroundColor = ConsoleColor.Green;
            foreach (Conductor cond in conductorList) 
            {                
                if (cond.net == "" || cond.net == null)
                {
                    p1 = (Pin)cond.pIn;
                    p2 = (Pin)cond.pOut;
                    if (s_Pin.Equals(p1))
                    {
                        cond.net = s_net;                    
                        Console.WriteLine("{0,6}  |{1,6} | присквоена цепь {2,5}", cond.pIn.connector, cond.pIn.pinNumber, s_net);
                        this.SetNet(cond.pOut, s_net);
                    }
                    if (s_Pin.Equals(p2))
                    {
                        cond.net = s_net;
                        Console.WriteLine("{0,6}  |{1,6} | присквоена цепь {2,5}", cond.pOut.connector, cond.pOut.pinNumber, s_net);
                        this.SetNet(cond.pIn, s_net);
                    }
                }
            }
        }
        public void SetNets()
        {
            if (conductorList != null && netList != null)
            {
                foreach (Net nT in netList)
                {
                    foreach (Pin p in nT.pinList)
                    {
                        SetNet(p, nT.name);
                    }
                }
            }
        }
        public void GenerateWaries() // создать провода
        {
            if (conductorList != null)
            {
                List<string> Numbers = new List<string>();
                wiresList = new List<Wire>();
                foreach (Conductor C1 in conductorList)
                {
                    if (Numbers.Any(N => N == C1.number) == false)
                        Numbers.Add(C1.number);
                }
                foreach (string n in Numbers)
                {
                    Wire oneWire = new Wire(conductorList, n);
                    wiresList.Add(oneWire); 
                }
            }
        }
        public void WriteWires()
        {
            if (wiresList != null)
            {
                foreach (Wire w in wiresList)
                {
                    w.WriteAtConsole();
                }
            }
        }
        public void WriteConductorsWithoutNet()
        {
 
        }
        
        public void SetNetsToExcel()
        {
 
        }
        public void ScrensToConductors()
        {
            foreach(Conductor gScreenConn in screenConductors) // для каждого вывода экрана
            {
                if (gScreenConn.pIn.connector.Length> 10)
                {
                    List<Pin> screenPins = SplitStringToPins(gScreenConn.pIn.connector);
                    foreach (Pin pn in screenPins)
                    {
                        foreach (Conductor C in conductorList)
                        {
                            if(C.pIn.Equals(pn))
                                C.pInCsreen = gScreenConn.pOut;
                        }
                    }
                }
                if (gScreenConn.pOut.connector.Length > 10)
                {
                    List<Pin> screenPins = SplitStringToPins(gScreenConn.pOut.connector);
                    foreach (Pin pn in screenPins)
                    {
                        foreach (Conductor C in conductorList)
                        {
                            if (C.pOut.Equals(pn))
                                C.pOutScreen = gScreenConn.pIn;
                        }
                    }
                }
                
            }
            
        }
        public void GenerateGroupsWires()
        {
            
            if (ConnectorList != null)
            {
                foreach (Conductor gCond in conductorList)
                {
                    if (gCond.pIn.connector.Length > 10)
                    {
                        List<Pin> onePin = SplitStringToPins(gCond.pIn.connector);

                    }
                }
            }
        }
    
    }
}

