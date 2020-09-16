using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.IO;
using OfficeOpenXml;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

//Nuget EPPlus and system.commandline.dragonfruit

//two different types of input spreadsheets - see details below
//1. newer format - 
//2. older format - 
namespace IPxact2XLS
{
    ///<Summary>
    /// register Access Type
    ///</Summary>
    public enum AccessType { RW, R }; //read-write, read-only

    ///<Summary>
    /// Field of register
    ///</Summary>
    public class field
    {
        public string name { get; set; }
        [JsonIgnore]
        public string description { get; set; }
        public int bitOffset { get; set; }
        public int value { get; set; }      //in spririt format field does not have reset value; but can be calculated from register reset value; it exist in ipxact format
        public int bitWidth { get; set; }
        [JsonIgnore]
        public AccessType Access { get; set; }
    }
    ///<Summary>
    /// register
    ///</Summary>
    public class register
    {
        public string name { get; set; }
        [JsonIgnore]
        public string description { get; set; }
        public int Offset { get; set; }
        [JsonIgnore]
        public int resetValue { get; set; }
        //public string resetMask { get; set; }
        public List<field> fields { get; set; }
        public register()
        {
            fields = new List<field>();
            
        }
    }
    class XMLReg2XLS
    {
        static int IntParse(string inputs)
        {
            //the input string from xls may contain non-ascii char (like some html char); 
            //it may look like "0546h", but it won't detect h is at the end; it causes failure; need to strip them
            //this works for detection 
            /*bool HasNonASCIIChars(string str)
            {
                return (System.Text.Encoding.UTF8.GetByteCount(str) != str.Length);
            }
            */
            var cleaned_s = System.Text.RegularExpressions.Regex.Replace(inputs, @"[^\u0000-\u007F]+", string.Empty);
            var s = cleaned_s.Trim().ToLower();

            int v = 0;
            string ss; //substring;
            try
            {

                if (s.Contains('h'))
                {
                    if (s.EndsWith('h'))
                    {
                        CultureInfo provider = CultureInfo.InvariantCulture;
                        //v = Convert.ToInt32(s.Substring(0, s.Length - 1), 16);
                        ss = s.Substring(0, s.Length - 1);
                        if (!int.TryParse(ss, NumberStyles.HexNumber, provider, out v))
                        {
                            // error detected!!! need to handle some how
                            Debug.WriteLine(ss);
                        }

                        //v = int.Parse(hs, System.Globalization.NumberStyles.HexNumber);
                    }
                    else //assume format is 'h00 
                    {
                        int p = s.IndexOf('h');
                        //v = Convert.ToInt32(s.Substring(p+1), 16);
                        v = int.Parse(s.Substring(p + 1), System.Globalization.NumberStyles.HexNumber);
                    }
                }
                else if (s.Contains("0x"))
                {
                    int p = s.IndexOf('x');
                    //v = Convert.ToInt32(s.Substring(p+1), 16);
                    v = int.Parse(s.Substring(p + 1), System.Globalization.NumberStyles.HexNumber);
                    //v = Convert.ToInt32(s, 16);
                }
                //this has to come after hex parsing; because 'b' is valid hex string can coexist with h, 0x
                else if (s.Contains('b'))
                {
                    if (s.EndsWith('b'))
                    {
                        v = Convert.ToInt32(s.Substring(0, s.Length - 1), 2);
                    }
                    else //assume format is 2'b00 
                    {
                        int p = s.IndexOf('b');

                        v = Convert.ToInt32(s.Substring(p + 1), 2);
                    }
                }
                else
                {
                    v = int.Parse(s);
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine("Integer Parsing error - " + s);
                throw ex;
            }
            return v;
        }

        static int GetFieldValue(int regValue, int start, int width)    //bitstart, fieldWidth
        {
            int fv = (regValue >> start) & ((1<<width) - 1);
            return fv;
        }

        static int GetRegValueFromFields(register r)
        {
            int rv = 0;
            foreach (var f in r.fields)
            {
                rv += f.value << f.bitOffset;
            }

            return rv;
        }

        //newer format - register have values; fields do not; fields get value from register
        static List<register> ProcessXML_SNPS(XElement root)
        {
            List<register> phyregs = new List<register>();
            //foreach(var x in root.Elements())
            //{
            //    Console.WriteLine(x.Name.LocalName);
            //}

            var ver = root.Elements().Where(x => x.Name.LocalName.Contains("version")).First().Value;
            Console.WriteLine("version " + ver);

            var mmaps = root.Elements().Where(x => x.Name.LocalName.Contains("memoryMaps")).First();
            var mmap = mmaps.Elements().First();
            var addrBlock = mmap.Elements().First(x => x.Name.LocalName.Contains("addressBlock"));

            //Console.WriteLine(addrBlock.Elements().Count());
            var regs = addrBlock.Elements().Where(x => string.Compare(x.Name.LocalName, "register", StringComparison.OrdinalIgnoreCase) ==0 );  //all the registers
            Console.WriteLine(regs.Count());
            foreach (var z in regs)
            {
                //skip ROM/RAM register - name starts with RAWMEM_DIG_ROM or RAWMEM_DIG_RAM
                var reg_name = z.Elements().First().Value;
                //Debug.WriteLine(reg_name);

                //var addr = z.Elements().First(x => x.Name.LocalName == "addressOffset").Value;
                //char msb = addr.ToCharArray()[2];

                if (!reg_name.StartsWith("RAWMEM_DIG_R"))
                {
                    var xr = new register();
                    xr.name = reg_name;
                    xr.description = z.Elements().First(x => x.Name.LocalName == "description").Value;
                    xr.Offset = IntParse(z.Elements().First(x => x.Name.LocalName == "addressOffset").Value);
                    xr.resetValue = IntParse(z.Elements().First(x => x.Name.LocalName == "reset").Elements().First().Value);
                    //xr.resetMask = z.Elements().First(x => x.Name.LocalName == "reset").Elements().Last().Value;

                    var fields = z.Elements().Where(x => x.Name.LocalName == "field");
                    foreach (var f in fields)
                    {
                        var n = f.Elements().First(x => x.Name.LocalName == "name").Value;
                        //description may not exist for some fields; must use FirstOrDefault()
                        string d = f.Elements().FirstOrDefault(x => x.Name.LocalName == "description")?.Value ?? string.Empty;

                        var offset = f.Elements().First(x => x.Name.LocalName == "bitOffset").Value;
                        //var resets = f.Elements().First(x => x.Name.LocalName == "resets");
                        //var resetvalue = resets.Descendants().First(x => x.Name.LocalName == "value").Value;
                        var w = f.Elements().First(x => x.Name.LocalName == "bitWidth").Value;
                        var acc = f.Elements().First(x => x.Name.LocalName == "access").Value;

                        field ff = new field
                        {
                            name = n,
                            description = d,
                            bitOffset = int.Parse(offset),
                            bitWidth = int.Parse(w)
                        };
                        ff.value = GetFieldValue(xr.resetValue, ff.bitOffset, ff.bitWidth);

                        if (acc == "read-only")
                        {
                            ff.Access = AccessType.R;
                        }
                        else if (acc == "read-write")
                        {
                            ff.Access = AccessType.RW;
                        }
                        else
                        {
                            throw new Exception("unexpected access type for field " + n);
                        }

                        xr.fields.Add(ff);
                    }
                    phyregs.Add(xr);
                }

            }
            return phyregs;
        }


        //older format - register have no values; fields have; registers get value from fields
        static List<register> ProcessXML_IPXact(XElement root)
        {
            List<register> phyregs = new List<register>();
            //var mmaps = root.Elements().Last();
            var mmaps = root.Elements().Where(x => x.Name.LocalName.Contains("memoryMaps")).First();
            var mmap = mmaps.Elements().First();
            var addrBlock = mmap.Elements().First(x => x.Name.LocalName.Contains("addressBlock"));

            //Console.WriteLine(addrBlock.Elements().Count());
            var regs = addrBlock.Elements().Where(x => x.Name.LocalName == "register");  //all the registers
            Console.WriteLine(regs.Count());
            foreach (var z in regs)
            {
                //skip ROM/RAM register - name starts with RAWMEM_DIG_ROM or RAWMEM_DIG_RAM
                var reg_name = z.Elements().First().Value;
                //Debug.WriteLine(reg_name);

                //var addr = z.Elements().First(x => x.Name.LocalName == "addressOffset").Value;
                //char msb = addr.ToCharArray()[2];

                if (!reg_name.StartsWith("RAWMEM_DIG_R"))
                {
                    var xr = new register();
                    xr.name = reg_name;
                    xr.description = z.Elements().First(x => x.Name.LocalName == "description").Value;
                    xr.Offset = IntParse(z.Elements().First(x => x.Name.LocalName == "addressOffset").Value);
                    

                    var fields = z.Elements().Where(x => x.Name.LocalName == "field");
                    foreach (var f in fields)
                    {
                        var n = f.Elements().First(x => x.Name.LocalName == "name").Value;
                        //description may not exist for some fields; must use FirstOrDefault()
                        string d = f.Elements().FirstOrDefault(x => x.Name.LocalName == "description")?.Value ?? string.Empty;

                        var offset = f.Elements().First(x => x.Name.LocalName == "bitOffset").Value;
                        var resets = f.Elements().First(x => x.Name.LocalName == "resets");
                        var resetvalue = IntParse(resets.Descendants().First(x => x.Name.LocalName == "value").Value);
                        //var resetmask = resets.Descendants().First(x => x.Name.LocalName == "mask").Value;
                        var w = f.Elements().First(x => x.Name.LocalName == "bitWidth").Value;
                        var acc = f.Elements().First(x => x.Name.LocalName == "access").Value;

                        field ff = new field
                        {
                            name = n,
                            description = d,
                            bitOffset = int.Parse(offset),
                            value = resetvalue,        
                           
                            bitWidth = int.Parse(w)
                        };
                        if (acc == "read-only")
                        {
                            ff.Access = AccessType.R;
                        }
                        else if (acc == "read-write")
                        {
                            ff.Access = AccessType.RW;
                        }
                        else
                        {
                            throw new Exception("unexpected access type for field " + n);
                        }

                        xr.fields.Add(ff);
                    }

                    //no "register" reset value in older format; only in fields; need to calculate

                    xr.resetValue = GetRegValueFromFields(xr);
                    phyregs.Add(xr);
                }

            }
            return phyregs;
        }
        static void Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.WriteLine("Usage: Program.exe [xml file path]");
                Environment.Exit(0);
            }

            string fn = args[0];
            string phyname = "";

            //test older format
            //fn = "dwc_usb31sspphy_phy_x2_ns.xml";
            if (!File.Exists(fn))
            {
                Console.WriteLine("Input file " + fn + " Not Found!");
            }
            else
            {
                List<register> phyregs = new List<register>();
                try
                {
                    XElement root;

                    if (true)  //load xml file directly
                    {
                        XDocument x = XDocument.Load(fn);
                        root = x.Root;
                    }
                    //XDocument.Load(fn) can run into problem if file format is not expected..
                    //alternative is to read in file, pre-process, then Parse
                    else
                    {
                        //File.rea
                        //string[] lines = File.ReadAllLines(fn);
                        

                        //root = x.Root;
                    }

                    phyname = root.Elements().First(x => x.Name.LocalName == "name").Value;

                    if (root.Name.NamespaceName.Contains("SPIRIT", StringComparison.CurrentCultureIgnoreCase))  //newer format
                    {
                        phyregs = ProcessXML_SNPS(root);
                    }
                    else if (root.Name.NamespaceName.Contains("Ipxact", StringComparison.CurrentCultureIgnoreCase)) //older format
                    {
                        phyregs = ProcessXML_IPXact(root);
                    }
                    else
                    {
                        Console.WriteLine("NameSpace not supported: " + root.Name.NamespaceName);
                        Environment.Exit(0);
                    }


                    Console.WriteLine(phyregs.Count());
                    //save to JSON
                    string jsonfile = Path.GetFileNameWithoutExtension(fn) +  ".json";

                    var options = new JsonSerializerOptions
                    {
                        WriteIndented = true
                    };
                    var jsonString = JsonSerializer.Serialize(phyregs, options);
                    File.WriteAllText(jsonfile, jsonString);


                    //save to spreadsheet
                    string fn2 = @"registers.xlsx";
                    var xlsfile = new FileInfo(fn2);
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (ExcelPackage p = new ExcelPackage(xlsfile))
                    {
                        ExcelWorksheet worksheet;
                        worksheet = p.Workbook.Worksheets.Where(x => x.Name == phyname).FirstOrDefault();
                        if (worksheet != null)
                        {
                            worksheet.Cells.Clear();     //if can't delete the sheet
                            // p.Workbook.Worksheets.Delete( wk );   //bug; can't delete
                        }
                        else
                        {
                            worksheet = p.Workbook.Worksheets.Add(phyname);
                        }
                        int i = 1; //row number
                        worksheet.InsertRow(1, 1);

                        worksheet.Cells[1, 1].Value = "Offset";
                        worksheet.Cells[1, 2].Value = "register Name";
                        worksheet.Cells[1, 3].Value = "register Description";
                        worksheet.Cells[1, 4].Value = "register Value";
                        worksheet.Cells[1, 5].Value = "Field Name";
                        worksheet.Cells[1, 6].Value = "Bit Start";
                        worksheet.Cells[1, 7].Value = "Bit Width";
                        worksheet.Cells[1, 8].Value = "Access";
                        worksheet.Cells[1, 9].Value = "Field Value";
                        worksheet.Cells[1, 10].Value = "Field Description";

                        i++;

                        foreach (var r in phyregs)
                        {
                            foreach (var f in r.fields)
                            {
                                worksheet.InsertRow(i, 1);
                                worksheet.Cells[i, 1].Value =  "0x" + r.Offset.ToString("X");
                                worksheet.Cells[i, 2].Value = r.name;
                                worksheet.Cells[i, 3].Value = r.description;
                                worksheet.Cells[i, 4].Value = "0x" + r.resetValue.ToString("X");
                                //worksheet.Cells[i, 5].Value = r.resetMask;

                                //field
                                worksheet.Cells[i, 5].Value = f.name;
                                worksheet.Cells[i, 6].Value = f.bitOffset;
                                worksheet.Cells[i, 7].Value = f.bitWidth;
                                worksheet.Cells[i, 8].Value = f.Access;
                                worksheet.Cells[i, 9].Value = "0x" + f.value.ToString("X");
                                
                                worksheet.Cells[i, 10].Value = f.description;
                                //
                                i++;
                            }
                        }
                        p.Save();
                    }

                    Console.WriteLine("It is finished !!!");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }


       
    }
}
