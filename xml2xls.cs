using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.IO;
using OfficeOpenXml;
using System.Globalization;


//Nuget EPPlus and system.commandline.dragonfruit

//two different types of input spreadsheets
//1. newer format - 

//2. older format - 
namespace IPxact2XLS
{
    ///<Summary>
    /// Register Access Type
    ///</Summary>
    public enum AccessType { RW, R }; //read-write, read-only

    ///<Summary>
    /// Field of Register
    ///</Summary>
    public class Field
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public int bitOffset { get; set; }
        public int resetValue { get; set; }      //in spririt format field does not have reset value; but can be calculated from register reset value; it exist in ipxact format
        public int bitWidth { get; set; }
        public AccessType Access { get; set; }
    }
    ///<Summary>
    /// Register
    ///</Summary>
    public class Register
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string addressOffset { get; set; }
        public int resetValue { get; set; }
        //public string resetMask { get; set; }
        public List<Field> Fields { get; set; }
        public Register()
        {
            Fields = new List<Field>();
            
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

        static int GetRegValueFromFields(Register r)
        {
            int rv = 0;
            foreach (var f in r.Fields)
            {
                rv += f.resetValue << f.bitOffset;
            }

            return rv;
        }

        //newer format
        static List<Register> ProcessXML_SNPS(XElement root)
        {
            List<Register> phyregs = new List<Register>();
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
                    var xr = new Register();
                    xr.Name = reg_name;
                    xr.Description = z.Elements().First(x => x.Name.LocalName == "description").Value;
                    xr.addressOffset = z.Elements().First(x => x.Name.LocalName == "addressOffset").Value;
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

                        Field ff = new Field
                        {
                            Name = n,
                            Description = d,
                            bitOffset = int.Parse(offset),
                            bitWidth = int.Parse(w)
                        };
                        ff.resetValue = GetFieldValue(xr.resetValue, ff.bitOffset, ff.bitWidth);

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

                        xr.Fields.Add(ff);
                    }
                    phyregs.Add(xr);
                }

            }
            return phyregs;
        }


        //older format
        static List<Register> ProcessXML_IPXact(XElement root)
        {
            List<Register> phyregs = new List<Register>();
            var mmaps = root.Elements().Last();
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
                    var xr = new Register();
                    xr.Name = reg_name;
                    xr.Description = z.Elements().First(x => x.Name.LocalName == "description").Value;
                    xr.addressOffset = z.Elements().First(x => x.Name.LocalName == "addressOffset").Value;
                    

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

                        Field ff = new Field
                        {
                            Name = n,
                            Description = d,
                            bitOffset = int.Parse(offset),
                            resetValue = resetvalue,        
                           
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

                        xr.Fields.Add(ff);
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
                List<Register> phyregs = new List<Register>();
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
                        worksheet.Cells[1, 2].Value = "Register Name";
                        worksheet.Cells[1, 3].Value = "Register Description";
                        worksheet.Cells[1, 4].Value = "Register Value";
                        worksheet.Cells[1, 5].Value = "Field Name";
                        worksheet.Cells[1, 6].Value = "Bit Start";
                        worksheet.Cells[1, 7].Value = "Bit Width";
                        worksheet.Cells[1, 8].Value = "Access";
                        worksheet.Cells[1, 9].Value = "Field Value";
                        worksheet.Cells[1, 10].Value = "Field Description";

                        i++;

                        foreach (var r in phyregs)
                        {
                            foreach (var f in r.Fields)
                            {
                                worksheet.InsertRow(i, 1);
                                worksheet.Cells[i, 1].Value = r.addressOffset.Replace("'h", "0x", StringComparison.OrdinalIgnoreCase); // "0x" + r.addressOffset.Substring(2);
                                worksheet.Cells[i, 2].Value = r.Name;
                                worksheet.Cells[i, 3].Value = r.Description;
                                worksheet.Cells[i, 4].Value = "0x" + r.resetValue.ToString("X");
                                //worksheet.Cells[i, 5].Value = r.resetMask;

                                //field
                                worksheet.Cells[i, 5].Value = f.Name;
                                worksheet.Cells[i, 6].Value = f.bitOffset;
                                worksheet.Cells[i, 7].Value = f.bitWidth;
                                worksheet.Cells[i, 8].Value = f.Access;
                                worksheet.Cells[i, 9].Value = "0x" + f.resetValue.ToString("X");
                                
                                worksheet.Cells[i, 10].Value = f.Description;
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
