using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Reflection;
using System.IO;
using OfficeOpenXml;


// Need EPPLus For Excel handling

namespace XML2XLS
{
    public enum accessType {RW, RO}; //read-write, read-only
    public class field
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public int bitOffset { get; set; }
        public string resetValue { get; set; }  //reset values which has value and mask properties in xml; mask=0 for reserved;
        public int bitWidth { get; set; }
        public accessType Access { get; set; }
    }

    public class register
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string addressOffset { get; set; }
        public List<field> Fields { get; set; }
        public register()
        {
            Fields = new List<field>();
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            List<register> regs = new List<register>();
            try
            {
                string fn = @"ipxact\register.xml";	//some path to file
                
                XDocument x = XDocument.Load(fn);
                var root = x.Root;
                var regs = root.Elements().ToList();  //all the registers

                foreach (var z in regs)
                {
                    var addr = z.Elements().First(x => x.Name.LocalName == "addressOffset").Value;
                    //remove the if to add all registers including ROM/RAM
                    char msb = addr.ToCharArray()[2];
                    
                    if (addr.Length == 6 && msb >= 'a')      //skip certain address range
                    {
                        //skip
                    }
                    else 
                    {
                        var xr = new register();
                        xr.Name = z.Elements().First(x => x.Name.LocalName == "name").Value;
                        xr.Description = z.Elements().First(x => x.Name.LocalName == "description").Value;
                        xr.addressOffset = addr; // z.Elements().First(x => x.Name.LocalName == "addressOffset").Value;

                        var fields = z.Elements().Where(x => x.Name.LocalName == "field");

                        foreach (var f in fields)
                        {
                            var n = f.Elements().First(x => x.Name.LocalName == "name").Value;
                            //description may not exist for some fields; must use FirstOrDefault()
                            string d = f.Elements().FirstOrDefault(x => x.Name.LocalName == "description")?.Value ?? string.Empty;

                            var offset = f.Elements().First(x => x.Name.LocalName == "bitOffset").Value;
                            var resets = f.Elements().First(x => x.Name.LocalName == "resets");
                            var resetvalue = resets.Descendants().First(x => x.Name.LocalName == "value").Value;
                            var w = f.Elements().First(x => x.Name.LocalName == "bitWidth").Value;
                            var acc = f.Elements().First(x => x.Name.LocalName == "access").Value;

                            field ff = new field
                            {
                                Name = n,
                                Description = d,
                                bitOffset = int.Parse(offset),
                                resetValue = resetvalue,
                                bitWidth = int.Parse(w)
                            };
                            if (acc == "read-only")
                            {
                                ff.Access = accessType.RO;
                            }
                            else if (acc == "read-write")
                            {
                                ff.Access = accessType.RW;
                            }
                            else
                            {
                                throw new Exception("unexpected access type for field " + n);
                            }

                            xr.Fields.Add(ff);
                        }
                        regs.Add(xr);
                    }
                    
                }
                Console.WriteLine(regs.Count());
                

                //save to spreadsheet
                string fn2 = @"registers.xlsx";
                var xlsfile = new FileInfo(fn2);
                using (ExcelPackage p = new ExcelPackage(xlsfile))
                {
                    var wk = p.Workbook.Worksheets.SingleOrDefault(x => x.Name == "xlsname");
                    if (wk != null) { p.Workbook.Worksheets.Delete(wk); }
                    ExcelWorksheet worksheet = p.Workbook.Worksheets.Add("xlsname");
                    
                    int i = 1; //row number
                    worksheet.InsertRow(1, 1);
                    
                    worksheet.Cells[1, 1].Value = "Offset";
                    worksheet.Cells[1, 2].Value = "Register Name";
                    worksheet.Cells[1, 3].Value = "Register Description";
                    worksheet.Cells[1, 4].Value = "Bit Start";
                    worksheet.Cells[1, 5].Value = "Bit Width";
                    worksheet.Cells[1, 6].Value = "Access";
                    worksheet.Cells[1, 7].Value = "Reset Value";
                    worksheet.Cells[1, 8].Value = "Field Name";
                    worksheet.Cells[1, 9].Value = "Field Description";

                    i++;

                    foreach (var r in regs)
                    {
                        foreach (var f in r.Fields)
                        {
                            worksheet.InsertRow(i, 1);
                            worksheet.Cells[i, 1].Value = r.addressOffset.Replace("'h", "0x", StringComparison.OrdinalIgnoreCase); // "0x" + r.addressOffset.Substring(2);
                            worksheet.Cells[i, 2].Value = r.Name;
                            worksheet.Cells[i, 3].Value = r.Description;

                            //field

                            worksheet.Cells[i, 4].Value = f.bitOffset;
                            worksheet.Cells[i, 5].Value = f.bitWidth;
                            worksheet.Cells[i, 6].Value = f.Access;
                            worksheet.Cells[i, 7].Value = f.resetValue.Replace("'h", "0x", StringComparison.OrdinalIgnoreCase);
                            worksheet.Cells[i, 8].Value = f.Name;
                            worksheet.Cells[i, 9].Value = f.Description;
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
