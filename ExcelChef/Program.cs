using ExcelChef.Parsers;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace ExcelChef
{
    public class Program
    {
        private readonly IParser _parser = new JsonParser();
        
        public Stream Input { private get; set; }

        public Stream Output { private get; set; }

        public Stream Template { private get; set; }

        public void Run()
        {
            IWorkbook template = WorkbookFactory.Create(Template);
            template.MissingCellPolicy = MissingCellPolicy.CREATE_NULL_AS_BLANK;
            IEnumerable<IInstruction> instructions = _parser.Parse(new StreamReader(Input));
            foreach (IInstruction instruction in instructions)
            {
                instruction.Execute(template);
            }
            template.Write(Output);
        }

        public static int Main(string[] args)
        {
            try
            {
                if (args.Length < 2) throw new Exception("No template file and output file specified");
                new Program
                {
                    Input = Console.OpenStandardInput(),
                    Output = new FileStream(args[1], FileMode.Create, FileAccess.ReadWrite),
                    Template = new FileStream(args[0], FileMode.Open, FileAccess.Read),
                }.Run();
                /* PROFILER
                new Program
                {
                    Input = new FileStream("instructions.json", FileMode.Open, FileAccess.Read),
                    Output = new FileStream("out.xlsx", FileMode.Create, FileAccess.ReadWrite),
                    Template = new FileStream("in.xlsx", FileMode.Open, FileAccess.Read),
                }.Run();
                */
                return 0;
            }
            catch (Exception exception) when (!Debugger.IsAttached)
            {
                Console.Error.WriteLine(exception);
                return 1;
            }
        }
    }
}
