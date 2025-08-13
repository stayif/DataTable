namespace PerillaTable
{
    static class Program
    {
        static void Main(string[] args)
        {
            DateTime time = DateTime.Now.ToUniversalTime();
            ExcelTool excelTool = new ExcelTool();
            FileTool fileTool = new FileTool();
            fileTool.GetAllFiles("bin/Debug/net6.0/Excel");

            //Config.I.exportType = Config.ExportType.Bytes;
            for (int i = 0; i < fileTool.fileList.Count; i++)
            {
                string item = fileTool.fileList[i];

                if (item.Contains("Enum.xlsx") || item.Contains("~$"))
                    continue;

                if (item.Contains(".xls"))
                {
                    try
                    {
                        excelTool.CreateDataTable(item);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex);
                        Console.ReadKey();
                    }
                }
            }

            Console.WriteLine("转换完成！耗时: " + (DateTime.Now.ToUniversalTime() - time).TotalSeconds + "秒");
            Console.WriteLine("按任意键退出...");
            Console.ReadKey();
        }
    }



}
