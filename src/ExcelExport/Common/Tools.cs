using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using ExcelExport.EF;
using ICSharpCode.SharpZipLib.Zip;
using  NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelExport.Common
{
    public static class Tools
    {
        public static MemoryStream ListToExcel<T> (IEnumerable<T> list, int rowsCount, int sheetCount = 10) where T: ViewModel
        {
            //check:sheetCount>0
            if (sheetCount <= 0)
            {
                return null;
            }

            var excelMs= new MemoryStream();
            HSSFWorkbook book =new HSSFWorkbook();

            var properties = typeof(T).GetProperties().ToList();
           var pageCount=list.Count()%rowsCount==0? list.Count()/rowsCount : list.Count()/rowsCount +1;

            for (int i = 1; i <= sheetCount; i++)
            {
                if (pageCount < i)
                {
                    continue;
                }

                var sheetDataList = list.Skip(rowsCount * (i - 1)).Take(rowsCount).ToArray();

                var sheet = book.CreateSheet($"sheet{i}");
                var row = sheet.CreateRow(0);

                //设置表头
                for (int j = 0; j < properties.Count; j++)
                {
                    var p = properties[j];
                    var displayName = p.GetCustomAttributes<DisplayNameAttribute>();
                    row.CreateCell(j).SetCellValue(displayName.FirstOrDefault().DisplayName);
                }

                //设置数据
                for (int k = 0; k < sheetDataList.Length; k++)
                {
                    var rowTemp = sheet.CreateRow(k + 1);
                    for (int a = 0; a < properties.Count; a++)
                    {
                        var p = properties[a];
                        rowTemp.CreateCell(a).SetCellValue(p.GetValue(sheetDataList[k]).ToString());
                    }
                }
                
            }
            book.Write(excelMs);
            excelMs.Seek(0, SeekOrigin.Begin);

            return excelMs;
        }

        public static void Export<T>(this IQueryable<T> iQueryable,  string fileName,int rowsCount,  int sheetCount=10,Func<IEnumerable<T>,IEnumerable<T>> sortFunc=null) where T :ViewModel
        {
            var fullName = fileName;
            var pageSize = rowsCount;           
            var fileCount = iQueryable.Count() % (pageSize*sheetCount) == 0
                ? (iQueryable.Count() / (pageSize * sheetCount))
                : (iQueryable.Count() / (pageSize * sheetCount) + 1);

            System.IO.MemoryStream ms = new System.IO.MemoryStream();

            var resp = HttpContext.Current.Response;
            resp.ContentEncoding = System.Text.Encoding.GetEncoding("GB2312");
            if (fileCount <= 1)
            {
                fullName += ".xls";
                resp.ContentType = "application/vnd.ms-excel";
            }
            else
            {
                fullName += ".zip";
                resp.ContentType = "application/zip";
            }

            resp.AppendHeader("Content-Disposition", "attachment;filename=" + fullName);

            var list = sortFunc==null?iQueryable:sortFunc(iQueryable);
            //var p = typeof(T).GetProperty(orderByName);
            using (ZipFile zip = ZipFile.Create(ms))
            {
                zip.BeginUpdate();
                for (int i = 0; i < fileCount; i++)
                {
                    
                    var excelMs = ListToExcel<T>(list.Skip(i * pageSize* sheetCount).Take(pageSize* sheetCount).ToList(), pageSize, sheetCount);
                    if (fileCount <= 1)
                    {
                        resp.BinaryWrite(excelMs.GetBuffer());
                        resp.End();
                        return;
                    }

                    zip.Add(new StreamDataSource(excelMs), $"{fileName}_{i + 1}.xls");
                }
                zip.CommitUpdate();
                ms.Seek(0,SeekOrigin.Begin);
            }
            resp.BinaryWrite(ms.GetBuffer());
            resp.End();
        }
    }


    public class StreamDataSource : IStaticDataSource
    {
        public byte[] bytes { get; set; }
        public StreamDataSource(MemoryStream ms)
        {
            bytes = ms.GetBuffer();
        }

        public Stream GetSource()
        {
            Stream s = new MemoryStream(bytes);
            return s;
        }
    }
}