using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using ICSharpCode.SharpZipLib.Zip;
using System.Configuration;

namespace WebApplication1
{
    public partial class FileWebForm : System.Web.UI.Page
    {
        //PDF存放地址
        private string pDFDocumentPath = string.Empty;
        /// <summary>
        /// PDF存放地址
        /// </summary>
        public string PDFDocumentPath
        {
            get
            {
                if (pDFDocumentPath == string.Empty)
                {
                    pDFDocumentPath = AppDomain.CurrentDomain.BaseDirectory + ConfigurationManager.AppSettings["PDFDocumentPath"].ToString();
                    //判断文件夹是否存在，不存在则创建
                    if (!System.IO.Directory.Exists(pDFDocumentPath))
                    {
                        System.IO.Directory.CreateDirectory(pDFDocumentPath);
                    }
                }
                return pDFDocumentPath + "\\";
            }
            set
            {
                pDFDocumentPath = value;
            }
        }

        //原始文件存放地址
        private string originalFilePath = string.Empty;
        /// <summary>
        /// 原始文件存放地址
        /// </summary>
        public string OriginalFilePath
        {
            get
            {
                if (originalFilePath == string.Empty)
                {
                    originalFilePath = AppDomain.CurrentDomain.BaseDirectory + ConfigurationManager.AppSettings["OriginalFilePath"].ToString();
                    //判断文件夹是否存在，不存在则创建
                    if (!System.IO.Directory.Exists(originalFilePath))
                    {
                        System.IO.Directory.CreateDirectory(originalFilePath);
                    }
                }
                return originalFilePath + "\\";
            }
            set
            {
                originalFilePath = value;
            }
        }

        //下载文件存放地址
        private string downloadFilePath = string.Empty;
        /// <summary>
        /// 下载文件存放地址
        /// </summary>
        public string DownloadFilePath
        {
            get
            {
                if (downloadFilePath == string.Empty)
                {
                    downloadFilePath = AppDomain.CurrentDomain.BaseDirectory + ConfigurationManager.AppSettings["DownloadFilePath"].ToString();
                    //判断文件夹是否存在，不存在则创建
                    if (!System.IO.Directory.Exists(downloadFilePath))
                    {
                        System.IO.Directory.CreateDirectory(downloadFilePath);
                    }
                }
                return downloadFilePath + "\\";
            }
            set
            {
                downloadFilePath = value;
            }
        }

        /// <summary>
        /// 页面初始化
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                GetDataBind();
            }
        }

        /// <summary>
        /// 获取绑定数据源
        /// </summary>
        public void GetDataBind()
        {
            System.Data.DataTable table = new System.Data.DataTable();
            table.Columns.Add("Title", typeof(string));
            table.Columns.Add("Address", typeof(string));
            DirectoryInfo root = new DirectoryInfo(PDFDocumentPath);
            foreach (FileInfo f in root.GetFiles())
            {
                DataRow row = table.NewRow();
                row["Title"] = Path.GetFileNameWithoutExtension(f.FullName);
                row["Address"] = "../" + ConfigurationManager.AppSettings["PDFDocumentPath"].ToString() + "/" + Path.GetFileName(f.FullName); 
                table.Rows.Add(row);
            }
            this.Repeater.DataSource = table;
            this.Repeater.DataBind();
        }

        /// <summary>
        /// Repeater点击事件
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        protected void Repeater_ItemCommand(object source, RepeaterCommandEventArgs e)
        {
            if (e.CommandName == "上传")
            {
                UploadFile();
            }
            if (e.CommandName == "下载")
            {
                DownloadFiles(e.CommandArgument.ToString());
            }
            if (e.CommandName == "删除")
            {
                //只删除PDF文件
                DeleteFile(PDFDocumentPath + Path.GetFileName(e.CommandArgument.ToString()));
            }
            if (e.CommandName == "合并")
            {
                Merge(e.CommandArgument.ToString());
            }
        }

        /// <summary>
        /// 下载
        /// </summary>
        protected void DownloadFiles(string filesName)
        {
            //删除之前下载的压缩包
            DirectoryInfo roots = new DirectoryInfo(DownloadFilePath);
            foreach (FileInfo f in roots.GetFiles())
            {
                DeleteFile(f.FullName);
            }
            List<string> filePathList = new List<string>();
            //添加下载文件地址（PDF）
            filePathList.Add(PDFDocumentPath + Path.GetFileName(filesName));
            //获取对应原文件地址并添加
            DirectoryInfo root = new DirectoryInfo(OriginalFilePath);
            foreach (FileInfo f in root.GetFiles())
            {
                if (Path.GetFileNameWithoutExtension(f.FullName) == Path.GetFileNameWithoutExtension(filesName))
                {
                    filePathList.Add(f.FullName);
                }
            }
            //存放下载压缩包地址
            string zipPath = DownloadFilePath + Path.GetFileNameWithoutExtension(filesName) + ".zip";
            ZipFiles(filePathList.ToArray(), "", zipPath);
            FileInfo fileInfo = new FileInfo(zipPath);
            Response.Clear();
            Response.Charset = "GB2312";
            Response.ContentEncoding = System.Text.Encoding.UTF8;
            Response.AddHeader("Content-Disposition", "attachment;filename=" + Server.UrlEncode(fileInfo.Name));
            Response.AddHeader("Content-Length", fileInfo.Length.ToString());
            Response.ContentType = "application/x-bittorrent";
            Response.WriteFile(fileInfo.FullName);
            Response.End();
        }

        /// <summary>
        /// 压缩文件
        /// </summary>
        /// <param name="dirPath">文件夹路径</param>
        /// <param name="password">压缩包设置密码(注：可为空)</param>
        /// <param name="zipFilePath">压缩包路径+名称+后缀(注：可为空,默认同目录)</param>
        /// <returns></returns>
        public string ZipFiles(string[] filePaths, string password, string zipFilePath)
        {
            try
            {
                using (ZipOutputStream s = new ZipOutputStream(System.IO.File.Create(zipFilePath)))
                {
                    s.SetLevel(9);
                    s.Password = password;
                    byte[] buffer = new byte[4096];
                    foreach (string file in filePaths)
                    {
                        ZipEntry entry = new ZipEntry(Path.GetFileName(file));
                        entry.DateTime = DateTime.Now;
                        s.PutNextEntry(entry);
                        using (FileStream fs = System.IO.File.OpenRead(file))
                        {
                            int sourceBytes;
                            do
                            {
                                sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                s.Write(buffer, 0, sourceBytes);
                            } while (sourceBytes > 0);
                        }
                    }
                    s.Finish();
                    s.Close();
                }
                return zipFilePath;
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

        /// <summary>
        /// 根据路径删除文件
        /// </summary>
        /// <param name="path"></param>
        public void DeleteFile(string path)
        {
            System.IO.File.Delete(path);
            GetDataBind();
        }

        /// <summary>
        /// 上传
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected string UploadFile()
        {
            if (File.HasFile)
            {
                string filePath = OriginalFilePath + File.PostedFile.FileName.Trim();
                File.SaveAs(filePath);
                string newfilePath = PDFDocumentPath + Path.GetFileNameWithoutExtension(filePath) + ".pdf".Trim();
                if (System.IO.Path.GetExtension(filePath) == ".docx")
                {
                    WordConvertPDF(filePath, newfilePath);
                }
                else if (System.IO.Path.GetExtension(filePath) == ".xlsx")
                {
                    ExcelConvertPDF(filePath, newfilePath);
                }
                else if (System.IO.Path.GetExtension(filePath) == ".txt")
                {
                    TxtConvertPDF(filePath, newfilePath);
                }
                else
                {
                    Response.Write("<p >不支持的文件" + System.IO.Path.GetExtension(filePath) + "!</p>");
                }
                GetDataBind();
                return newfilePath;
            }
            else
            {
                Response.Write("<p >请选择文件!</p>");
                return string.Empty;
            } 
        }

        /// <summary>
        /// 合并PDF文件
        /// </summary>
        private void Merge(string filesName)
        {
            string mergeFilePath = UploadFile();
            if (mergeFilePath != string.Empty)
            {
                MergePdf(PDFDocumentPath + Path.GetFileName(filesName), mergeFilePath, PDFDocumentPath + Path.GetFileNameWithoutExtension(filesName) + "-" + Path.GetFileName(mergeFilePath));
                GetDataBind();
            }
        }

       /// <summary>
        /// 合并PDF文件
       /// </summary>
       /// <param name="originalFilePath">原始文件地址</param>
       /// <param name="mergeFilePath">合并文件地址</param>
       /// <param name="newFilePath">生成文件地址</param>
        private void MergePdf(string originalFilePath, string mergeFilePath, string newFilePath)
        {
            string[] fileList = { originalFilePath, mergeFilePath };
            List<PdfReader> readerList = new List<PdfReader>();
            PdfReader reader = null;
            iTextSharp.text.Rectangle rec = new iTextSharp.text.Rectangle(1660, 1000);
            iTextSharp.text.Document document = new iTextSharp.text.Document(rec);
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(newFilePath, FileMode.Create));
            document.Open();
            PdfContentByte cb = writer.DirectContent;
            PdfImportedPage newPage;
            foreach (var item in fileList)
            {
                reader = new PdfReader(item);
                int iPageNum = reader.NumberOfPages;
                for (int j = 1; j <= iPageNum; j++)
                {
                    document.NewPage();
                    newPage = writer.GetImportedPage(reader, j);
                    cb.AddTemplate(newPage, 0, 0);
                }
                readerList.Add(reader);
            }
            document.Close();
            foreach (var item in readerList)
            {
                item.Dispose();
            }
        }

        #region 将word文档转换成PDF格式

        /// <summary>
        /// 将word文档转换成PDF格式
        /// </summary>
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param>
        /// <returns></returns>
        private bool WordConvertPDF(string sourcePath, string targetPath)
        {
            bool result;
            Microsoft.Office.Interop.Word.WdExportFormat exportFormat = Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF;   //PDF格式
            object paramMissing = Type.Missing;
            Microsoft.Office.Interop.Word.ApplicationClass wordApplication = new Microsoft.Office.Interop.Word.ApplicationClass();
            Microsoft.Office.Interop.Word.Document wordDocument = null;
            try
            {
                object paramSourceDocPath = sourcePath;
                string paramExportFilePath = targetPath;

                Microsoft.Office.Interop.Word.WdExportFormat paramExportFormat = exportFormat;
                bool paramOpenAfterExport = false;
                Microsoft.Office.Interop.Word.WdExportOptimizeFor paramExportOptimizeFor =
                        Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint;
                Microsoft.Office.Interop.Word.WdExportRange paramExportRange = Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument;
                int paramStartPage = 0;
                int paramEndPage = 0;
                Microsoft.Office.Interop.Word.WdExportItem paramExportItem = Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent;
                bool paramIncludeDocProps = true;
                bool paramKeepIRM = true;
                Microsoft.Office.Interop.Word.WdExportCreateBookmarks paramCreateBookmarks =
                        Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks;
                bool paramDocStructureTags = true;
                bool paramBitmapMissingFonts = true;
                bool paramUseISO19005_1 = false;

                wordDocument = wordApplication.Documents.Open(
                        ref paramSourceDocPath, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing);

                if (wordDocument != null)
                {
                    wordDocument.ExportAsFixedFormat(paramExportFilePath,
                            paramExportFormat, paramOpenAfterExport,
                            paramExportOptimizeFor, paramExportRange, paramStartPage,
                            paramEndPage, paramExportItem, paramIncludeDocProps,
                            paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                            paramBitmapMissingFonts, paramUseISO19005_1,
                            ref paramMissing);
                    result = true;
                }
                else
                    result = false;
            }
            catch (Exception ex)
            {
                Response.Write("<p>Word转PDF出现问题，详情："+ex.ToString()+"</p>");
                result = false;
            }
            finally
            {
                if (wordDocument != null)
                {
                    wordDocument.Close(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordDocument = null;
                }
                if (wordApplication != null)
                {
                    wordApplication.Quit(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordApplication = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }

        #endregion

        #region 将excel文档转换成PDF格式

        /// <summary>
        /// 将excel文档转换成PDF格式
        /// </summary>
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param>
        /// <returns></returns>
        private bool ExcelConvertPDF(string sourcePath, string targetPath)
        {
            bool result;
            Microsoft.Office.Interop.Excel.XlFixedFormatType targetType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF; //PDF格式
            object missing = Type.Missing;
            Microsoft.Office.Interop.Excel.ApplicationClass application = null;
            Microsoft.Office.Interop.Excel.Workbook workBook = null;
            try
            {
                application = new Microsoft.Office.Interop.Excel.ApplicationClass();
                object target = targetPath;
                object type = targetType;
                workBook = application.Workbooks.Open(sourcePath, missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing, missing, missing, missing, missing);
                if (workBook != null)
                {
                    workBook.ExportAsFixedFormat(targetType, target, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, true, false, missing, missing, missing, missing);
                    result = true;
                }
                else
                    result = false;
            }
            catch (Exception ex)
            {
                Response.Write("<p>Excel转PDF出现问题，详情：" + ex.ToString() + "</p>");
                result = false;
            }
            finally
            {
                if (workBook != null)
                {
                    workBook.Close(true, missing, missing);
                    workBook = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }

        #endregion

        #region 将txt文档转换成PDF格式

        /// <summary>
        /// 将txt文档转换成PDF格式
        /// </summary>
        /// <param name="sourcePath">源文件路径</param>
        /// <param name="targetPath">目标文件路径</param>
        /// <returns></returns>
        private bool TxtConvertPDF(string sourcePath, string targetPath)
        {
            try
            {
                //第一个参数是txt文件物理路径
                string[] lines = System.IO.File.ReadAllLines(sourcePath,Encoding.GetEncoding("GB2312"));
                //iTextSharp.text.PageSize.A4    自定义页面大小
                iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 20, 20, 20);
                PdfWriter pdfwriter = PdfWriter.GetInstance(doc, new FileStream(targetPath, FileMode.Create));
                doc.Open();
                BaseFont baseFont = BaseFont.CreateFont("C:\\WINDOWS\\FONTS\\STSONG.TTF",BaseFont.IDENTITY_H,BaseFont.NOT_EMBEDDED);
                iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLUE);
                iTextSharp.text.Paragraph paragraph;
                foreach (string line in lines)
                {
                    paragraph = new iTextSharp.text.Paragraph(line, font);
                    doc.Add(paragraph);
                }
                doc.Close();
                return true;
            }
            catch (Exception ex)
            {
                Response.Write("<p>Txt转PDF出现问题，详情：" + ex.ToString() + "</p>");
                return false;
            }
           
        }

        #endregion
    }
}