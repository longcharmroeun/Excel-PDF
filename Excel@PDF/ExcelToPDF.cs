using OfficeOpenXml;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_PDF
{
    public class ExcelToPDF
    {
        private Workbook Workbook;
        /// <summary>
        /// Load Excel From Stream
        /// </summary>
        /// <param name="stream">Excel Data</param>
        public ExcelToPDF(Stream stream)
        {
            this.Workbook = new Workbook();
            this.Workbook.LoadFromStream(stream);
        }
        /// <summary>
        /// Load Excel From File
        /// </summary>
        /// <param name="file">Fileinfo</param>
        public ExcelToPDF(FileInfo file)
        {
            this.Workbook = new Workbook();
            this.Workbook.LoadFromFile(file.FullName);
        }
        /// <summary>
        /// Save File Asynchronous
        /// </summary>
        /// <param name="filename">FileName</param>
        public void SaveFileAsync(string filename)
        {
            Task.Run(new Action(async () =>
            {
                List<Stream> streams = new List<Stream>();
                for (int i = 0; i < this.Workbook.Worksheets.Count; i++)
                {
                    Stream stream = new MemoryStream();
                    this.Workbook.Worksheets[i].SaveToPdfStream(stream);
                    streams.Add(stream);
                }
                PdfDocument pdf = await this.CombinePDFAsyn(streams);
                pdf.Save(filename);
            }));
        }
        /// <summary>
        /// Save File
        /// </summary>
        /// <param name="filename">FileName</param>
        public async void SaveFile(string filename)
        {
            List<Stream> streams = new List<Stream>();
            for (int i = 0; i < this.Workbook.Worksheets.Count; i++)
            {
                Stream stream = new MemoryStream();
                this.Workbook.Worksheets[i].SaveToPdfStream(stream);
                streams.Add(stream);
            }
            PdfDocument pdf = await this.CombinePDFAsyn(streams);
            pdf.Save(filename);
        }
        /// <summary>
        /// Save PDF To Stream
        /// </summary>
        /// <returns></returns>
        public async Task<PdfDocument> SaveToStreamAsync()
        {
            List<Stream> streams = new List<Stream>();
            for (int i = 0; i < this.Workbook.Worksheets.Count; i++)
            {
                Stream stream = new MemoryStream();
                this.Workbook.Worksheets[i].SaveToPdfStream(stream);
                streams.Add(stream);
            }
            return await this.CombinePDFAsyn(streams);
        }

        private Task<PdfDocument> CombinePDFAsyn(List<Stream> streams)
        {
            return Task.Run(new Func<PdfDocument>(() =>
            {
                using (PdfDocument outPdf = new PdfDocument())
                {
                    foreach (var item in streams)
                    {
                        using (PdfDocument one = PdfReader.Open(item, PdfDocumentOpenMode.Import))
                            CopyPages(one, outPdf);
                    }
                    return outPdf;

                    void CopyPages(PdfDocument from, PdfDocument to)
                    {
                        for (int i = 0; i < from.PageCount; i++)
                        {
                            to.AddPage(from.Pages[i]);
                        }
                    }
                }
            }));
        }
    }
}
