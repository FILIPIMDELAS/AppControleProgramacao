using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using Microsoft.AspNetCore.Mvc;
using Programação.Models;
using iText.Layout.Properties;
using iText.Kernel.Font;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Colors;
using iText.Commons.Actions.Contexts;
using OfficeOpenXml;
using iText.Layout.Borders;
using System.IO;
using System.IO.Compression;
using Amazon.CloudWatchEvents.Model;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Collections;
using Org.BouncyCastle.Ocsp;


[Route("api/export")]
[ApiController]

public class ExportPDF : ControllerBase
{

    [HttpGet]
    public async Task<List<(string fileName, byte[] pdfBytes)>> ExportPDFEng(IList<IFormFile> arquivos)
    {
        List<GetDate> Date = new List<GetDate>();
        foreach (var item in arquivos)
        {
            using (var strem = new MemoryStream())
            {
                await item.CopyToAsync(strem);
                using (var packege = new ExcelPackage(strem))
                {
                    var worksheet = packege.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;

                    for (int i = 7; i <= rowCount; i++)
                    {
                        GetDate newDate = new GetDate();
                        newDate.Codigo = worksheet.Cells[i, 3].Text; ;
                        newDate.Cordenador = worksheet.Cells[i, 4].Text;
                        newDate.Engenheiro = worksheet.Cells[i, 5].Text;
                        newDate.CodigoCed = worksheet.Cells[i, 6].Text;
                        newDate.Responsavel = worksheet.Cells[i, 7].Text;
                        newDate.Empresa = worksheet.Cells[i, 8].Text;
                        newDate.Equipe = worksheet.Cells[i, 9].Text;
                        newDate.LogisticaInterna = worksheet.Cells[i, 10].Text;
                        newDate.Pacote = worksheet.Cells[i, 11].Text;
                        newDate.Servico = worksheet.Cells[i, 12].Text;
                        newDate.Seg = worksheet.Cells[i, 13].Text;
                        newDate.Ter = worksheet.Cells[i, 14].Text;
                        newDate.Qua = worksheet.Cells[i, 15].Text;
                        newDate.Qui = worksheet.Cells[i, 16].Text;
                        newDate.Sex = worksheet.Cells[i, 17].Text;
                        newDate.Sab = worksheet.Cells[i, 18].Text;
                        newDate.Dom = worksheet.Cells[i, 19].Text;
                        newDate.Seg2 = worksheet.Cells[i, 20].Text;
                        newDate.Ter2 = worksheet.Cells[i, 21].Text;
                        newDate.Qua2 = worksheet.Cells[i, 22].Text;
                        newDate.Qui2 = worksheet.Cells[i, 23].Text;
                        newDate.Sex2 = worksheet.Cells[i, 24].Text;
                        newDate.Sab2 = worksheet.Cells[i, 25].Text;
                        newDate.Dom2 = worksheet.Cells[i, 26].Text;
                        newDate.Seg3 = worksheet.Cells[i, 27].Text;
                        newDate.Ter3 = worksheet.Cells[i, 28].Text;
                        newDate.Qua3 = worksheet.Cells[i, 29].Text;
                        newDate.Qui3 = worksheet.Cells[i, 30].Text;
                        newDate.Sex3 = worksheet.Cells[i, 31].Text;

                        Date.Add(newDate);
                    }
                }
            }
        }

        string calibriFontPath = @"C:/Users/fcarmo/Documents/Programação/calibri.ttf";
        string calibriFontPathBold = @"C:/Users/fcarmo/Documents/Programação/calibri-bold.ttf";

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var EmpresaUnicos = Date
            .Select(d => d.Empresa)
            .Where(engenheiro => !string.IsNullOrEmpty(engenheiro))
            .Distinct()
            .ToList();

        var CordenadorUnicos = Date
            .Select(d => d.Cordenador)
            .Where(engenheiro => !string.IsNullOrEmpty(engenheiro))
            .Distinct()
            .ToList();

        var EncarregadoUnicos = Date
            .Select(d => d.Responsavel)
            .Where(engenheiro => !string.IsNullOrEmpty(engenheiro))
            .Distinct()
            .ToList();

        var EngenheirosUnicos = Date
            .Select(d => d.Engenheiro)
            .Where(engenheiro => !string.IsNullOrEmpty(engenheiro))
            .Distinct()
            .ToList();

        var DateOffVazio = Date
            .Where(d => (!string.IsNullOrEmpty(d.Engenheiro) &&
                        !string.IsNullOrEmpty(d.Cordenador)) ||
                        d.Codigo == ".")
            .ToList();

        DateTime today = DateTime.Today;
        var seg = GetNextMonday(today);

        DateTime GetNextMonday(DateTime currentDate)
        {
            int daysUntilNextMonday = ((int)DayOfWeek.Monday - (int)currentDate.DayOfWeek + 7) % 7;

            if (daysUntilNextMonday == 0)
            {
                daysUntilNextMonday = 7;
            }

            return currentDate.AddDays(daysUntilNextMonday);
        }

        var pdfList = new List<(string fileName, byte[] pdfBytes)>();

        foreach (var empresa in EmpresaUnicos)
        {
            List<GetDate> FilteredEmpresa = Date.Where(x => x.Empresa == empresa || x.Codigo == ".").ToList();

            using (var memoryStream = new MemoryStream())
            {
                var writer = new PdfWriter(memoryStream);
                var pdf = new PdfDocument(writer);
                pdf.SetDefaultPageSize(PageSize.A4.Rotate());
                var document = new Document(pdf);

                var calibriFont = PdfFontFactory.CreateFont(calibriFontPath, PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);
                var calibriFontBold = PdfFontFactory.CreateFont(calibriFontPathBold, PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);
                document.SetFont(calibriFont);
                document.SetMargins(10, 10, 10, 10);

                var empresaNome = empresa.Replace(" ", "_");
                var currentDate = $"{seg:dd.MM.yyyy}";
                var fileName = $"PPC_SEMANAL_EMP_REV {currentDate} - {empresaNome}.pdf";

                fileName = fileName.Replace("/", "_").Replace("\\", "_");

                fileName = $"Empresa/{fileName}";

                float[] columnWidths = { 70f, 70F, 120f, 290f, 30f, 30f, 30f, 30f, 30f, 5f, 5f, 30f, 30f, 30f, 30f, 30f };

                Paragraph paragraph = new Paragraph("PROGRAMAÇÃO OBRA")
                    .SetFontSize(10)
                    .SetFont(calibriFontBold)
                    .SetFontColor(ColorConstants.BLACK)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetBackgroundColor(new DeviceRgb(242, 242, 242))
                    .SetWidth(pdf.GetDefaultPageSize().GetWidth());

                document.Add(paragraph);

                Paragraph spaces = new Paragraph("")
                    .SetFontSize(5)
                    .SetFont(calibriFontBold)
                    .SetFontColor(ColorConstants.BLACK)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetBackgroundColor(new DeviceRgb(242, 242, 242))
                    .SetWidth(pdf.GetDefaultPageSize().GetWidth());

                document.Add(spaces);

                Table tabela = new Table(columnWidths);

                Cell cell = new Cell();

                cell.SetBorder(Border.NO_BORDER).SetBackgroundColor(new DeviceRgb(255, 255, 255));

                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell(1, 5)
                        .Add(new Paragraph("SEMANA ATUAL"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell(1, 5)
                        .Add(new Paragraph("SEMANA SEGUINTE"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );

                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg:dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(1):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(2):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(3):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(4):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(7):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(8):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(9):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(10):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(11):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );

                tabela.AddCell(
                    new Cell().Add(new Paragraph("Empresa"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("Equipe"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("Pacote de Trabalho"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("Serviço"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEG2"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("TER2"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUA2"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUI2"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEX2"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEG2"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("TER3"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUA3"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUI3"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEX3"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(10)
                );

                for (int i = 0; i < FilteredEmpresa.Count; i++)
                {
                    if (i > 0)
                    {
                        if (FilteredEmpresa[i - 1].Codigo == "." && FilteredEmpresa[i].Codigo == ".")
                        {
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                        }
                        else
                        {
                            var corDeFundoPadrao = new DeviceRgb(217, 225, 242);
                            var corDeFundoLaranja = new DeviceRgb(252, 228, 214);
                            var corDeFundoSabado = new DeviceRgb(217, 217, 217);
                            var corDeFundoDomingo = new DeviceRgb(255, 133, 133);

                            if (FilteredEmpresa[i].Codigo == ".")
                            {
                                corDeFundoPadrao = new DeviceRgb(255, 255, 255);
                                corDeFundoLaranja = new DeviceRgb(255, 255, 255);
                            }

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEmpresa[i].Empresa))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(8)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEmpresa[i].Equipe))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(8)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEmpresa[i].Pacote))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(8)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEmpresa[i].Servico))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(8)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEmpresa[i].Seg2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(8)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEmpresa[i].Ter2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(8)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEmpresa[i].Qua2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(8)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEmpresa[i].Qui2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(8)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEmpresa[i].Sex2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(8)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph())
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoSabado)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(8)
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetHeight(10)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph())
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoDomingo)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(8)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEmpresa[i].Seg3))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(8)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEmpresa[i].Ter3))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(8)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEmpresa[i].Qua3))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(8)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEmpresa[i].Qui3))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(8)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEmpresa[i].Sex3))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(8)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                        }
                    }
                }

                document.Add(tabela);
                document.Close();

                pdfList.Add((fileName, memoryStream.ToArray()));
            }
        }

        foreach (var engenheiro in EngenheirosUnicos)
        {
            List<GetDate> FilteredEngenheiro = Date.Where(x => x.Engenheiro == engenheiro || x.Codigo == ".").ToList();

            using (var memoryStream = new MemoryStream())
            {
                var writer = new PdfWriter(memoryStream);
                var pdf = new PdfDocument(writer);
                pdf.SetDefaultPageSize(PageSize.A4.Rotate());
                var document = new Document(pdf);

                var calibriFont = PdfFontFactory.CreateFont(calibriFontPath, PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);
                var calibriFontBold = PdfFontFactory.CreateFont(calibriFontPathBold, PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);
                document.SetFont(calibriFont);
                document.SetMargins(10, 10, 10, 10);

                var engenheiroNome = engenheiro.Replace(" ", "_");
                var currentDate = $"{seg:dd.MM.yyyy}";
                var fileName = $"PPC_SEMANAL_ENG_REV {currentDate} - {engenheiroNome}.pdf";

                fileName = fileName.Replace("/", "_").Replace("\\", "_");

                fileName = $"Engenheiro/{fileName}";

                float[] columnWidths = { 30f, 45f, 45f, 45f, 110f, 270f, 20f, 20f, 20f, 20f, 20f, 20f, 20f, 20f, 20f, 20f, 10f, 20f, 20f, 20f, 20f, 20f };

                Paragraph paragraph = new Paragraph("PROGRAMAÇÃO OBRA")
                    .SetFontSize(10)
                    .SetFont(calibriFontBold)
                    .SetFontColor(ColorConstants.BLACK)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetBackgroundColor(new DeviceRgb(242, 242, 242))
                    .SetWidth(pdf.GetDefaultPageSize().GetWidth());

                document.Add(paragraph);

                Paragraph spaces = new Paragraph("")
                    .SetFontSize(5)
                    .SetFont(calibriFontBold)
                    .SetFontColor(ColorConstants.BLACK)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetBackgroundColor(new DeviceRgb(242, 242, 242))
                    .SetWidth(pdf.GetDefaultPageSize().GetWidth());

                document.Add(spaces);

                Paragraph Cord = new Paragraph($" ENGENHEIRO: {engenheiro}           PPC SEMANAL:  0%")
                    .SetFontSize(8)
                    .SetFont(calibriFontBold)
                    .SetFontColor(ColorConstants.BLACK)
                    .SetTextAlignment(TextAlignment.LEFT)
                    .SetBackgroundColor(new DeviceRgb(242, 242, 242))
                    .SetWidth(170);

                document.Add(Cord);

                Table tabela = new Table(columnWidths);

                Cell cell = new Cell();

                cell.SetBorder(Border.NO_BORDER).SetBackgroundColor(new DeviceRgb(255, 255, 255));

                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell(1, 5)
                        .Add(new Paragraph("SEMANA PLANEJADA"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell(1, 5)
                        .Add(new Paragraph("SEMANA ACOMPANHADA"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell(1, 5)
                        .Add(new Paragraph("PROBLEMAS"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );

                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg:dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(1):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(2):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(3):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(4):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg:dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(1):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(2):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(3):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(4):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg:dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(1):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(2):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(3):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(4):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );

                tabela.AddCell(
                    new Cell().Add(new Paragraph("Coordenador"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("Responsavel"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("Empresa"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("Equipe"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("Pacote"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("Serviço"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEG"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("TER"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUA"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUI"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEX"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEG"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("TER"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUA"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUI"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEX"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("EXEC%"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(255, 255, 255))
                        .SetFontColor(ColorConstants.BLACK)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEG"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("TER"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUA"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUI"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEX"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );

                for (int i = 0; i < FilteredEngenheiro.Count; i++)
                {
                    if (i > 0)
                    {
                        if (FilteredEngenheiro[i - 1].Codigo == "." && FilteredEngenheiro[i].Codigo == ".")
                        {
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                        }
                        else
                        {
                            var corDeFundoPadrao = new DeviceRgb(217, 225, 242);
                            var corDeFundoLaranja = new DeviceRgb(252, 228, 214);
                            var corDeFundo0 = new DeviceRgb(0, 0, 0);

                            if (FilteredEngenheiro[i].Codigo == ".")
                            {
                                corDeFundoPadrao = new DeviceRgb(255, 255, 255);
                                corDeFundoLaranja = new DeviceRgb(255, 255, 255);
                                corDeFundo0 = new DeviceRgb(255, 255, 255);
                            }

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEngenheiro[i].Cordenador))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetHeight(10)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEngenheiro[i].Responsavel))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEngenheiro[i].Empresa))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEngenheiro[i].Equipe))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEngenheiro[i].Pacote))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEngenheiro[i].Servico))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEngenheiro[i].Seg2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEngenheiro[i].Ter2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEngenheiro[i].Qua2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEngenheiro[i].Qui2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEngenheiro[i].Sex2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph("0%"))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(corDeFundo0)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                        }
                    }
                }

                document.Add(tabela);
                document.Close();

                pdfList.Add((fileName, memoryStream.ToArray()));
            }
        }

        foreach (var cordenador in CordenadorUnicos)
        {
            List<GetDate> FilteredCordenador = Date.Where(x => x.Cordenador == cordenador || x.Codigo == ".").ToList();

            using (var memoryStream = new MemoryStream())
            {
                var writer = new PdfWriter(memoryStream);
                var pdf = new PdfDocument(writer);
                pdf.SetDefaultPageSize(PageSize.A4.Rotate());
                var document = new Document(pdf);

                var calibriFont = PdfFontFactory.CreateFont(calibriFontPath, PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);
                var calibriFontBold = PdfFontFactory.CreateFont(calibriFontPathBold, PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);
                document.SetFont(calibriFont);
                document.SetMargins(10, 10, 10, 10);

                var CordenadorNome = cordenador.Replace(" ", "_");
                var currentDate = $"{seg:dd.MM.yyyy}";
                var fileName = $"PPC_SEMANAL_COORD_REV {currentDate} - {CordenadorNome}.pdf";

                fileName = fileName.Replace("/", "_").Replace("\\", "_");

                fileName = $"Coordenador/{fileName}";

                float[] columnWidths = { 45f, 45f, 45f, 110f, 270f, 20f, 20f, 20f, 20f, 20f, 20f, 20f, 20f, 20f, 20f, 10f, 20f, 20f, 20f, 20f, 20f };

                Paragraph paragraph = new Paragraph("PROGRAMAÇÃO OBRA")
                    .SetFontSize(10)
                    .SetFont(calibriFontBold)
                    .SetFontColor(ColorConstants.BLACK)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetBackgroundColor(new DeviceRgb(242, 242, 242))
                    .SetWidth(pdf.GetDefaultPageSize().GetWidth());

                document.Add(paragraph);

                Paragraph spaces = new Paragraph("")
                    .SetFontSize(5)
                    .SetFont(calibriFontBold)
                    .SetFontColor(ColorConstants.BLACK)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetBackgroundColor(new DeviceRgb(242, 242, 242))
                    .SetWidth(pdf.GetDefaultPageSize().GetWidth());

                document.Add(spaces);

                Paragraph Cord = new Paragraph($" COORDENADOR: {cordenador}           PPC SEMANAL:  0%")
                    .SetFontSize(8)
                    .SetFont(calibriFontBold)
                    .SetFontColor(ColorConstants.BLACK)
                    .SetTextAlignment(TextAlignment.LEFT)
                    .SetBackgroundColor(new DeviceRgb(242, 242, 242))
                    .SetWidth(170);

                document.Add(Cord);

                Table tabela = new Table(columnWidths);

                Cell cell = new Cell();

                cell.SetBorder(Border.NO_BORDER).SetBackgroundColor(new DeviceRgb(255, 255, 255));

                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell(1, 5)
                        .Add(new Paragraph("SEMANA PLANEJADA"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell(1, 5)
                        .Add(new Paragraph("SEMANA ACOMPANHADA"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell(1, 5)
                        .Add(new Paragraph("PROBLEMAS"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );

                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg:dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(1):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(2):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(3):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(4):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg:dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(1):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(2):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(3):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(4):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg:dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(1):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(2):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(3):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(4):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );

                tabela.AddCell(
                    new Cell().Add(new Paragraph("Responsavel"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("Empresa"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("Equipe"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("Pacote"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("Serviço"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEG"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("TER"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUA"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUI"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEX"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEG"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("TER"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUA"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUI"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEX"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("EXEC%"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(255, 255, 255))
                        .SetFontColor(ColorConstants.BLACK)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEG"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("TER"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUA"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUI"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEX"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );

                for (int i = 0; i < FilteredCordenador.Count; i++)
                {
                    if (i > 0)
                    {
                        if (FilteredCordenador[i - 1].Codigo == "." && FilteredCordenador[i].Codigo == ".")
                        {
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                        }
                        else
                        {
                            var corDeFundoPadrao = new DeviceRgb(217, 225, 242);
                            var corDeFundoLaranja = new DeviceRgb(252, 228, 214);
                            var corDeFundo0 = new DeviceRgb(0, 0, 0);

                            if (FilteredCordenador[i].Codigo == ".")
                            {
                                corDeFundoPadrao = new DeviceRgb(255, 255, 255);
                                corDeFundoLaranja = new DeviceRgb(255, 255, 255);
                                corDeFundo0 = new DeviceRgb(255, 255, 255);
                            }

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredCordenador[i].Responsavel))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredCordenador[i].Empresa))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredCordenador[i].Equipe))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredCordenador[i].Pacote))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredCordenador[i].Servico))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredCordenador[i].Seg2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredCordenador[i].Ter2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredCordenador[i].Qua2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredCordenador[i].Qui2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredCordenador[i].Sex2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph("0%"))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(corDeFundo0)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                        }
                    }
                }

                document.Add(tabela);
                document.Close();

                pdfList.Add((fileName, memoryStream.ToArray()));
            }
        }

        foreach (var encarregado in EncarregadoUnicos)
        {
            List<GetDate> FilteredEncarregado = Date.Where(x => x.Responsavel == encarregado || x.Codigo == ".").ToList();

            using (var memoryStream = new MemoryStream())
            {
                var writer = new PdfWriter(memoryStream);
                var pdf = new PdfDocument(writer);
                pdf.SetDefaultPageSize(PageSize.A4.Rotate());
                var document = new Document(pdf);

                var calibriFont = PdfFontFactory.CreateFont(calibriFontPath, PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);
                var calibriFontBold = PdfFontFactory.CreateFont(calibriFontPathBold, PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);
                document.SetFont(calibriFont);
                document.SetMargins(10, 10, 10, 10);

                var EncarregadoNome = encarregado.Replace(" ", "_");
                var currentDate = $"{seg:dd.MM.yyyy}";
                var fileName = $"PPC_SEMANAL_ENC_REV {currentDate} - {EncarregadoNome}.pdf";

                fileName = fileName.Replace("/", "_").Replace("\\", "_");

                fileName = $"Encarregado/{fileName}";

                float[] columnWidths = { 30f, 45f, 45f, 45f, 110f, 270f, 20f, 20f, 20f, 20f, 20f, 20f, 20f, 20f, 20f, 20f, 10f, 20f, 20f, 20f, 20f, 20f };

                Paragraph paragraph = new Paragraph("PROGRAMAÇÃO OBRA")
                    .SetFontSize(10)
                    .SetFont(calibriFontBold)
                    .SetFontColor(ColorConstants.BLACK)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetBackgroundColor(new DeviceRgb(242, 242, 242))
                    .SetWidth(pdf.GetDefaultPageSize().GetWidth());

                document.Add(paragraph);

                Paragraph spaces = new Paragraph("")
                    .SetFontSize(5)
                    .SetFont(calibriFontBold)
                    .SetFontColor(ColorConstants.BLACK)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetBackgroundColor(new DeviceRgb(242, 242, 242))
                    .SetWidth(pdf.GetDefaultPageSize().GetWidth());

                document.Add(spaces);

                Paragraph Cord = new Paragraph($" ENCARREGADO: {encarregado}           PPC SEMANAL:  0%")
                    .SetFontSize(8)
                    .SetFont(calibriFontBold)
                    .SetFontColor(ColorConstants.BLACK)
                    .SetTextAlignment(TextAlignment.LEFT)
                    .SetBackgroundColor(new DeviceRgb(242, 242, 242))
                    .SetWidth(170);

                document.Add(Cord);

                Table tabela = new Table(columnWidths);

                Cell cell = new Cell();

                cell.SetBorder(Border.NO_BORDER).SetBackgroundColor(new DeviceRgb(255, 255, 255));

                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell(1, 5)
                        .Add(new Paragraph("SEMANA PLANEJADA"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell(1, 5)
                        .Add(new Paragraph("SEMANA ACOMPANHADA"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell(1, 5)
                        .Add(new Paragraph("PROBLEMAS"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );

                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg:dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(1):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(2):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(3):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(4):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg:dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(1):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(2):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(3):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(4):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(cell);
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg:dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(1):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(2):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(3):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph($"{seg.AddDays(4):dd/MM}"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );

                tabela.AddCell(
                    new Cell().Add(new Paragraph("Engenheiro"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );

                tabela.AddCell(
                    new Cell().Add(new Paragraph("Responsavel"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("Empresa"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("Equipe"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("Pacote"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("Serviço"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEG"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("TER"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUA"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUI"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEX"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEG"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("TER"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUA"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUI"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEX"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("EXEC%"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(255, 255, 255))
                        .SetFontColor(ColorConstants.BLACK)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEG"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("TER"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUA"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("QUI"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );
                tabela.AddCell(
                    new Cell().Add(new Paragraph("SEX"))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBackgroundColor(new DeviceRgb(175, 26, 23))
                        .SetFontColor(ColorConstants.WHITE)
                        .SetFontSize(8)
                );

                for (int i = 0; i < FilteredEncarregado.Count; i++)
                {
                    if (i > 0)
                    {
                        if (FilteredEncarregado[i - 1].Codigo == "." && FilteredEncarregado[i].Codigo == ".")
                        {
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                            tabela.AddCell(cell);
                        }
                        else
                        {
                            var corDeFundoPadrao = new DeviceRgb(217, 225, 242);
                            var corDeFundoLaranja = new DeviceRgb(252, 228, 214);
                            var corDeFundo0 = new DeviceRgb(0, 0, 0);

                            if (FilteredEncarregado[i].Codigo == ".")
                            {
                                corDeFundoPadrao = new DeviceRgb(255, 255, 255);
                                corDeFundoLaranja = new DeviceRgb(255, 255, 255);
                                corDeFundo0 = new DeviceRgb(255, 255, 255);
                            }

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEncarregado[i].Engenheiro))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEncarregado[i].Responsavel))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEncarregado[i].Empresa))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEncarregado[i].Equipe))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEncarregado[i].Pacote))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEncarregado[i].Servico))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );

                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEncarregado[i].Seg2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEncarregado[i].Ter2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEncarregado[i].Qua2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEncarregado[i].Qui2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(FilteredEncarregado[i].Sex2))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph("0%"))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoPadrao)
                                    .SetFontColor(corDeFundo0)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                            tabela.AddCell(
                                new Cell().Add(new Paragraph(""))
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .SetBackgroundColor(corDeFundoLaranja)
                                    .SetFontColor(ColorConstants.BLACK)
                                    .SetFontSize(6)
                                    .SetTextAlignment(TextAlignment.CENTER)
                            );
                        }
                    }
                }

                document.Add(tabela);
                document.Close();

                pdfList.Add((fileName, memoryStream.ToArray()));
            }
        }

        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Relatório");

            worksheet.Cells[1, 1].Value = "Codigo";
            worksheet.Cells[1, 2].Value = "Cordenadoro";
            worksheet.Cells[1, 3].Value = "Engenheiro";
            worksheet.Cells[1, 4].Value = "CodigoCedo";
            worksheet.Cells[1, 5].Value = "Responsavel";
            worksheet.Cells[1, 6].Value = "Empresao";
            worksheet.Cells[1, 7].Value = "Equipe";
            worksheet.Cells[1, 8].Value = "LogisticaInternao";
            worksheet.Cells[1, 9].Value = "Pacote";
            worksheet.Cells[1, 10].Value = "Servicoo";
            worksheet.Cells[1, 11].Value = "Seg";
            worksheet.Cells[1, 12].Value = "Ter";
            worksheet.Cells[1, 13].Value = "Qua";
            worksheet.Cells[1, 14].Value = "Qui";
            worksheet.Cells[1, 15].Value = "Sex";
            worksheet.Cells[1, 16].Value = "Sab";
            worksheet.Cells[1, 17].Value = "Dom";
            worksheet.Cells[1, 18].Value = "Seg2";
            worksheet.Cells[1, 19].Value = "Ter2";
            worksheet.Cells[1, 20].Value = "Qua2";
            worksheet.Cells[1, 21].Value = "Qui2";
            worksheet.Cells[1, 22].Value = "Sex2";
            worksheet.Cells[1, 23].Value = "Sab2";
            worksheet.Cells[1, 24].Value = "Dom2";
            worksheet.Cells[1, 25].Value = "Seg3";
            worksheet.Cells[1, 26].Value = "Ter3";
            worksheet.Cells[1, 27].Value = "Qua3";
            worksheet.Cells[1, 28].Value = "Qui3";
            worksheet.Cells[1, 29].Value = "Sex3";
            int row = 2;

            foreach (var dado in DateOffVazio)
            {
                if (dado.Codigo == ".")
                {
                    worksheet.Cells[row, 1].Value = ".";
                }
                else
                {
                    worksheet.Cells[row, 1].Value = row - 1;
                }
                worksheet.Cells[row, 2].Value = dado.Cordenador;
                worksheet.Cells[row, 3].Value = dado.Engenheiro;
                worksheet.Cells[row, 4].Value = dado.CodigoCed;
                worksheet.Cells[row, 5].Value = dado.Responsavel;
                worksheet.Cells[row, 6].Value = dado.Empresa;
                worksheet.Cells[row, 7].Value = dado.Equipe;
                worksheet.Cells[row, 8].Value = dado.LogisticaInterna;
                worksheet.Cells[row, 9].Value = dado.Pacote;
                worksheet.Cells[row, 10].Value = dado.Servico;
                worksheet.Cells[row, 11].Value = dado.Seg;
                worksheet.Cells[row, 12].Value = dado.Ter;
                worksheet.Cells[row, 13].Value = dado.Qua;
                worksheet.Cells[row, 14].Value = dado.Qui;
                worksheet.Cells[row, 15].Value = dado.Sex;
                worksheet.Cells[row, 16].Value = dado.Sab;
                worksheet.Cells[row, 17].Value = dado.Dom;
                worksheet.Cells[row, 18].Value = dado.Seg2;
                worksheet.Cells[row, 19].Value = dado.Ter2;
                worksheet.Cells[row, 20].Value = dado.Qua2;
                worksheet.Cells[row, 21].Value = dado.Qui2;
                worksheet.Cells[row, 22].Value = dado.Sex2;
                worksheet.Cells[row, 23].Value = dado.Sab2;
                worksheet.Cells[row, 24].Value = dado.Dom2;
                worksheet.Cells[row, 25].Value = dado.Seg3;
                worksheet.Cells[row, 26].Value = dado.Ter3;
                worksheet.Cells[row, 27].Value = dado.Qua3;
                worksheet.Cells[row, 28].Value = dado.Qui3;
                worksheet.Cells[row, 29].Value = dado.Sex3;
                row++;
            }

            var currentDate = $"{seg:dd.MM.yyyy}";
            var fileName = $"ALL PPC_SEMANAL_REV {currentDate}.xlsx";

            fileName = fileName.Replace("/", "_").Replace("\\", "_");

            pdfList.Add((fileName, package.GetAsByteArray()));
        }

        return pdfList;
    }


    [HttpPost]
    public async Task<IActionResult> ExportPdf( IList<IFormFile> arquivos)
    {
        Console.WriteLine(arquivos.Count);
        var pdfFiles = await ExportPDFEng(arquivos);

        using (var zipMemoryStream = new MemoryStream())
        {
            using (var archive = new ZipArchive(zipMemoryStream, ZipArchiveMode.Create, true))
            {
                foreach (var (fileName, pdfBytes) in pdfFiles)
                {
                    var zipEntry = archive.CreateEntry(fileName);

                    using (var entryStream = zipEntry.Open())
                    {
                        await entryStream.WriteAsync(pdfBytes, 0, pdfBytes.Length);
                    }
                }
            }

            zipMemoryStream.Seek(0, SeekOrigin.Begin);
            return File(zipMemoryStream.ToArray(), "application/zip", "Programação_PDF.zip");
        }
    }
}