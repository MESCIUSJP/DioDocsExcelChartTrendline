// See https://aka.ms/new-console-template for more information
using GrapeCity.Documents.Excel;
using GrapeCity.Documents.Excel.Drawing;
using System.Drawing;

Console.WriteLine("ワークシートにチャートを追加して近似曲線を設定する");

// ワークシートにデータを追加
var workbook = new Workbook();
IWorksheet worksheet = workbook.Worksheets[0];
worksheet.Range["A1:B7"].Value = new object[,]
{
               {null, "2023年"},
               {"4月", 5200},
               {"5月", 1400},
               {"6月", 6600},
               {"7月", 4200},
               {"8月", 7300},
               {"9月", 5100},
};

// チャートを追加
IShape shape = worksheet.Shapes.AddChart(ChartType.LineMarkers, worksheet.Range["E2:K13"]);
shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B7"], RowCol.Columns, true, true);
// チャートのタイトルを設定
shape.Chart.HasTitle = true;
shape.Chart.ChartTitle.Text = "2023年度上期の売上";

// 近似曲線を設定
ISeries series1 = shape.Chart.SeriesCollection[0];
series1.Trendlines.Add();
series1.Trendlines[0].Type = TrendlineType.Linear;
series1.Trendlines[0].DisplayEquation = true;
series1.Trendlines[0].DisplayRSquared = true;

// 近似曲線のデータラベルをカスタマイズ
IDataLabel trendlineDataLabel = series1.Trendlines[0].DataLabel;
trendlineDataLabel.Font.Color.RGB = Color.White;
trendlineDataLabel.Font.Size = 9;
trendlineDataLabel.Font.Bold = true;
trendlineDataLabel.Font.Name = "ＭＳ 明朝";
trendlineDataLabel.Format.Fill.Color.RGB = Color.Red;
trendlineDataLabel.Format.Fill.Transparency = 0.5f;
trendlineDataLabel.Format.Line.Color.RGB = Color.FromArgb(255, 64, 64);
trendlineDataLabel.Text = "上昇傾向";

//ワークシートに印刷範囲を設定します
worksheet.PageSetup.Orientation = PageOrientation.Landscape;
worksheet.PageSetup.PrintArea = "$A$1:$L$14";

// Excelファイルに保存
workbook.Save("CreateTrendline.xlsx");

// PDFファイルに保存
workbook.Save("CreateTrendline.pdf");