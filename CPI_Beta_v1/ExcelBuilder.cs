using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using CPI_Beta_v1.Properties;
using Excel = Microsoft.Office.Interop.Excel;

namespace CPI_Beta_v1
{
    public class ExcelBuilder
    {
        readonly string[] _months = { "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO" };
        private int _lastInterventionPosition;
        readonly string[] _cellPosition = { "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P" /*, "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ" */};

        /// <summary>
        /// Generates an Excel Table Plan for all interventions for a specific year
        /// </summary>
        /// <param name="EquipmentsList">List with all the equipments and their interventions</param>
        /// <param name="interventionYear">Plan year</param>
        public void GenerateExcel(List<Equipment> EquipmentsList, int interventionYear, string tipoRelatorio)
        {
            try
            {
                //Start Excel and get Application object
                var oXl = new Excel.Application { Visible = false };

                //Get a new workbook
                Excel._Workbook oWb = oXl.Workbooks.Add(Missing.Value);

                var oSheet = (Excel._Worksheet)oWb.ActiveSheet;

                //Initial positon to start writing the equipments (one per row)
                _lastInterventionPosition = EquipmentsList.Count + 3;

                //Builds an multidimensional array with the identifier and description of each equipment
                var equipmentInfo = new string[EquipmentsList.Count, 4];
                var index = 0;

                foreach (var eq in EquipmentsList)
                {
                    equipmentInfo[index, 0] = eq.NumeroInventario;
                    equipmentInfo[index, 1] = eq.EqDescription;
                    equipmentInfo[index, 2] = eq.NumeroSerie;
                    equipmentInfo[index, 3] = eq.Periodicidade;
                    index++;
                }

                //Headers 2
                oSheet.Cells[2, 1] = "EQUIPAMENTO";
                oSheet.Cells[2, 5] = "MANUTENÇÕES PREVENTIVAS";

                //Headers 3
                oSheet.Cells[3, 1] = "Nº INVENTÁRIO";
                oSheet.Cells[3, 2] = "DESIGNAÇÃO";
                oSheet.Cells[3, 3] = "Nº SÉRIE";
                oSheet.Cells[3, 4] = "PERIODICIDADE";

                #region STYLES
                //Header 2
                var oRng = oSheet.Range["A2", "D2"];
                oRng.EntireColumn.AutoFit();
                Header2Style(oRng);

                oRng = oSheet.Range["E2", "P2"];
                oRng.EntireColumn.AutoFit();
                Header2Style(oRng);

                //Header 3
                oRng = oSheet.Range["A3", "D3"];
                oRng.EntireColumn.ColumnWidth = 20;
                Header3Style(oRng);

                //Fill A4:D... with the information from the equipment array
                oSheet.Range["A4", "D" + _lastInterventionPosition].Value2 = equipmentInfo;
                oRng = oSheet.Range["A4", "D" + _lastInterventionPosition];
                CommonStyle(oRng);

                //Specific modifications
                oRng = oSheet.Range["A1"];
                oRng.EntireColumn.AutoFit();

                //oRng = oSheet.Range["D4", "D" + _lastInterventionPosition];
                //oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                oRng = oSheet.Range["A2", "D2"];
                oRng.Merge();

                oRng = oSheet.Range["E2", "P2"];
                oRng.Merge();
                #endregion

                ////Configurations to get all weeks of this year
                //var jan1 = new DateTime(interventionYear, 1, 1);
                //var startOfFirstWeek = jan1.AddDays(1 - (int)(jan1.DayOfWeek));

                //will store the last column of the table
                string finalCell;

                //Reset variable to handle cells positions ahead
                index = 0;

                //For each months of the year get the correspondent weeks
                for (var j = 1; j <= 12; j++)
                {
                //    //Finds the correspondent weeks for a month
                //    var weeksOfMonth =
                //   Enumerable
                //       .Range(0, 54)
                //       .Select(y => new
                //       {
                //           weekStart = startOfFirstWeek.AddDays(y * 7)
                //       })
                //       .TakeWhile(x => x.weekStart.Year <= jan1.Year)
                //       .Select(x => new
                //       {
                //           x.weekStart,
                //           weekFinish = x.weekStart.AddDays(6)
                //       })
                //       .SkipWhile(x => x.weekFinish < jan1.AddDays(1))
                //       .Select((x, y) => new
                //       {
                //           x.weekStart,
                //           x.weekFinish,
                //           weekNum = y + 1
                //       }).Where(x => (x.weekStart.Month == j && x.weekStart.Year == interventionYear) ||
                //           (x.weekFinish.Month == j && x.weekFinish.Year == interventionYear))
                //  .Select(x => x.weekNum).ToList();

                    //Save the initial cell position where the month begins
                    var initialCell = _cellPosition[index] + "3";

                    ////Writes every week in excel
                    //foreach (var week in weeksOfMonth)
                    //{
                    //    oRng = oSheet.Range[_cellPosition[index] + "3"].Cells[1, 1];
                    //    //oRng.Cells.Value2 = week;
                    //    Header3Style(oRng);
                    index++;
                    //}
                    ////Save the final cell position where the month ends
                    //finalCell = _cellPosition[index - 1] + "2";

                    //Merge the initial cell with the final, write the month name and apply the correspondent style
                    oRng = oSheet.Range[initialCell, initialCell];
                    //oRng.Merge();
                    //oRng.ColumnWidth = 2;
                    oRng.EntireColumn.ColumnWidth = 16;
                    oRng.Cells[1, 1] = _months[j - 1];
                    Header2Style(oRng);
                }

                //Apply final style to the entire table
                finalCell = _cellPosition[index - 1] + (EquipmentsList.Count + 3);
                oRng = oSheet.Range["A1", finalCell];
                oRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                //Style for the title of the table
                oRng = oSheet.Range["A1", _cellPosition[index - 1] + "1"];
                oRng.Merge();
                oRng.Cells[1, 1] = "PLANO DE INTERVENÇÃO PREVENTIVA SIE " + interventionYear + " POR " + tipoRelatorio.ToUpper();
                Header1Style(oRng);

                //---------------------------------------------------------------------------------------------------

                //The interventions start row
                var startPosition = 4;

                //Bidimensional array that will store the periodicity of each intervention
                //var periodicity = new int[EquipmentsList.Count, 1];

                //var indexPeriodicity = 0;

                foreach (var eq in EquipmentsList)
                {
                    var lastCell = 0;
                    //var month = 0;

                    ////Estimation of the  periodicity of an intervention
                    //if (intervention.MarkedInterventionsList.Count == 1)
                    //{
                    //    periodicity[indexPeriodicity, 0] = 365;
                    //}
                    //else
                    //{
                    //    var list = intervention.MarkedInterventionsList.Take(2).ToList();
                    //    periodicity[indexPeriodicity, 0] = Math.Abs((int)(list.First() - list.Last()).TotalDays);

                    //}
                    //indexPeriodicity++;

                    foreach (var intervention in eq.InterventionsList)
                    {
                        //for (var j = lastCell; j < index; j++)
                        //{
                        //    ////Estimation of the  periodicity of an intervention
                        //    //var value2 = oSheet.Range[_cellPosition[j] + "2"].Value2;
                        //    //if (value2 != null)
                        //    //{
                        //    //    month++;
                        //    //}
                        //    if (intervention.Item1.Value.Month != month) continue;

                        //    //Write the specific day of the intervention in the correspondent week and style it
                        //    oSheet.Range[_cellPosition[j] + startPosition].Cells[1, 1] = interventionDate.Day;
                        //    oRng = oSheet.Range[_cellPosition[j] + startPosition];
                        //    oRng.Cells.Style = "Good";
                        //    oRng.Font.Name = "Calibri";
                        //    oRng.Font.Size = 10;
                        //    oRng.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;

                        //    //Store the last cell position that the next iteration will start for the weeks
                        //    lastCell = j + 1;
                        //    break;
                        //}

                        for (int j = 0; j < _months.Length; j++)
                        {
                            if (intervention.Item1.HasValue && intervention.Item1.Value.Month == j + 1)
                            {
                                oRng = oSheet.Range[_cellPosition[j] + startPosition];
                                oRng.Cells.Style = "Good";
                                oRng.Cells.Interior.Color = Color.FromArgb(217, 225, 242);
                            }
                            if (intervention.Item2.HasValue && intervention.Item2.Value.Month == j + 1)
                            {
                                switch (intervention.Item3.ToString())
                                {
                                    case "Aprovado":
                                        oSheet.Range[_cellPosition[j] + startPosition].Cells[1, 1] = intervention.Item2.Value.Day.ToString() + " - Aprovado";
                                        oRng.Font.Color = Color.Green;
                                        break;
                                    case "Aprovado Condicionado":
                                        oSheet.Range[_cellPosition[j] + startPosition].Cells[1, 1] = intervention.Item2.Value.Day.ToString() + " - A. Condicionado";
                                        oRng.Font.Color = Color.FromArgb(47, 117, 181);
                                        break;
                                    case "Não Aprovado":
                                        oSheet.Range[_cellPosition[j] + startPosition].Cells[1, 1] = intervention.Item2.Value.Day.ToString() + " - N. Aprovado";
                                        oRng.Font.Color = Color.Red;
                                        break;
                                }
       
                                oRng = oSheet.Range[_cellPosition[j] + startPosition];
                                //oRng.Cells.Style = "Good";
                                oRng.Font.Name = "Calibri";
                                oRng.Font.Size = 10;
                                oRng.Font.Bold = true;
                                oRng.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;                
                            }
                        }
                    }
                    startPosition++;
                }


                //FormatPeriodicity(interventionsList, oSheet, periodicity);
                oXl.Visible = true;

            }
            catch (Exception theException)
            {
                var errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, Resources.ExcelBuilder_GenerateExcel_Error);
            }
        }

        //private static void FormatPeriodicity(List<Intervention> interventionsList, Excel._Worksheet oSheet, int[,] periodicity)
        //{
        //    var oRng = oSheet.Range["C4", "C" + (interventionsList.Count + 3)];
        //    oRng.Value2 = periodicity;
        //    oRng.Font.Name = "Calibri";
        //    oRng.Font.Size = 10;
        //    oRng.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
        //}

        private static void Header1Style(Excel.Range oRng)
        {
            CommonHeadersStyle(oRng);
            oRng.Font.Color = Color.White;
            oRng.Cells.Interior.Color = Color.FromArgb(49, 134, 155);

        }


        private static void Header3Style(Excel.Range oRng)
        {

            CommonHeadersStyle(oRng);
            oRng.Cells.Interior.Color = Color.FromArgb(217, 217, 217);
        }

        private static void Header2Style(Excel.Range oRng)
        {
            CommonHeadersStyle(oRng);
            oRng.Font.Color = Color.FromArgb(74, 134, 168);
            oRng.Cells.Interior.Color = Color.FromArgb(183, 222, 232);
        }

        private static void CommonHeadersStyle(Excel.Range oRng)
        {
            CommonStyle(oRng);
            oRng.Font.Bold = true;
        }

        private static void CommonStyle(Excel.Range oRng)
        {
            oRng.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            oRng.Font.Name = "Calibri";
            oRng.Font.Size = 10;
        }

        private static Int16 GetWeek(DateTime date1)
        {
            var dfi = DateTimeFormatInfo.CurrentInfo;
            if (dfi == null) return -1;
            var cal = dfi.Calendar;

            return (short)cal.GetWeekOfYear(date1, dfi.CalendarWeekRule,
                dfi.FirstDayOfWeek);
        }
    }
}
