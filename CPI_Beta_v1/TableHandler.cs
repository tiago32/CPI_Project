using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CPI_Beta_v1
{
    public class TableHandler
    {
        public List<Intervention> BuildInterventions(DataTable dt)
        {
            var majorList = new List<Intervention>();

            ////The first line of the file is the header.
            //var first = true;



            //Split each line.
            foreach (DataRow row in dt.Rows)
            {
                ////To discard the header file.
                //if (first || array[2] == string.Empty)
                //{
                //    first = false;
                //    continue;
                //}
                //Find an intervention by the identifier in majorList to add dates intervention. 
                //If the result is null then create a new one.
                var intervention = majorList.FirstOrDefault(m => m.ID_Manutencao.Equals(row["id_equipamento"].ToString()));
                if (intervention == null)
                {
                    intervention = new Intervention { ID_Manutencao = row["id_manutencao"].ToString(), ID_Equipamento = row["id_equipamento"].ToString(), Description = row["designacao"].ToString() };
                    //try
                    //{
                    //    //Selects the number contained in the identifier to serve in the ranking list.
                    //    intervention.NumberId = Int16.Parse(Regex.Match(array[0], @"\d+").Value);
                    //}
                    //catch (FormatException exception)
                    //{

                    //    Console.WriteLine(exception.Message);
                    //}

                    majorList.Add(intervention);
                }
                try
                {
                    ////Parse the date in portuguese culture.
                    //_cultureInfo = new CultureInfo("pt-PT");

                    ////Date format pattern.
                    //_ddMmmYy = "dd-MMM-yy";

                    intervention.MarkedInterventionsList.Add(DateTime.Parse(row["data_agendada"].ToString()));
                }
                catch (FormatException exception)
                {

                    Console.WriteLine(exception.Message);
                }
            }
            //Sorts the list in ascending order of NumberId.
            return majorList.OrderBy(x => x.NumberId).ToList();

        }
    }
}
