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
        /// <summary>
        /// Creates a list of equipments and their interventions from a DataTable with the information
        /// </summary>
        /// <param name="dt">DataTable containing the equipments</param>
        /// <returns>Equipment list with the interventions</returns>
        public List<Equipment> BuildInterventions(DataTable dt)
        {
            //Main list to be returned
            var majorList = new List<Equipment>();

            //Process each DataTable row
            foreach (DataRow row in dt.Rows)
            {
                //Find an equipment by the identifier in majorList to add intervention dates 
                //If the result is null then create a new one
                var eq = majorList.FirstOrDefault(e => e.ID_Equipamento.Equals(row["ID Equipamento"].ToString()));
                if (eq == null)
                {
                    eq = new Equipment { ID_Equipamento = row["ID Equipamento"].ToString(), NumeroInventario = row["Número de Inventário"].ToString(), EqDescription = row["Designação"].ToString(), NumeroSerie = row["Número de Série"].ToString(), Periodicidade = row["Periodicidade de Manutenção"].ToString() };
                    majorList.Add(eq);
                }
                //Intervention Information (Scheduled Date, Performed Date, Decision)
                DateTime? scheduled; if (string.IsNullOrEmpty(row["data_agendada"].ToString())) { scheduled = null; } else { scheduled = DateTime.Parse(row["data_agendada"].ToString()); }
                DateTime? performed; if (string.IsNullOrEmpty(row["data_realizada"].ToString())) { performed = null; } else { performed = DateTime.Parse(row["data_realizada"].ToString()); }
                eq.InterventionsList.Add(new Tuple<DateTime?, DateTime?, string>(scheduled, performed, row["decisao"].ToString()));
            }
            //Sorts the list in ascending order of NumeroInventario
            return majorList.OrderBy(x => x.NumeroInventario).ToList();

        }
    }
}
