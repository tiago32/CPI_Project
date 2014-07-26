using System;
using System.Collections.Generic;

namespace CPI_Beta_v1
{
    public class Equipment
    {

        private String _id_equipamento;
        private String _numeroInventario;
        private String _eqdescription;
        private String _numeroSerie;
        private String _periodicidade;
        private List<Tuple<DateTime?, DateTime?, String>> _interventionsList; //Tuple<Scheduled Date, Performed Date, Decision>


        public Equipment()
        {
            _id_equipamento = string.Empty;
            _eqdescription = string.Empty;
            _interventionsList = new List<Tuple<DateTime?, DateTime?, string>>();
            _numeroInventario = string.Empty;
            _numeroSerie = string.Empty;
            _periodicidade = string.Empty;
        }

        public string ID_Equipamento
        {
            get { return _id_equipamento; }
            set { _id_equipamento = value; }
        }

        public string EqDescription
        {
            get { return _eqdescription; }
            set { _eqdescription = value; }
        }

        public List<Tuple<DateTime?, DateTime?, String>> InterventionsList
        {
            get { return _interventionsList; }
            set { _interventionsList = value; }
        }

        public string NumeroInventario
        {
            get { return _numeroInventario; }
            set { _numeroInventario = value; }
        }

        public string NumeroSerie
        {
            get { return _numeroSerie; }
            set { _numeroSerie = value; }
        }

        public string Periodicidade
        {
            get { return _periodicidade; }
            set { _periodicidade = value; }
        }
    }
}
