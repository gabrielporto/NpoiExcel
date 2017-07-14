using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace NPOI
{
    class Ilidio : Excel
    {
        public DataTable ProcessarExecel(string filemane, string codigBusca)
        {
            Carregar(filemane);
            DataTable dt = GetDataTable();

            //DataTable codigos ira ter os valores dos codigos e suas linhas para serem somadas 
            DataTable soma = new DataTable();

            //Colocar colunas nos DT
            soma = addColumns(soma);

            //Verificar em todas as linhas do DT embusca de um codigo
            for (int rowDt = 0; rowDt < dt.Rows.Count; rowDt++)
            {
                //caso ache esse codigo colocar o valor no DT soma
                if (codigBusca == dt.Rows[rowDt][0].ToString())
                {
                    //adicionar valor ao DT soma
                    soma.Rows.Add(dt.Rows[rowDt][0], dt.Rows[rowDt][1], dt.Rows[rowDt][2], dt.Rows[rowDt][3], dt.Rows[rowDt][4], dt.Rows[rowDt][5],
                        dt.Rows[rowDt][6], dt.Rows[rowDt][7], dt.Rows[rowDt][8], dt.Rows[rowDt][9],
                        dt.Rows[rowDt][10], dt.Rows[rowDt][11], dt.Rows[rowDt][12], dt.Rows[rowDt][13], dt.Rows[rowDt][14]);
                }
            }

            //Add um linha ao DT soma para fazer a soma
            soma.Rows.Add(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

            //valores que nao mudam 
            soma.Rows[soma.Rows.Count - 1][0] = soma.Rows[0][0];
            soma.Rows[soma.Rows.Count - 1][1] = soma.Rows[0][1];

            double testeSoma1;
            double testeSoma2;
            double testeSoma3;


            //passar em todas as linha para fazer a soma
            for (int rowCodigos = 0; rowCodigos < soma.Rows.Count - 1; rowCodigos++)
            {
                for (int colCodigos = 2; colCodigos < soma.Columns.Count; colCodigos++)
                {
                    if (string.IsNullOrEmpty(soma.Rows[rowCodigos][colCodigos].ToString()))
                    {
                        soma.Rows[rowCodigos][colCodigos] = 0;
                    }

                    testeSoma1 = Convert.ToDouble(soma.Rows[rowCodigos][colCodigos]);
                    testeSoma2 = Convert.ToDouble(soma.Rows[soma.Rows.Count - 1][colCodigos]);
                    testeSoma3 = Convert.ToDouble(soma.Rows[rowCodigos][colCodigos]) + Convert.ToDouble(soma.Rows[soma.Rows.Count - 1][colCodigos]);

                    soma.Rows[soma.Rows.Count - 1][colCodigos] = Convert.ToDouble(soma.Rows[rowCodigos][colCodigos]) + Convert.ToDouble(soma.Rows[soma.Rows.Count - 1][colCodigos]);
                }
            }

            return soma;
        }

        //Colocar colunas no DT
        private DataTable addColumns(DataTable dt)
        {
            dt.Columns.Add("Cód. Red.");
            dt.Columns.Add("TOTAL DESPESAS CRYOBRÁS");
            dt.Columns.Add("Janeiro");
            dt.Columns.Add("Fevereiro");
            dt.Columns.Add("Março");
            dt.Columns.Add("Abril");
            dt.Columns.Add("Maio");
            dt.Columns.Add("Junho");
            dt.Columns.Add("Julho");
            dt.Columns.Add("Agosto");
            dt.Columns.Add("Setembro");
            dt.Columns.Add("Outubro");
            dt.Columns.Add("Novembro");
            dt.Columns.Add("Dezembro");
            dt.Columns.Add("2017");
            return dt;
        }


    }

}
