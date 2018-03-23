using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TShark;

namespace ITOne
{
    class FrmListaConfigMetas : _FrmListagens
    {
        /// <summary>
        /// form de listagem de funções
        /// </summary>
        /// <param name="addOn"></param>
        public FrmListaConfigMetas(Addon addOn, Dictionary<string, dynamic> ExtraParams = null): base(addOn, ExtraParams)
        {
            this.FormId = "FrmListaConfigMetas";

            // Form de popup
            this.ExtraParams["FORM_POPUP"] = "FrmConfigMetas";

            // Sobreescreve parametros:
            this.FormParams.Title = "Lista de Configurações de Metas e Comissões por Função";
            this.FormParams.Controls["hd01"].Label = "Configurações de Comissões e Metas por Função";
            this.FormParams.Controls["mtxLista"].Height = 370;

            // Exibe ou esconde pesquisa por data:
            this.FormParams.Controls["DtDe"].Visible = false;
            this.FormParams.Controls["DtAte"].Visible = false;
        }


        /// <summary>
        /// Matrix de listagem
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void mtxListaOnCreate(SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            this.InsereTodasFuncoes();

            // SQL da listagem
            this.ExtraParams["SQL_DT_LISTAGEM"] =
                " SELECT " +
                "    tb1.Code, tb1.U_nome, tb1.U_obs " +
                "  FROM [@UPD_IT_FUNCOES] tb1 " +
                "  WHERE 1 = 1 ";

            // SQL de pesquisa
            //this.ExtraParams["SQL_WHERE_DT_DE"] = "";
            //this.ExtraParams["SQL_WHERE_DT_ATE"] = "";
            this.ExtraParams["SQL_WHERE_FILTRO"] =
                " AND ( " +
                "   tb1.Code LIKE '%_TEXTO_%' " +
                "   OR tb1.U_nome LIKE '%_TEXTO_%' " +
                " )";

            // SQL de Order By
            this.ExtraParams["SQL_ORDER_DT_LISTAGEM"] = " tb1.U_nome ASC ";


            // Cria a matrix
            this.SetupMatrix(evObj.ItemUID, "DT_LISTAGEM", new List<ColumnDefinition>()
            {
                new ColumnDefinition() { Width = 3,     Id = "hash",    Caption = "#", Bind = false},
                new ColumnDefinition() { Percent = 30,  Id = "U_nome",  Caption = "Função"},
                new ColumnDefinition(){  Percent = 70,  Id = "U_obs",   Caption = "Observações" },

                // coluna invisível
                new ColumnDefinition() { Percent = 8,   Id = "Code",    Visible = false,},
            }, true, this.ParseSQL());

            // Abertura de popup em dblClick na matriz
            this.SetOnDblClick("mtxLista");
        }

        
        #region :: Regras de Negócio

        public void InsereTodasFuncoes()
        {
            string sql = "SELECT COUNT(*) as count FROM [@UPD_IT_FUNCOES]";
            SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);

            int count = rs.Fields.Item("count").Value;
            if (count > 0)
                return;

            // Todas as funções da ITOne
            List<Dictionary<string, dynamic>> funcoes =
                new List<Dictionary<string, dynamic>>()
                {
                    new Dictionary<string,dynamic>(){
                        {"U_nome","Arquiteto de Soluções"},
                        {"U_ativo","S"},
                    },
                    new Dictionary<string,dynamic>(){
                        {"U_nome","Assistente de Vendas"},
                        {"U_ativo","S"},
                    },
                    new Dictionary<string,dynamic>(){
                        {"U_nome","Inside Sales"},
                        {"U_ativo","S"},
                    },
                    new Dictionary<string,dynamic>(){
                        {"U_nome","Alocações Cross"},
                        {"U_ativo","S"},
                    },
                    new Dictionary<string,dynamic>(){
                        {"U_nome","Gerente de Contas"},
                        {"U_ativo","S"},
                    },
                };

            int i = 0;
            foreach (var funcao in funcoes)
            {
                i++;
                string new_code = "000" + i;
                this.DtSources.udoInsert("@UPD_IT_FUNCOES", funcao, out new_code);
            }
        }

        #endregion
    }
}
