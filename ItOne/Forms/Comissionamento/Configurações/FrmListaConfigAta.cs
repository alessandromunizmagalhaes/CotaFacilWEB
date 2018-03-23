using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TShark;

namespace ITOne
{
    class FrmListaConfigAta : _FrmListagens
    {
        /// <summary>
        /// form de listagem de funções
        /// </summary>
        /// <param name="addOn"></param>
        public FrmListaConfigAta(Addon addOn, Dictionary<string, dynamic> ExtraParams = null): base(addOn, ExtraParams)
        {
            this.FormId = "FrmListaConfigAta";

            // Form de popup
            this.ExtraParams["FORM_POPUP"] = "FrmConfigAta";

            // Sobreescreve parametros:
            this.FormParams.Title = "Lista de Configurações de Replicação de Ata";
            this.FormParams.Controls["hd01"].Label = "Configurações de Replicação de Ata";
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

            // SQL da listagem
            this.ExtraParams["SQL_DT_LISTAGEM"] =
                " SELECT " +
                "    tb1.Code, tb1.U_desc, tb1.U_percent, tb1.U_obs " +
                "  FROM [@UPD_IT_CONFIG_ATA] tb1 " +
                "  WHERE 1 = 1 ";

            // SQL de pesquisa
            //this.ExtraParams["SQL_WHERE_DT_DE"] = "";
            //this.ExtraParams["SQL_WHERE_DT_ATE"] = "";
            this.ExtraParams["SQL_WHERE_FILTRO"] =
                " AND ( " +
                "   tb1.Code LIKE '%_TEXTO_%' " +
                "   OR tb1.U_desc LIKE '%_TEXTO_%' " +
                " )";

            // SQL de Order By
            this.ExtraParams["SQL_ORDER_DT_LISTAGEM"] = " tb1.U_desc ASC ";


            // Cria a matrix
            this.SetupMatrix(evObj.ItemUID, "DT_LISTAGEM", new List<ColumnDefinition>()
            {
                new ColumnDefinition() { Width = 3,     Id = "hash",        Caption = "#", Bind = false},
                new ColumnDefinition() { Percent = 30,  Id = "U_desc",      Caption = "Descrição"},
                new ColumnDefinition() { Percent = 10,  Id = "U_percent",   Caption = "Porcentagem"},
                new ColumnDefinition(){  Percent = 55,  Id = "U_obs",       Caption = "Observações" },

                // coluna invisível
                new ColumnDefinition() { Percent = 8,   Id = "Code",    Visible = false,},
            }, true, this.ParseSQL());

            // Abertura de popup em dblClick na matriz
            this.SetOnDblClick("mtxLista");
        }

        
        #region :: Regras de Negócio

        #endregion
    }
}
