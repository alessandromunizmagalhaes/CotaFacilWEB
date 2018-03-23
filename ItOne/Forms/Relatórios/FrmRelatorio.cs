using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TShark;
using System.IO;

namespace ITOne
{
    class FrmRelatorio : TShark.Forms
    {
        SAPbouiCOM.DataTable dt = null;
        SAPbouiCOM.DataTable dtTemp = null;
        SAPbouiCOM.Grid oGrid = null;
        
        public FrmRelatorio(Addon addOn, Dictionary<string, dynamic> ExtraParams = null): base(addOn, ExtraParams)
        {
            //Define o id do form como o nome da classe
            this.FormId = "FrmRelatorio";

            //Define as configurações do form
            this.FormParams = new FormParams()
            {
                Title = "Relatório de Comissão",

                //Definição de tamanho e posição do Form
                Bounds = new Bounds(){
                    Top = 25,
                    Left = 200,
                    Width = 1100,
                    Height = 520
                },

                #region :: Layout Componentes

                Linhas = new Dictionary<string,int>(){
                    {"space", 10},{"de", 8},{"ate", 8},{"search", 30},{"ver_por", 10},{"acelera", 12},{"tipo_rel", 10},{"btnPesq", 12},
                    {"hd01", 100},
                    {"grid", 100},
                    {"space1", 70},{"btnExpand", 15},{"btnColapse", 15},
                },

                Buttons = new Dictionary<string, int>(){
                    {"btnFechar", 20},{"space", 60},{"btnExport", 20},
                },

                #endregion


                #region :: Propriedades Componentes

                Controls = new Dictionary<string,CompDefinition>(){
                    
                    #region :: Cabeçalho
                
                    {"de", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "De",
                        UserDataType = SAPbouiCOM.BoDataType.dt_DATE,
                        onExitHandler = "OnChangePeriodo"
                    }},
                    {"ate", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Até",
                        UserDataType = SAPbouiCOM.BoDataType.dt_DATE,
                        onExitHandler = "OnChangePeriodo"
                    }},
                    {"search", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Pesquisa",
                    }},
                    {"ver_por", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                        Label = "Ver Por",
                        PopulateItens = new Dictionary<string,string>(){
                            {"1","Nº Lead"},
                            {"2","Comissionado"},
                            {"3","Função"},
                            {"4","Time"},
                        }
                    }},
                    {"acelera", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                        Label = "Tipo de Aceleração",
                        PopulateItens = new Dictionary<string,string>(){
                            {"1","Aceleração retroativa"},
                            {"3","Sem aceleração"},
                        }
                    }},
                    {"tipo_rel", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                        Label = "Tipo Relatório",
                        PopulateItens = new Dictionary<string,string>(){
                            {"1","Processado"},
                            {"2","Faturado"},
                            {"3","Realizado"},
                        }
                    }},
                    {"btnPesq", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "OK",
                        marginTop = 10,
                    }},

                    #endregion
                    

                    #region :: Grid

                    {"hd01", new CompDefinition(){
                        Label = "Relatório de Comissão",
                        Type = SAPbouiCOM.BoFormItemTypes.it_RECTANGLE,
                        Height = 1
                    }},
                    {"grid", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_GRID,
                        Height = 370
                    }},
                    {"btnExpand", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Expandir Todos"
                    }},
                    {"btnColapse", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Fechar Todos"
                    }},

                    #endregion
                    
                    
                    #region :: Botões Padrões

                    {"btnFechar", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Fechar"
                    }},
                    {"btnExport", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Exportar para Excel"
                    }},
                    
                    #endregion

                },

                #endregion

            };
        }


        #region :: Eventos do Formulário

        /// <summary>
        /// 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void FrmRelatorioOnFormOpen(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            DateTime dt_de  = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime dt_ate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.DaysInMonth(DateTime.Now.Year,DateTime.Now.Month));

            this.UpdateUserDataSource(new Dictionary<string, dynamic>()
            {
                {"de",dt_de.ToString("yyyyMMdd")},
                {"ate",dt_ate.ToString("yyyyMMdd")},
                {"ver_por","1"},
                {"tipo_rel","1"},
                {"acelera","1"},
            });
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void FrmRelatorioOnFormClose(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        #endregion


        #region :: Eventos dos Componentes

        /// <summary>
        /// Insere um ítem.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnPesqOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            this.Pesquisar();
        }

        /// <summary>
        /// Insere um ítem.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnExportOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            this.Addon.StatusInfo("Exportando para Excel...");
            this.SapForm.Freeze(true);
            try
            {
                string filename = "rel_comissao.csv";
                if (((Addon)this.Addon).ExportarCSV(filename, this.dtTemp, this.oGrid))
                {
                    string path = Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.System)) + "\\tmp\\" + filename;
                    this.Addon.StatusInfo("Arquivo exportado com sucesso.\nO arquivo se encontra em ' " + path + "'", true);
                }
            }
            catch (Exception ex)
            {
                this.Addon.StatusErro("Erro ao exportar CSV. " + ex.Message, true);
            }
            finally
            {
                this.SapForm.Freeze(false);
            }
        }

        /// <summary>
        /// Insere um ítem.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnExpandOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            this.ExpandirGrid();
        }

        /// <summary>
        /// Insere um ítem.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnColapseOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            this.FecharGrid();
        }

        /// <summary>
        /// Insere um ítem.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnChangePeriodo(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            this.FiltroPossuiPeriodoValido();
        }

        #endregion


        #region :: Grid

        /// <summary>
        /// Matriz de Comunicação
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void gridOnCreate(SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            
            this.dt     = this.SapForm.DataSources.DataTables.Add("RELATORIO");
            this.dtTemp = this.SapForm.DataSources.DataTables.Add("TEMP");
            this.oGrid  = ((SAPbouiCOM.Grid)this.GetItem(evObj.ItemUID).Specific);
            
            this.oGrid.DataTable     = this.dt;
            this.GetItem(evObj.ItemUID).Enabled = false;
        }

        #endregion


        #region :: Regras de Negócio

        /// <summary>
        /// Efetua a pesquisa de acordo com os parametros de filtragem.
        /// </summary>
        public void Pesquisar()
        {
            if (!this.FiltroPossuiPeriodoValido(true))
                return;
            
            this.SapForm.Freeze(true);

            try
            {
                string select = "",where = "", select_union = "", where_pesq = "";
                string aceleracao = this.GetValue("acelera");

                #region :: SQL Processado

                string join_processado = " LEFT JOIN OINV tb6 (NOLOCK) ON (tb1.U_UPD_IT_LEAD = tb6.U_UPD_IT_LEAD AND tb6.DocStatus != 'C' AND tb6.CANCELED = 'N' ) ";
                string where_processado = " AND tb1.U_UPD_IT_STATUS = 'P' AND tb1.DocStatus = 'O' ";

                #endregion


                #region :: SQL Faturado

                string join_faturado = " LEFT JOIN OINV tb6 (NOLOCK) ON (tb1.U_UPD_IT_LEAD = tb6.U_UPD_IT_LEAD AND tb6.DocStatus != 'C' AND tb6.CANCELED = 'N') ";
                string where_faturado = " AND tb1.U_UPD_IT_STATUS = 'F' ";

                #endregion  


                #region :: SQL Realizado

                string join_realizado   = " LEFT JOIN OINV tb6 (NOLOCK) ON (tb1.U_UPD_IT_LEAD = tb6.U_UPD_IT_LEAD AND tb6.DocStatus != 'C' AND tb6.CANCELED = 'N' ) ";
                string where_realizado = " AND tb1.U_UPD_IT_STATUS = 'R' ";

                #endregion


                #region :: SELECT Por Lead

                string select_por_lead =
                    "   SELECT " +
                    "       outro.[Número Lead], outro.Comissionado " +
                    "       ,FORMAT(CASE WHEN 3 = " + aceleracao + " THEN outro.perc_fixo WHEN outro.U_ata > 0 THEN 0.0 WHEN outro.U_percom > 0 THEN 0.0 ELSE tb9.perc_comissao END,'N','pt-br') + '%' AS 'Comissão %' " +
                    "       ,FORMAT(outro.[Renda Comissionável],'C','pt-br') as 'Renda Comissionável' " +
                    "       ,FORMAT(CASE WHEN 3 = " + aceleracao + " THEN outro.[Renda Comissionável] * (outro.perc_fixo/100) WHEN outro.U_ata > 0 THEN 0.0 WHEN outro.U_percom > 0 THEN 0.0 WHEN tb9.perc_comissao > 0 THEN outro.[Renda Comissionável] * (tb9.perc_comissao/100) ELSE 0 END,'C','pt-br') AS 'Comissão' " +
                    "       ,FORMAT(0.0,'C','pt-br') AS 'Ajustes de Ata' " +
                    "       ,FORMAT(0.0,'C','pt-br') AS 'SPLIT' " +
                    "       ,FORMAT(CASE WHEN 3 = " + aceleracao + " THEN outro.[Renda Comissionável] * (outro.perc_fixo/100) WHEN outro.U_ata > 0 THEN 0.0 WHEN outro.U_percom > 0 THEN 0.0 WHEN tb9.perc_comissao > 0 THEN outro.[Renda Comissionável] * (tb9.perc_comissao/100) ELSE 0 END,'C','pt-br') AS 'Comissão Final' " +
                    "       , FORMAT(outro.[Valor Pedido],'C','pt-br') as 'Valor Pedido' " +
                    "       ,outro.[Pedido de Venda], outro.[Nota Fiscal de Saída] " +
                    "       ,outro.Operação, outro.Função, outro.Time " +
                    
                    // campos invisiveis
                    "       ,outro.U_ata, outro.U_gerente, outro.U_percom, outro.U_split, outro.Empresa ";

                string select_por_lead_union =
                    "SELECT " +
                    "   tb1.U_UPD_IT_LEAD as 'Número Lead' " +
                    "   ,'PARCEIRO - ' + tb3.CardName as 'Comissionado' " +
                    "   , FORMAT( tb1.U_UPD_IT_PERCENT, 'N', 'pt-br' )  + '%' as 'Comissão %' " +
                    "   , FORMAT( tb1.U_UPD_IT_RENDA, 'C', 'pt-br' ) as 'Renda Comissionável' " +
                    "   , FORMAT( ((tb1.U_UPD_IT_RENDA * tb1.U_UPD_IT_PERCENT) / 100), 'C', 'pt-br' ) as 'Comissão' " +
                    "   , FORMAT(0.0,'C','pt-br') AS 'Ajustes de Ata' " +
                    "   , FORMAT(0.0,'C','pt-br') AS 'SPLIT' " +
                    "   , FORMAT( ((tb1.U_UPD_IT_RENDA * tb1.U_UPD_IT_PERCENT) / 100), 'C', 'pt-br' ) as 'Comissão Final' " +
                    "   , FORMAT(tb1.DocTotal,'C','pt-br') AS 'Valor Pedido' " +
                    "   , tb1.DocNum AS 'Pedido de Venda' " +
                    "   , CASE WHEN tb6.Serial IS NULL THEN tb1.U_UPD_IT_SERIAL ELSE tb6.Serial END AS 'Nota Fiscal de Saída' " +
                    "   , ' ' as 'Operação' " +
                    "   , 'Parceiro' as 'Função' " +
                    "   , tb10.firstName + ' ' + tb10.lastName AS 'Time' " +
                    "   , 0.0 as 'U_ata' " +
                    "   , ' ' as 'U_gerente' " +
                    "   , 0.0 as 'U_percom' " +
                    "   , ' ' as 'U_split' " +
                    "   , ' ' as 'Empresa' ";

                #endregion


                #region :: SELECT por Comissionado

                string select_por_comissionado =
                    "   SELECT " +
                    "       outro.Comissionado, outro.[Número Lead] " +
                    "       ,FORMAT(CASE WHEN 3 = " + aceleracao + " THEN outro.perc_fixo WHEN outro.U_ata > 0 THEN 0.0 WHEN outro.U_percom > 0 THEN 0.0 ELSE tb9.perc_comissao END,'N','pt-br') + '%' AS 'Comissão %' " +
                    "       ,FORMAT(outro.[Renda Comissionável],'C','pt-br') as 'Renda Comissionável' " +
                    "       ,FORMAT(CASE WHEN 3 = " + aceleracao + " THEN outro.[Renda Comissionável] * (outro.perc_fixo/100) WHEN outro.U_ata > 0 THEN 0.0 WHEN outro.U_percom > 0 THEN 0.0 WHEN tb9.perc_comissao > 0 THEN outro.[Renda Comissionável] * (tb9.perc_comissao/100) ELSE 0 END,'C','pt-br') AS 'Comissão' " +
                    "       ,FORMAT(0.0,'C','pt-br') AS 'Ajustes de Ata' " +
                    "       ,FORMAT(0.0,'C','pt-br') AS 'SPLIT' " +
                    "       ,FORMAT(CASE WHEN 3 = " + aceleracao + " THEN outro.[Renda Comissionável] * (outro.perc_fixo/100) WHEN outro.U_ata > 0 THEN 0.0 WHEN outro.U_percom > 0 THEN 0.0 WHEN tb9.perc_comissao > 0 THEN outro.[Renda Comissionável] * (tb9.perc_comissao/100) ELSE 0 END,'C','pt-br') AS 'Comissão Final' " +
                    "       , FORMAT(outro.[Valor Pedido],'C','pt-br') as 'Valor Pedido' " +
                    "       ,outro.[Pedido de Venda], outro.[Nota Fiscal de Saída] " +
                    "       ,outro.Operação, outro.Função, outro.Time " +

                    // campos invisiveis
                    "       ,outro.U_ata, outro.U_gerente, outro.U_percom, outro.U_split, outro.Empresa ";

                string select_por_comissionado_union =
                    "SELECT " +
                    "   'PARCEIRO - ' + tb3.CardName as 'Comissionado' " +
                    "   ,tb1.U_UPD_IT_LEAD as 'Número Lead' " +
                    "   , FORMAT( tb1.U_UPD_IT_PERCENT, 'N', 'pt-br' )  + '%' as 'Comissão %' " +
                    "   , FORMAT( tb1.U_UPD_IT_RENDA, 'C', 'pt-br' ) as 'Renda Comissionável' " +
                    "   , FORMAT( ((tb1.U_UPD_IT_RENDA * tb1.U_UPD_IT_PERCENT) / 100), 'C', 'pt-br' ) as 'Comissão' " +
                    "   , FORMAT(0.0,'C','pt-br') AS 'Ajustes de Ata' " +
                    "   , FORMAT(0.0,'C','pt-br') AS 'SPLIT' " +
                    "   , FORMAT( ((tb1.U_UPD_IT_RENDA * tb1.U_UPD_IT_PERCENT) / 100), 'C', 'pt-br' ) as 'Comissão Final' " +
                    "   , FORMAT(tb1.DocTotal,'C','pt-br') AS 'Valor Pedido' " +
                    "   , tb1.DocNum AS 'Pedido de Venda' " +
                    "   , CASE WHEN tb6.Serial IS NULL THEN tb1.U_UPD_IT_SERIAL ELSE tb6.Serial END AS 'Nota Fiscal de Saída' " +
                    "   , ' ' as 'Operação' " +
                    "   , 'Parceiro' as 'Função' " +
                    "   , tb10.firstName + ' ' + tb10.lastName AS 'Time' " +
                    "   , 0.0 as 'U_ata' " +
                    "   , ' ' as 'U_gerente' " +
                    "   , 0.0 as 'U_percom' " +
                    "   , ' ' as 'U_split' " +
                    "   , ' ' as 'Empresa' ";

                #endregion


                #region :: SELECT por Função
                
                string select_por_funcao =
                    "   SELECT " +
                    "       outro.Função, outro.Comissionado, outro.[Número Lead] " +
                    "       ,FORMAT(CASE WHEN 3 = " + aceleracao + " THEN outro.perc_fixo WHEN outro.U_ata > 0 THEN 0.0 WHEN outro.U_percom > 0 THEN 0.0 ELSE tb9.perc_comissao END,'N','pt-br') + '%' AS 'Comissão %' " +
                    "       ,FORMAT(outro.[Renda Comissionável],'C','pt-br') as 'Renda Comissionável' " +
                    "       ,FORMAT(CASE WHEN 3 = " + aceleracao + " THEN outro.[Renda Comissionável] * (outro.perc_fixo/100) WHEN outro.U_ata > 0 THEN 0.0 WHEN outro.U_percom > 0 THEN 0.0 WHEN tb9.perc_comissao > 0 THEN outro.[Renda Comissionável] * (tb9.perc_comissao/100) ELSE 0 END,'C','pt-br') AS 'Comissão' " +
                    "       ,FORMAT(0.0,'C','pt-br') AS 'Ajustes de Ata' " +
                    "       ,FORMAT(0.0,'C','pt-br') AS 'SPLIT' " +
                    "       ,FORMAT(CASE WHEN 3 = " + aceleracao + " THEN outro.[Renda Comissionável] * (outro.perc_fixo/100) WHEN outro.U_ata > 0 THEN 0.0 WHEN outro.U_percom > 0 THEN 0.0 WHEN tb9.perc_comissao > 0 THEN outro.[Renda Comissionável] * (tb9.perc_comissao/100) ELSE 0 END,'C','pt-br') AS 'Comissão Final' " +
                    "       , FORMAT(outro.[Valor Pedido],'C','pt-br') as 'Valor Pedido' " +
                    "       ,outro.[Pedido de Venda], outro.[Nota Fiscal de Saída] " +
                    "       ,outro.Operação,  outro.Time " +

                    // campos invisiveis
                    "       ,outro.U_ata, outro.U_gerente, outro.U_percom, outro.U_split, outro.Empresa ";

                string select_por_funcao_union =
                    "SELECT " +
                    "   'Parceiro' as 'Função' " +
                    "   ,'PARCEIRO - ' + tb3.CardName as 'Comissionado' " +
                    "   , tb1.U_UPD_IT_LEAD as 'Número Lead' " +
                    "   , FORMAT( tb1.U_UPD_IT_PERCENT, 'N', 'pt-br' )  + '%' as 'Comissão %' " +
                    "   , FORMAT( tb1.U_UPD_IT_RENDA, 'C', 'pt-br' ) as 'Renda Comissionável' " +
                    "   , FORMAT( ((tb1.U_UPD_IT_RENDA * tb1.U_UPD_IT_PERCENT) / 100), 'C', 'pt-br' ) as 'Comissão' " +
                    "   , FORMAT(0.0,'C','pt-br') AS 'Ajustes de Ata' " +
                    "   , FORMAT(0.0,'C','pt-br') AS 'SPLIT' " +
                    "   , FORMAT( ((tb1.U_UPD_IT_RENDA * tb1.U_UPD_IT_PERCENT) / 100), 'C', 'pt-br' ) as 'Comissão Final' " +
                    "   , FORMAT(tb1.DocTotal,'C','pt-br') AS 'Valor Pedido' " +
                    "   , tb1.DocNum AS 'Pedido de Venda' " +
                    "   , CASE WHEN tb6.Serial IS NULL THEN tb1.U_UPD_IT_SERIAL ELSE tb6.Serial END AS 'Nota Fiscal de Saída' " +
                    "   , ' ' as 'Operação' " +
                    "   , tb10.firstName + ' ' + tb10.lastName AS 'Time' " +
                    "   , 0.0 as 'U_ata' " +
                    "   , ' ' as 'U_gerente' " +
                    "   , 0.0 as 'U_percom' " +
                    "   , ' ' as 'U_split' " + 
                    "   , ' ' as 'Empresa'";

                #endregion


                #region :: SELECT por Time

                string select_por_time =
                    "   SELECT " +
                    "       outro.Time, outro.Comissionado, outro.[Número Lead] " +
                    "       ,FORMAT(CASE WHEN 3 = " + aceleracao + " THEN outro.perc_fixo WHEN outro.U_ata > 0 THEN 0.0 WHEN outro.U_percom > 0 THEN 0.0 ELSE tb9.perc_comissao END,'N','pt-br') + '%' AS 'Comissão %' " +
                    "       ,FORMAT(outro.[Renda Comissionável],'C','pt-br') as 'Renda Comissionável' " +
                    "       ,FORMAT(CASE WHEN 3 = " + aceleracao + " THEN outro.perc_fixo WHEN outro.U_ata > 0 THEN 0.0 WHEN outro.U_percom > 0 THEN 0.0 WHEN tb9.perc_comissao > 0 THEN outro.[Renda Comissionável] * (tb9.perc_comissao/100) ELSE 0 END,'C','pt-br') AS 'Comissão' " +
                    "       ,FORMAT(0.0,'C','pt-br') AS 'Ajustes de Ata' " +
                    "       ,FORMAT(0.0,'C','pt-br') AS 'SPLIT' " +
                    "       ,FORMAT(CASE WHEN 3 = " + aceleracao + " THEN outro.perc_fixo WHEN outro.U_ata > 0 THEN 0.0 WHEN outro.U_percom > 0 THEN 0.0 WHEN tb9.perc_comissao > 0 THEN outro.[Renda Comissionável] * (tb9.perc_comissao/100) ELSE 0 END,'C','pt-br') AS 'Comissão Final' " +
                    "       , FORMAT(outro.[Valor Pedido],'C','pt-br') as 'Valor Pedido' " +
                    "       ,outro.[Pedido de Venda], outro.[Nota Fiscal de Saída] " +
                    "       ,outro.Operação, outro.Função " +

                    // campos invisiveis
                    "       ,outro.U_ata, outro.U_gerente, outro.U_percom, outro.U_split, outro.Empresa ";

                string select_por_time_union =
                    "SELECT " +
                    "   tb10.firstName + ' ' + tb10.lastName AS 'Time' " +
                    "   , 'PARCEIRO - ' + tb3.CardName as 'Comissionado' " +
                    "   , tb1.U_UPD_IT_LEAD as 'Número Lead' " +
                    "   , FORMAT( tb1.U_UPD_IT_PERCENT, 'N', 'pt-br' )  + '%' as 'Comissão %' " +
                    "   , FORMAT( tb1.U_UPD_IT_RENDA, 'C', 'pt-br' ) as 'Renda Comissionável' " +
                    "   , FORMAT( ((tb1.U_UPD_IT_RENDA * tb1.U_UPD_IT_PERCENT) / 100), 'C', 'pt-br' ) as 'Comissão' " +
                    "   , FORMAT(0.0,'C','pt-br') AS 'Ajustes de Ata' " +
                    "   , FORMAT(0.0,'C','pt-br') AS 'SPLIT' " +
                    "   , FORMAT( ((tb1.U_UPD_IT_RENDA * tb1.U_UPD_IT_PERCENT) / 100), 'C', 'pt-br' ) as 'Comissão Final' " +
                    "   , FORMAT(tb1.DocTotal,'C','pt-br') AS 'Valor Pedido' " +
                    "   , tb1.DocNum AS 'Pedido de Venda' " +
                    "   , CASE WHEN tb6.Serial IS NULL THEN tb1.U_UPD_IT_SERIAL ELSE tb6.Serial END AS 'Nota Fiscal de Saída' " +
                    "   , ' ' as 'Operação' " +
                    "   , 'Parceiro' as 'Função' " +
                    "   , 0.0 as 'U_ata' " +
                    "   , ' ' as 'U_gerente' " +
                    "   , 0.0 as 'U_percom' " +
                    "   , ' ' as 'U_split' " +
                    "   , ' ' as 'Empresa'";

                #endregion


                // Definindo filtros
                string de       = this.GetValue("de");
                string ate      = this.GetValue("ate");
                string ver_por  = this.GetValue("ver_por");
                string tipo_rel = this.GetValue("tipo_rel");
                string pesquisa = this.GetValue("search");

                DateTime data_de = this.Addon.ToDatetime(de);
                de = data_de.ToString("yyy-MM-dd");
                where += " AND CONVERT(DATE,tb1.DocDate,103) >= '" + de + "'";

                DateTime data_ate = this.Addon.ToDatetime(ate);
                ate = data_ate.ToString("yyy-MM-dd");
                where += " AND CONVERT(DATE,tb1.DocDate,103) <= '" + ate + "'";

                int ano = data_de.Year;

                if (!String.IsNullOrEmpty(pesquisa))
                    where_pesq += this.MontaWherePesquisa(pesquisa);

                if (ver_por == "1")
                {
                    select = select_por_lead;
                    select_union = select_por_lead_union;
                }
                else if (ver_por == "2")
                {
                    select = select_por_comissionado;
                    select_union = select_por_comissionado_union;
                }
                else if (ver_por == "3")
                {
                    select = select_por_funcao;
                    select_union = select_por_funcao_union;
                }
                else if (ver_por == "4")
                {
                    select = select_por_time;
                    select_union = select_por_time_union;
                }
                else
                {
                    select = select_por_lead;
                    select_union = select_por_lead_union;
                }



                #region :: SQL WITH 1

                string sql_with1 =
                "WITH base as ( " +
                "   SELECT  " +
                "       tb1.U_UPD_IT_LEAD AS 'Número Lead' " +
                "       , tb5.firstName + ' ' + tb5.lastName AS 'Comissionado' " +
                "       , CASE " +
			    "           WHEN tb1.U_UPD_IT_PERCENT > 0 " +
                "               THEN ((tb1.U_UPD_IT_RENDA * ( 100 - tb1.U_UPD_IT_PERCENT ) ) / 100) " +
			    "               ELSE tb1.U_UPD_IT_RENDA  " +
                "       END AS 'Renda Comissionável' " +
                "       , tb1.DocTotal AS 'Valor Pedido' " +
                "       , tb1.DocNum AS 'Pedido de Venda' " +
                "       , CASE WHEN tb6.Serial IS NULL THEN tb1.U_UPD_IT_SERIAL ELSE tb6.Serial END AS 'Nota Fiscal de Saída' " +
                "       , tb2.U_upd_3_operacao AS 'Operação' " +
                "       , tb4.U_nome AS 'Função' " +
                "       , tb4.U_comissao AS 'perc_fixo' " +
                "       , tb4.Code as 'code_funcao' " +
                "       , tb3.U_percom , tb3.U_ata, tb3.U_gerente, tb3.U_split, tb11.CardName as 'Empresa' " +
                "       , tb10.firstName + ' ' + tb10.lastName AS 'Time' " +
                "       , tb7.U_meta as 'Meta' " +
                "       , tb8.vendas " +
                "   FROM ORDR tb1 (NOLOCK) " +
                "   INNER JOIN OPR2 tb2 (NOLOCK) ON (tb1.U_UPD_IT_LEAD = tb2.U_upd_1_nlead) " +
                "   INNER JOIN [@UPD_IT_PARTICIP] tb3 (NOLOCK) ON (tb1.DocNum = tb3.U_docentry) " +
                "   INNER JOIN [@UPD_IT_FUNCOES] tb4  (NOLOCK) ON (tb3.U_funcao = tb4.Code) " +
                "   INNER JOIN OHEM tb5 (NOLOCK) ON (tb3.U_empid = tb5.empID) ";

                sql_with1 += tipo_rel == "1" ? join_processado : (tipo_rel == "2" ? join_faturado : join_realizado);
                sql_with1 +=
                "   LEFT JOIN " +
                "   ( " +
                "       SELECT " +
		        "           Code, U_acelera, U_empid,SUM(U_meta) as U_meta " +
	            "       FROM [@UPD_IT_METAS]  " +
	            "       WHERE YEAR(U_dtinicio) = " + ano + " AND YEAR(U_dtfim) = " + ano + " " +
                "       GROUP BY Code, U_acelera, U_empid " +
                "   ) tb7 ON (tb7.Code = tb4.Code AND tb3.U_empid = tb7.U_empid AND YEAR(CONVERT(DATE, tb1.DocDate,103)) = " + ano + " ) " +
                "   LEFT JOIN " +
                "   ( " +
                "           SELECT " +
                "               tbParticip.U_empid, CONVERT(DATE, tbORDR.DocDate, 103) AS DocDate, DocTotal AS vendas " +
                "           FROM [@UPD_IT_PARTICIP] (NOLOCK) tbParticip " +
                "           INNER JOIN ORDR (NOLOCK) tbORDR ON (tbParticip.U_docentry = tbORDR.DocNum) " +
                "   ) tb8 ON (tb8.U_empid = tb5.empid AND YEAR(CONVERT(DATE,tb8.DocDate,103)) = " + ano + " ) " +
                "   LEFT JOIN OHEM tb10 (NOLOCK) ON (tb2.U_UPD_IT_TIME = tb10.empID) " +
                "   INNER JOIN OCRD tb11 (NOLOCK) ON (tb11.CardCode = tb1.CardCode ) " +
                
                //só traz Sem aceleração/Aceleração Retroativa
                "   WHERE tb2.U_upd_2_status = 'WON' AND tb7.U_acelera <> 2 ";
                sql_with1 += where + where_pesq;
                sql_with1 += tipo_rel == "1" ? where_processado : (tipo_rel == "2" ? where_faturado : where_realizado);

                #endregion


                #region :: SQL WITH 2

                string sql_with2 =
                "), " +
                "outro as ( " +
	            "   SELECT " +
		        "       base.[Número Lead], base.Comissionado, base.[Renda Comissionável], base.[Valor Pedido], base.[Pedido de Venda], " +
                "       base.[Nota Fiscal de Saída],base.Operação, base.Função, base.Time, " +
                "       base.code_funcao, base.Meta,base.U_percom,base.U_ata,base.U_gerente,base.U_split,base.perc_fixo, base.Empresa, " +
		        "       SUM(base.vendas) as total_vendas " +
                "   FROM  base " +
                "   GROUP BY    base.[Número Lead], base.Comissionado, base.[Renda Comissionável], base.[Valor Pedido], base.[Pedido de Venda], " +
                "               base.[Nota Fiscal de Saída],base.Operação, base.Função, base.Time, " +
                "               base.code_funcao, base.Meta,base.U_percom,base.U_ata,base.U_gerente, base.U_split, base.perc_fixo, base.Empresa " +
                ")";

                #endregion


                #region :: SQL Base

                #region :: estava assim, cheguei de férias, vi o bug e corrigi.
                //AND outro.Meta IS NOT NULL 
                //a função exercida existe, mas o comissionado não está configurado, logo nao tem meta. se não tem meta não pode calcular.
                //o else retornava 0 quando null, e aí trazia sempre a primeira comissão da função.
                /*string sql_base = 
                    select +
                    "   FROM  outro"  +
                    "   LEFT JOIN " +
                    "   ( " +
                    "       SELECT Code, U_comissao AS perc_comissao, U_piso, U_teto " +
                    "           FROM [@UPD_IT_COMISSAO] (NOLOCK) " +
                    "   ) tb9 ON  " +
                    "   ( " +
                    "       tb9.Code = outro.code_funcao AND ((CASE WHEN outro.Meta > 0 THEN ((outro.total_vendas * 100.00)/outro.Meta) ELSE 0 END) BETWEEN U_piso AND U_teto) " +
                    "   )";*/
                #endregion

                string sql_base =
                    select +
                    "   FROM  outro" +
                    "   LEFT JOIN " +
                    "   ( " +
                    "       SELECT Code, U_comissao AS perc_comissao, U_piso, U_teto " +
                    "           FROM [@UPD_IT_COMISSAO] (NOLOCK) " +
                    "   ) tb9 ON  " +
                    "   ( " +
                    "       tb9.Code = outro.code_funcao AND outro.Meta IS NOT NULL AND ((CASE WHEN outro.Meta > 0 THEN ((outro.total_vendas * 100.00)/outro.Meta) ELSE 0 END) BETWEEN U_piso AND U_teto) " +
                    "   )";

                string order_by =
                        " ORDER BY outro.[Número Lead] ASC ";

                #endregion


                #region :: SQL BASE UNION

                string 
                    sql_union =
                    " UNION ALL " +
                    select_union +
                    "FROM ORDR tb1 (NOLOCK) " +
                    "INNER JOIN OPR2 tb2 (NOLOCK) ON (tb1.U_UPD_IT_LEAD = tb2.U_upd_1_nlead) " +
                    "INNER JOIN OCRD tb3 (NOLOCK) ON ( tb1.U_UPD_IT_PARCEIRO = tb3.CardCode ) ";
                    sql_union += tipo_rel == "1" ? join_processado : (tipo_rel == "2" ? join_faturado : join_realizado);
                    sql_union +=
                    "LEFT JOIN OHEM tb10 (NOLOCK) ON (tb2.U_UPD_IT_TIME = tb10.empID) " +
                    "WHERE tb2.U_upd_2_status = 'WON' ";
                    sql_union += where;
                    sql_union += tipo_rel == "1" ? where_processado : (tipo_rel == "2" ? where_faturado : where_realizado);

                #endregion


                string sql = sql_with1 + sql_with2 + sql_base + sql_union + order_by;

                this.dt.ExecuteQuery(sql);

                ((Addon)this.Addon).SalvarSQLArquivoTXT(sql);


                #region :: Ajustes de Colunas

                // Configurações das Colunas.
                this.oGrid.Columns.Item("Comissão %").RightJustified            = true;
                this.oGrid.Columns.Item("Comissão").RightJustified              = true;
                this.oGrid.Columns.Item("Renda Comissionável").RightJustified   = true;
                this.oGrid.Columns.Item("Valor Pedido").RightJustified          = true;
                this.oGrid.Columns.Item("Ajustes de Ata").RightJustified        = true;
                this.oGrid.Columns.Item("SPLIT").RightJustified                 = true;
                this.oGrid.Columns.Item("Comissão Final").RightJustified        = true;
                
                this.oGrid.Columns.Item("U_ata").Visible        = false;
                this.oGrid.Columns.Item("U_gerente").Visible    = false;
                this.oGrid.Columns.Item("U_percom").Visible     = false;
                this.oGrid.Columns.Item("U_split").Visible      = false;

                ((SAPbouiCOM.EditTextColumn)this.oGrid.Columns.Item("Pedido de Venda")).LinkedObjectType = "17";

                // Se for processado, não tem que aparecer a coluna Nota Fiscal de Saída
                if( tipo_rel == "1" )
                {
                    this.oGrid.Columns.Item("Nota Fiscal de Saída").Visible = false;
                }

                this.oGrid.CommonSetting.FixedColumnsCount = 2;
                this.oGrid.CollapseLevel = 1;

                #endregion


                #region :: Pós Processamento

                Dictionary<dynamic, Dictionary<string, double>> soma_lead = new Dictionary<dynamic, Dictionary<string, double>>() { };

                string campo_base = ver_por == "1" ? "Número Lead" : (ver_por == "2" ? "Comissionado" : (ver_por == "3" ? "Função" : (ver_por == "4" ? "Time" : "Número Lead")));

                
                // utilizados para replicação de ata
                Dictionary<int, Dictionary<string, double>> grupos = new Dictionary<int, Dictionary<string, double>>() { };
                Dictionary<int, List<Dictionary<string, double>>> atas = new Dictionary<int,List<Dictionary<string,double>>>(){};
                
                // utilizados para split
                Dictionary<int, Dictionary<string, double>> split_ref = new Dictionary<int, Dictionary<string, double>>() { };
                Dictionary<int, List<Dictionary<string, double>>> splits = new Dictionary<int, List<Dictionary<string, double>>>() { };

                #region :: PRIMEIRO LOOPING PARA ORGANIZAR ATA/SPLIT

                for (int i = 0; i < this.dt.Rows.Count; i++ )
                {
                    /* BUSCANDO VALORES BASE PARA CÁLCULOS */
                    double comissao = 0, comissao_perc = 0;

                    string str_comissao         = this.dt.GetValue("Comissão", i);
                    string str_comissao_perc    = this.dt.GetValue("Comissão %", i);

                    dynamic val_base = this.dt.GetValue(campo_base, i);
                    Double.TryParse(str_comissao.Replace("R$",""), out comissao);
                    Double.TryParse(str_comissao_perc.Replace("%", ""), out comissao_perc);

                    /* ORGANIZAÇÃO DA REPLICAÇÃO DE ATA */
                    int docnum      = this.dt.GetValue("Pedido de Venda", i);
                    double ata      = this.dt.GetValue("U_ata", i);
                    string gerente  = this.dt.GetValue("U_gerente", i);

                    if(!grupos.ContainsKey(docnum))
                    {
                        grupos.Add(docnum,new Dictionary<string,double>(){
                        });

                        atas.Add(docnum, new List<Dictionary<string, double>>() { });
                    }

                    if( gerente == "S" )
                    {
                        grupos[docnum]["ndx"] = i;
                        grupos[docnum]["vlr"] = comissao;
                    }
                    else if( ata > 0 )
                    {
                        atas[docnum].Add(new Dictionary<string, double>()
                        {
                            {"ndx",i},
                            {"vlr",ata}
                        });
                    }


                    /* ORGANIZAÇÃO DE SPLIT */

                    double percom = this.dt.GetValue("U_percom", i);
                    string split = this.dt.GetValue("U_split", i);

                    if (!split_ref.ContainsKey(docnum))
                    {
                        split_ref.Add(docnum, new Dictionary<string, double>()
                        {
                        });

                        splits.Add(docnum, new List<Dictionary<string, double>>() { });
                    }

                    if (split == "S")
                    {
                        split_ref[docnum]["ndx"] = i;
                        split_ref[docnum]["vlr"] = comissao;
                    }
                    else if (percom > 0)
                    {
                        splits[docnum].Add(new Dictionary<string, double>()
                        {
                            {"ndx",i},
                            {"vlr",percom}
                        });
                    }
                }

                #endregion

                #region :: LOOPING PARA ATUALIZAR VALORES DE ATA

                /* ATUALIZANDO NO DATATABLE A REPLICAÇÃO DE ATA */
                foreach (var ata in atas)
                {
                    // se não tiver valor no dicionário de atas, então passa pra frente
                    if (ata.Value.Count == 0)
                    {
                        continue;
                    }

                    double soma_debito_ata = 0.0;
                    double comissao_base = 0;
                    int r = Convert.ToInt32(grupos[ata.Key]["ndx"]);
                    string str_com_final = this.dt.GetValue("Comissão Final", r);
                    Double.TryParse(str_com_final.Replace("R$",""), out comissao_base);

                    foreach (var valores in ata.Value)
                    {
                        // inicializando valores e calculando
                        double comissao_final = 0;
                        double ajuste_ata = comissao_base * valores["vlr"] / 100;
                        int row = Convert.ToInt32(valores["ndx"]);
                        string str_com_final2 = this.dt.GetValue("Comissão Final", row);
                        Double.TryParse(str_com_final2.Replace("R$",""), out comissao_final);
                        double comissao_final_ajustado = comissao_final + ajuste_ata;

                        // atualizando colunas no dt
                        this.dt.SetValue("Ajustes de Ata", row, "+ R$ " + string.Format("{0:#,0.00}", ajuste_ata));
                        this.dt.SetValue("Comissão Final", row, "R$ " + string.Format("{0:#,0.00}", comissao_final_ajustado));

                        // incrementando tudo que eu já debitei deste grupo
                        soma_debito_ata += ajuste_ata;
                    }

                    double comissao_final_ajustado2 = comissao_base - soma_debito_ata;

                    // atualizando colunas no dt
                    this.dt.SetValue("Ajustes de Ata", r, "- R$ " + string.Format("{0:#,0.00}", soma_debito_ata));
                    this.dt.SetValue("Comissão Final", r, "R$ " + string.Format("{0:#,0.00}", comissao_final_ajustado2));
                }

                #endregion

                #region :: LOOPING PARA ATUALIZAR VALORES DE SPLIT

                /* ATUALIZANDO NO DATATABLE OS SPLITS */
                foreach (var split in splits)
                {
                    // se não tiver valor no dicionário de atas, então passa pra frente
                    if (split.Value.Count == 0)
                    {
                        continue;
                    }

                    double soma_debito_split = 0.0;
                    double comissao_base = 0;
                    int r = Convert.ToInt32(split_ref[split.Key]["ndx"]);
                    string str_com_final = this.dt.GetValue("Comissão Final", r);
                    Double.TryParse(str_com_final.Replace("R$", ""), out comissao_base);

                    foreach (var valores in split.Value)
                    {
                        // inicializando valores e calculando
                        double comissao_final = 0;
                        double ajuste_split = comissao_base * valores["vlr"] / 100;
                        int row = Convert.ToInt32(valores["ndx"]);
                        string str_com_final2 = this.dt.GetValue("Comissão Final", row);
                        Double.TryParse(str_com_final2.Replace("R$", ""), out comissao_final);
                        double comissao_final_ajustado = comissao_final + ajuste_split;

                        // atualizando colunas no dt
                        this.dt.SetValue("SPLIT", row, "+ R$ " + string.Format("{0:#,0.00}", ajuste_split));
                        this.dt.SetValue("Comissão Final", row, "R$ " + string.Format("{0:#,0.00}", comissao_final_ajustado));

                        // incrementando tudo que eu já debitei deste grupo
                        soma_debito_split += ajuste_split;
                    }

                    double comissao_final_ajustado2 = comissao_base - soma_debito_split;

                    // atualizando colunas no dt
                    this.dt.SetValue("SPLIT", r, "- R$ " + string.Format("{0:#,0.00}", soma_debito_split));
                    this.dt.SetValue("Comissão Final", r, "R$ " + string.Format("{0:#,0.00}", comissao_final_ajustado2));
                }

                #endregion

                #region :: LOOPING PARA ORGANIZAR OS TOTALIZADORES

                for (int i = 0; i < this.dt.Rows.Count; i++)
                {
                    /* ORGANIZAÇÃO DOS TOTALIZADORES */
                    double comissao = 0, comissao_perc = 0;

                    string str_comissao = this.dt.GetValue("Comissão Final", i);
                    string str_comissao_perc = this.dt.GetValue("Comissão %", i);

                    dynamic val_base = this.dt.GetValue(campo_base, i);
                    Double.TryParse(str_comissao.Replace("R$", ""), out comissao);
                    Double.TryParse(str_comissao_perc.Replace("%", ""), out comissao_perc);

                    if (!soma_lead.ContainsKey(val_base))
                    {
                        soma_lead.Add(val_base, new Dictionary<string, double>(){
                            {"comissao",comissao},
                            //{"comissao_perc",comissao_perc},
                        });
                    }
                    else
                    {
                        soma_lead[val_base]["comissao"] = soma_lead[val_base]["comissao"] + comissao;
                        //soma_lead[val_base]["comissao_perc"] = soma_lead[val_base]["comissao_perc"] + comissao_perc;
                    }
                }

                #endregion

                #region :: LOOPING PARA ATUALIZAR TOTALIZADORES JÁ ORGANIZADOS

                this.dtTemp.CopyFrom(this.dt);

                /* APLICANDO OS TOTALIZADORES ORGANIZADOS */
                foreach (var val_base in soma_lead)
                {
                    this.dt.Rows.Add();
                    int row = this.dt.Rows.Count - 1;
                    this.dt.SetValue(campo_base, row, val_base.Key);
                    this.dt.SetValue("Comissão Final", row, "R$ " + string.Format("{0:#,0.00}", soma_lead[val_base.Key]["comissao"]));
                    //this.dt.SetValue("Comissão %", row, string.Format("{0:#,0.00}", soma_lead[val_base.Key]["comissao_perc"]) + "%");
                }

                #endregion

                #endregion


                //clique no menu "Ajustar Colunas"
                this.Addon.SBO_Application.ActivateMenuItem("1300");
            }
            catch( Exception ex )
            {
                this.Addon.StatusErro("Erro ao efetuar a pesquisa.\nErro: " + ex.Message);
            }
            finally
            {
                this.SapForm.Freeze(false);
            }
        }

        /// <summary>
        ///  De acordo com o que foi pesquisa, verifica se consegue dar um parse em int.
        ///  se não conseguir, tenta quebrar por virgula e verifica se é uma lista de nomes.
        /// </summary>
        /// <param name="pesquisa"></param>
        /// <returns></returns>
        public string MontaWherePesquisa(string pesquisa)
        {
            string in_lead = "", like_comissionado = "", where = "", between_lead = "";
            int lead_global = 0;

            if (pesquisa.Contains(",") || pesquisa.Contains(";"))
            {
                string[] pieces = pesquisa.Split(new[] { ",",";" }, StringSplitOptions.None);
                foreach (string piece in pieces)
                {
                    int lead = 0;
                    if (Int32.TryParse(piece, out lead))
                    {
                        in_lead += "," + lead;
                    }
                    else
                    {
                        like_comissionado +=
                            !String.IsNullOrEmpty(like_comissionado) ? " OR tb5.firstName LIKE '%" + piece + "%' OR tb5.lastName LIKE '%" + piece + "%' "
                            : " tb5.firstName LIKE '%" + piece + "%' OR tb5.lastName LIKE '%" + piece + "%' ";
                    }
                }
            }
            else if (pesquisa.Contains("-") || pesquisa.Contains(":"))
            {
                string[] pieces = pesquisa.Split(new[] { "-",":" }, StringSplitOptions.None);

                int lead1 = 0, lead2 = 0;
                if (Int32.TryParse(pieces[0], out lead1) && Int32.TryParse(pieces[1], out lead2))
                {
                    between_lead = " tb1.U_UPD_IT_LEAD BETWEEN " + lead1 + " AND  " + lead2 + " ";
                }
            }
            else if (Int32.TryParse(pesquisa, out lead_global))
            {
                where += " AND tb1.U_UPD_IT_LEAD = " + lead_global + " ";
            }
            else
            {
                like_comissionado += " tb5.firstName LIKE '%" + pesquisa + "%' OR tb5.lastName LIKE '%" + pesquisa + "%' ";
            }

            //Se não estiver vazio, a primeira posição sempre vai ser uma ,
            if(!String.IsNullOrEmpty(in_lead))
            {
                //a primeira posição sempre vai ser uma virgula, então sempre tira ela.
                in_lead = in_lead.Remove(0, 1);
                where += " AND tb1.U_UPD_IT_LEAD IN ( " + in_lead + " ) ";
            }

            if( !String.IsNullOrEmpty(like_comissionado))
            {
                where += " AND ( " + like_comissionado + " ) ";
            }

            if( !String.IsNullOrEmpty(between_lead) )
            {
                where += " AND " + between_lead;
            }
            
            return where;
        }

        /// <summary>
        /// Expandir o grid
        /// </summary>
        public void ExpandirGrid()
        {
            this.oGrid.Rows.ExpandAll();
        }

        /// <summary>
        /// Fechar o grid
        /// </summary>
        public void FecharGrid()
        {
            this.oGrid.Rows.CollapseAll();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool FiltroPossuiPeriodoValido( bool veio_da_pesquisa = false )
        {
            string de = this.GetValue("de");
            string ate = this.GetValue("ate");
            bool res = true;

            if (!String.IsNullOrEmpty(de) && !String.IsNullOrEmpty(ate))
            {
                int ano_de = this.Addon.ToDatetime(de).Year;
                int ano_ate = this.Addon.ToDatetime(ate).Year;

                if (ano_de != ano_ate)
                {
                    res = false;
                    
                    // Se veio da pesquisa, então tem que dar o POPUP, se veio do change, não pode dar o popup senão fica preso porcausa dos eventos de perda de foco, etc.
                    this.Addon.StatusErro("Não é possível filtrar por um período com ano distinto.", veio_da_pesquisa);
                    
                    if( veio_da_pesquisa )
                        this.GetItem("de").Click();
                }
            }
            else
            {
                this.Addon.StatusErro("Defina um período 'De' e 'Até'");
            }

            return res;
        }

        #endregion

    }
}
