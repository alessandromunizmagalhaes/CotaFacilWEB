using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TShark;
using System.IO;
using System.Windows;

namespace ITOne
{
    class FrmRelatorioEquipe : TShark.Forms
    {
        SAPbouiCOM.DataTable dt = null;
        SAPbouiCOM.DataTable dtTemp = null;
        SAPbouiCOM.Grid oGrid   = null;

        public string nome_para_meta_do_time = "META DO TIME";
        public string nome_para_total_por_comissionado = "";
        
        public FrmRelatorioEquipe(Addon addOn, Dictionary<string, dynamic> ExtraParams = null): base(addOn, ExtraParams)
        {
            //Define o id do form como o nome da classe
            this.FormId = "FrmRelatorioEquipe";

            //Define as configurações do form
            this.FormParams = new FormParams()
            {
                Title = "Relatório de Perfomance de Equipe",

                //Definição de tamanho e posição do Form
                Bounds = new Bounds(){
                    Top = 25,
                    Left = 280,
                    Width = 1000,
                    Height = 515
                },

                #region :: Layout Componentes

                Linhas = new Dictionary<string,int>(){
                    {"space", 10},{"chkWON", 6},{"chkCOMMIT", 7},{"chkSU", 10},{"chkUPSIDE", 7},{"de", 8},{"ate", 8},{"search", 30},{"btnPesq", 14},
                    {"hd01", 100},
                    {"grid", 100},
                    {"edCota", 9},{"edWON", 9},{"edCOMMIT", 9},{"edSU", 9},{"edUPSIDE", 9},{"edResult", 9},{"edMetaTime", 9},{"space2", 10},{"btnExpand", 13},{"btnColapse", 13},
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
                        UserDataType = SAPbouiCOM.BoDataType.dt_DATE
                    }},
                    {"ate", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Até",
                        UserDataType = SAPbouiCOM.BoDataType.dt_DATE
                    }},
                    {"search", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Pesquisar por Vendedor",
                    }},
                    {"chkWON", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX,
                        Label = "WON",
                    }},
                    {"chkCOMMIT", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX,
                        Label = "COMMIT",
                    }},
                    {"chkSU", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX,
                        Label = "STRONG UPSIDE",
                    }},
                    {"chkUPSIDE", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX,
                        Label = "UPSIDE",
                    }},
                    {"btnPesq", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "OK",
                        marginTop = 10,
                    }},

                    #endregion
                    

                    #region :: Grid

                    {"hd01", new CompDefinition(){
                        Label = "Relatório de Perfomance de Equipe",
                        Type = SAPbouiCOM.BoFormItemTypes.it_RECTANGLE,
                        Height = 1
                    }},
                    {"grid", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_GRID,
                        Height = 370
                    }},
                    {"btnExpand", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Expandir",
                        marginTop = 10
                    }},
                    {"btnColapse", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Fechar",
                        marginTop = 10
                    }},

                    #endregion

                    
                    #region :: Totalizadores
                    
                    {"edCota", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Cota",
                        Enabled = false
                    }},
                    {"edWON", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "WON",
                        Enabled = false
                    }},
                    {"edCOMMIT", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "COMMIT",
                        Enabled = false
                    }},
                    {"edSU", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "STRONG UPSIDE",
                        Enabled = false
                    }},
                    {"edUPSIDE", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "UPSIDE",
                        Enabled = false
                    }},
                    {"edResult", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Resultado Previsto",
                        Enabled = false
                    }},
                    {"edMetaTime", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Meta Time",
                        Enabled = false
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
        public void FrmRelatorioEquipeOnFormOpen(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            DateTime dt_de  = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime dt_ate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.DaysInMonth(DateTime.Now.Year,DateTime.Now.Month));

            this.UpdateUserDataSource(new Dictionary<string, dynamic>()
            {
                {"de",dt_de.ToString("yyyyMMdd")},
                {"ate",dt_ate.ToString("yyyyMMdd")},
                {"edCota","R$ " + string.Format("{0:#,0.00}", 0.0)},
                {"edWON","R$ " + string.Format("{0:#,0.00}", 0.0)},
                {"edCOMMIT","R$ " + string.Format("{0:#,0.00}", 0.0)},
                {"edSU","R$ " + string.Format("{0:#,0.00}", 0.0)},
                {"edUPSIDE","R$ " + string.Format("{0:#,0.00}", 0.0)},
                {"edResult","R$ " + string.Format("{0:#,0.00}", 0.0)},
                {"edMetaTime","R$ " + string.Format("{0:#,0.00}", 0.0)},
            });
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void FrmRelatorioEquipeOnFormClose(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
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
                string filename = "rel_performance_equipe.csv";
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
            
            this.dt     = this.SapForm.DataSources.DataTables.Add("PESQUISA");
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
            this.SapForm.Freeze(true);

            try
            {
                string where = " WHERE 1 = 1 ";

                // Definindo filtros
                // F na frente pq é de filtro.
                string fwon     = this.GetValue("chkWON");
                string fcommit  = this.GetValue("chkCOMMIT");
                string fsu      = this.GetValue("chkSU");
                string fupside  = this.GetValue("chkUPSIDE");
                string de       = this.GetValue("de");
                string ate      = this.GetValue("ate");
                string pesquisa = this.GetValue("search");

                //se não tiver nenhum, é pq tem que trazer todos. setando todos pra Y.
                if (fwon != "Y" && fcommit != "Y" && fsu != "Y" && fupside != "Y")
                {
                    fwon = fcommit = fsu = fupside = "Y";
                }

                if (!String.IsNullOrEmpty(de))
                {
                    DateTime data_de = this.Addon.ToDatetime(de);
                    de = data_de.ToString("yyy-MM-dd");
                    where += " AND CONVERT(DATE,tb1.OpenDate,103) >= '" + de + "'";
                }

                if (!String.IsNullOrEmpty(ate))
                {
                    DateTime data_ate = this.Addon.ToDatetime(ate);
                    ate = data_ate.ToString("yyy-MM-dd");
                    where += " AND CONVERT(DATE,tb1.OpenDate,103) <= '" + ate + "'";
                }

                if (!String.IsNullOrEmpty(pesquisa))
                    where += this.MontaWherePesquisa(pesquisa);

                #region :: SELECT

                string select_won       = fwon == "Y"       ? 
                    " ,FORMAT( CASE WHEN [WON] > 0 THEN [WON] ELSE 0 END, 'C','pt-br' ) as [WON] " +
                    " ,FORMAT(  " +
                    "       CASE WHEN [Cota] > 0 AND [WON] > 0 " +
                    "           THEN " +
                    "               ( [WON] * 100 / [Cota] ) " +
                    "           ELSE 0 " +
                    "       END, 'N','pt-br' ) + '%' as [% Cota Parcial] "
                    : "";
                string select_commit    = fcommit == "Y"    ? " ,FORMAT( CASE WHEN [COMMIT] > 0 THEN [COMMIT] ELSE 0 END, 'C','pt-br' ) as [COMMIT] " : "";
                string select_su        = fsu == "Y"        ? " ,FORMAT( CASE WHEN [STRONG UPSIDE] > 0 THEN [STRONG UPSIDE] ELSE 0 END, 'C','pt-br' ) as [STRONG UPSIDE] " : "";
                string select_upside    = fupside == "Y"    ? " ,FORMAT( CASE WHEN [UPSIDE] > 0 THEN [UPSIDE] ELSE 0 END, 'C','pt-br' ) as [UPSIDE] " : "";
                
                string soma_won         = fwon      == "Y" ? " ISNULL([WON],0) "            : " 0 ";
                string soma_commit      = fcommit   == "Y" ? " ISNULL([COMMIT],0) "         : " 0 ";
                string soma_su          = fsu       == "Y" ? " ISNULL([STRONG UPSIDE],0) "  : " 0 ";
                string soma_upside      = fupside   == "Y" ? " ISNULL([UPSIDE],0) "         : " 0 ";
                
                string select_soma      = " ( " + soma_won + " + " + soma_commit + " + " + soma_su + " + " + soma_upside + " ) ";
                string select_soma_perc = " ( " + soma_won + " + " + soma_commit + " + " + soma_su + " + " + soma_upside + " ) ";

                string status_won         = fwon      == "Y" ? ",[WON]" : "";
                string status_commit      = fcommit   == "Y" ? ",[COMMIT]" : "";
                string status_su          = fsu       == "Y" ? ",[STRONG UPSIDE]" : "";
                string status_upside      = fupside   == "Y" ? ",[UPSIDE]" : "";
                string status = status_won + status_commit + status_su + status_upside;
                status = status.Remove(0,1);

                string where_won        = fwon == "Y" ? ",'WON'" : "";
                string where_commit     = fcommit == "Y" ? ",'COMMIT'" : "";
                string where_su         = fsu == "Y" ? ",'STRONG UPSIDE'" : "";
                string where_upside     = fupside == "Y" ? ",'UPSIDE'" : "";
                string where_sts = where_won + where_commit + where_su + where_upside;
                where_sts = where_sts.Remove(0, 1);

                string select_default = 
                "SELECT " +
                "   ISNULL([Time],'SEM TIME') as Time " +
                "   ,[Account Manager] " +
                "   ,FORMAT( CASE WHEN [Cota] > 0 THEN [Cota] ELSE 0 END, 'C','pt-br' ) as [Cota] " +
                    select_won +
                    select_commit +
                    select_su +
                    select_upside +
                "   ,FORMAT( " + select_soma + " , 'C','pt-br') as [Resultado Previsto] " +
                "   ,FORMAT(  " +
                "       CASE WHEN [Cota] > 0  " +
                "           THEN " +
                "               ( " + select_soma_perc + " * 100 / [Cota] ) " +
                "           ELSE 0 " +
                "       END, 'N','pt-br' ) + '%' as [% Cota Previsto] " +
                "   , meta_time ";

                #endregion


                #region :: SQL BASE

                string sql_base =
                "FROM " +
                "( " +
	            "   SELECT  " +
		        "       tb4.firstName + ' ' + tb4.lastName as [Account Manager] " +
		        "       , tb3.firstName + ' ' + tb3.lastName as [Time] " +
		        "       , SUM(U_upd_12_renda) as renda , tb2.U_upd_2_status " +
                "       , tb3.U_UPD_IT_META_TIME as 'meta_time' " +
		        "       , AVG(tb4.U_UPD_IT_META) as [Cota] " +
	            "   FROM OOPR tb1 " +
	            "   INNER JOIN OPR2 tb2(NOLOCK) ON (tb2.OpportId = tb1.OpprId) " +
	            "   LEFT JOIN OHEM tb3(NOLOCK) ON (tb2.U_UPD_IT_TIME = tb3.empID) " +
	            "   INNER JOIN OHEM tb4(NOLOCK) ON (tb1.SlpCode  = tb4.salesPrson) " +
                    where +
                "   AND tb2.U_upd_2_status IN ( " + where_sts + " ) " +
                "   GROUP BY tb3.firstName + ' ' + tb3.lastName, tb4.firstName + ' '	+ tb4.lastName, tb2.U_upd_2_status, tb3.U_UPD_IT_META_TIME " +
                ") up " +
                "PIVOT  " +
                "( " +
	            "   SUM(renda) " +
                "   FOR U_upd_2_status IN ( " + status + " ) " +
                ") AS pvt";

                #endregion


                string sql = select_default + sql_base;

                this.dt.ExecuteQuery(sql);
                this.oGrid.Rows.CollapseAll();
                this.oGrid.CollapseLevel = 1;

                ((Addon)this.Addon).SalvarSQLArquivoTXT(sql);
                
                Dictionary<dynamic, Dictionary<string, double>> soma_nvl1 = new Dictionary<dynamic, Dictionary<string, double>>() { };
                Dictionary<string, double> totais = new Dictionary<string, double>() { 
                    {"cota",0.0},
                    {"won",0.0},
                    {"commit",0.0},
                    {"strong_upside",0.0},
                    {"upside",0.0},
                    {"result",0.0},
                    {"meta_time",0.0},
                };
                 
                for (int i = 0; i < this.dt.Rows.Count; i++)
                {
                    double cota = 0, won = 0, commit = 0, strong_upside = 0, upside = 0, result = 0;

                    string str_cota             = this.dt.GetValue("Cota", i);
                    string str_won              = fwon      == "Y" ? this.dt.GetValue("WON", i) : "0";
                    string str_commit           = fcommit   == "Y" ? this.dt.GetValue("COMMIT", i) : "0";
                    string str_strong_upside    = fsu       == "Y" ? this.dt.GetValue("STRONG UPSIDE", i) : "0";
                    string str_upside           = fupside   == "Y" ? this.dt.GetValue("UPSIDE", i) : "0";
                    string str_result           = this.dt.GetValue("Resultado Previsto", i);
                    double meta_time            = this.dt.GetValue("meta_time", i);
                    dynamic val_base            = this.dt.GetValue("Time", i);
                    
                    Double.TryParse(str_cota.Replace("R$",""), out cota);
                    Double.TryParse(str_won.Replace("R$",""), out won);
                    Double.TryParse(str_commit.Replace("R$",""), out commit);
                    Double.TryParse(str_strong_upside.Replace("R$",""), out strong_upside);
                    Double.TryParse(str_upside.Replace("R$",""), out upside);
                    Double.TryParse(str_result.Replace("R$", ""), out result);

                    if(!soma_nvl1.ContainsKey(val_base))
                    {
                        soma_nvl1.Add(val_base, new Dictionary<string, double>()
                        {
                            {"cota",cota},
                            {"won",won},
                            {"commit",commit},
                            {"strong_upside",strong_upside},
                            {"upside",upside},
                            {"result",result},
                            {"meta_time",meta_time},
                        });

                        totais["meta_time"] = totais["meta_time"] + meta_time;
                    }
                    else
                    {
                        soma_nvl1[val_base]["cota"]             = soma_nvl1[val_base]["cota"] + cota;
                        soma_nvl1[val_base]["won"]              = soma_nvl1[val_base]["won"] + won;
                        soma_nvl1[val_base]["commit"]           = soma_nvl1[val_base]["commit"] + commit;
                        soma_nvl1[val_base]["strong_upside"]    = soma_nvl1[val_base]["strong_upside"] + strong_upside;
                        soma_nvl1[val_base]["upside"]           = soma_nvl1[val_base]["upside"] + upside;
                        soma_nvl1[val_base]["result"]           = soma_nvl1[val_base]["result"] + result;
                    }

                    totais["cota"]          = totais["cota"] + cota;
                    totais["won"]           = totais["won"] + won;
                    totais["commit"]        = totais["commit"] + commit;
                    totais["strong_upside"] = totais["strong_upside"] + strong_upside;
                    totais["upside"]        = totais["upside"] + upside;
                    totais["result"]        = totais["result"] + result;
                }

                this.dtTemp.CopyFrom(this.dt);

                foreach (var val_base in soma_nvl1)
                {
                    this.dt.Rows.Add();
                    int row = this.dt.Rows.Count - 1;
                    this.dt.SetValue("Time", row, val_base.Key);
                    this.dt.SetValue("Cota", row, "R$ " + string.Format("{0:#,0.00}", soma_nvl1[val_base.Key]["cota"]));
                    this.dt.SetValue("Account Manager", row, this.nome_para_total_por_comissionado);
                    
                    if( fwon == "Y" )
                        this.dt.SetValue("WON", row, "R$ " + string.Format("{0:#,0.00}", soma_nvl1[val_base.Key]["won"]));

                    if(fcommit == "Y")
                        this.dt.SetValue("COMMIT", row, "R$ " + string.Format("{0:#,0.00}", soma_nvl1[val_base.Key]["commit"]));

                    if(fsu == "Y")
                        this.dt.SetValue("STRONG UPSIDE", row, "R$ " + string.Format("{0:#,0.00}", soma_nvl1[val_base.Key]["strong_upside"]));

                    if(fupside == "Y")
                        this.dt.SetValue("UPSIDE", row, "R$ " + string.Format("{0:#,0.00}", soma_nvl1[val_base.Key]["upside"]));

                    this.dt.SetValue("Resultado Previsto", row, "R$ " + string.Format("{0:#,0.00}", soma_nvl1[val_base.Key]["result"]));


                    /* NOVA LINHA DA META DO TIME */
                    this.dt.Rows.Add();
                    int r = this.dt.Rows.Count - 1;
                    this.dt.SetValue("Time", r, val_base.Key);
                    this.dt.SetValue("Account Manager", r, this.nome_para_meta_do_time);
                    this.dt.SetValue("Cota", r, "R$ " + string.Format("{0:#,0.00}", soma_nvl1[val_base.Key]["meta_time"]));
                }

                this.oGrid.Columns.Item("Cota").RightJustified = true;

                if( fwon == "Y" )
                {
                    this.oGrid.Columns.Item("WON").RightJustified               = true;
                    this.oGrid.Columns.Item("% Cota Parcial").RightJustified    = true;
                }
                
                if( fcommit == "Y" )
                {
                    this.oGrid.Columns.Item("COMMIT").RightJustified = true;
                }

                if( fsu == "Y" )
                {
                    this.oGrid.Columns.Item("STRONG UPSIDE").RightJustified = true;
                }

                if( fupside == "Y" )
                {
                    this.oGrid.Columns.Item("UPSIDE").RightJustified = true;
                }

                this.oGrid.Columns.Item("Resultado Previsto").RightJustified = true;

                this.oGrid.Columns.Item("meta_time").Visible = false;

                this.UpdateUserDataSource(new Dictionary<string, dynamic>()
                {
                    {"edCota","R$ " + string.Format("{0:#,0.00}", totais["cota"])},
                    {"edWON","R$ " + string.Format("{0:#,0.00}", totais["won"])},
                    {"edCOMMIT","R$ " + string.Format("{0:#,0.00}", totais["commit"])},
                    {"edSU","R$ " + string.Format("{0:#,0.00}", totais["strong_upside"])},
                    {"edUPSIDE","R$ " + string.Format("{0:#,0.00}", totais["upside"])},
                    {"edResult","R$ " + string.Format("{0:#,0.00}", totais["result"])},
                    {"edMetaTime","R$ " + string.Format("{0:#,0.00}", totais["meta_time"])},
                });

                this.ColorirGrid();
                
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
            string like_comissionado = "", where = "";

            if (pesquisa.Contains(",") || pesquisa.Contains(";"))
            {
                string[] pieces = pesquisa.Split(new[] { ",",";" }, StringSplitOptions.None);
                foreach (string piece in pieces)
                {
                    like_comissionado +=
                        !String.IsNullOrEmpty(like_comissionado) ? " OR tb4.firstName + ' ' + tb4.lastName LIKE '%" + piece + "%' "
                        : " tb4.firstName + ' ' + tb4.lastName LIKE '%" + piece + "%' ";
                }
            }
            else
            {
                like_comissionado += " tb4.firstName + ' ' + tb4.lastName LIKE '%" + pesquisa + "%' ";
            }

            if( !String.IsNullOrEmpty(like_comissionado))
            {
                where += " AND ( " + like_comissionado + " ) ";
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
        /// colore as linhas totalizadoras do grid
        /// </summary>
        public void ColorirGrid()
        {
            int verde = 147 | (247 << 8) | (169 << 16);
            int laranja = 247 | (219 << 8) | (146 << 16);

            for (int i = 1; i < this.oGrid.Rows.Count; i++)
            {
                int row_in_dt = this.oGrid.GetDataTableRowIndex(i);

                if (row_in_dt == -1)
                    continue;

                string acct_manager = this.dt.GetValue("Account Manager", row_in_dt);
                string perc_cota    = this.dt.GetValue("% Cota Previsto", row_in_dt);

                if (acct_manager == this.nome_para_total_por_comissionado && String.IsNullOrEmpty(perc_cota))
                {
                    this.oGrid.CommonSetting.SetRowBackColor(i + 1, verde);
                }
                else if (acct_manager == this.nome_para_meta_do_time)
                {
                    this.oGrid.CommonSetting.SetRowBackColor(i + 1, laranja);
                }
            }
        }

        #endregion
    }
}
