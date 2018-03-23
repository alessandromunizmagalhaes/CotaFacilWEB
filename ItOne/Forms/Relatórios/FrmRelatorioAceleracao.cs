using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TShark;
using System.IO;
using System.Windows;

namespace ITOne
{
    class FrmRelatorioAceleracao : TShark.Forms
    {
        SAPbouiCOM.DataTable dt = null;
        SAPbouiCOM.DataTable dtTemp = null;
        SAPbouiCOM.Grid oGrid   = null;
        Dictionary<string, List<List<double>>> faixa_comissoes = new Dictionary<string, List<List<double>>>() { };
        Dictionary<string, Dictionary<string, double>> totalizadores = new Dictionary<string, Dictionary<string, double>>() { };
        
        public FrmRelatorioAceleracao(Addon addOn, Dictionary<string, dynamic> ExtraParams = null): base(addOn, ExtraParams)
        {
            //Define o id do form como o nome da classe
            this.FormId = "FrmRelatorioAceleracao";

            //Define as configurações do form
            this.FormParams = new FormParams()
            {
                Title = "Relatório de Aceleração Não Retroativa",

                //Definição de tamanho e posição do Form
                Bounds = new Bounds(){
                    Top = 25,
                    Left = 330,
                    Width = 1000,
                    Height = 500
                },

                #region :: Layout Componentes

                Linhas = new Dictionary<string,int>(){
                    {"space", 18},{"de", 10},{"ate", 10},{"search", 40},{"tipo_rel", 12},{"btnPesq", 10},
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
                        Label = "Pesquisar por Comissionado",
                    }},
                    {"tipo_rel", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                        Label = "Tipo Relatório",
                        PopulateItens = new Dictionary<string,string>(){
                            {"1","Comissão Total"},
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
                        Label = "Relatório de Comissão com Aceleração Não Retroativa",
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
        public void FrmRelatorioAceleracaoOnFormOpen(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            DateTime dt_de  = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime dt_ate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.DaysInMonth(DateTime.Now.Year,DateTime.Now.Month));

            this.UpdateUserDataSource(new Dictionary<string, dynamic>()
            {
                {"de",dt_de.ToString("yyyyMMdd")},
                {"ate",dt_ate.ToString("yyyyMMdd")},
                {"tipo_rel","1"},
            });
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void FrmRelatorioAceleracaoOnFormClose(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
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
                string filename = "rel_comissao_nao_retroativo.csv";
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
        /// 
        /// </summary>
        /// <returns></returns>
        public bool FiltroPossuiPeriodoValido(bool veio_da_pesquisa = false)
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

                    if (veio_da_pesquisa)
                        this.GetItem("de").Click();
                }
            }
            else
            {
                this.Addon.StatusErro("Defina um período 'De' e 'Até'");
            }

            return res;
        }

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
                string where = "";

                // Definindo filtros
                string de       = this.GetValue("de");
                string ate      = this.GetValue("ate");
                string pesquisa = this.GetValue("search");
                string tipo_rel = this.GetValue("tipo_rel");

                DateTime data_de = this.Addon.ToDatetime(de);
                de = data_de.ToString("yyy-MM-dd");
                where += " AND CONVERT(DATE,tb3.DocDate,103) >= '" + de + "'";

                DateTime data_ate = this.Addon.ToDatetime(ate);
                ate = data_ate.ToString("yyy-MM-dd");
                where += " AND CONVERT(DATE,tb3.DocDate,103) <= '" + ate + "'";

                int ano = data_de.Year;

                if (!String.IsNullOrEmpty(pesquisa))
                    where += this.MontaWherePesquisa(pesquisa);


                #region :: SQL Processado

                string join_processado = "  ";
                string where_processado = "  ";

                #endregion


                #region :: SQL Faturado

                string join_faturado = "  ";
                string where_faturado = "  ";

                #endregion


                #region :: SQL Realizado

                string join_realizado = " LEFT JOIN OINV tb6 (NOLOCK) ON (tb3.U_UPD_IT_LEAD = tb6.U_UPD_IT_LEAD AND tb6.DocStatus != 'C' AND tb6.CANCELED = 'N' ) ";
                string where_realizado = " AND tb3.U_UPD_IT_STATUS = 'R' ";

                #endregion


                string sql =
                    "SELECT " +
                    "   CASE WHEN tb8.firstname IS NULL THEN 'SEM TIME' ELSE tb8.firstname END AS 'Time' " +
                    "   , COALESCE(tb2.firstName,'') + ' ' + COALESCE(tb2.lastName,'') as 'Comissionado' " +
                    "   , CONVERT(VARCHAR(250),tb3.U_UPD_IT_LEAD) AS 'Nº Lead' " +
                    "   ,FORMAT(tb3.U_UPD_IT_RENDA,'C','pt-br') as 'Renda Comissionável' " +
                    "   ,FORMAT(ISNULL(tb4.U_meta,0.0),'C','pt-br') as 'Meta' " +
                    "   ,'% ' + FORMAT(0.0,'N','pt-br') as 'Comissão (%)' " +
                    "   ,FORMAT(0.0,'C','pt-br') as 'Valor Comissionado' " +
                    "   ,FORMAT(0.0,'C','pt-br') as 'Comissão Final' " +
                    "   ,tb3.DocEntry as 'Pedido de Venda' " +
                    "   , tb5.U_nome as 'Função' " +
                    "   , tb5.Code as 'code_funcao' " +
                    "FROM [@UPD_IT_PARTICIP] tb1 (NOLOCK) " +
                    "INNER JOIN OHEM (NOLOCK) tb2 ON (tb2.empID = tb1.U_empid) " +
                    "INNER JOIN ORDR (NOLOCK) tb3 ON (tb1.U_docentry = tb3.DocEntry) " +
                    "INNER JOIN  " +
                    "(  " +
                    "    SELECT Code, U_empid,SUM(U_meta) as U_meta  FROM [@UPD_IT_METAS]  " +
                    "    WHERE YEAR(U_dtinicio) =  " + ano + "  AND YEAR(U_dtfim) =  " + ano + " AND U_acelera = 2 " +
                    "    GROUP BY Code,U_empid  " +
                    ") tb4 ON (tb4.Code = tb1.U_funcao AND tb1.U_empid = tb4.U_empid )  " +
                    "INNER JOIN [@UPD_IT_FUNCOES] (NOLOCK) tb5 ON (tb1.U_funcao = tb5.Code) ";

                    sql += tipo_rel == "1" ? join_processado : (tipo_rel == "2" ? join_faturado : join_realizado);
                    sql += 
                    "INNER JOIN OPR2 tb7 (NOLOCK) ON (tb7.U_upd_1_nlead = tb3.U_UPD_IT_LEAD) " +
		            "LEFT JOIN OHEM tb8  (NOLOCK) ON (tb8.empID = tb7.U_UPD_IT_TIME )";

                    where =
                        " WHERE 1 = 1 " + where;
                    where += tipo_rel == "1" ? where_processado : (tipo_rel == "2" ? where_faturado : where_realizado);
                
                string order = 
                    " ORDER BY COALESCE(tb2.firstName,'') + ' ' + COALESCE(tb2.lastName,'') ASC, tb3.U_UPD_IT_LEAD ASC ";

                this.dt.ExecuteQuery(sql + where + " AND 1 = 2 " + order);
                this.dtTemp.ExecuteQuery(sql + where + order);
                
                this.oGrid.CollapseLevel = 2;

                #region :: Configuração das colunas
                
                this.oGrid.Columns.Item("Renda Comissionável").RightJustified = true;
                this.oGrid.Columns.Item("Meta").RightJustified = true;
                this.oGrid.Columns.Item("Comissão (%)").RightJustified = true;
                this.oGrid.Columns.Item("Valor Comissionado").RightJustified = true;
                this.oGrid.Columns.Item("Comissão Final").RightJustified = true;
                this.oGrid.Columns.Item("code_funcao").Visible = false;
                ((SAPbouiCOM.EditTextColumn)this.oGrid.Columns.Item("Pedido de Venda")).LinkedObjectType = "17";
                
                #endregion

                #region :: Pós Processamento

                this.SetFaixasDeComissao();
                this.dt.Rows.Remove(this.dt.Rows.Count - 1);
                this.totalizadores = new Dictionary<string, Dictionary<string, double>>() { };

                string funcao = "";
                string comissionado = "";

                double teto_meta = 0.0;
                List<List<double>> tetos = new List<List<double>>() { };
                for (int i = 0; i < this.dtTemp.Rows.Count; i++)
                {
                    /* BUSCANDO VALORES BASE PARA CÁLCULOS */
                    double val_pedido = 0, meta = 0;

                    string time                 = this.dtTemp.GetValue("Time", i);
                    string novo_comissionado    = this.dtTemp.GetValue("Comissionado", i);
                    string str_valor_pedido     = this.dtTemp.GetValue("Renda Comissionável", i);
                    string str_meta             = this.dtTemp.GetValue("Meta", i);
                    string nova_funcao          = this.dtTemp.GetValue("code_funcao", i);
                    string lead                 = this.dtTemp.GetValue("Nº Lead", i);
                    int pedido                  = this.dtTemp.GetValue("Pedido de Venda", i);

                    Double.TryParse(str_valor_pedido.Replace("R$", ""), out val_pedido);
                    Double.TryParse(str_meta.Replace("R$", ""), out meta);
                    double val_pedido_ref = val_pedido;

                    if (funcao != nova_funcao)
                    {
                        comissionado = "";
                        funcao = nova_funcao;
                        tetos = this.GetTetos(funcao, meta);
                    }

                    if (comissionado != novo_comissionado)
                    {
                        comissionado = novo_comissionado;
                        teto_meta = 0.0;
                        tetos = this.GetTetos(funcao, meta);
                    }

                    //se não tem teto/meta/comissão(%)
                    if (tetos.Count == 0)
                    {
                        #region :: Preenchendo os valores para atualizar datatable

                        Dictionary<string, dynamic> values = new Dictionary<string, dynamic>() {
                            {"time",time},
                            {"comissionado",comissionado},
                            {"lead",lead},
                            {"val_pedido_ref",val_pedido_ref},
                            {"meta",meta},
                            {"comissao_perc",0},
                            {"valor_comissionado",0},
                            {"comissao_final",0},
                            {"funcao",this.dtTemp.GetValue("Função", i)},
                            {"pedido",pedido},
                        };

                        #endregion

                        this.AtualizarDataTable(values);
                        this.AtualizarDataTableBase(values, i);
                        continue;
                    }
                        

                    if (teto_meta == 0)
                    {
                        //recuperar PRÓXIMO valor do teto da meta
                        teto_meta = tetos[0][0];
                    }

                    while (val_pedido > 0)
                    {
                        double valor_comissionado = val_pedido < teto_meta ? val_pedido : teto_meta;
                        val_pedido -= valor_comissionado;
                        teto_meta -= valor_comissionado;

                        #region :: Preenchendo os valores para atualizar datatable

                        Dictionary<string, dynamic> values = new Dictionary<string, dynamic>() {
                            {"time",time},
                            {"comissionado",comissionado},
                            {"lead",lead},
                            {"val_pedido_ref",val_pedido_ref},
                            {"meta",meta},
                            {"comissao_perc",tetos[0][1]},
                            {"valor_comissionado",valor_comissionado},
                            {"comissao_final",((tetos[0][1] * valor_comissionado) / 100)},
                            {"funcao",this.dtTemp.GetValue("Função", i)},
                            {"pedido",pedido},
                        };

                        #endregion

                        this.AtualizarDataTable(values);
                        this.AtualizarDataTableBase(values, i);

                        if (teto_meta == 0)
                        {
                            tetos.RemoveAt(0);
                            
                            //recuperar PRÓXIMO valor do teto da meta
                            teto_meta = tetos[0][0];
                        }
                    }
                }

                #endregion

                this.InserirTotalizadores();
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
                        !String.IsNullOrEmpty(like_comissionado) ? " OR tb2.firstName LIKE '%" + piece + "%' OR tb2.lastName LIKE '%" + piece + "%' "
                        : " tb2.firstName LIKE '%" + piece + "%' OR tb2.lastName LIKE '%" + piece + "%' ";
                }
            }
            else
            {
                like_comissionado += " tb2.firstName LIKE '%" + pesquisa + "%' OR tb2.lastName LIKE '%" + pesquisa + "%' ";
            }

            if( !String.IsNullOrEmpty(like_comissionado))
            {
                where += " AND ( " + like_comissionado + " ) ";
            }
            
            return where;
        }

        /// <summary>
        /// Retorna todas as faixas de comissão por função
        /// </summary>
        /// <returns></returns>
        public void SetFaixasDeComissao()
        {
            this.faixa_comissoes = new Dictionary<string, List<List<double>>>();

            SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery("SELECT Code,U_piso, U_teto , U_comissao FROM [@UPD_IT_COMISSAO]");

            while(!rs.EoF)
            {
                string funcao   = rs.Fields.Item("Code").Value;
                double piso     = rs.Fields.Item("U_piso").Value;
                double teto     = rs.Fields.Item("U_teto").Value;
                double comissao = rs.Fields.Item("U_comissao").Value;

                if (!this.faixa_comissoes.ContainsKey(funcao))
                {
                    this.faixa_comissoes.Add(funcao, new List<List<double>>() { });
                }

                this.faixa_comissoes[funcao].Add(new List<double>(){
                    piso, teto, comissao
                });
                
                rs.MoveNext();
            }
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
        public List<List<double>> GetTetos(string funcao, double meta)
        {
            List<List<double>> ret = new List<List<double>>(){};

            if(this.faixa_comissoes.ContainsKey(funcao) && meta > 0)
            {
                foreach (var faixa in this.faixa_comissoes[funcao])
                {
                    double total_teto = ((meta * faixa[1]) / 100) - ((meta * faixa[0]) / 100);
                    double comissao = faixa[2];

                    ret.Add(new List<double>()
                    {
                        total_teto,
                        comissao
                    });
                }
            }

            return ret;
        }

        /// <summary>
        /// 
        /// </summary>
        public void AtualizarDataTable( Dictionary<string, dynamic> values )
        {
            // gerando nova linha
            this.dt.Rows.Add();
            int row = this.dt.Rows.Count - 1;
            this.dt.SetValue("Time", row, values["time"]);
            this.dt.SetValue("Comissionado", row, values["comissionado"]);
            this.dt.SetValue("Nº Lead", row, values["lead"]);
            this.dt.SetValue("Renda Comissionável", row, "R$ " + string.Format("{0:#,0.00}", values["val_pedido_ref"]));
            this.dt.SetValue("Meta", row, "R$ " + string.Format("{0:#,0.00}", values["meta"]));
            this.dt.SetValue("Comissão (%)", row, string.Format("{0:#,0.00}", values["comissao_perc"]) + "%");
            this.dt.SetValue("Valor Comissionado", row, "R$ " + string.Format("{0:#,0.00}", values["valor_comissionado"]));
            this.dt.SetValue("Comissão Final", row, "R$ " + string.Format("{0:#,0.00}", values["comissao_final"]));
            this.dt.SetValue("Função", row, values["funcao"]);
            this.dt.SetValue("Pedido de Venda", row, values["pedido"]);

            if (values["comissao_final"] > 0)
            {
                if (!this.totalizadores.ContainsKey(values["time"]))
                {
                    this.totalizadores.Add(values["time"], new Dictionary<string, double>()
                    {
                        {values["comissionado"], values["comissao_final"]},
                    });
                }
                else
                {
                    if (!this.totalizadores[values["time"]].ContainsKey(values["comissionado"]))
                    {
                        this.totalizadores[values["time"]].Add(values["comissionado"], values["comissao_final"]);
                    }
                    else
                    {
                        this.totalizadores[values["time"]][values["comissionado"]] = this.totalizadores[values["time"]][values["comissionado"]] + values["comissao_final"];
                    }
                }
            }
        }

        /// <summary>
        /// atualiza o datatable base
        /// FUNÇÃO UTILIZADA SÓ PARA PREENCHER O DATATABLE BASE, PARA EXPORTAÇÃO DO CSV
        /// NA EXPORTAÇÃO DO CSV, NÃO SE LEVA AS LINHAS DE TOTALIZADORES
        /// PARA EVITAR IF'S NO CÓDIGO DE EXPORTAÇÃO, ESTAMOS ATUALIZANDO O DATATABLE BAE
        /// </summary>
        public void AtualizarDataTableBase(Dictionary<string, dynamic> values, int row)
        {
            // não precisa de atulizar as linhas que já tem valores né
            //this.dtTemp.SetValue("Time", row, values["time"]);
            //this.dtTemp.SetValue("Comissionado", row, values["comissionado"]);
            //this.dtTemp.SetValue("Nº Lead", row, values["lead"]);
            this.dtTemp.SetValue("Renda Comissionável", row, "R$ " + string.Format("{0:#,0.00}", values["val_pedido_ref"]));
            this.dtTemp.SetValue("Meta", row, "R$ " + string.Format("{0:#,0.00}", values["meta"]));
            this.dtTemp.SetValue("Comissão (%)", row, string.Format("{0:#,0.00}", values["comissao_perc"]) + "%");
            this.dtTemp.SetValue("Valor Comissionado", row, "R$ " + string.Format("{0:#,0.00}", values["valor_comissionado"]));
            this.dtTemp.SetValue("Comissão Final", row, "R$ " + string.Format("{0:#,0.00}", values["comissao_final"]));
            //this.dtTemp.SetValue("Função", row, values["funcao"]);
        }

        /// <summary>
        /// 
        /// </summary>
        public void InserirTotalizadores()
        {
            foreach(var time in this.totalizadores)
            {
                foreach(var total in time.Value)
                {
                    this.dt.Rows.Add();
                    int row = this.dt.Rows.Count - 1;
                    this.dt.SetValue("Time", row, time.Key);
                    this.dt.SetValue("Comissionado", row, total.Key);
                    this.dt.SetValue("Comissão Final", row, "R$ " + string.Format("{0:#,0.00}", total.Value));
                }
            }
        }

        /// <summary>
        /// colore as linhas totalizadoras do grid
        /// </summary>
        public void ColorirGrid()
        {
            int verde = 147 | (247 << 8) | (169 << 16);

            for (int i = 1; i < this.oGrid.Rows.Count; i++)
            {
                int row_in_dt = this.oGrid.GetDataTableRowIndex(i);

                if (row_in_dt == -1)
                    continue;

                string lead = this.dt.GetValue("Nº Lead", row_in_dt);

                if (String.IsNullOrEmpty(lead))
                {
                    this.oGrid.CommonSetting.SetRowBackColor(i + 1, verde);
                }
            }
        }

        #endregion
    }
}
