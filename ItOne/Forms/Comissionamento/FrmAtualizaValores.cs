using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TShark;

namespace ITOne
{
    class FrmAtualizaValores : TShark.Forms
    {

        #region :: Definições globais do Form

        /// <summary>
        /// Dicionário que contém as informações básicas de uma matriz, para ser usada em todo o contexto do form.
        /// dbdatasource: DbDataSource associado a matriz (caso seja UDO Child).
        /// datatable: DataTable associado a matriz ( caso exista )
        /// nome: Nome do Componente da matriz.
        /// specific: Retorna o objeto SAP da matriz (SAPbouiCOM.Matrix)
        /// sql: Sql inicial para dar carga a matriz.
        /// </summary>
        public Dictionary<string, dynamic> mtx = new Dictionary<string, dynamic>() { };

        #endregion

        public FrmAtualizaValores(Addon addOn, Dictionary<string, dynamic> ExtraParams = null): base(addOn, ExtraParams)
        {
            //Define o id do form como o nome da classe
            this.FormId = "FrmAtualizaValores";

            //Define as configurações do form
            this.FormParams = new FormParams()
            {
                Title = "Ajustes Manuais de Comissionamento",
                Focus = "U_lead",

                //Definição de tamanho e posição do Form
                Bounds = new Bounds(){
                    Top = 60,
                    Left = 430,
                    Width = 800,
                    Height = 440
                },

                #region :: Layout Componentes

                Linhas = new Dictionary<string,int>(){
                    {"hd01", 100},
                    {"U_lead", 15},{"U_docentry", 20},{"btnPesq", 10},{"space", 55},
                    {"space1", 100},
                    {"hd02", 100},
                    {"matrix", 100},
                    {"space2", 65},{"btnAdd", 20},{"btnRmv", 15},
                },

                Buttons = new Dictionary<string, int>(){
                    {"btnFechar", 20},{"space", 80},
                },

                #endregion


                #region :: Propriedades Componentes

                Controls = new Dictionary<string,CompDefinition>(){
                    
                    #region :: Cabeçalho

                    {"hd01", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_RECTANGLE,
                        Label = "Pesquisar por",
                        Height = 1,
                    }},
                    {"U_lead", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Nº do Lead",
                        //onKeyDownHandler = "OnKeyDown"
                    }},
                    {"U_docentry", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Nº do Pedido de Venda",
                        //onKeyDownHandler = "OnKeyDown"
                    }},
                    {"btnPesq", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        marginTop = 10,
                        Caption = "Pesquisar",
                    }},

                    #endregion


                    #region :: Matriz

                    {"hd02", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_RECTANGLE,
                        Label = "Participantes e suas comissões",
                        Height = 1,
                    }},
                    {"matrix", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_MATRIX,
                        Height = 270,
                        Enabled = false
                    }},
                    {"btnAdd", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "SPLIT"
                    }},
                    {"btnRmv", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Remover"
                    }},

                    #endregion

                    
                    #region :: Botões Padrões
 
                    {"btnFechar", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Fechar"
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
        public void FrmAtualizaValoresOnFormOpen(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void FrmAtualizaValoresOnFormClose(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        #endregion


        #region :: Eventos dos Componentes

        /// <summary>
        ///
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnKeyDown(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = false;

            // Só deixa passar se for ENTER ou TAB
            if (evObj.CharPressed != 13 && evObj.CharPressed != 9)
                return;

            this.PesquisarParticipantes();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnPesqOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = false;

            this.PesquisarParticipantes();
        }

        /// <summary>
        /// Insere participante.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnAddOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            int docentry = this.BuscaDocEntryParaAddParticipante();

            if( docentry > 0 )
            {
                Dictionary<string, dynamic> val = new Dictionary<string, dynamic>() { 
                    {"U_docentry",docentry.ToString()}
                };

                SAPbouiCOM.DataTable dt = this.SapForm.DataSources.DataTables.Item(this.mtx["datatable"]);
                int lead = dt.GetValue("lead", 0);
                if(lead > 0)
                {
                    val.Add("lead", lead.ToString());
                }
                
                this.Addon.OpenForm("FrmAddParticipante", this,val);
            }
        }

        /// <summary>
        /// Insere participante.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnRmvOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            this.DeleteMatrixOnServer("matrix", "@UPD_IT_PARTICIP");
        }

        #endregion


        #region :: Matriz

        /// <summary>
        /// Matriz de Comunicação
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void matrixOnCreate(SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            //Declarações básicas da matriz.
            this.mtx["dbdatasource"]  = "";
            this.mtx["datatable"]     = "LISTA";
            this.mtx["nome"]          = evObj.ItemUID;
            this.mtx["sql"]           = 
                "   SELECT " +
                "       tb1.Code, tb1.U_docentry, tb4.U_nome as funcao, tb3.firstName + ' ' + tb3.lastName as nome, tb1.U_percom, tb2.U_UPD_IT_LEAD as lead " +
                "   FROM [@UPD_IT_PARTICIP] (NOLOCK) tb1 " +
                "   INNER JOIN ORDR tb2 (NOLOCK) ON ( tb2.DocNum = tb1.U_docentry ) " +
                "   INNER JOIN OHEM tb3 (NOLOCK) ON ( tb1.U_empid = tb3.empID ) " +
                "   INNER JOIN [@UPD_IT_FUNCOES] tb4 (NOLOCK) ON ( tb1.U_funcao = tb4.Code ) " +
                "   WHERE U_oculto <> 'S' ";

            this.mtx["specific"] = this.SetupMatrix(evObj.ItemUID, this.mtx["datatable"], new List<ColumnDefinition>() { 
                new ColumnDefinition() { Width = 3 ,        Id = "hash",        Caption = "#", Bind = false},
                new ColumnDefinition() { Percent = 10,      Id = "lead",        Caption = "Nº Lead"},
                new ColumnDefinition() { Percent = 10,      Id = "U_docentry",  Caption = "Pedido Venda", LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Order },
                new ColumnDefinition() { Percent = 30,      Id = "funcao",      Caption = "Função",  
                    Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                    PopulateSQL = "SELECT Code, U_nome FROM [@UPD_IT_FUNCOES] WHERE U_ativo = 'S' ORDER BY U_nome",
                },
                new ColumnDefinition() { Percent = 35,      Id = "nome",     Caption = "Colaborador"},
                new ColumnDefinition() { Percent = 10,      Id = "U_percom",    Caption = "Comissão (%)", RightJustified = true  },
                
                //colunas invisiveis
                new ColumnDefinition() { Id = "Code",       Visible = false},

            }, true, this.mtx["sql"] + " AND 1 = 2 ");
        }

        #endregion


        #region :: Regras de Negócio

        /// <summary>
        /// Pesquisa os participantes de acordo com o num do lead ou numero do pedido.
        /// </summary>
        public void PesquisarParticipantes()
        {
            this.SapForm.Freeze(true);

            try
            {
                string lead     = this.GetValue("U_lead");
                string docentry = this.GetValue("U_docentry");

                if (String.IsNullOrEmpty(lead) && String.IsNullOrEmpty(docentry))
                    return;

                string where = "";

                if (!String.IsNullOrEmpty(lead))
                {
                    where += " AND tb2.U_UPD_IT_LEAD =  " + lead;
                }
                if (!String.IsNullOrEmpty(docentry))
                {
                    where += " AND tb1.U_docentry =  " + docentry;
                }

                string sql = this.mtx["sql"] + where;
                SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                this.RefreshMatrix(this.mtx["nome"], this.mtx["datatable"], sql);

                if (rs.RecordCount == 0)
                {
                    string msg = "";
                    msg =
                        !String.IsNullOrEmpty(lead) ?
                        "Não foram encontrados Pedidos de Venda com o Lead " + lead + " e que contenham participantes " :
                        "Não foram encontrados participantes no Pedido de Venda " + docentry;

                    this.Addon.StatusErro(msg, true);
                }
                else
                {
                    this.Addon.StatusInfo("Dados encontrados com sucesso.");
                }
            }
            finally
            {
                this.SapForm.Freeze(false);
            }
        }

        /// <summary>
        /// Método que verifica se possue dados suficientes para adicionar novo participante.
        /// 1) Deve existir ao menos uma linha no datatable, para que possa existir um DocEntry de referencia para adicionar o participante
        /// 2) Se nao existir nenhuma linha no dt, tenta pegar do campo de busca.
        /// </summary>
        /// <returns>DocEntry a inserir participante</returns>
        public int BuscaDocEntryParaAddParticipante()
        {
            int docentry = 0;

            SAPbouiCOM.DataTable dt = this.SapForm.DataSources.DataTables.Item(this.mtx["datatable"]);
            if( dt.IsEmpty )
            {
                // Se tiver vazio o datatable, tenta pegar do valor do campo de busca e associar o participante a este pedido buscado.
                string str_docentry = this.GetValue("U_docentry");
                if( Int32.TryParse(str_docentry, out docentry))
                {
                    SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rs.DoQuery("SELECT COUNT(*) FROM ORDR (NOLOCK) WHERE DocNum = " + docentry);

                    if( rs.RecordCount > 0 )
                    {
                        if (this.Addon.SBO_Application.MessageBox("Deseja adicionar Participante para o Pedido de Venda " + docentry + "?", 2, "Sim", "Não") == 1)
                        {
                            return docentry;
                        }
                        else
                        {
                            return 0;
                        }
                    }
                    else
                    {
                        this.Addon.StatusErro("Pedido de Venda " + docentry + " inexistente.", true);
                        return 0;
                    }
                }
                else
                {
                    this.Addon.StatusErro("Não foi possível encontrar um Pedido de Venda para adicionar um Participante.",true);
                }
            }
            else
            {
                // Se não estiver vazio, aceita o docentry da primeira linha do datatable.
                docentry = dt.GetValue("U_docentry",0);
            }

            return docentry;
        }

        #endregion

    }
}
