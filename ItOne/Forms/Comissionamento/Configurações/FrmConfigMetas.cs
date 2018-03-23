using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TShark;

namespace ITOne
{
    class FrmConfigMetas : TShark.Forms
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
        public Dictionary<string, dynamic> mtxMetas = new Dictionary<string, dynamic>() { };
        public Dictionary<string, dynamic> mtxComissoes = new Dictionary<string, dynamic>() { };

        #endregion

        public FrmConfigMetas(Addon addOn, Dictionary<string, dynamic> ExtraParams = null): base(addOn, ExtraParams)
        {
            //Define o id do form como o nome da classe
            this.FormId = "FrmConfigMetas";

            //Define as configurações do form
            this.FormParams = new FormParams()
            {
                Title = "Configuração de Metas e Comissões por Função",
                MainDatasource = "@UPD_IT_FUNCOES",
                ExtraDatasources = new List<string>() {
                    "@UPD_IT_METAS",
                },
                BusinessObjectId = "UPD_IT_FUNCOESO",
                BrowseByComp = "Code",
                Focus = "U_nome",

                //Definição de tamanho e posição do Form
                Bounds = new Bounds(){
                    Top = 60,
                    Left = 480,
                    Width = 660,
                    Height = 470
                },

                #region :: Layout Componentes

                Linhas = new Dictionary<string,int>(){
                    {"hd01", 100},
                    {"Code", 10},{"U_ativo", 16},{"U_nome", 40},{"U_mapfield", 18},{"U_percent",16},
                    {"space", 85},{"U_comissao",15},
                    {"Tabs", 100},
                },

                Tabs = new tabParams
                {
                    Height = 320,
                    Tabs = new Dictionary<String, Dictionary<String, Int32>>(){
                        {"Metas" , new Dictionary<String,Int32>(){
                            {"hd02", 100},
                            {"mtxMetas", 100},
                            {"space",65},{"btnAddMeta", 20}, {"btnRmvMeta",15},
                        }},
                        {"Comissões" , new Dictionary<String,Int32>(){
                            {"hd03", 100},
                            {"mtxComissa", 100},
                            {"space",65},{"btnAddCom", 20}, {"btnRmvCom",15},
                        }},
                        {"Observações" , new Dictionary<String,Int32>(){
                            {"U_obs",100},
                        }},
                    },
                },

                Buttons = new Dictionary<string, int>(){
                    {"1", 20},{"2", 20},{"space", 40},
                },

                #endregion


                #region :: Propriedades Componentes

                Controls = new Dictionary<string,CompDefinition>(){
                    
                    #region :: Cabeçalho

                    {"hd01", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_RECTANGLE,
                        Label = "Definição de Função",
                        Height = 1,
                    }},
                    {"Code", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Código",
                        BindTo = "Code",
                        Enabled = false,
                    }},
                    {"U_ativo", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                        Label = "Participa de Comissão?",
                        BindTo = "U_ativo",
                        PopulateItens = new Dictionary<string,string>(){
                            {"S","Sim"},
                            {"N","Não"},
                        }
                    }},
                    {"U_nome", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Nome da Função",
                        BindTo = "U_nome",
                    }},
                    {"U_mapfield", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Coluna Mapeada",
                        BindTo = "U_mapfield",
                    }},
                    {"U_percent", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Porcentagem Ata",
                        BindTo = "U_percent",
                    }},
                    {"U_comissao", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Comissão Fixa (%)",
                        BindTo = "U_comissao",
                    }},

                    #endregion


                    #region :: Aba de Metas

                    {"hd02", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_RECTANGLE,
                        Label = "Configuração de Metas da Função",
                        Height = 1,
                    }},
                    {"mtxMetas", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_MATRIX,
                        Height = 250
                    }},
                    {"btnAddMeta", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Adicionar Meta"
                    }},
                    {"btnRmvMeta", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Remover"
                    }},

                    #endregion


                    #region :: Aba de Comissões

                    {"hd03", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_RECTANGLE,
                        Label = "Configuração de Comissões da Função",
                        Height = 1,
                    }},
                    {"mtxComissa", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_MATRIX,
                        Height = 250
                    }},
                    {"btnAddCom", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Adicionar Comissão"
                    }},
                    {"btnRmvCom", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Remover"
                    }},

                    #endregion


                    #region :: Aba de Observações

                    {"U_obs", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EXTEDIT,
                        Label = "Observações",
                        BindTo = "U_obs",
                        Height = 250
                    }},

                    #endregion

                    
                    #region :: Botões Padrões
 
                    {"1", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                    }},
                    {"2", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
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
        public void FrmConfigMetasOnFormOpen(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void FrmConfigMetasOnFormClose(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            
            // Atualiza o grid de listagem
            if (this.Oppener.GetType().Name == "FrmListaConfigMetas" && ((FrmListaConfigMetas)this.Oppener).SapForm != null)
            {
                ((FrmListaConfigMetas)this.Oppener).RefreshListagem();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void FrmConfigMetasOnDataSave(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            if( this.MetasValidas() )
            {
                this.CopyTableToDatasource(this.mtxMetas["datatable"], this.mtxMetas["dbdatasource"]);
                this.CopyTableToDatasource(this.mtxComissoes["datatable"], this.mtxComissoes["dbdatasource"]);
                
                BubbleEvent = true;
            }
            else
            {
                this.Addon.StatusErro("Não é possível criar um período de Meta com anos distintos.",true);
                BubbleEvent = false;
            }
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void FrmConfigMetasOnRefresh(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this.RefreshMtxMetas();
            this.RefreshMtxComissa();
        }

        #endregion


        #region :: Matriz de Metas

        /// <summary>
        /// Matriz de Comunicação
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void mtxMetasOnCreate(SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            //Declarações básicas da matriz.
            this.mtxMetas["dbdatasource"]  = "@UPD_IT_METAS";
            this.mtxMetas["datatable"]     = "METAS";
            this.mtxMetas["nome"]          = evObj.ItemUID;
            this.mtxMetas["sql"]           = 
                "SELECT " +
	            "   * " +
                "FROM [@UPD_IT_METAS] " +
                "WHERE 1 = 1 ";
            this.mtxMetas["order"]         = " ORDER BY U_dtinicio DESC, U_dtfim DESC ";

            this.mtxMetas["specific"] = this.SetupMatrix(evObj.ItemUID, this.mtxMetas["datatable"], new List<ColumnDefinition>() { 
                new ColumnDefinition() { Width = 3 ,        Id = "hash",        Caption = "#", Bind = false,    Enabled = false},
                new ColumnDefinition() { Percent = 50,      Id = "U_empid",      Caption = "Colaborador", 
                    Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                    PopulateSQL = "SELECT empID, firstName + ' ' + lastName FROM OHEM WHERE Active = 'Y' ORDER BY firstName, lastName"
                },
                new ColumnDefinition() { Percent = 12,      Id = "U_dtinicio",  Caption = "Data Inicial",   DisplayDesc = false},
                new ColumnDefinition() { Percent = 12,      Id = "U_dtfim",     Caption = "Data Final",        DisplayDesc = false},
                new ColumnDefinition() { Percent = 12,      Id = "U_meta",      Caption = "Meta"},
                new ColumnDefinition() { Percent = 15,      Id = "U_acelera",      Caption = "Tipo de Aceleração", Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, 
                    PopulateItens = ((Addon)this.Addon).tipos_aceleracao
                },

                //colunas invisiveis
                new ColumnDefinition() { Id = "Code", Visible = false},
                new ColumnDefinition() { Id = "LineId", Visible = false},
            }, true, this.mtxMetas["sql"] + " AND 1 = 2 ");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnAddMetaOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            //adicionando uma linha.
            this.InsertOnMatrix(this.mtxMetas["nome"], new Dictionary<string, dynamic>()
            {
            });
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnRmvMetaOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this.DeleteOnMatrix(this.mtxMetas["nome"]);
        }

        /// <summary>
        /// 
        /// </summary>
        public void RefreshMtxMetas()
        {
            string code = this.GetValue(this.FormParams.MainDatasource,"Code");
            string sql = this.mtxMetas["sql"] + " AND Code = '" + code + "' " + this.mtxMetas["order"];
            
            this.RefreshMatrix(this.mtxMetas["nome"], this.mtxMetas["datatable"],sql);
        }

        #endregion


        #region :: Matriz de Comissões

        /// <summary>
        /// Matriz de Comunicação
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void mtxComissaOnCreate(SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            //Declarações básicas da matriz.
            this.mtxComissoes["dbdatasource"]   = "@UPD_IT_COMISSAO";
            this.mtxComissoes["datatable"]      = "COMISSAO";
            this.mtxComissoes["nome"]           = evObj.ItemUID;
            this.mtxComissoes["sql"]            =
                "SELECT " +
                "   * " +
                "FROM [@UPD_IT_COMISSAO] " +
                "WHERE 1 = 1 ";
            this.mtxComissoes["order"] = "  ";

            this.mtxComissoes["specific"] = this.SetupMatrix(evObj.ItemUID, this.mtxComissoes["datatable"], new List<ColumnDefinition>() { 
                new ColumnDefinition() { Width = 3 ,        Id = "hash",        Caption = "#", Bind = false,    Enabled = false},
                new ColumnDefinition() { Percent = 20,      Id = "U_overcota",      Caption = "Limite Over Cota", Type = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX},
                new ColumnDefinition() { Percent = 23,      Id = "U_piso",      Caption = "Limite Mínimo (%)"},
                new ColumnDefinition() { Percent = 23,      Id = "U_teto",      Caption = "Limite Máximo (%)"},
                new ColumnDefinition() { Percent = 23,      Id = "U_comissao",  Caption = "Comissão (%)"},
                
                //colunas invisiveis
                new ColumnDefinition() { Id = "Code", Visible = false},
                new ColumnDefinition() { Id = "LineId", Visible = false},
            }, true, this.mtxComissoes["sql"] + " AND 1 = 2 ");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnAddComOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            //adicionando uma linha.
            this.InsertOnMatrix(this.mtxComissoes["nome"], new Dictionary<string, dynamic>()
            {
            });
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnRmvComOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this.DeleteOnMatrix(this.mtxComissoes["nome"]);
        }

        /// <summary>
        /// 
        /// </summary>
        public void RefreshMtxComissa()
        {
            string code = this.GetValue(this.FormParams.MainDatasource, "Code");
            string sql = this.mtxComissoes["sql"] + " AND Code = '" + code + "' " + this.mtxComissoes["order"];

            this.RefreshMatrix(this.mtxComissoes["nome"], this.mtxComissoes["datatable"], sql);
        }

        #endregion


        #region :: Regras de Negócio

        /// <summary>
        /// Método que valida as datas do matriz de Metas.
        /// Não pode ter meta com perídoo em que os anos são diferentes.
        /// </summary>
        /// <returns></returns>
        public bool MetasValidas()
        {
            bool res = true;

            SAPbouiCOM.DataTable dt = this.SapForm.DataSources.DataTables.Item(this.mtxMetas["datatable"]);
            SAPbouiCOM.Matrix mtx = this.mtxMetas["specific"];
            mtx.FlushToDataSource();

            for (int i = 0; i < dt.Rows.Count && !dt.IsEmpty; i++ )
            {
                DateTime dt_inicio  = dt.GetValue("U_dtinicio",i);
                DateTime dt_fim     = dt.GetValue("U_dtfim", i);

                if (dt_inicio.Year != dt_fim.Year)
                {
                    res = false;
                    break;
                }
            }

            return res;
        }

        #endregion

    }
}
