using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TShark;

namespace ITOne
{
    class _FrmListagens : TShark.Forms
    {
        /// <summary>
        /// Form template de listageNS
        /// </summary>
        /// <param name="addOn"></param>
        public _FrmListagens(Addon addOn, Dictionary<string, dynamic> ExtraParams = null)
            : base(addOn, ExtraParams)
        {
            // Default de listagem
            this.ExtraParams["SQL_ORDER_DT_LISTAGEM"] = " Code ";

            this.FormParams = new FormParams()
            {

                //Definição de tamanho e posição do Form
                Bounds = new Bounds()
                {
                    Top = 50,
                    Left = 370,
                    Width = 820,
                    Height = 490
                },

                #region Layout Componentes

                Linhas = new Dictionary<string, int>()
                {
                    {"space", 21}, {"DtDe", 12}, {"DtAte", 12}, {"edFiltro", 40}, {"btnBuscar", 15},
                    {"hd01", 100},
                    {"mtxLista", 100}
                },

                Buttons = new Dictionary<string, int>
                {
                    {"btnClose", 20}, {"space", 50}, {"btnNovo", 30}
                },

                #endregion

                #region Propriedade dos Componentes

                Controls = new Dictionary<string, CompDefinition>()
                {
                    {"DtDe", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "De",
                        UserDataType = SAPbouiCOM.BoDataType.dt_DATE,
                        onKeyDownHandler = "OnKeyDownRefresh"
                    }},
                    {"DtAte", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Até",
                        UserDataType = SAPbouiCOM.BoDataType.dt_DATE,
                        onKeyDownHandler = "OnKeyDownRefresh"
                    }},
                    {"edFiltro", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        UserDataSize = 200,
                        Label = "Digite sua Pesquisa:",
                        onKeyDownHandler = "OnKeyDownRefresh"
                    }},

                    {"btnBuscar", new CompDefinition(){
                         Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                         Caption = "Pesquisar",
                         marginTop = 10,
                     }},
                     
                    {"hd01", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_RECTANGLE,
                        Label = " Cadastrados",
                        Height = 1,
                    }},
                    {"mtxLista", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_MATRIX,
                        Enabled = false,
                        Height = 370
                    }},

                    {"btnClose", new CompDefinition(){
                         Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                         Caption = "Fechar"
                    }},

                    {"btnNovo", new CompDefinition(){
                         Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                         Caption = "Cadastrar Novo"
                    }}
                },

                #endregion

            };
        }


        #region :: Métodos onCreate

        #endregion


        #region :: Eventos do formulário

        #endregion


        #region :: Eventos de componentes

        /// <summary>
        /// Abre um form de motorista de acordo com a linha da matrix clicada
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void mtxListaOnDblClick(SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this.Addon.OpenFormUDOFind(
                this.ExtraParams["FORM_POPUP"],   // Formulário de popup definido em classes filhas 
                this.getCellValue("mtxLista", this.FormId, "Code"), this
            );
        }

        /// <summary>
        /// Abre form de motorista em modo de inserção
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public virtual void btnNovoOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this.Addon.OpenFormUDOAdd(
                this.ExtraParams["FORM_POPUP"],   // Formulário de popup definido em classes filhas 
                this
            );
        }

        /// <summary>
        /// OnKeydown dos campos de pesquisa
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnKeyDownRefresh(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            // Pesquisa ao teclar ENTER
            if(evObj.CharPressed == 13)
            {
                this.GetItem("btnBuscar").Click();
            }
        }

        /// <summary>
        /// Botão de filtragem
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnBuscarOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this.RefreshListagem();
        }

        #endregion


        #region :: Regras de Negócio

        /// <summary>
        /// Executa refresh em matrix de listagem 
        /// </summary>
        new public void RefreshListagem()
        {
            this.RefreshMatrix("mtxLista", "DT_LISTAGEM", this.ParseSQL());
        }

        /// <summary>
        /// Processa o SQL da pesquisa
        /// </summary>
        /// <returns></returns>
        public string ParseSQL()
        {
            string sql = this.ExtraParams["SQL_DT_LISTAGEM"];

            // Filtra a partir de:
            string de = this.GetValue("dtDe");
            if(!String.IsNullOrEmpty(de) && this.ExtraParams.ContainsKey("SQL_WHERE_DT_DE") && !String.IsNullOrEmpty(this.ExtraParams["SQL_WHERE_DT_DE"]))
            {
                sql += Convert.ToString(this.ExtraParams["SQL_WHERE_DT_DE"]).Replace("_DT_DE_", de);
            }

            // Filtra até data:
            string ate = this.GetValue("dtAte");
            if(!String.IsNullOrEmpty(ate) && this.ExtraParams.ContainsKey("SQL_WHERE_DT_ATE") && !String.IsNullOrEmpty(this.ExtraParams["SQL_WHERE_DT_ATE"]))
            {
                sql += Convert.ToString(this.ExtraParams["SQL_WHERE_DT_ATE"]).Replace("_DT_ATE_", ate);
            }

            // Filtra texto:
            string termo = this.GetValue("edFiltro").Trim();
            if(!String.IsNullOrEmpty(termo) && this.ExtraParams.ContainsKey("SQL_WHERE_FILTRO") && !String.IsNullOrEmpty(this.ExtraParams["SQL_WHERE_FILTRO"]))
            {
                sql += Convert.ToString(this.ExtraParams["SQL_WHERE_FILTRO"]).Replace("_TEXTO_", termo);
            }

            if(this.ExtraParams.ContainsKey("SQL_ORDER_DT_LISTAGEM"))
            {
                sql += " ORDER BY " + this.ExtraParams["SQL_ORDER_DT_LISTAGEM"]; // Ordenação de dados definido em classes filhas 
            }

            return sql;
        }


        public override void PostParams(Dictionary<string, dynamic> oppener_params)
        {
            this.Addon.StatusInfo(oppener_params["teste"]);
        }

        #endregion

    }
}
