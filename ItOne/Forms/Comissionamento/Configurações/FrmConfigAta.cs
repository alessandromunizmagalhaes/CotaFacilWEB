using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TShark;

namespace ITOne
{
    class FrmConfigAta : TShark.Forms
    {

        public FrmConfigAta(Addon addOn, Dictionary<string, dynamic> ExtraParams = null): base(addOn, ExtraParams)
        {
            //Define o id do form como o nome da classe
            this.FormId = "FrmConfigAta";

            //Define as configurações do form
            this.FormParams = new FormParams()
            {
                Title = "Configuração de Replicação de Ata",
                MainDatasource = "@UPD_IT_CONFIG_ATA",
                ExtraDatasources = new List<string>() {
                    "@UPD_IT_METAS",
                },
                BusinessObjectId = "UPD_IT_CONFIG_ATAO",
                BrowseByComp = "Code",
                Focus = "U_desc",

                //Definição de tamanho e posição do Form
                Bounds = new Bounds(){
                    Top = 80,
                    Left = 430,
                    Width = 660,
                    Height = 320
                },

                #region :: Layout Componentes

                Linhas = new Dictionary<string,int>(){
                    {"hd01", 100},
                    {"Code", 10},{"U_desc", 70},{"U_percent", 20},
                    {"U_obs", 100},
                },

                Buttons = new Dictionary<string, int>(){
                    {"1", 20},{"2", 20},{"space", 40},
                },

                #endregion


                #region :: Propriedades Componentes

                Controls = new Dictionary<string,CompDefinition>(){

                    {"hd01", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_RECTANGLE,
                        Label = "Configuração de Replicação de Ata",
                        Height = 1,
                    }},
                    {"Code", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Código",
                        BindTo = "Code",
                        Enabled = false,
                    }},
                    {"U_desc", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Descrição",
                        BindTo = "U_desc",
                    }},
                    {"U_percent", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Porcentagem de Comissão",
                        BindTo = "U_percent",
                    }},
                    {"U_obs", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EXTEDIT,
                        Label = "Observações",
                        BindTo = "U_obs",
                        Height = 200
                    }},

                    
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
        public void FrmConfigAtaOnFormOpen(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void FrmConfigAtaOnFormClose(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            
            // Atualiza o grid de listagem
            if (this.Oppener.GetType().Name == "FrmListaConfigAta" && ((FrmListaConfigAta)this.Oppener).SapForm != null)
            {
                ((FrmListaConfigAta)this.Oppener).RefreshListagem();
            }
        }

        #endregion


        #region :: Regras de Negócio



        #endregion

    }
}
