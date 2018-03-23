using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TShark;

namespace ITOne
{
    class FrmAtualizaComissaoParticipante : TShark.Forms
    {

        public FrmAtualizaComissaoParticipante(Addon addOn, Dictionary<string, dynamic> ExtraParams = null): base(addOn, ExtraParams)
        {
            //Define o id do form como o nome da classe
            this.FormId = "FrmAtualizaComissaoParticipante";

            //Define as configurações do form
            this.FormParams = new FormParams()
            {
                Title = "Atualização de Valores de Comissão",

                //Definição de tamanho e posição do Form
                Bounds = new Bounds(){
                    Top = 100,
                    Left = 420,
                    Width = 650,
                    Height = 110
                },

                #region :: Layout Componentes

                Linhas = new Dictionary<string,int>(){
                    {"hd01", 100},
                    {"U_funcao", 35},{"U_empid", 35},{"U_vlcom", 15},{"U_percom", 15},
                },

                Buttons = new Dictionary<string, int>(){
                    {"btnUpdate", 20},{"btnFechar", 20},{"space", 60},
                },

                #endregion


                #region :: Propriedades Componentes

                Controls = new Dictionary<string,CompDefinition>(){
                    
                    #region :: Cabeçalho

                    {"hd01", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_RECTANGLE,
                        Label = "Atualização de Comissão do Participante",
                        Height = 1,
                    }},
                    {"U_funcao", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                        Label = "Função",
                        PopulateSQL = "SELECT Code, U_nome FROM [@UPD_IT_FUNCOES] WHERE U_ativo = 'S' ORDER BY U_nome",
                        Enabled = false,
                    }},
                    {"U_empid", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                        Label = "Colaborador",
                        PopulateSQL = "SELECT empID, firstName + ' ' + lastName FROM OHEM WHERE Active = 'Y' ORDER BY firstName, lastName",
                        Enabled = false,
                    }},
                    {"U_vlcom", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Comissão",
                        UserDataType = SAPbouiCOM.BoDataType.dt_PRICE
                    }},
                    {"U_percom", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Comissão (%)",
                        UserDataType = SAPbouiCOM.BoDataType.dt_PRICE
                    }},

                    #endregion

                    
                    #region :: Botões Padrões
 
                    {"btnUpdate", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Atualizar"
                    }},
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
        public void FrmAtualizaComissaoParticipanteOnFormOpen(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            int linha = this.ExtraParams["linha"];

            if( linha > 0 )
            {
                this.UpdateUserDataSource(new Dictionary<string, dynamic>()
                {
                    {"U_funcao",this.ExtraParams["funcao"]},
                    {"U_empid",this.ExtraParams["empid"]},
                });
            }
            else
            {
                this.Addon.StatusErro("Não foi possível encontrar a linha do participante. O Form será fechado.", true);
                this.SapForm.Close();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void FrmAtualizaComissaoParticipanteOnFormClose(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
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
        public void btnUpdateOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            this.AtualizarComissaoParticipantes();
        }

        #endregion


        #region :: Regras de Negócio

        /// <summary>
        /// 
        /// </summary>
        public void AtualizarComissaoParticipantes()
        {
            string vlcom    = this.GetValue("U_vlcom").Replace(".","").Replace(",",".");
            string percom   = this.GetValue("U_percom").Replace(".", "").Replace(",", ".");
            
            int linha = this.ExtraParams["linha"];

            ((UserFields)this.Oppener).AtualizaComissaoParticipante(linha, new Dictionary<string, dynamic>() { 
                {"U_vlcom",vlcom},
                {"U_percom",percom},
            });

            this.SapForm.Close();
        }

        #endregion

    }
}
