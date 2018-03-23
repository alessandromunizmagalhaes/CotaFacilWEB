using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TShark;

namespace ITOne
{
    class FrmAddParticipante : TShark.Forms
    {
        public FrmAddParticipante(Addon addOn, Dictionary<string, dynamic> ExtraParams = null): base(addOn, ExtraParams)
        {
            //Define o id do form como o nome da classe
            this.FormId = "FrmAddParticipante";

            //Define as configurações do form
            this.FormParams = new FormParams()
            {
                Title = "Adicionar novo Participante",

                //Definição de tamanho e posição do Form
                Bounds = new Bounds(){
                    Top = 110,
                    Left = 445,
                    Width = 750,
                    Height = 200
                },

                #region :: Layout Componentes

                Linhas = new Dictionary<string,int>(){
                    {"hd01", 100},
                    {"lead", 15},{"pedido", 15},{"space",70},
                    {"space1", 100},
                    {"hd02", 100},
                    {"U_split", 35},{"U_empid", 28},{"U_funcao", 25},{"U_percom", 12},
                },

                Buttons = new Dictionary<string, int>(){
                    {"btnInserir", 20},{"btnFechar", 20},{"space", 60},
                },

                #endregion


                #region :: Propriedades Componentes

                Controls = new Dictionary<string,CompDefinition>(){
                    
                    #region :: Cabeçalho
                    
                    {"hd01", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_RECTANGLE,
                        Label = "Identificação",
                        Height = 1,
                    }},
                    {"lead", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Nº Lead",
                        Enabled = false,
                    }},
                    {"pedido", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Pedido de Venda",
                        Enabled = false,
                    }},
                    {"hd02", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_RECTANGLE,
                        Label = "Adicionar novo Participante",
                        Height = 1,
                    }},
                    {"U_split", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                        Label = "Participante de Origem",
                    }},
                    {"U_empid", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                        Label = "Novo Participante",
                        PopulateSQL = "SELECT empID, firstName + ' ' + lastName FROM OHEM WHERE Active = 'Y' ORDER BY firstName, lastName",
                    }},
                    {"U_funcao", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                        Label = "Função",
                        PopulateSQL = "SELECT Code, U_nome FROM [@UPD_IT_FUNCOES] WHERE U_ativo = 'S' ORDER BY U_nome",
                    }},
                    {"U_percom", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Comissão (%)",
                        UserDataType = SAPbouiCOM.BoDataType.dt_PRICE
                    }},

                    #endregion

                    
                    #region :: Botões Padrões
 
                    {"btnInserir", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Inserir"
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
        public void FrmAddParticipanteOnFormOpen(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            Dictionary<string,dynamic> val = new Dictionary<string,dynamic>(){};

            if( this.ExtraParams.ContainsKey("U_docentry"))
            {
                val.Add("pedido", this.ExtraParams["U_docentry"]);
            }
            if (this.ExtraParams.ContainsKey("lead"))
            {
                val.Add("lead", this.ExtraParams["lead"]);
            }

            // Preenchendo com participantes de origem, somente os que participam do pedido de venda.
            string sql_combo_participantes =
                    "SELECT " +
                    "       tb3.empID, tb3.firstName + ' ' + tb3.lastName + ' - ' + tb2.U_nome " +
                    "   FROM [@UPD_IT_PARTICIP] tb1 (NOLOCK) " +
                    "   INNER JOIN [@UPD_IT_FUNCOES] tb2 ON ( tb1.U_funcao = tb2.Code ) " +
                    "   INNER JOIN OHEM tb3 ON ( tb1.U_empid = tb3.empID ) " +
                    "   WHERE U_docentry = " + this.ExtraParams["U_docentry"];
            this.populateCombo("U_split", this.FormId, sql_combo_participantes);

            // buscando qual é o código do vendedor/gerente de contas do pedido.
            string sql =
                "SELECT empID FROM ORDR (nolock) tb1 INNER JOIN OHEM tb2 ON ( tb1.SlpCode = tb2.salesPrson )  WHERE DocEntry = " + this.ExtraParams["U_docentry"];
            SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);

            if( rs.RecordCount > 0 )
            {
                int empid = rs.Fields.Item("empID").Value;
                val.Add("U_split", empid.ToString());
            }

            this.UpdateUserDataSource(val);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void FrmAddParticipanteOnFormClose(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
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
        public void btnInserirOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            this.AdicionarParticipante();
        }

        #endregion


        #region :: Regras de Negócio

        /// <summary>
        /// 
        /// </summary>
        public void AdicionarParticipante()
        {
            string split = this.GetValue("U_split");
            if (String.IsNullOrEmpty(split))
            {
                this.Addon.StatusErro("Defina o Participante Relacionado.", true);
                this.GetItem("U_split").Click();
                return;
            }
            
            string empid = this.GetValue("U_empid");
            if (String.IsNullOrEmpty(empid))
            {
                this.Addon.StatusErro("Defina o Novo participante.", true);
                this.GetItem("U_empid").Click();
                return;
            }

            string funcao = this.GetValue("U_funcao");
            if (String.IsNullOrEmpty(funcao))
            {
                this.Addon.StatusErro("Defina a função do Novo Participante.", true);
                this.GetItem("U_funcao").Click();
                return;
            }

            string percom = this.GetValue("U_percom").Replace(".", "").Replace(",", ".");
              
            if( this.InsertOnServer("@UPD_IT_PARTICIP", new Dictionary<string, dynamic>()
            {
                {"U_docentry",this.ExtraParams["U_docentry"]},
                {"U_funcao",funcao},
                {"U_empid",empid},
                {"U_percom",percom},
            }) )
            {
                // garantindo que, sempre que criar um SPLIT o empregado de origem vai ser o selecionado.
                string update =
                    "UPDATE [@UPD_IT_PARTICIP] SET U_split = NULL WHERE U_docentry = " + this.ExtraParams["U_docentry"] + ";" +
                    "UPDATE [@UPD_IT_PARTICIP] SET U_split = 'S' WHERE U_docentry = " + this.ExtraParams["U_docentry"] + " AND U_empid = '" + split + "' ";
                SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(update);
                
                this.Addon.StatusInfo("Participante inserido com sucesso.", true);
                ((FrmAtualizaValores)this.Oppener).PesquisarParticipantes();
                this.SapForm.Close();
            }
            else
            {
                this.Addon.StatusErro("Não foi possível inserir participante.", true);
            }
        }

        #endregion

    }
}
