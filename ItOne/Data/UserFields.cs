using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TShark;

namespace ITOne
{
    class UserFields : TShark.UserFields
    {
        public UserFields(FastOne addOn): base(addOn)
        {
            this.recreate = false;            
        }

        #region :: Form 10016 - CFL na Oportunidade

        public void sapForm10016()
        {
            string form = "10016";
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_CLICK, form, "5", "Item5OnClick", false);
        }

        /// <summary>
        /// Antes de dar o choose from list
        /// pega qual é a linha da oportunidade que está sendo dado o choose
        /// e qual é o código da oportunidade e salva em uma propriedade do Addon pra ser trabalhado quando abre o pedido de venda.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void Item5OnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            ((Addon)this.Addon).active = true;
        }

        #endregion


        #region :: Form 320 - Oportunidade de Venda

        public void sapForm320()
        {
            userFieldsParams["320"] = new List<userFieldsParams>(){
                /*
                new userFieldsParams(){
                    fieldId = "UPD_IT_CONFIG_ATA",
                    tableId = "OPR1",
                    field = new fieldParams(){
                        descricao = "Config Ata",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                        size = 30
                    },
                    comp = new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                        PopulateSQL = "SELECT Code, U_desc FROM [@UPD_IT_CONFIG_ATA] ORDER BY U_desc ASC"
                    }
                },*/
                new userFieldsParams(){
                    fieldId = "UPD_IT_TIME",
                    tableId = "OPR2",
                    field = new fieldParams(){
                        descricao = "Time",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                        size = 30
                    },
                    comp = new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                        PopulateSQL = "SELECT empID, firstName + ' ' + lastName FROM OHEM WHERE Active = 'Y' ORDER BY firstName, lastName"
                    }
                },
                new userFieldsParams(){
                    fieldId = "UPD_IT_PERCENT",
                    tableId = "OPR2",
                    field = new fieldParams(){
                        descricao = "Porcentagem Parceiro",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Float,
                        subtipo = SAPbobsCOM.BoFldSubTypes.st_Percentage,
                    },
                    comp = new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                    }
                },

                #region :: Criação de Componentes

                
                new userFieldsParams(){
                    itemRef = "2",
                    fieldId = "btnSPLIT",
                    comp = new CompDefinition(){    
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "SPLIT",
                        FromPane = 0,
                        ToPane = 0,
                        Bounds = new Bounds(){ Top = 0, Left = 200, Width = 150 },
                    }
                },

                #endregion

            };
            
            string form = "320";
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST, form, "56", "Item56OnChoose", false);
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST, form, "56", "Item56OnAfterChoose", true);
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT, form, "91", "MtxItemLeadOnChange", false);

            // Registra o evento para colocação dos comps nos formulários:
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_FORM_LOAD, form, form, "Form320OnLoad", false);
        }



        /// <summary>
        /// atualiza Id da oportunidade
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void Form320OnLoad(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                SAPbouiCOM.Form frm = this.Addon.SBO_Application.Forms.GetFormByTypeAndCount(evObj.FormType, evObj.FormTypeCount);
                SAPbouiCOM.Matrix mtx = frm.Items.Item("91").Specific;
                SAPbouiCOM.Column col = mtx.Columns.Item("U_UPD_IT_TIME");
                this.populateColumn("91", ref col, "SELECT empID, firstName + ' ' + lastName FROM OHEM WHERE Active = 'Y' ORDER BY firstName, lastName");

            } catch(Exception e)
            {
                this.Addon.StatusErro(e.Message);
            }
        }

        /// <summary>
        /// Antes de dar o choose from list
        /// pega qual é a linha da oportunidade que está sendo dado o choose
        /// e qual é o código da oportunidade e salva em uma propriedade do Addon pra ser trabalhado quando abre o pedido de venda.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void Item56OnChoose(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Matrix mtx = ((SAPbouiCOM.Matrix)this.GetItem("56", this.Addon.SBO_Application.Forms.ActiveForm).Specific);

            if( evObj.ColUID == "15" && mtx.GetCellSpecific("14",evObj.Row).Value == "17" )
            {
                SAPbouiCOM.Form frm = this.Addon.SBO_Application.Forms.GetFormByTypeAndCount(evObj.FormType, evObj.FormTypeCount);
                if (frm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    string operacao = frm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE ? "Adicione" : "Atualize";
                    this.Addon.StatusErro(operacao + " a Oportunidade de Venda para associá-la a um Pedido de Venda", true);
                    BubbleEvent = false;
                    return;
                }
                
                int opprid = 0,lead_ata = 0, slp_code = 0;

                string str_opprid = frm.DataSources.DBDataSources.Item("OOPR").GetValue("OpprId",0);
                Int32.TryParse(str_opprid, out opprid);
                /*
                string str_lead_ata = mtx.GetCellSpecific("U_UPD_IT_LEAD_ATA", evObj.Row).Value;
                Int32.TryParse(str_lead_ata, out lead_ata);
                */
                string str_slp_code = this.GetValue("OOPR", "SlpCode");
                Int32.TryParse(str_slp_code, out slp_code);

                ((Addon)this.Addon).lineid      = evObj.Row - 1;
                ((Addon)this.Addon).opprid      = opprid;
                ((Addon)this.Addon).lead_ata    = lead_ata;
            }
        }

        /// <summary>
        /// atualiza Id da oportunidade
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void MtxItemLeadOnChange(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            // Se o combo foi na coluna Fornecedor
            if( evObj.ColUID == "1")
            {
                SAPbouiCOM.Form frm = this.Addon.SBO_Application.Forms.GetFormByTypeAndCount(evObj.FormType, evObj.FormTypeCount);
                SAPbouiCOM.Matrix mtx = frm.Items.Item("91").Specific;
                string time = mtx.GetCellSpecific("U_UPD_IT_TIME", evObj.Row).Value;
                string slpcode = frm.DataSources.DBDataSources.Item("OOPR").GetValue("SlpCode",0);

                if (String.IsNullOrEmpty(time) && !String.IsNullOrEmpty(slpcode))
                {
                    // pegando o gerente do funcionário
                    string sql =
                    " SELECT manager FROM OHEM (NOLOCK) WHERE salesprson = " + slpcode;
                    SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rs.DoQuery(sql);

                    int gerente = rs.Fields.Item("manager").Value;

                    if (rs.RecordCount > 0 && gerente > 0)
                    {
                        mtx.SetCellWithoutValidation(evObj.Row, "U_UPD_IT_TIME", gerente.ToString());
                    }
                }
            }
        }

        /// <summary>
        /// Depois de dar o choosefromlist
        /// depois que finalizou o choose, tem que limpar a propriedade
        /// senão, quando for aberto o pedido, a propriedade ainda vai existir 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void Item56OnAfterChoose(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            ((Addon)this.Addon).resetarDadosOppr();
        }


        /// <summary>
        /// Antes de dar o choose from list
        /// pega qual é a linha da oportunidade que está sendo dado o choose
        /// e qual é o código da oportunidade e salva em uma propriedade do Addon pra ser trabalhado quando abre o pedido de venda.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnSPLITOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            this.Addon.OpenForm("FrmAtualizaValores");
        }

        #endregion


        #region :: Form 139 - Pedido de Venda

        public void sapForm139()
        {
            string form = "139";
            int pane    = 2993;
            

            userFieldsParams[form] = new List<userFieldsParams>(){

                #region :: Aba de Comissionamento


                #region :: Cabeçalho

                new userFieldsParams(){
                    itemRef = "112",
                    fieldId = "tabComiss",
                    comp = new CompDefinition(){    
                        Type = SAPbouiCOM.BoFormItemTypes.it_FOLDER,
                        Caption = "Comissionamento",
                        Pane = pane,
                    }
                },

                new userFieldsParams(){
                    itemRef = "38",
                    fieldId = "hdComiss",
                    comp = new CompDefinition(){    
                        Type = SAPbouiCOM.BoFormItemTypes.it_RECTANGLE,
                        Label = "Dados do Comissionamento",
                        FromPane = pane,
                        ToPane = pane,
                        Bounds = new Bounds(){ Height = 1, Width = 500 }
                    }
                },

                #endregion


                #region :: Matriz

                new userFieldsParams(){
                    itemRef = "hdComiss",
                    fieldId = "hdParticip",
                    comp = new CompDefinition(){    
                        Type = SAPbouiCOM.BoFormItemTypes.it_RECTANGLE,
                        Label = "Participantes do Comissionamento",
                        FromPane = pane,
                        ToPane = pane,
                        Bounds = new Bounds(){ Height = 1, Width = 500, Top = 56 }
                    }
                },
                new userFieldsParams(){
                    itemRef = "hdParticip",
                    fieldId = "mtxPart",
                    comp = new CompDefinition(){    
                        Type = SAPbouiCOM.BoFormItemTypes.it_MATRIX,
                        FromPane = pane,
                        ToPane = pane,
                        Bounds = new Bounds(){ Height = 127, Width = 500, Top = 7 },
                        Columns = new columnParams(){
                            Widths = new List<int>(){
                                //4,30,33,14,15,1
                                4,46,46,1,1,1
                            }
                        },
                    }
                },
                new userFieldsParams(){
                    itemRef = "mtxPart",
                    fieldId = "btnAddPart",
                    comp = new CompDefinition(){    
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Adicionar Participante",
                        FromPane = pane,
                        ToPane = pane,
                        Bounds = new Bounds(){ Top = 30, Left = 515, Width = 120 },
                    }
                },
                new userFieldsParams(){
                    itemRef = "btnAddPart",
                    fieldId = "btnRmvPart",
                    comp = new CompDefinition(){    
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Remover",
                        FromPane = pane,
                        ToPane = pane,
                        Bounds = new Bounds(){ Top = 30, Width = 120 },
                    }
                },

                #endregion


                #endregion


                #region :: Criação de Campos de Usuário

                new userFieldsParams(){
                    fieldId = "UPD_IT_LEAD",
                    tableId = "ORDR",
                    field = new fieldParams(){
                        descricao = "Número Lead",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                        size = 11,
                    },
                    itemRef = "hdComiss",
                    comp = new CompDefinition(){    
                        Id = "edLead",
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Nº do Lead",
                        FromPane = pane,
                        ToPane = pane,
                        Bounds = new Bounds(){ Width = 100, Top = 17 },
                        onKeyDownHandler = "edLeadOnKeyDown"
                    }
                },
                new userFieldsParams(){
                    fieldId = "UPD_IT_RENDA",
                    tableId = "ORDR",
                    field = new fieldParams(){
                        descricao = "Valor Renda",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Float,
                        subtipo = SAPbobsCOM.BoFldSubTypes.st_Price, 
                    },
                    itemRef = "edLead",
                    comp = new CompDefinition(){    
                        Id = "edRenda",
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Renda Comissionável",
                        FromPane = pane,
                        ToPane = pane,
                        Bounds = new Bounds(){ Left = 110, Width = 110 }
                    }
                },
                /*new userFieldsParams(){
                    fieldId = "UPD_IT_TIME",
                    tableId = "ORDR",
                    field = new fieldParams(){
                        descricao = "Time/Distrito",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                    },
                    itemRef = "edRenda",
                    comp = new CompDefinition(){    
                        Id = "cbTime",
                        Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                        PopulateSQL = "SELECT empID, firstName + ' ' + lastName FROM OHEM WHERE Active = 'Y' ORDER BY firstName, lastName",
                        Label = "Time Comercial",
                        FromPane = pane,
                        ToPane = pane,
                        Bounds = new Bounds(){ Left = 130, Width = 140 }
                    }
                },*/

                new userFieldsParams(){
                    fieldId = "UPD_IT_LEAD_ATA",
                    tableId = "ORDR",
                    field = new fieldParams(){
                        descricao = "Lead Ata",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                        size = 11,
                    },
                    itemRef = "edRenda",
                    comp = new CompDefinition(){    
                        Id = "edLeadAta",
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Replicar Ata do Lead:",
                        FromPane = pane,
                        ToPane = pane,
                        Bounds = new Bounds(){ Left = 150, Width = 110 },
                        onKeyDownHandler = "edLeadAtaOnKeyDown"
                    }
                },
                new userFieldsParams(){
                    fieldId = "UPD_IT_NOTA",
                    tableId = "ORDR",
                    field = new fieldParams(){
                        descricao = "DocEntry Nota",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                    },
                },
                new userFieldsParams(){
                    fieldId = "UPD_IT_SERIAL",
                    tableId = "ORDR",
                    field = new fieldParams(){
                        descricao = "Serial Nota",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                    },
                },
                
                //(P)rocessado
                //(F)aturado
                //(R)ealizado
                new userFieldsParams(){
                    fieldId = "UPD_IT_STATUS",
                    tableId = "ORDR",
                    field = new fieldParams(){
                        descricao = "Status Comissionamento",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                        size = 1
                    },
                },
                
                new userFieldsParams(){
                    fieldId = "UPD_IT_PARCEIRO",
                    tableId = "ORDR",
                    field = new fieldParams(){
                        descricao = "Parceiro",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                        size = 50,
                    },
                    itemRef = "edLeadAta",
                    comp = new CompDefinition(){    
                        Id = "edParc",
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = " ",
                        FromPane = pane,
                        ToPane = pane,
                        Bounds = new Bounds(){ Left = 130, Width = 140 },
                        Visible = false
                    }
                },

                new userFieldsParams(){
                    fieldId = "UPD_IT_PERCENT",
                    tableId = "ORDR",
                    field = new fieldParams(){
                        descricao = "Percent Parceiro",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Float,
                        subtipo = SAPbobsCOM.BoFldSubTypes.st_Percentage
                    },
                    itemRef = "edParc",
                    comp = new CompDefinition(){    
                        Id = "edPerc",
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = " ",
                        FromPane = pane,
                        ToPane = pane,
                        Bounds = new Bounds(){ Left = 100, Width = 100 },
                        Visible = false
                    }
                },

                #endregion

            };

            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD, form, form, "sapFormOnRefresh",true);
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, form, form, "OnORDRDataAdd",true);
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE, form, form, "SaveMatrix", true);
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_FORM_DRAW, form, form, "sapFormActivate");
        }


        #region :: Métodos do formulário

        /// <summary>
        ///
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void sapFormActivate(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = false;

            this.ImportarParticipantes();
        }

        /// <summary>
        /// atualiza Id da oportunidade
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void sapFormOnRefresh(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                // Refresh matrix
                this.RefreshMtxPart(evObj);
            }
            catch (Exception e)
            {

            }
        }

        /// <summary>
        /// After Insert
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnORDRDataAdd(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = false;

            if (evObj.ActionSuccess)
            {
                this.SaveMatrix(ref evObj, out BubbleEvent);

                string docentry = this.Addon.SBO_Application.Forms.Item(evObj.FormUID).DataSources.DBDataSources.Item("ORDR").GetValue("DocNum", 0);
                SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string update = 
                    " UPDATE ORDR SET U_UPD_IT_STATUS = 'P' WHERE DocNum = " + docentry;
                rs.DoQuery(update);

                BubbleEvent = true;
            }
        }

        /// <summary>
        /// Garante que as alterações sejam salvas no banco.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void SaveMatrix(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = false;

            if (evObj.ActionSuccess)
            {
                this.SaveMatrixToServer("mtxPart", "@UPD_IT_PARTICIP");


                string lead     = this.Addon.SBO_Application.Forms.Item(evObj.FormUID).DataSources.DBDataSources.Item("ORDR").GetValue("U_UPD_IT_LEAD", 0);
                string docentry = this.Addon.SBO_Application.Forms.Item(evObj.FormUID).DataSources.DBDataSources.Item("ORDR").GetValue("DocNum", 0);

                string sql = "SELECT Code FROM [@UPD_IT_PARTICIP] WHERE U_docentry = " + docentry + " AND U_oculto = 'S' ";
                SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                if (!String.IsNullOrEmpty(lead) && rs.RecordCount == 0)
                {
                    this.InserirParticipanteOculto();
                }

                BubbleEvent = true;
            }
        }

        /// <summary>
        /// Carga e refresh da matriz
        /// </summary>
        public void RefreshMtxPart(SAPbouiCOM.BusinessObjectInfo evObj)
        {
            try
            {
                SAPbouiCOM.Form oForm = this.Addon.SBO_Application.Forms.Item(evObj.FormUID);
                string docentry = oForm.DataSources.DBDataSources.Item("ORDR").GetValue("DocNum", 0);
                
                // Condições
                SAPbouiCOM.Conditions conds = new SAPbouiCOM.Conditions();
                
                SAPbouiCOM.Condition cond = conds.Add();
                cond.Alias      = "U_docentry";
                cond.Operation  = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                cond.CondVal    = docentry;

                cond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                SAPbouiCOM.Condition outroCond = conds.Add();
                outroCond.Alias     = "U_oculto";
                outroCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                outroCond.CondVal   = "S";

                // Refresh na matrix
                this.RefreshMatrix(ref oForm, "mtxPart", "@UPD_IT_PARTICIP", conds);
            }
            catch (Exception e)
            {
                this.Addon.StatusAlerta(e.Message);
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void edLeadOnKeyDown(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = false;

            // Só deixa passar se for ENTER ou TAB
            if (evObj.CharPressed != 13 && evObj.CharPressed != 9)
                return;

            int lead = 0;

            SAPbouiCOM.Form frm = this.Addon.SBO_Application.Forms.GetFormByTypeAndCount(evObj.FormType, evObj.FormTypeCount);
            string str_lead = frm.DataSources.DBDataSources.Item("ORDR").GetValue("U_UPD_IT_LEAD", 0);
            
            if( Int32.TryParse(str_lead, out lead) )
            {
                this.BuscarDadosLead(lead);
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void edLeadAtaOnKeyDown(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = false;

            // Só deixa passar se for ENTER ou TAB
            if (evObj.CharPressed != 13 && evObj.CharPressed != 9)
                return;

            int lead = 0;
            
            SAPbouiCOM.Form frm = this.Addon.SBO_Application.Forms.GetFormByTypeAndCount(evObj.FormType, evObj.FormTypeCount);
            string str_lead = frm.DataSources.DBDataSources.Item("ORDR").GetValue("U_UPD_IT_LEAD", 0);

            if (Int32.TryParse(str_lead, out lead))
            {
                this.BuscarDadosLeadAta(lead);
            }
        }

        #endregion


        #region :: Métodos da matriz

        /// <summary>
        /// Criando a matriz de Itens
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void mtxPartOnCreate(SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form frm = this.Addon.SBO_Application.Forms.GetFormByTypeAndCount(evObj.FormType, evObj.FormTypeCount);
            frm.DataSources.DBDataSources.Add("@UPD_IT_PARTICIP");

            this.SetupMatrix(evObj.ItemUID, "@UPD_IT_PARTICIP", new List<ColumnDefinition>(){
                
                new ColumnDefinition(){Id = "hash",         Caption = "#", Enabled = false, Bind = false,},
                new ColumnDefinition(){Id = "U_funcao",     Caption = "Função",
                    Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                    PopulateSQL = "SELECT Code, U_nome FROM [@UPD_IT_FUNCOES] WHERE U_ativo = 'S' ORDER BY U_nome"
                },
                new ColumnDefinition(){Id = "U_empid",      Caption = "Colaborador", 
                    Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                    PopulateSQL = "SELECT empID, firstName + ' ' + lastName FROM OHEM WHERE Active = 'Y' ORDER BY firstName, lastName"
                },

                //campos invisiveis
                new ColumnDefinition(){Id = "U_docentry",   Visible = false, },
                new ColumnDefinition(){Id = "Code",         Visible = false, },
                new ColumnDefinition(){Id = "U_vlcom",      Visible = false, Caption = "Comissão",  },
                new ColumnDefinition(){Id = "U_percom",     Visible = false, Caption = "Comissão (%)", },
                new ColumnDefinition(){Id = "U_ata",        Visible = false, Caption = "(%) Ata", },
                new ColumnDefinition(){Id = "U_gerente",    Visible = false, Caption = "Flag Vendedor", },
            });
        }

        /// <summary>
        /// Criando a matriz de Itens
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void mtxPartOnDblClick(SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)this.GetItem(evObj.ItemUID).Specific;

            string funcao   = mtx.GetCellSpecific("U_funcao", evObj.Row).Value;
            string empid    = mtx.GetCellSpecific("U_empid", evObj.Row).Value;

            this.Addon.OpenForm("FrmAtualizaComissaoParticipante",this, new Dictionary<string, dynamic>()
            {
                {"linha",evObj.Row},
                {"funcao",funcao},
                {"empid",empid},
            });
        }

        /// <summary>
        /// Método responsável pela atualização das comissões do participante.
        /// Utilizando um método dentro de userfields, pois seria dificil passar um this.Oppener pro outro form e ele atualizar uma linha
        /// de uma matriz dentro de um form padrão sap.
        /// </summary>
        /// <param name="linha"></param>
        /// <param name="valores"></param>
        public void AtualizaComissaoParticipante( int linha, Dictionary<string,dynamic> valores )
        {
            this.UpdateOnMatrix("mtxPart",valores,linha);

            if( this.SapForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE )
            {
                this.SapForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            }
            
            // Forçando um click na matriz só pra atualizar os dados
            this.GetItem("mtxPart").Click();
        }

        /// <summary>
        /// Insere um ítem.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnAddPartOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form frm = this.Addon.SBO_Application.Forms.GetFormByTypeAndCount(evObj.FormType, evObj.FormTypeCount);
            dynamic docentry = frm.DataSources.DBDataSources.Item("ORDR").GetValue("DocNum", 0);

            this.InsertOnMatrix("mtxPart", new Dictionary<string, dynamic>(){
                {"U_docentry", docentry}, 
            },frm);
        }

        /// <summary>
        /// Insere um ítem.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnRmvPartOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            
            this.DeleteMatrixOnServer("mtxPart", "@UPD_IT_PARTICIP");
        }

        #endregion


        #region :: Regras de Negócio

        /// <summary>
        /// Importa participantes da aba de Etapa de Negócios (OPR1)
        /// Descobre qual a função de cada participante
        /// Descobre qual o empID de cada participante
        /// </summary>
        public void ImportarParticipantes()
        {
            this.SapForm.Freeze(true);
            
            try
            {
                int opprid      = ((Addon)this.Addon).opprid;
                int lineid      = ((Addon)this.Addon).lineid;
                int lead_ata    = ((Addon)this.Addon).lead_ata;
                dynamic docnum = this.GetValue("ORDR", "DocNum");

                if (opprid == 0 || ((Addon)this.Addon).active == false)
                    return;

                string sql =
                    "   SELECT " +
                    "       tb1.U_PROJETO, tbProj.Code as code_proj_func, tbProj.U_nome as nome_proj_func , ohemProj.empID as empid_proj " +
                    "       , tb1.U_Assistente, tbAssist.Code as code_assist_func, tbAssist.U_nome as nome_assist_func, ohemAssist.empID empid_assist " +
                    "       , tb1.U_inside_sales, tbInside.Code as code_inside_func, tbInside.U_nome as nome_inside_func, ohemInside.empID empid_inside " +
                    "       , tb1.SlpCode, tbSlp.Code as code_slp_func, tbSlp.U_nome as nome_slp_func, ohemSlp.empID empid_slp " +
                    "       , tb3.U_empid as empid_arquiteto_ata " +
                    "   FROM OPR1 tb1 (NOLOCK) " +

                    "   LEFT JOIN [@UPD_IT_FUNCOES] tbProj (NOLOCK) ON ( LOWER(tbProj.U_mapfield) = 'u_projeto' AND tbProj.U_ativo = 'S' ) " +
                    "   LEFT JOIN OHEM ohemProj (NOLOCK) ON ( ohemProj.firstName + ' ' + ohemProj.lastName LIKE '%' + tb1.U_PROJETO + '%' ) " +

                    "   LEFT JOIN [@UPD_IT_FUNCOES] tbAssist (NOLOCK) ON ( LOWER(tbAssist.U_mapfield) = 'u_assistente' AND tbAssist.U_ativo = 'S' ) " +
                    "   LEFT JOIN OHEM ohemAssist (NOLOCK) ON ( ohemAssist.firstName + ' ' + ohemAssist.lastName LIKE '%' + tb1.U_Assistente + '%' ) " +

                    "   LEFT JOIN [@UPD_IT_FUNCOES] tbInside (NOLOCK) ON ( LOWER(tbInside.U_mapfield) = 'u_inside_sales' AND tbInside.U_ativo = 'S' ) " +
                    "   LEFT JOIN OHEM ohemInside (NOLOCK) ON ( ohemInside.firstName + ' ' + ohemInside.lastName LIKE '%' + tb1.U_inside_sales + '%' ) " +

                    "   LEFT JOIN [@UPD_IT_FUNCOES] tbSlp (NOLOCK) ON ( LOWER(tbSlp.U_mapfield) = 'slpcode' AND tbSlp.U_ativo = 'S' ) " +
                    "   LEFT JOIN OHEM ohemSlp (NOLOCK) ON ( ohemSlp.salesPrson = tb1.SlpCode ) " +

                    "   LEFT JOIN ORDR tb2 (NOLOCK) ON (tb2.U_UPD_IT_LEAD = '" + lead_ata + "') " +
                    "   LEFT JOIN [@UPD_IT_PARTICIP] tb3 (NOLOCK) ON (tb2.DocNum = tb3.U_docentry AND tb3.U_funcao = tbProj.Code) " +

                    "   WHERE " +
                    "       OpprId = '" + opprid + "' AND " +
                    "       Line = " + lineid;

                ((Addon)this.Addon).resetarDadosOppr();

                SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);

                if (rs.RecordCount > 0)
                {
                    string code_proj    = rs.Fields.Item("code_proj_func").Value;
                    int empid_proj      = rs.Fields.Item("empid_proj").Value;
                    
                    string nome_comiss_oopr_proj     = rs.Fields.Item("U_PROJETO").Value;
                    string nome_funcao_oopr_proj     = rs.Fields.Item("nome_proj_func").Value;

                    if (!String.IsNullOrEmpty(code_proj) && empid_proj > 0)
                    {
                        Dictionary<string, dynamic> dic_proj = new Dictionary<string, dynamic>() { 
                            {"U_docentry", docnum},
                            {"U_funcao", code_proj},
                            {"U_empid", empid_proj.ToString()},
                        };

                        this.InsertOnMatrix("mtxPart", dic_proj);
                    }
                    else if (!String.IsNullOrEmpty(nome_comiss_oopr_proj))
                    {
                        this.Addon.StatusErro("Foi encontrado '" + nome_comiss_oopr_proj + "' como " + nome_funcao_oopr_proj + ",\n porém o mesmo não foi encontrado como um Colaborador no sistema. ", true);
                    }

                    string code_assist  = rs.Fields.Item("code_assist_func").Value;
                    int empid_assist    = rs.Fields.Item("empid_assist").Value;
                    
                    string nome_comiss_oopr_assist = rs.Fields.Item("U_Assistente").Value;
                    string nome_funcao_oopr_assist = rs.Fields.Item("nome_assist_func").Value;

                    if (!String.IsNullOrEmpty(code_assist) && empid_assist > 0)
                    {
                        Dictionary<string, dynamic> dic_assist = new Dictionary<string, dynamic>() { 
                            {"U_docentry", docnum},
                            {"U_funcao", code_assist},
                            {"U_empid", empid_assist.ToString()},
                        };

                        this.InsertOnMatrix("mtxPart", dic_assist);
                    }
                    else if (!String.IsNullOrEmpty(nome_comiss_oopr_assist))
                    {
                        this.Addon.StatusErro("Foi encontrado '" + nome_comiss_oopr_assist + "' como " + nome_funcao_oopr_assist + ",\n porém o mesmo não foi encontrado como um Colaborador no sistema. ", true);
                    }

                    string code_inside  = rs.Fields.Item("code_inside_func").Value;
                    int empid_inside    = rs.Fields.Item("empid_inside").Value;

                    string nome_comiss_oopr_inside = rs.Fields.Item("U_inside_sales").Value;
                    string nome_funcao_oopr_inside = rs.Fields.Item("nome_inside_func").Value;

                    if (!String.IsNullOrEmpty(code_inside) && empid_inside > 0)
                    {
                        Dictionary<string, dynamic> dic_inside = new Dictionary<string, dynamic>() { 
                            {"U_docentry", docnum},
                            {"U_funcao", code_inside},
                            {"U_empid", empid_inside.ToString()},
                        };

                        this.InsertOnMatrix("mtxPart", dic_inside);
                    }
                    else if (!String.IsNullOrEmpty(nome_comiss_oopr_inside))
                    {
                        this.Addon.StatusErro("Foi encontrado '" + nome_comiss_oopr_inside + "' como " + nome_funcao_oopr_inside + ",\n porém o mesmo não foi encontrado como um Colaborador no sistema. ", true);
                    }

                    string code_slp = rs.Fields.Item("code_slp_func").Value;
                    int empid_slp   = rs.Fields.Item("empid_slp").Value;

                    if (!String.IsNullOrEmpty(code_slp) && empid_slp > 0)
                    {
                        // flagando como gerente de vendas da venda.
                        Dictionary<string, dynamic> dic_slp = new Dictionary<string, dynamic>() { 
                        {"U_docentry", docnum},
                        {"U_funcao", code_slp},
                        {"U_empid", empid_slp.ToString()},
                        {"U_gerente", "S"},
                    };

                        this.InsertOnMatrix("mtxPart", dic_slp);
                    }
                }
                /*
                SAPbobsCOM.Recordset rstime = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rstime.DoQuery("SELECT U_UPD_IT_TIME FROM OPR2 WHERE OpprId = '" + opprid + "'");
                int time = rstime.Fields.Item("Industry").Value;

                SAPbouiCOM.ComboBox cbTime = this.GetItem("cbTime").Specific;
                cbTime.Select(time.ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                 */
            }
            catch( Exception e )
            {
                this.Addon.DesenvTimeError(e, "Erro ao importar dados da Oportunidade de Venda.");
            }
            finally
            {
                ((Addon)this.Addon).resetarDadosOppr();
                this.SapForm.Freeze(false);
            }
        }

        /// <summary>
        /// Verifica se existe algum participante oculo
        /// Se tiver, associa ele como um dos participantes da venda
        /// Seta um campo de U_oculto = 'S' pra ele não aparecer na tela.
        /// </summary>
        public void InserirParticipanteOculto()
        {
            string sql = "SELECT empID, U_UPD_IT_FUNCAO FROM OHEM (NOLOCK) WHERE U_UPD_IT_FUNCAO IS NOT NULL ";
            SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);

            if( rs.RecordCount > 0 )
            {
                string docentry = this.GetValue("ORDR", "DocNum");

                while( !rs.EoF )
                {
                    int empid = rs.Fields.Item("empID").Value;
                    string funcao = rs.Fields.Item("U_UPD_IT_FUNCAO").Value;

                    this.InsertOnServer("@UPD_IT_PARTICIP", new Dictionary<string, dynamic>()
                    {
                        {"U_funcao",funcao},
                        {"U_empid",empid},
                        {"U_oculto","S"},
                        {"U_docentry",docentry}
                    });
                    
                    rs.MoveNext();
                }
            }
        }

        /// <summary>
        /// Busca dados de renda do lead.
        /// Importa o campo de "alocações cross"
        /// </summary>
        public void BuscarDadosLead(int lead)
        {
            this.Addon.StatusInfo("Importando dados do lead " + lead + " ... ");
            
            string sql = 
                "   SELECT " +
                "       tb2.DocNum, tbMemo.Code as code_funcao, tbMemo.U_nome as funcao " +
                "       , ohemMemo.empID, ohemMemo.firstName + ' ' + ohemMemo.lastName as 'nome'" +
                "       , tb1.U_upd_12_renda as 'renda' " +
                "       , tb1.U_UPD_IT_PERCENT as 'percent_pn' " +
                "       , tb3.ChnCrdCode as 'pn' " +
                "       , tb1.Memo as 'nome_campo_valor_cross' " +
                "   FROM OPR2 tb1 (NOLOCK) " +
                "   LEFT JOIN ORDR tb2 (NOLOCK) ON (tb2.U_UPD_IT_LEAD = " + lead + ") " +
                "   LEFT JOIN OOPR tb3 (NOLOCK) ON (tb1.opportid = tb3.opprid ) " +
                "   LEFT JOIN [@UPD_IT_FUNCOES] tbMemo (NOLOCK) ON ( LOWER(tbMemo.U_mapfield) = 'memo' AND tbMemo.U_ativo = 'S' ) " +
                "   LEFT JOIN OHEM ohemMemo (NOLOCK) ON ( ohemMemo.firstName + ' ' + ohemMemo.lastName LIKE '%' + tb1.Memo + '%' ) " +
                "   WHERE tb1.U_upd_1_nlead = " + lead;

            SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);

            int empID               = rs.Fields.Item("empID").Value;
            int docnum              = rs.Fields.Item("DocNum").Value;
            double renda            = rs.Fields.Item("renda").Value;
            string code_funcao      = rs.Fields.Item("code_funcao").Value;
            string funcao           = rs.Fields.Item("funcao").Value;
            string nome_colab       = rs.Fields.Item("nome").Value;
            string str_renda        = renda.ToString().Replace(".", "").Replace(",", ".");
            string str_renda_antes  = this.GetValue("ORDR", "U_UPD_IT_RENDA");
            string docnum_atual     = this.GetValue("ORDR", "DocNum");
            double percent_pn       = rs.Fields.Item("percent_pn").Value;
            string pn               = rs.Fields.Item("pn").Value;
            
            //valor que está na coluna Memo da tabela OPR2. Pode ser que não encontre o empID pelo nome estar errado.
            string nome_campo_valor_cross = rs.Fields.Item("nome_campo_valor_cross").Value;            

            if( rs.RecordCount > 0 )
            {
                if (renda > 0)
                {
                    SAPbouiCOM.EditText edRenda     = ((SAPbouiCOM.EditText)this.GetItem("edRenda").Specific);
                    SAPbouiCOM.EditText edLead      = ((SAPbouiCOM.EditText)this.GetItem("edLead").Specific);
                    SAPbouiCOM.EditText edNumAtCard = ((SAPbouiCOM.EditText)this.GetItem("14").Specific);

                    // Se não existe um pedido com este lead ou ele está atualizando o próprio pedido.
                    if( docnum == 0 || docnum_atual == docnum.ToString() )
                    {
                        if (!String.IsNullOrEmpty(str_renda_antes) && str_renda_antes != "0.0")
                        {
                            if (this.Addon.SBO_Application.MessageBox("Já existe uma renda comissionável para este pedido.\nDeseja atualizá-la?", 2, "Sim", "Não") == 1)
                            {
                                edRenda.Value = str_renda;
                                edNumAtCard.Value = lead.ToString();
                            }
                        }
                        else
                        {
                            this.Addon.StatusInfo("A Renda total do lead é de R$" + string.Format("{0:#,0.00}", renda) + ".\nConfirme o valor da renda para este pedido.", true);
                            edRenda.Value = str_renda;
                            edNumAtCard.Value = lead.ToString();
                        }
                    }
                    else
                    {
                        this.Addon.StatusInfo("O Pedido de Venda " + docnum + " já possui o lead " + lead + " associado.", true);
                        edRenda.Value = "0.0";
                        edNumAtCard.Value = "";
                        edLead.Value = "";
                    }
                }
                else
                {                    
                    this.Addon.StatusErro("Não foi encontrado uma renda para este lead.", true);
                }

                // Adicionando o empregado encontrado e sua função.
                if (!String.IsNullOrEmpty(code_funcao))
                {
                    if( empID > 0 )
                    {
                        bool ja_existe_cross = false;

                        SAPbouiCOM.Matrix mtx = this.GetItem("mtxPart").Specific;
                        for (int i = 1; i <= mtx.RowCount && ja_existe_cross == false; i++)
                        {
                            string funcaoLinha = mtx.GetCellSpecific("U_funcao", i).Value;
                            string empidLinha = mtx.GetCellSpecific("U_empid", i).Value;

                            ja_existe_cross = funcaoLinha == code_funcao && empidLinha == empID.ToString() ? true : false;
                        }

                        if (!ja_existe_cross)
                        {
                            this.InsertOnMatrix("mtxPart", new Dictionary<string, dynamic>(){
                                {"U_docentry", docnum_atual}, 
                                {"U_funcao", code_funcao }, 
                                {"U_empid", empID.ToString() }, 
                            });
                        }
                    }
                    else if (!String.IsNullOrEmpty(nome_campo_valor_cross))
                    {
                        this.Addon.StatusErro("Foi encontrado '" + nome_campo_valor_cross + "' como " + funcao + ",\n porém o mesmo não foi encontrado como um Colaborador no sistema. ",true);
                    }
                }

                if( !String.IsNullOrEmpty(pn) && percent_pn > 0 )
                {
                    SAPbouiCOM.EditText edParceiro = ((SAPbouiCOM.EditText)this.GetItem("edParc").Specific);
                    SAPbouiCOM.EditText edPercent = ((SAPbouiCOM.EditText)this.GetItem("edPerc").Specific);

                    edParceiro.Value = pn;
                    edPercent.Value = percent_pn.ToString().Replace(".", "").Replace(",", ".");
                }

                this.Addon.StatusInfo("Ok");
            }
            else
            {
                this.Addon.StatusErro("Lead " + lead + " não encontrado.", true);
            }
        }

        /// <summary>
        /// Busca dados do lead base de ata.
        /// </summary>
        public void BuscarDadosLeadAta(int lead_ata)
        {
            this.Addon.StatusInfo("Importando dados do lead de ata " + lead_ata + " ... ");

            string sql =
                "   SELECT " +
                "       tb2.U_funcao, tb2.U_empid " +
	            "       ,tb3.U_percent, tb3.U_nome " +
                "       ,tb4.U_upd_3_operacao, tb4.U_upd_2_status" +
                "   FROM ORDR tb1 " +
                "   INNER JOIN [@UPD_IT_PARTICIP] tb2 (NOLOCK) ON ( tb1.DocNum = tb2.U_docentry ) " +
                "   INNER JOIN [@UPD_IT_FUNCOES] tb3 (NOLOCK) ON ( tb2.U_funcao = tb3.Code ) " +
                "   INNER JOIN OPR2 tb4 (NOLOCK) ON ( tb4.U_upd_1_nlead = " + lead_ata + "  ) " +
                "   WHERE 1 = 1 " +
                "       AND tb1.U_UPD_IT_LEAD = " + lead_ata +
                "       AND tb3.U_percent IS NOT NULL ";

            SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);
            int encontrados = 0;

            // LIMPANDO SE EXISTIR UMA COLUNA COM % DE ATA PARA GARANTIR QUE REMOVEU TODOS QUE JÁ EXISTIAM.
            SAPbouiCOM.Matrix mtx = this.GetItem("mtxPart").Specific;
            for (int i = 1; i <= mtx.RowCount; i++)
            {
                double percent_ata = 0.0;
                string str_percent_ata = mtx.GetCellSpecific("U_ata", i).Value;
                Double.TryParse(str_percent_ata, out percent_ata);
                if (percent_ata > 0)
                {
                    mtx.DeleteRow(i);
                }
            }

            string msg = "";

            if (rs.RecordCount > 0)
            {                
                while( !rs.EoF )
                {
                    int empid       = rs.Fields.Item("U_empid").Value;
                    string funcao   = rs.Fields.Item("U_funcao").Value;
                    double percent  = rs.Fields.Item("U_percent").Value;
                    string operacao = rs.Fields.Item("U_upd_3_operacao").Value;
                    string status   = rs.Fields.Item("U_upd_2_status").Value;

                    if( status != "WON" )
                    {
                        msg = "O status do Lead " + lead_ata + " é '" + status + "', portanto sua Ata não pode ser replicada.";
                        break;
                    }
                    else if( operacao != "VENDA" )
                    {
                        msg = "A operação do Lead " + lead_ata + " é '" + operacao + "', portanto sua Ata não pode ser replicada.";
                        break;
                    }

                    if (percent > 0)
                    {
                        this.InsertOnMatrix("mtxPart", new Dictionary<string, dynamic>(){
                            {"U_docentry", this.GetValue("ORDR","DocNum")}, 
                            {"U_funcao", funcao }, 
                            {"U_empid", empid.ToString() },
                            {"U_ata", percent },
                            {"U_percom", 0.0 },
                        });

                        encontrados++;
                    }
                    
                    rs.MoveNext();
                }

                if( encontrados > 0 )
                {
                    this.Addon.StatusInfo("Ok");
                }
                else
                {
                    if( String.IsNullOrEmpty(msg) )
                    {
                        msg = "Nenhuma das Funções exercidas no Lead " + lead_ata + " possuem Porcentagem de Ata configurada. ";
                    }

                    this.Addon.StatusErro(msg,true);
                }
            }
            else
            {
                this.Addon.StatusErro("Lead de Ata " + lead_ata + " não encontrado.", true);
            }
        }

        #endregion

        #endregion


        #region :: Form 60100 - Cadastro de Colaboradores

        public void sapForm60100()
        {

            userFieldsParams["60100"] = new List<userFieldsParams>(){

                #region :: Criação de Campos de Usuário

                //função
                new userFieldsParams(){
                    itemRef = "185",
                    fieldId = "UPD_IT_FUNCAO",
                    tableId = "OHEM",
                    field = new fieldParams(){
                        descricao = "Função",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                        size = 30,
                    },
                    comp = new CompDefinition(){
                        Id = "edFunc",
                        Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX,
                        Label = "Função de Comissionamento",
                        Bounds = new Bounds(){ Width = 200, Top = 40, },
                        PopulateSQL = "SELECT Code, U_nome FROM [@UPD_IT_FUNCOES] WHERE U_ativo = 'S'",
                        Visible = ((Addon)this.Addon).mostrar_menu
                    }
                },
                
                //META -> ANTES O CAMPO CHAMAVA META, PEDIRAM PRA MUDAR O LABEL PARA "COTA", POREM O NOME DO CAMPO CONTINUOU META.
                new userFieldsParams(){
                    itemRef = "49",
                    fieldId = "UPD_IT_META",
                    tableId = "OHEM",
                    field = new fieldParams(){
                        descricao = "Meta",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Float,
                        subtipo = SAPbobsCOM.BoFldSubTypes.st_Price
                    },
                    comp = new CompDefinition(){
                        Id = "edMeta",
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Cota",
                        Bounds = new Bounds(){ Width = 90, Left = 110, },
                        Visible = ((Addon)this.Addon).mostrar_menu
                    }
                },

                //meta do time
                new userFieldsParams(){
                    itemRef = "edMeta",
                    fieldId = "UPD_IT_META_TIME",
                    tableId = "OHEM",
                    field = new fieldParams(){
                        descricao = "Meta Time",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Float,
                        subtipo = SAPbobsCOM.BoFldSubTypes.st_Price
                    },
                    comp = new CompDefinition(){
                        Id = "edMetaTime",
                        Type = SAPbouiCOM.BoFormItemTypes.it_EDIT,
                        Label = "Meta (Time)",
                        Bounds = new Bounds(){ Width = 90, Left = 100, },
                        Visible = ((Addon)this.Addon).mostrar_menu
                    }
                },

                #endregion

            };
        }

        #endregion

            
        #region :: Form 179 - Dev. Nota Fiscal de Saída

        public void sapForm179()
        {
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, "179", "179", "OnORINDataAdd", true);
        }

        /// <summary>
        /// After Insert
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnORINDataAdd(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = false;

            if (evObj.ActionSuccess)
            {
                string docentry = this.Addon.GetEventObjectKey(evObj);

                SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string update =
                    " UPDATE " +
                    "   tb4 SET  U_UPD_IT_STATUS = 'P', U_UPD_IT_NOTA = NULL, U_UPD_IT_SERIAL = NULL " +
                    "FROM RIN1 tb1 (NOLOCK) " +
                    "   INNER JOIN INV1 tb3 ON ( tb1.BaseRef = tb3.DocEntry AND tb1.BaseType = 13 )  " +
                    "   INNER JOIN ORDR tb4 ON ( tb3.BaseRef = tb4.DocEntry AND tb3.BaseType = 17 )" +
                    "WHERE  " +
                    "   tb1.DocEntry = " + docentry;
                rs.DoQuery(update);

                BubbleEvent = true;
            }
        }

        #endregion


        #region :: Form 133 - Nota Fiscal de Saída

        public void sapForm133()
        {
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, "133", "133", "OnOINVDataAdd", true);
        }

        /// <summary>
        /// After Insert
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnOINVDataAdd(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = false;

            if (evObj.ActionSuccess)
            {
                string docentry = this.Addon.GetEventObjectKey(evObj);

                SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string update =
                    "UPDATE tb2 " +
	                "   SET   " +
		            "       U_UPD_IT_STATUS = CASE WHEN tb3.DocStatus = 'C' THEN 'R' ELSE 'F' END,   " +
                    "       U_UPD_IT_NOTA = tb3.DocEntry,   " +
                    "       U_UPD_IT_SERIAL = tb3.Serial   " +
                    "FROM INV1 tb1 (NOLOCK)  " +
                    "INNER JOIN ORDR tb2 ON ( tb1.BaseRef = tb2.DocEntry AND tb1.BaseType = 17 )  " +
                    "INNER JOIN OINV tb3 ON ( tb1.DocEntry = tb3.DocEntry ) " +
                    "WHERE  " +
                    "   tb1.DocEntry = " + docentry;
                rs.DoQuery(update);

                BubbleEvent = true;
            }
        }

        #endregion


        #region :: Form 170 - Contas a Receber

        public void sapForm170()
        {
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, "170", "170", "OnORCTDataAdd", true);
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE, "170", "170", "OnORCTDataUpdate", true);
        }

        /// <summary>
        /// After Insert
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnORCTDataAdd(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = false;

            if (evObj.ActionSuccess)
            {
                string docentry = this.Addon.GetEventObjectKey(evObj);

                SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string update =
                    " UPDATE tb4 SET U_UPD_IT_STATUS = 'R' " +
	                "   FROM RCT2 tb1 (NOLOCK) " +
                    "   INNER JOIN OINV tb2 ON ( tb1.DocEntry = tb2.DocEntry )  " +
	                "   INNER JOIN INV1 tb3 ON ( tb2.DocEntry = tb3.DocEntry )  " +
                    "   INNER JOIN ORDR tb4 ON ( tb3.BaseRef = tb4.DocEntry AND tb3.BaseType = 17 )  " +
                    "WHERE  " +
	                "   tb1.DocNum = " + docentry + " " +
                    "   AND tb1.InvType = 13 " +
	                "   AND tb2.DocStatus = 'C'";
                rs.DoQuery(update);

                BubbleEvent = true;
            }
        }

        /// <summary>
        /// After Insert
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnORCTDataUpdate(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = false;

            if (evObj.ActionSuccess)
            {
                if (this.Addon.SBO_Application.Forms.Item(evObj.FormUID).DataSources.DBDataSources.Item("ORCT").GetValue("Canceled", 0) == "Y")
                {
                    string docentry = this.Addon.GetEventObjectKey(evObj);

                    SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string update =
                        " UPDATE tb4 SET U_UPD_IT_STATUS = 'F' " +
                        "   FROM RCT2 tb1 (NOLOCK) " +
                        "   INNER JOIN OINV tb2 ON ( tb1.DocEntry = tb2.DocEntry )  " +
                        "   INNER JOIN INV1 tb3 ON ( tb2.DocEntry = tb3.DocEntry )  " +
                        "   INNER JOIN ORDR tb4 ON ( tb3.BaseRef = tb4.DocEntry AND tb3.BaseType = 17 )  " +
                        "WHERE  " +
                        "   tb1.DocNum = " + docentry + " " +
                        "   AND tb1.InvType = 13 ";
                    rs.DoQuery(update);
                }

                BubbleEvent = true;
            }
        }

        #endregion
    }
}
