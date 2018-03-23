using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TShark
{

    /// <summary>
    /// Estrutura configuração para campos definidos 
    /// por usuário.
    /// By Labs - 01/2013
    /// </summary>
    public class userFieldsParams
    {
        public string fieldId;
        public string tableId;
        public string itemRef;
        public string heightRef;
        public fieldParams field;
        public CompDefinition comp;
        public Dictionary<string, string> values;
        public Dictionary<int, fieldParams> versoes;
        public bool valid_values = true;

        public Dictionary<string, string> PopulateItens;
    }

    public class userFieldsLayout
    {

    }

    /// <summary>
    /// Classe para gerir e implementar campos de usuários
    /// em forms padrão SAP.
    /// By Labs - 08/2013
    /// </summary>
    public class UserFields : Forms
    {

        /// <summary>
        /// Armazena campos definidos por usuário.
        /// By Labs - 12/2012
        /// </summary>
        public Dictionary<string, List<userFieldsParams>> userFieldsParams;
        //public Dictionary<string, List<userFieldsParams>> userFieldsParams;

        public bool recreate = false;
        internal bool OnFormSAP = true;

        /// <summary>
        /// Construtor
        /// </summary>
        public UserFields(FastOne addon) :
            base(addon)
        {

            // Registro dos parametros de campos de usuário
            this.userFieldsParams = new Dictionary<string, List<userFieldsParams>>();

            this.FormParams = new FormParams
            {
                Controls = new Dictionary<string, CompDefinition>()
            };

        }

        /// <summary>
        ///  Cria e registra campos de usuário, definidos para
        ///  serem exibidos em form padrão SAP. Ao registrar, os 
        ///  campos serão criados e exibidos nos forms correspondentes.
        ///  By Labs - 01/2013
        /// </summary>
        /// <todo>Verificar a possibilidade de registrar eventos padrão de form tipo BeforeUpdate, etc...</todo>
        internal void registerUserFields()
        {
            string formRef;
            List<string> forms = new List<string>();
            foreach(KeyValuePair<string, List<userFieldsParams>> usrFieldParam in this.userFieldsParams)
            {

                formRef = usrFieldParam.Key;
                if(forms.IndexOf(formRef) < 0)
                {
                    forms.Add(formRef);

                    // Registra o evento para colocação dos comps nos formulários:
                    this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_FORM_LOAD, formRef, formRef, "UserFieldsOnCreateHandler", false);

                    // Garante o redimensionamento:
                    this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE, formRef, formRef, "SystemFormResizeHandler", true);

                    // Ajusta this.SapForm quando o form recebe foco
                    this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE, formRef, formRef, "SystemFormOnActivateHandler");
                }

                // Cria o campo na tabela de destino:
                foreach(userFieldsParams usrField in usrFieldParam.Value)
                {
                    if(!String.IsNullOrEmpty(usrField.tableId) && !String.IsNullOrEmpty(usrField.fieldId))
                    {
                        if(this.recreate)
                        {
                            this.Addon.DtSources.removeField(usrField.tableId, usrField.fieldId);
                        }

                        if(usrField.comp != null && usrField.comp.Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX && usrField.valid_values)
                        {
                            if(usrField.values == null)
                            {
                                usrField.values = new Dictionary<string, string>(){
                                    {"O", "Selecione..."}
                                };
                            }
                        }

                        this.Addon.DtSources.addField(usrField.tableId, usrField.fieldId, usrField.field, 
                            (usrField.values != null
                                ? usrField.values        // <--- DEPRECATED
                                : usrField.PopulateItens
                            )
                        );
                    }
                }
            }
            GC.Collect();
        }

        /// <summary>
        /// Posiciona e exibe os campos de usuário registrados
        /// nos forms SAP.
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="evObj"></param>
        public void SystemFormSetup(ref SAPbouiCOM.ItemEvent evObj)
        {
            string frmId = evObj.FormTypeEx;
            string last_field = "";
            SAPbouiCOM.Form form = this.Addon.SBO_Application.Forms.GetForm(frmId, evObj.FormTypeCount);
            this.SapForm = form;

            // Ajusta os componentes do form:
            foreach(userFieldsParams usrField in this.userFieldsParams[frmId])
            {
                try
                {
                    if(usrField.comp != null)
                    {
                        last_field = usrField.fieldId;

                        // Pega o componente de referência no form:
                        SAPbouiCOM.Item ctrlRef = form.Items.Item(usrField.itemRef);
                        if (ctrlRef != null)
                        {

                            // Bind
                            if (!String.IsNullOrEmpty(usrField.tableId)) {
                                usrField.comp.BindTo = "U_" + usrField.fieldId;
                                usrField.comp.BindTable = usrField.tableId;
                            } else
                            {
                                usrField.comp.BindTo = "_no_bind_";
                            }

                            usrField.comp.FromPane = (usrField.comp.FromPane > 0 ? usrField.comp.FromPane : ctrlRef.FromPane);
                            usrField.comp.ToPane = (usrField.comp.ToPane > 0 ? 
                                usrField.comp.ToPane 
                                : (usrField.comp.FromPane > 0 
                                    ? usrField.comp.FromPane 
                                    : ctrlRef.ToPane
                                )
                            );
                            SAPbouiCOM.Item ctrl = this.makeComp(usrField.fieldId, ref form, ref usrField.comp);

                            /*if (!usrField.comp.Enabled)
                            {
                                try
                                {
                                    ctrl.Enabled = false;
                                    ctrl.SetAutoManagedAttribute(
                                        SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                                        -1,
                                        SAPbouiCOM.BoModeVisualBehavior.mvb_False
                                    );

                                }
                                catch {}
                            }*/

                            if (ctrl == null)
                            {
                                this.Addon.StatusErro(this.Addon.AddonInfo.Descricao + ": Não foi possível criar o componente '" + usrField.fieldId + "'");

                            } else
                            {
                                if(usrField.comp.Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                                {
                                    try
                                    {
                                        ((SAPbouiCOM.ComboBox)ctrl.Specific).ValidValues.Add("", "Selecione...");
                                    } catch { }
                                }

                                // Nova tab
                                if(usrField.comp.Type == SAPbouiCOM.BoFormItemTypes.it_FOLDER)
                                {
                                    ctrl.Width = 350;
                                    ctrl.AffectsFormMode = false;
                                    ((SAPbouiCOM.Folder)ctrl.Specific).Caption = usrField.comp.Caption;
                                    ((SAPbouiCOM.Folder)ctrl.Specific).GroupWith(usrField.itemRef);
                                    ((SAPbouiCOM.Folder)ctrl.Specific).Pane = (usrField.comp.Pane > 0 ? usrField.comp.Pane : 999);
                                    ((SAPbouiCOM.Folder)ctrl.Specific).AutoPaneSelection = true;
                                    this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, frmId, usrField.fieldId, "swapTabs");
                                }
                            }

                        } else
                        {
                            this.Addon.StatusErro(this.Addon.AddonInfo.Descricao + ": Não foi possível encontrar o componente de referência '" + usrField.itemRef + "'");
                        }
                    }
                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, " - Erro criando campo de usuário " + last_field);
                }
            }
            GC.Collect();
        }

        /// <summary>
        /// Posiciona e exibe os campos de usuário registrados
        /// nos forms SAP.
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="evObj"></param>
        public void SystemFormResize(ref SAPbouiCOM.ItemEvent evObj)
        {
            SAPbouiCOM.Form form = this.Addon.SBO_Application.Forms.GetForm(evObj.FormTypeEx, evObj.FormTypeCount);
            string last_field = "";

            // Ajusta os componentes do form:
            foreach(userFieldsParams usrField in this.userFieldsParams[evObj.FormTypeEx])
            {
                if(usrField.comp != null)
                {
                    try
                    {
                        //   form.Freeze(true);
                        last_field = (String.IsNullOrEmpty(usrField.comp.Id) ? usrField.fieldId : usrField.comp.Id);

                        // Pega o componente de referência no form:
                        SAPbouiCOM.Item ctrlRef = form.Items.Item(usrField.itemRef);

                        // Pega o componente:
                        SAPbouiCOM.Item ctrl = null;
                        try
                        {
                            ctrl = form.Items.Item((String.IsNullOrEmpty(usrField.comp.Id) ? usrField.fieldId : usrField.comp.Id));
                        } catch(Exception e) { }


                        if(ctrl != null)
                        {
                            int t = ctrlRef.Top  + usrField.comp._getTop();
                            int l = ctrlRef.Left + usrField.comp._getLeft();
                            int w = usrField.comp._getWidth();
                            int h = usrField.comp._getHeight();

                            // Posição
                            ctrl.Top = t;
                            ctrl.Left = l;

                            // Tamanho  
                            if(usrField.comp.Bounds.PinBottom)
                            {
                                var ch = (String.IsNullOrEmpty(usrField.heightRef) 
                                    ? form.ClientHeight
                                    : form.Items.Item(usrField.heightRef).Height
                                );
                                h = (int)(((float)(ch - t) / 100) * h);
                            }
                            if(h > 0)
                            {
                                ctrl.Height = h;
                            }
                            
                            if(usrField.comp.Bounds.PinRight)
                            {
                                /*var cw = (String.IsNullOrEmpty(usrField.heightRef)
                                    ? form.ClientHeight
                                    : form.Items.Item(usrField.heightRef).Height
                                );*/
                                w = (int)(((float)(form.ClientWidth - l) / 100) * w);
                            }
                            if(w > 0)
                            {
                                ctrl.Width = w;
                            }

                            // Label
                            if(!String.IsNullOrEmpty(usrField.comp.Label))
                            {
                                SAPbouiCOM.Item lbl = form.Items.Item(ctrl.LinkTo);
                                this.calcLabelPos(lbl, ctrl, usrField.comp.LblAlign, usrField.comp.LblSpace);
                            }

                            // Se tiver colunas, ajusta de acordo:
                            if(usrField.comp.Columns != null && usrField.comp.Columns.Widths != null)
                            {
                                this.resizeColumns(ctrl.Width, ref usrField.comp.Columns, ref ctrl);
                            }

                        }

                    } catch(Exception e)
                    {
                        this.Addon.DesenvTimeError(e, " - Erro posicionando campo de usuário " + last_field);

                    } finally
                    {
                        //  form.Freeze(false);
                    }
                }
            }
            GC.Collect();
        }

        /// <summary>
        /// Recupera ou um form SAP padrão ou o proprio form instanciado
        /// </summary>
        /// <param name="frmId">Id do Form</param>
        /// <param name="frmCount"></param>
        /// <returns>Retorna um form sap</returns>
        new public SAPbouiCOM.Form getForm(string frmId = "", int frmCount = 0)
        {
            if(String.IsNullOrEmpty(frmId) && this.SapForm != null)
            {
                return this.SapForm;
            } else
            {
                return this.Addon.getForm(frmId, frmCount);
            }
        }

        /// <summary>
        /// Retorna o valor de um campo em um client dataset.
        /// </summary>
        /// <param name="table">Não esquecer o @.</param>
        /// <param name="field">Campo que se deseja recuperar o valor.</param>
        /// <param name="RecNo">Se nao informado, assume o ultimo registro.</param>
        new public string GetValue(string table, string field, int RecNo = -1)
        {
            string res = "";
            try
            {
                SAPbouiCOM.Form form = this.getForm();
                if(form != null)
                {
                    res = this.GetValue(form, table, field, RecNo);
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Recuperando valor na tabela " + table + " / field " + field);
            }
            return res.Trim();
        }

        /// <summary>
        /// Recupera um valor em UserDataSources de um form.
        /// </summary>
        /// <param name="field"></param>
        /// <param name="form"></param>
        /// <returns></returns>
        new public string GetValue(string field, SAPbouiCOM.Form form = null)
        {
            string res = "";
            try
            {
                if(form == null)
                {
                    form = this.getForm();
                }
                res = form.DataSources.UserDataSources.Item(field).Value;

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Recuperando valor em UserDataSources / field " + field);
            }
            return res.Trim();
        }

        /// <summary>
        /// Remove todos os rows de um client dataset.
        /// </summary>
        /// <param name="table">Não esquecer o arroba.</param>
        new public void ClearValues(string table)
        {
            try
            {
                SAPbouiCOM.Form form = this.getForm();
                if(form != null)
                {
                    this.ClearValues(form, table);
                }
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Limpando a tabela " + table);
            }
        }

        /// <summary>
        /// Retorna a quantidade de registros em um client dataset.
        /// </summary>
        /// <param name="table">Não esquecer o arroba.</param>
        new public int GetCount(string table)
        {
            int res = 0;
            try
            {
                SAPbouiCOM.Form form = this.getForm();
                if(form != null)
                {
                    res = this.GetCount(form, table);
                }
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Recuperando a quantidade de registros na tabela " + table);
            }

            return res;
        }



        #region :: Manipulação de Dados


        #region :: Not Suported On SAP FORM

        internal void NotSuportedInSAPForm()
        {
            this.Addon.ShowMessage("Erro de Desenvolvimento! Esta função não é suportada em Forms internos do SAP, portanto, NÃO UTILIZE EM USERFIELDS! ");
        }

        /// <summary>
        /// NÃO SUPORTADA EM USERFIELDS!!
        /// </summary>
        // new public SAPbouiCOM.Matrix SetupMatrix(string mtxId, string tbId, List<ColumnDefinition> columns, bool UsingDataTable = false) { this.NotSuportedInSAPForm(); return null; }

        /// <summary>
        /// NÃO SUPORTADA EM USERFIELDS!!
        /// </summary>
        new public void InsertOnClient(Dictionary<string, dynamic> values, string table, SAPbouiCOM.Form form) { this.NotSuportedInSAPForm(); }

        /// <summary>
        /// NÃO SUPORTADA EM USERFIELDS!!
        /// </summary>
        new public void InsertOnClient(Dictionary<string, dynamic> values, string table, string formId = "", int formCount = 0) { this.NotSuportedInSAPForm(); }

        /// <summary>
        /// NÃO SUPORTADA EM USERFIELDS!!
        /// </summary>
        new public void UpdateOnClient(Dictionary<string, dynamic> values, string table, int RecNo = -1, string formId = "", int formCount = 0) { this.NotSuportedInSAPForm(); }

        /// <summary>
        /// NÃO SUPORTADA EM USERFIELDS!!
        /// </summary>
        new public void UpdateOnClient(Dictionary<string, dynamic> values, string table, SAPbouiCOM.Form form, int RecNo = -1) { this.NotSuportedInSAPForm(); }

        /// <summary>
        /// NÃO SUPORTADA EM USERFIELDS!!
        /// </summary>
        new public void DeleteOnClient(string table, int RecNo, string formId = "", int formCount = 0) { this.NotSuportedInSAPForm(); }

        #endregion


        /// <summary>
        /// Atualiza campos em um row de uma matriz em Form Padrão SAP.
        /// </summary>
        /// <param name="mtxId">Identificador da matriz</param>
        /// <param name="values">Valores a serem atualizados ({ColID, value})</param>
        /// <param name="RecNo">Se informado, é o número do row da matriz, se não é atualizado o row selecionado.</param>
        /// <param name="formId">Se informado, identifica o form via "Forms.Item()"</param>
        /// <param name="formCount">Se informado, identifica o form via "Forms.GetForm()"</param>
        public void UpdateOnMatrix(string mtxId, Dictionary<string, dynamic> values, int RecNo = 1, string formId = "", int formCount = 0)
        {
            try
            {
                SAPbouiCOM.Matrix mtx = this.GetItem(mtxId, formId, formCount).Specific;
                if(mtx != null)
                {
                    if(RecNo < 1)
                    {
                        RecNo = mtx.GetNextSelectedRow();
                    }

                    if(RecNo < 1)
                    {
                        RecNo = 1;
                    }

                    foreach(KeyValuePair<string, dynamic> v in values)
                    {
                        SAPbouiCOM.Cell cell = mtx.Columns.Item(v.Key).Cells.Item(RecNo);
                        try
                        {
                            cell.Specific.Value = v.Value;
                        } catch (Exception e)
                        {
                            try
                            {
                                cell.Specific.Select(v.Value);
                            } catch (Exception e2)
                            {
                                try
                                {
                                    cell.Specific.Check(RecNo, v.Value);
                                } catch(Exception e3)
                                {

                                }
                            }
                        }
                    }

                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Alterando valor na matirz (A matriz usa DBDataSource?) " + mtxId);
            }
        }

        /// <summary>
        /// Salva alterações feitas em um matrix com dataset não UDO em um formulário padrão SAP.
        /// </summary>
        /// <param name="formId"></param>
        /// <param name="mtxId"></param>
        /// <param name="table"></param>
        new public void SaveToDataSource(string mtxId, string table)
        {
            try
            {
                SAPbouiCOM.Matrix matrix = this.GetItem(mtxId).Specific;
                matrix.FlushToDataSource();
                this.Addon.DtSources.saveUserDataSource(table);
            } catch(Exception e)
            {
                this.Addon.StatusErro(e.Message);
            }
        }


        #endregion

    }

}
