using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Timers;

namespace TShark {

    /// <summary>
    /// Classe para criação de forms
    /// </summary>
    public class Forms {

        /// <summary>
        /// Id do form
        /// </summary>
        public string FormId;

        /// <summary>
        /// Armazena o código de um form UDO / noObject.
        /// </summary>
        public string UDOCode = "";

        /// <summary>
        /// Objeto form SAP
        /// </summary>
        public SAPbouiCOM.Form SapForm;

        /// <summary>
        /// Referência ao form que abriu este form, quando houver.
        /// </summary>
        public Object Oppener = null;

        public bool LoadingFromXML = false;
        internal SAPbouiCOM.ItemEvent itemEventForXML = null;

        public SAPbouiCOM.Matrix Matrix = null;
        public SAPbouiCOM.Button btnSalvar;

        public bool CodeAfter = false;

        /// <summary>
        /// Armazena as colunas e sql por matriz, que precisam ser atualizadas com SQL, 
        ///   string: matrix, string: colId, List: sql, first_key, first_value
        /// </summary>
        internal Dictionary<string, Dictionary<string, List<string>>> MatrixRefreshColumns;
        internal Dictionary<string, SetupMatrixParams> MatrixParams;
        internal Dictionary<string, List<string>> MatrixEmptyRows;
        internal Dictionary<string, List<string>> MatrixUniqueRows;

        /// <summary>
        /// Armazena os parâmetros de configuração do form
        /// </summary>
        public FormParams FormParams;


        /// <summary>
        /// Parametros extras passados ao form
        /// </summary>
        public Dictionary<string, dynamic> ExtraParams;

        /// <summary>
        /// Acesso a classe datasources
        /// </summary>
        public Datasources DtSources;

        /// <summary>
        /// Status do form
        /// </summary>
        public FormStatus Status = FormStatus.frmNull;

        /// <summary>
        /// Helper para nomear labels
        /// </summary>
        private int lblIdCount = 0;

        private bool HasToValidate = false;
        public bool ValidateError = false;
        private List<string> ToValidate;

        /// <summary>
        /// Objeto SAP Application
        /// By Labs - 10/2012
        /// </summary>
        public FastOne Addon;

        /// <summary>
        /// Tags de espaçamento
        /// </summary>
        internal List<string> spaces;

        /// <summary>
        /// Listage de métodos para verificação de eventos declarados no form
        /// </summary>
        internal Dictionary<string, string> EventMethods;

        public SAPbouiCOM.Item FirstTab = null;

        internal Timer timerUDOFind = null;
        internal Timer timerUDOAdd = null;
        internal bool InInsertMode = false;
        internal bool OnFormSAP = false;

        /// <summary>
        /// Inicializa a classe.
        /// By Labs - 10/2012
        /// </summary>
        /// <param name="SBO_App"></param>
        public Forms(FastOne addon, Dictionary<string, dynamic> ExtraParams = null) {

            // Referencia ao addon
            this.Addon = addon;

            // Referencía dtsources
            this.DtSources = addon.DtSources;

            // Colunas de matrizes para atualização:
            this.MatrixRefreshColumns = new Dictionary<string, Dictionary<string, List<string>>>();
            this.MatrixParams = new Dictionary<string, SetupMatrixParams>();
            this.MatrixEmptyRows = new Dictionary<string, List<string>>();
            this.MatrixUniqueRows = new Dictionary<string, List<string>>();

            if(ExtraParams == null)
            {
                this.ExtraParams = new Dictionary<string, dynamic>();
            } else
            {
                this.ExtraParams = ExtraParams;
            }

            // Timer de imagens
            this.timerUDOFind = new Timer();
            this.timerUDOFind.Elapsed += delegate { GotoFormCode(); };
            this.timerUDOFind.Interval = 700;
            this.timerUDOFind.Enabled = false;

            
            this.timerUDOAdd = new Timer();
            this.timerUDOAdd.Elapsed += delegate { FormUDOSetAddMode(); };
            this.timerUDOAdd.Interval = 500;
            this.timerUDOAdd.Enabled = false;

            
            // Lista de eventos (metodos) declarados no form
            this.EventMethods = new Dictionary<string, string>();

            // Tags de espaçamento
            this.spaces = new List<string>() { 
                "space1", "space2", "space3", "space4", "space5", 
                "space6", "space7", "space8", "space9", "space0",
                "_space", "space_", "_space_", "space"
            };
        }


        #region :: Funções Paraquedas

        /// <summary>
        /// Função genérica para executar RefreshListagem em Oppener.
        /// Deve ser implementada em forms Oppener.
        /// </summary>
        public virtual void RefreshListagem() { }

        /// <summary>
        /// Função genérica para passar parametros para Oppener.
        /// Deve ser implementada em forms Oppener.
        /// </summary>
        public virtual void PostParams(Dictionary<string, dynamic> oppener_params) { }

        #endregion


        #region :: Reinicialização de Forms

        /// <summary>
        /// Inicializa os datasources do form e retorna
        /// o total de registros do MainDatasource.
        /// </summary>
        /// <param name="query_open"></param>
        internal int InitFormData(SAPbouiCOM.Conditions query_open = null)
        {
            int t = 0;
            if(this.SapForm != null)
            {
                try
                {
                    this.ToValidate = new List<string>();
                    SAPbouiCOM.DBDataSource db = null;
                    if(!String.IsNullOrEmpty(this.FormParams.MainDatasource))
                    {
                        try
                        {
                            if(query_open == null)
                            {
                                query_open = this.Addon.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                                SAPbouiCOM.Condition oCond = query_open.Add();
                                oCond.Alias = "Code";
                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_IS_NULL;
                                
                            }

                            db = this.SapForm.DataSources.DBDataSources.Item(this.FormParams.MainDatasource);
                            db.Query(query_open);
                            //t = db.Size;

                            this.SapForm.SupportedModes = -1;
                            this.SapForm.AutoManaged = true;
                            this.SapForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE; // (t == 0 ? SAPbouiCOM.BoFormMode.fm_ADD_MODE : SAPbouiCOM.BoFormMode.fm_OK_MODE);
                            
                        } catch(Exception e)
                        {
                            this.Addon.DesenvTimeError(e, " - Erro abrindo o DBDatasource " + this.FormParams.MainDatasource);
                        }
                    }

                    // Acrescenta os datasources extras
                    if(this.FormParams.ExtraDatasources != null)
                    {
                        foreach(string dts in this.FormParams.ExtraDatasources)
                        {
                            this.SapForm.DataSources.DBDataSources.Item(dts).Query(query_open);
                        }
                    }
                    this.Status = FormStatus.frmDtsStarted;
                } catch(Exception e)
                {

                }
            }

            // Retorna total de registros
            return t;
        }

        /// <summary>
        /// Recria os links do form quando carregado via XML.
        /// By Labs - 05/2013
        /// </summary>
        internal bool RebuildLinks(bool initdata, SAPbouiCOM.Conditions query_open = null)
        {

            // Remove eventos do form, se ele estiver sendo chamado de novo
           // this.Addon.ClearFormEvents(this.FormId);

            // Parametros do form:
            FormParams FormParams = this.FormParams;

            // Recupera a área dos componentes:
            this.FormParams.Area = this.SapForm.Items.Item("FormArea");

            // Recupera a área de referencia dos botões:
            this.FormParams.ButtonArea = this.SapForm.Items.Item("ButtonArea");

            // Acrescenta o datasource principal
            /*if((FormParams.MainDatasource != null) && (FormParams.MainDatasource != ""))
            {
                try
                {
                    this.SapForm.DataSources.DBDataSources.Add(FormParams.MainDatasource);
                    this.Status = FormStatus.frmDtsCreated;
                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, " - Erro adicionando " + FormParams.MainDatasource + " ao DBDatasources do form " + this.FormId);
                }
            }

            // Acrescenta os datasources extras
            if(FormParams.ExtraDatasources != null)
            {
                foreach(string dts in FormParams.ExtraDatasources)
                {
                    try
                    {
                        this.SapForm.DataSources.DBDataSources.Add(dts);
                        this.Status = FormStatus.frmDtsCreated;
                    } catch(Exception e)
                    {
                        this.Addon.DesenvTimeError(e, " - Erro adicionando " + dts + " ao DBDatasources do form " + this.FormId);
                    }
                }
            }
            */
            if(initdata)
            {
                this.InitFormData(query_open);
            }

            // Ajusta componentes:
            SAPbouiCOM.Item oItem = null;
            foreach(KeyValuePair<string, CompDefinition> comp in FormParams.Controls)
            {
                oItem = this.SapForm.Items.Item(comp.Key);
                if(oItem != null)
                {
                    comp.Value.Id = comp.Key;
                    /*this.setupComp(this.SapForm, comp.Value, oItem);

                    if(comp.Value.Type == SAPbouiCOM.BoFormItemTypes.it_MATRIX)
                    {
                     //   this.resizeColumns(oItem.Width, ref comp.Value.Columns, ref oItem);
                    }
                    */

                    // Executa refresh:
                    bool BubbleEvent = true;
                    FastOneItemEvent evObj = new FastOneItemEvent()
                    {
                        FormUID = this.FormId,
                        ItemUID = oItem.UniqueID
                    };
                    if(this.EventMethods.ContainsKey(oItem.UniqueID + "OnRefresh"))
                    {
                        this.Addon.ExecEvent(oItem.UniqueID + "OnRefresh", out BubbleEvent, new Object[] { evObj, BubbleEvent });
                    }

                    if(this.EventMethods.ContainsKey(oItem.UniqueID + "OnCreate"))
                    {
                        this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_CLICK, this.FormId, oItem.UniqueID, "GetItemEventForXML");
                        oItem.Click();

                        this.Addon.ExecEvent(oItem.UniqueID + "OnCreate", out BubbleEvent, new Object[] { this.itemEventForXML, BubbleEvent });
                    }

                    // Validação
                    if(comp.Value.NonEmpty || comp.Value.Validate != null)
                    {
                        this.HasToValidate = true;
                        this.ToValidate.Add(comp.Key);
                    }

                    // Actions
                    this.setCompActions(comp.Key, this.FormId, comp.Value);
                }
            }

            // Ajusta Tabs:
            if(FormParams.Tabs.Tabs != null)
            {
                string tabId = "";
                int i = 1;
                this.FormParams.TabArea = this.SapForm.Items.Item("TabArea");
                foreach(KeyValuePair<string, Dictionary<string, int>> tab in FormParams.Tabs.Tabs)
                {
                    tabId = "tab" + i;

                    oItem = this.SapForm.Items.Item(tabId);
                    ((SAPbouiCOM.Folder)oItem.Specific).Pane = i;

                    // Registra swap:
                    this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, this.FormId, tabId, "swapTabs");

                    // Registra actions:
                    if(FormParams.Tabs.actions != null)
                    {
                        foreach(action action in FormParams.Tabs.actions)
                        {
                            this.Addon.RegisterEvent(action.EventType, this.FormId, tabId, action.EventHandler);
                        }
                    }
                    i++;
                }
            }

            // Registra actions:
            /*if(FormParams.FormActions != null)
            {
                foreach(action action in FormParams.FormActions)
                {
                    this.Addon.RegisterEvent(action.EventType, this.FormId, this.FormId, action.EventHandler);
                }
            }
            this.Status = FormStatus.frmControlsCreated;
            */

            // Registra os eventos padrão:
            // this.registerFormEvents();

            // Retorna
            GC.Collect();
            return true;
        }

        internal void GetItemEventForXML(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this.itemEventForXML = evObj;
        }

        #endregion


        #region :: Criação de Forms

        /// <summary>
        /// Cria um novo formulário com base nos parametros passados.
        /// By Labs - 12/2012
        /// </summary>
        public bool makeForm(bool initdata, SAPbouiCOM.Conditions query_open = null)
        {

            // Remove eventos do form, se ele estiver sendo chamado de novo
            //this.Addon.ClearFormEvents(this.FormId);

            // Cria o novo form:
            FormParams FormParams = this.FormParams;
            SAPbouiCOM.FormCreationParams cp = ((SAPbouiCOM.FormCreationParams)(this.Addon.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
            cp.UniqueID = this.FormId;
            cp.FormType = this.FormId;
            cp.BorderStyle = FormParams.BorderStyle;
            if(!string.IsNullOrEmpty(FormParams.BusinessObjectId))
            {
                cp.ObjectType = FormParams.BusinessObjectId;
            }

            try
            {
                this.SapForm = this.Addon.SBO_Application.Forms.AddEx(cp);
                this.SapForm.Freeze(true);
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Erro ao criar o form '" + this.FormId + "'");
                return false;
            }

            // Define o básico:
            this.SapForm.Title = FormParams.Title +
                                 " :: " + this.Addon.AddonInfo.Descricao +
                                 " :: v" + this.Addon.AddonInfo.Versao + "." + this.Addon.AddonInfo.Release + "." + this.Addon.AddonInfo.Revisao;
                                 
            this.SapForm.Top = FormParams.Bounds.Top;
            this.SapForm.Left = FormParams.Bounds.Left;
            this.SapForm.ClientWidth = FormParams.Bounds.Width;
            this.SapForm.ClientHeight = FormParams.Bounds.Height;

            // Acrescenta o datasource principal
            if(!string.IsNullOrEmpty(FormParams.MainDatasource))
            {
                try
                {
                    this.SapForm.DataSources.DBDataSources.Add(FormParams.MainDatasource);
                    this.Status = FormStatus.frmDtsCreated;
                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, " - Erro adicionando " + FormParams.MainDatasource + " ao DBDatasources do form " + this.FormId);
                }
            }

            // Acrescenta os datasources extras
            if(FormParams.ExtraDatasources != null)
            {
                foreach(string dts in FormParams.ExtraDatasources)
                {
                    try
                    {
                        this.SapForm.DataSources.DBDataSources.Add(dts);
                        this.Status = FormStatus.frmDtsCreated;
                    } catch(Exception e)
                    {
                        this.Addon.DesenvTimeError(e, " - Erro adicionando " + dts + " ao DBDatasources do form " + this.FormId);
                    }
                }
            }

            if(initdata)
            {
                this.InitFormData(query_open);
            }

            // Cria a área dos componentes:
            this.FormParams.Area = this.SapForm.Items.Add("FormArea", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            this.FormParams.Area.Visible = false;

            // Cria a área de referencia dos botões:
            this.FormParams.ButtonArea = this.SapForm.Items.Add("ButtonArea", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            this.FormParams.ButtonArea.Visible = false;

            // Cria componentes:
            SAPbouiCOM.Item oItem = null;
            CompDefinition ctrlParams;
            foreach(KeyValuePair<string, CompDefinition> comp in FormParams.Controls)
            {
                ctrlParams = comp.Value;
                try
                {
                    oItem = this.makeComp(comp.Key, ref this.SapForm, ref ctrlParams);
                    //oItem.Visible = ctrlParams.Visible;
                    //oItem.Enabled = ctrlParams.Enabled;

                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, " - Erro criando o componente " + comp.Key);
                }
            }

            // Cria Tabs:
            if(FormParams.Tabs.Tabs != null)
            {
                this.createTabs(ref FormParams.Tabs);
            }

            // Conecta o navegador padrão
            if(!String.IsNullOrEmpty(FormParams.BrowseByComp))
            {
                try
                {
                    this.SapForm.DataBrowser.BrowseBy = FormParams.BrowseByComp;
                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, " - Erro definido BrowseBy: " + FormParams.BrowseByComp);
                }
            }

            // Registra actions:
            if(FormParams.FormActions != null)
            {
                foreach(action action in FormParams.FormActions)
                {
                    this.Addon.RegisterEvent(action.EventType, this.FormId, this.FormId, action.EventHandler);
                }
            }
            this.Status = FormStatus.frmControlsCreated;

            // Registra os eventos padrão do form
            //this.registerFormEvents();

            // Retorna
            GC.Collect();

            return true;
        }

        /// <summary>
        /// Calcula posicionamento dos componentes no form.
        /// By Labs - 12/2012
        /// </summary>
        private void formResize(string formId)
        {

            this.SapForm.Freeze(true);
            int Top = 6;

            // Garante espaço para botões
            int saveHeight = 0;
            if(this.FormParams.Buttons != null)
            {
                saveHeight = 38;
            }

            // Ajusta a área de componentes:
            if(this.FormParams.Area != null)
            {
                if(this.FormParams.Linhas != null)
                {
                    this.FormParams.Area.Top = Top;
                    this.FormParams.Area.Left = 0;
                    this.FormParams.Area.Width = this.FormParams.Bounds.Width - 2;
                    this.FormParams.Area.Height = (this.FormParams.Bounds.Height - Top - 2) - saveHeight;
                    Top += this.FormParams.Area.Height + 5;

                    // Posiciona os componentes do form:
                    this.layoutForm(ref this.FormParams.Area, this.FormParams.Linhas);
                } else
                {
                    this.FormParams.Area.Visible = false;
                }
            }

            // Botões padrão
            if(this.FormParams.Buttons != null && this.FormParams.ButtonArea != null)
            {
                this.FormParams.ButtonArea.Top = this.FormParams.Bounds.Height - saveHeight;
                this.FormParams.ButtonArea.Left = 0;
                this.FormParams.ButtonArea.Width = this.FormParams.Bounds.Width - 2;
                this.FormParams.ButtonArea.Height = 1;
                this.FormParams.ButtonArea.Visible = true;

                // Posiciona os botões:
                try
                {
                    this.layoutForm(ref this.FormParams.ButtonArea, this.FormParams.Buttons, 8);
                } catch { }
            }

            this.SapForm.Freeze(false);
            GC.Collect();
        }

        /// <summary>
        ///  (Re) Posiciona os componentes no formulário.
        ///  By Labs - 01/2013
        /// </summary>
        /// <param name="formId"></param>
        /// <param name="areaRef"></param>
        /// <param name="comps"></param>
        private void layoutForm(ref SAPbouiCOM.Item areaRef, Dictionary<string, int> comps, int refY = 0)
        {
            GC.Collect();

            // Define bases:
            int Top = areaRef.Top + refY;
            int entreComps = 3;
            int entreLinhas = 18;
            int entreLinhasMenor = 5;

            int maxWidth = areaRef.Width - entreComps - 2;
            int maxHeight = areaRef.Height - entreComps;

            int linhaY = areaRef.Top + 2 + refY;
            int linhaX = areaRef.Left + entreComps + 2;
            int percentUsado = 0;

            int compHeight = 0;
            int compWidth = 0;
            int w = 0;
            int y = 0;

            int force_x = 0;
            int force_y = 0;

            // Posiciona:
            CompDefinition ctrl = null;
            bool isButton = false;
            bool isTabs = false;
            bool isCheck = false;
            bool hasLabel = false;
            bool usaEntreLinhasMenor = true;
            columnParams Columns = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Item oLabel = null;
            SAPbouiCOM.Form frm = this.SapForm;
            compAlign align = compAlign.calLeft;
            foreach(KeyValuePair<string, int> comp in comps)
            {
                //try
               // {
                    oItem = null;
                    compWidth = (comp.Value > 100 ? 100 : comp.Value);
                    align = compAlign.calLeft;

                    this.Addon.SBO_Application.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, true);

                    // Tabs
                    if(comp.Key == "Tabs")
                    {
                        oItem = this.FormParams.TabArea;
                        oItem.Height = this.FormParams.Tabs.Height;
                        isTabs = true;
                        isButton = false;
                        hasLabel = false;
                        Columns = null;
                        force_x = this.FormParams.Tabs.force_left;
                        force_y = this.FormParams.Tabs.force_top;

                        // Espaçamento:
                    } else if(this.spaces.IndexOf(comp.Key) != -1)
                    {
                        if(comp.Value == 100)
                        {
                            y += 14;
                            linhaY += 14;
                        }

                        // Componente normal
                    } else
                    {
                        oItem = frm.Items.Item(comp.Key);
                        ctrl = this.FormParams.Controls[comp.Key];

                        if(ctrl._getHeight() > 0)
                        {
                            oItem.Height = ctrl._getHeight();
                        }
                        isTabs = false;
                        isButton = (ctrl.Type == SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                        isCheck = (ctrl.Type == SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);

                        hasLabel = (ctrl.Label != null);
                        Columns = ctrl.Columns;
                        force_x = ctrl.ForceLeft;
                        force_y = ctrl.ForceTop;
                        align = ctrl.Align;

                        /*if(ctrl.marginTop > 0)
                        {
                            y += ctrl.marginTop;
                           // linhaY += ctrl.marginTop;
                        }*/
                    }

                    // Muda de linha
                    if((compWidth + percentUsado) > 100)
                    {
                        linhaX = areaRef.Left + entreComps + 2;
                        linhaY += compHeight + (usaEntreLinhasMenor ? entreLinhasMenor : entreLinhas);
                        percentUsado = 0;
                        compHeight = 0;
                        usaEntreLinhasMenor = true;
                    }

                    if(align == compAlign.calRight)
                    {
                        linhaX = (areaRef.Width - entreComps - 2) - linhaX;
                    }

                    w = (maxWidth * comp.Value) / 100;
                    if(linhaX + w > maxWidth)
                    {
                        w -= (linhaX + w) - maxWidth;
                    }

                    // Posicionamento forçado
                    if(force_x > 0)
                    {
                        force_x = (maxWidth * force_x) / 100;
                    }
                    if(force_y > 0)
                    {
                        force_y = (maxHeight * force_y) / 100;
                    }

                    if(oItem != null)
                    {

                        // Label
                        y = linhaY; // (isTabs ? linhaY + 20 : linhaY);
                        if(hasLabel)
                        {
                            y += (isButton ? 11 : 14);
                            if((oItem.LinkTo != "") && !isButton)
                            {
                                oLabel = frm.Items.Item(oItem.LinkTo);
                                oLabel.Top = linhaY;
                                oLabel.Left = linhaX;
                                oLabel.LinkTo = comp.Key;
                                // oLabel.Width = w;
                            }
                            usaEntreLinhasMenor = (ctrl.LblAlign != labelAlign.lblTop);

                            // Não tem Label:
                        } else
                        {
                            y += (isButton
                                ? (ctrl.Align == compAlign.calBottom ? 11 : 0)
                                : (isTabs ? 14 : 0)
                            );
                        }

                        // Ajusta a posição por referencia:
                        if(ctrl != null && !String.IsNullOrEmpty(ctrl.itemRef))
                        {
                            SAPbouiCOM.Item ctrlRef = frm.Items.Item(ctrl.itemRef);
                            oItem.Top = ctrlRef.Top + ctrl._getTop();
                            oItem.Left = ctrlRef.Left + ctrl._getLeft();
                            oItem.Width = ctrl._getWidth();
                            if(oItem.Width == 0)
                            {
                                oItem.Width = (isCheck ? 15 : w);
                            }

                        } else
                        {
                            oItem.Top = y + force_y + (ctrl != null ? ctrl.marginTop : 0);       // (force_y > 0 ? force_y : y);
                            oItem.Left = linhaX + force_x; // (force_x > 0 ? force_x : linhaX);
                            oItem.Width = (isCheck ? 15 : w);
                        }


                        // Guarda o Height maior
                        if(oItem.Height > compHeight)
                        {
                            compHeight = oItem.Height;
                        }


                        if(hasLabel)
                        {
                            if(String.IsNullOrWhiteSpace(ctrl.Label))
                            {
                                oLabel.Width = 0;
                            }
                            //oLabel.BackColor = Color.Yellow;
                            this.calcLabelPos(oLabel, oItem, ctrl.LblAlign, ctrl.LblSpace, isCheck ? w : ctrl.LblWidth);
                        }

                        // Tem likedButton?
                        if(ctrl != null && ctrl.Type == SAPbouiCOM.BoFormItemTypes.it_EDIT)
                        {
                            this.CalcLinkButtonPos(oItem, frm);
                        }

                        // Se tiver colunas, ajusta de acordo:
                        if(Columns != null && Columns.Widths != null)
                        {
                            this.resizeColumns(w, ref Columns, ref oItem);
                        }

                        // Ajusta os componentes do Tab
                        if(isTabs)
                        {
                            int i = 1;
                            int l = linhaX;
                            foreach(KeyValuePair<string, Dictionary<string, int>> tab in this.FormParams.Tabs.Tabs)
                            {

                                // Botão de tab
                                oItem = this.SapForm.Items.Item("tab" + i);
                                oItem.Left = l;
                                oItem.Top = y - 19;
                                oItem.Width = 200;
                                i++;
                                l += 100;

                                // Componentes
                                this.layoutForm(ref this.FormParams.TabArea, tab.Value, 15);
                            }
                            linhaY += (entreLinhas / 2);
                        }
                    }
               // } catch { }

                // Incrementa:
                percentUsado += compWidth;
                linhaX += w + entreComps;

            }
            GC.Collect();
        }

        internal void CalcLinkButtonPos(SAPbouiCOM.Item refItem, SAPbouiCOM.Form frm)
        {
            SAPbouiCOM.Item link = null;
            try
            {
                string lkid = refItem.UniqueID.Remove(0, 2);
                lkid = "lk" + lkid;
                link = frm.Items.Item(lkid);
                link.Top = refItem.Top + 1;
                link.Left = (refItem.Left + refItem.Width) - (link.Width + 1);
                link.Height = 12;
            } catch { }
        }

        /// <summary>
        /// Aplica tamanho de colunas
        /// </summary>
        /// <param name="w"></param>
        /// <param name="Columns"></param>
        /// <param name="ctrl"></param>
        public void resizeColumns(int w, ref columnParams Columns, ref SAPbouiCOM.Item ctrl)
        {
            int i = 0;
            int n = 0;
            SAPbouiCOM.Column column = null;
            foreach(int col_width in Columns.Widths)
            {
                // 0 significa não recalcular e usar o que tiver sido definido no componente
                if(col_width > 0)
                {
                    try
                    {
                        column = (ctrl.Specific).Columns.Item(i);
                        column.Width = (int)((float)(w-n) * col_width) / 100;
                    } catch(Exception e)
                    {
                       // this.Addon.DesenvTimeError(e, " - Erro ajustando as colunas de " + ctrl.UniqueID + "\n - Verifique se existe a coluna na definição de 'Controls' do componente: " + ctrl.UniqueID);
                    }
                    if(i > 1) n = 1;
                }
                i++;
            }
        }

        /// <summary>
        /// Calcula e acerta o posicionamento de um Label em referencia ao seu ctrl
        /// de acordo como lblAlign
        /// </summary>
        /// <param name="lbl"></param>
        /// <param name="ctrl"></param>
        /// <param name="align"></param>
        /// <param name="lblSpace"></param>
        internal void calcLabelPos(SAPbouiCOM.Item lbl, SAPbouiCOM.Item ctrl, labelAlign align, int lblSpace, int lblWidth = -1)
        {
            try
            {
                if(lblWidth > -1)
                {
                    lbl.Width = lblWidth;
                }

                switch(align)
                {
                    case labelAlign.lblLeft:
                        lbl.Top = ctrl.Top;
                        lbl.Left = ctrl.Left;
                        lbl.Width += lblSpace;
                        ctrl.Left += lbl.Width + 4;
                        ctrl.Width -= lbl.Width + 4;
                        break;

                    case labelAlign.lblRight:
                        lbl.Top = ctrl.Top; // + 4;
                        lbl.Left = ctrl.Left + lblSpace;
                        break;

                    default:
                        lbl.Top = ctrl.Top - 14;
                        lbl.Left = ctrl.Left;
                        if(lblWidth == -1)
                        {
                            lbl.Width = ctrl.Width;
                        }
                        break;
                }

                lbl.FromPane = ctrl.FromPane;
                lbl.ToPane = ctrl.ToPane;

            } catch(Exception e)
            {
            }
        }

        #endregion


        #region :: Eventos

        /// <summary>
        /// Registra os eventos padrão de um form
        /// By Labs - 07/2013
        /// </summary>
        internal void registerFormEvents()
        {

            #region :: Eventos de Form

            // O form é criado na UI SAP: OnFormCreate
            this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_FORM_LOAD, this.FormId, this.FormId, "OnFormCreate", false);

            // O form é setado visible: OnBeforeFormOpen / OnFormOpen
            this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE, this.FormId, this.FormId, "OnBeforeFormOpen", false);
            this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE, this.FormId, this.FormId, "OnFormOpen");

            // Um novo registro UDO foi carregado no form -- via browse, link button, ou find (FormDataEvent): OnRefresh
            this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD, this.FormId, this.FormId, "OnRefresh");

            // Um registro UDO no form vai ser inserido no banco -- et_FORM_DATA_ADD (FormDataEvent): OnDataAdd / OnAfterDataAdd
            this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, this.FormId, this.FormId, "OnDataAdd", false);
            this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, this.FormId, this.FormId, "OnAfterDataAdd");

            // Um registro UDO no form vai ser atualizado no banco -- et_FORM_DATA_UPDATE (FormDataEvent): OnDataUpdate / OnAfterDataUpdate
            this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE, this.FormId, this.FormId, "OnDataUpdate", false);
            this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE, this.FormId, this.FormId, "OnAfterDataUpdate");

            // Um registro UDO no form vai ser inserido no banco -- et_FORM_DATA_DELETE (FormDataEvent): OnBeforeDelete / OnAfterDelete
            this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE, this.FormId, this.FormId, "OnBeforeDelete", false);
            this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE, this.FormId, this.FormId, "OnAfterDelete");

            // Uma operação (BEFORE) de DATA_ADD ou de DATA_UPDATE irá acontecer com o registro UDO (FormDataEvent): OnDataSave / OnAfterDataSave
            /** Esse evento é virtual, chamado a partir de OnDataAdd e de OnDataUpdate **/

            // O form é fechado na UI SAP: OnFormClose / OnAfterFormClose
            this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD, this.FormId, this.FormId, "OnFormClose", false);
            this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD, this.FormId, this.FormId, "OnAfterFormClose");

            // Evento da AÇÃO de inserir em um UDO:
            if(this.EventMethods.ContainsKey(this.FormId + "OnInsertUDO"))
            {
                this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_MENU_CLICK, this.FormId, this.FormId, this.FormId + "OnInsertUDO");
            }

            #endregion


            #region :: Eventos Opcionais

            // O form recebeu foco
            if(this.EventMethods.ContainsKey(this.FormId + "OnFormFocus"))
            {
                this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE, this.FormId, this.FormId, "OnFormFocus");
            }

            // O form perdeu foco
            if(this.EventMethods.ContainsKey(this.FormId + "OnFormLostFocus"))
            {
                this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE, this.FormId, this.FormId, "OnFormLostFocus");
            }

            #endregion
                    
            // Resize do form
            this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE, this.FormId, this.FormId, "FormResizeHandler");

        }

        /// <summary>
        /// Handler para o evento form_resize.
        /// By Labs - 12/2012
        /// </summary>
        public void FormResizeHandler(SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            string frmId = evObj.FormTypeEx;
            if(this.SapForm != null && !this.LoadingFromXML)
            {
                this.FormParams.Bounds.Width = this.SapForm.ClientWidth;
                this.FormParams.Bounds.Height = this.SapForm.ClientHeight;
                this.formResize(frmId);
            }
        }


        /// <summary>
        /// Executa um handler na classe filha
        /// </summary>
        /// <param name="handler"></param>
        /// <param name="BubbleEvent"></param>
        /// <param name="evento"></param>
        internal void _executeFormEvent(string handler, out bool BubbleEvent, object[] evento)
        {
            BubbleEvent = true;
            if(this.EventMethods.ContainsKey(handler))
            {
                this.Addon.ExecEvent(handler, out BubbleEvent, evento);
            }
        }


        #region :: Eventos Internos

        /// <summary>
        /// Chamado na criação do Form pela UI (et_FORM_LOAD)
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnFormCreate(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            this.SapForm = this.getForm(evObj.FormTypeEx, evObj.FormTypeCount);

            if(this.LoadingFromXML)
            {
                foreach(KeyValuePair<string, CompDefinition> comp in FormParams.Controls)
                {
                    if(this.EventMethods.ContainsKey(comp.Key + "OnCreate"))
                    {
                        this.Addon.ExecEvent(comp.Key + "OnCreate", out BubbleEvent, new Object[] { evObj, BubbleEvent });
                    }

                    if(this.EventMethods.ContainsKey(comp.Key + "OnRefresh"))
                    {
                        // Executa refresh:
                        FastOneItemEvent evObj2 = new FastOneItemEvent()
                        {
                            FormUID = this.FormId,
                            ItemUID = comp.Key
                        };
                        this.Addon.ExecEvent(comp.Key + "OnRefresh", out BubbleEvent, new Object[] { evObj2, BubbleEvent });
                    }
                }
            }

            // Executa o handler no form
            this._executeFormEvent(this.FormId + "OnFormCreate", out BubbleEvent, new object[] { evObj, BubbleEvent });
        }

        /// <summary>
        /// Chamado antes do form ficar visibel (et_FORM_VISIBLE)
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnBeforeFormOpen(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            this._executeFormEvent(this.FormId + "OnBeforeFormOpen", out BubbleEvent, new object[] { evObj, BubbleEvent });
        }

        /// <summary>
        /// Chamado depois do form ficar visibel (et_FORM_VISIBLE)
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnFormOpen(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                this.SapForm.Freeze(true);
                if(!String.IsNullOrEmpty(this.UDOCode) && !String.IsNullOrEmpty(this.FormParams.BrowseByComp) && this.SapForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    this.SapForm.Items.Item(this.FormParams.BrowseByComp).Specific.Value = this.UDOCode;
                }

                // Desabilita comps
                foreach(KeyValuePair<string, CompDefinition> c in this.FormParams.Controls)
                {
                    if(!c.Value.Enabled)
                    {
                        try
                        {
                            this.SapForm.Items.Item(c.Key).Enabled = false;
                            this.SapForm.Items.Item(c.Key).SetAutoManagedAttribute(
                                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                                c.Value.ModeMask,
                                SAPbouiCOM.BoModeVisualBehavior.mvb_False
                            );

                        } catch(Exception e)
                        {

                        }
                    }
                }

                // Ajusta Tab
                if(this.FirstTab != null)
                {
                    this.FirstTab.Click();
                }

                if(this.InInsertMode)
                {
                    this.SapForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

                    // Insert UDO
                    if(!String.IsNullOrEmpty(this.FormParams.BusinessObjectId))
                    {
                        this.FormUDOSetAddMode();

                    // Insert NoObject
                    } else if(!String.IsNullOrEmpty(this.FormParams.MainDatasource))
                    {
                        if(this.ExtraParams.ContainsKey("INSERT_PARAMS"))
                        {
                            this.InsertOnClient(this.ExtraParams["INSERT_PARAMS"], this.FormParams.MainDatasource);
                        } else
                        {
                            SAPbouiCOM.DBDataSource dts = this.SapForm.DataSources.DBDataSources.Item(this.FormParams.MainDatasource);
                            dts.InsertRecord(dts.Size);
                            dts.Offset = dts.Size - 1;
                        }
                        this.SapForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    }
                }

                if (this.SapForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE && !String.IsNullOrEmpty(this.UDOCode)){
                    this.SapForm.Items.Item("1").Click();
                }

                // Executa o handler no form
                this._executeFormEvent(this.FormId + "OnFormOpen", out BubbleEvent, new object[] { evObj, BubbleEvent });
                
                // Seta foco no componente definido em CompDefinition
                if(!String.IsNullOrEmpty(this.FormParams.Focus))
                {
                    this.SapForm.Items.Item(this.FormParams.Focus).Click();
                }

            } catch
            {

            } finally
            {
                this.SapForm.Freeze(false);
            }
        }

        /// <summary>
        /// Chamado no refresh de um form UDO (quando se navega de um registro para outro, ou da a primeira carga) 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnRefresh(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                this.GetFormCode(evObj);
                this._executeFormEvent(this.FormId + "OnRefresh", out BubbleEvent, new object[] { evObj, BubbleEvent });

                if(!String.IsNullOrEmpty(this.FormParams.BrowseByComp) && !this.FormParams.Controls[this.FormParams.BrowseByComp].Enabled)
                {
                    this.SapForm.Items.Item(this.FormParams.BrowseByComp).Enabled = false;
                }
            } catch { }
        }

        /// <summary>
        /// Chamado antes do registro UDO for ser inserido ou alterado no banco.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnDataSave(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
        
        }

        /// <summary>
        /// Chamado após do registro UDO for ser inserido ou alterado no banco.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnAfterDataSave(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

        }

        /// <summary>
        /// Chamado antes do registro UDO for ser inserido no banco, executando as validações.
        /// Dispara também o evento OnDataSave.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnDataAdd(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if(this.CodeAfter)
            {
                string code = this.getNextCode(this.FormParams.MainDatasource, 5);
                //this.Addon.StatusAlerta(code);

                this.UpdateOnClient(new Dictionary<string, dynamic>() { 
                    {"Code", code},
                    {"DocEntry", code},
                }, this.FormParams.MainDatasource);

                this.UDOCode = code;
            }

            this._executeFormEvent(this.FormId + "OnDataAdd", out BubbleEvent, new object[] { evObj, BubbleEvent });
            this._executeFormEvent(this.FormId + "OnDataSave", out BubbleEvent, new object[] { evObj, BubbleEvent });
            
            // Executa validações
            if(BubbleEvent)
            {
                BubbleEvent = this.ValidateHandler();
            }
        }

        /// <summary>
        /// Chamado depois que o registro UDO foi inserido
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnAfterDataAdd(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this.InInsertMode = false;

            this._executeFormEvent(this.FormId + "OnAfterDataAdd", out BubbleEvent, new object[] { evObj, BubbleEvent });
            this._executeFormEvent(this.FormId + "OnAfterDataSave", out BubbleEvent, new object[] { evObj, BubbleEvent });
            if(BubbleEvent)
            {
                // Vai para o registro UDO recém inserido
                if(!String.IsNullOrEmpty(this.FormParams.BusinessObjectId))
                {
                    try
                    {
                        this.timerUDOFind.Start();

                    } catch(Exception e)
                    {
                        this.Addon.DesenvTimeError(e, " em formOnUDOAdd de " + evObj.FormTypeEx);
                    }
                }

            }
        }
        
        /// <summary>
        /// Chamado antes do registro UDO ser atualizado no banco, executando as validações.
        /// Dispara também o evento OnDataSave.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnDataUpdate(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this._executeFormEvent(this.FormId + "OnDataUpdate", out BubbleEvent, new object[] { evObj, BubbleEvent });
            this._executeFormEvent(this.FormId + "OnDataSave", out BubbleEvent, new object[] { evObj, BubbleEvent });
            
            // Executa validações
            if(BubbleEvent)
            {
                BubbleEvent = this.ValidateHandler();
            }

            if(!BubbleEvent)
            {
                this.Addon.StatusErro("Nenhum dado pode ser salvo.");
            }
        }

        /// <summary>
        /// Chamado depois do update no registro UDO ser feito
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnAfterDataUpdate(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this._executeFormEvent(this.FormId + "OnAfterDataUpdate", out BubbleEvent, new object[] { evObj, BubbleEvent });
            this._executeFormEvent(this.FormId + "OnAfterDataSave", out BubbleEvent, new object[] { evObj, BubbleEvent });

            foreach(string dtsId in this.FormParams.SaveDatasources)
            {
                this.Addon.DtSources.saveUserDataSource(dtsId, evObj.FormUID);
            }
        }

        /// <summary>
        /// Executa as validações antes de salvar dados no banco.
        /// </summary>
        internal bool ValidateHandler()
        {
            bool res = true;

            // Limpa matrizes
            foreach(KeyValuePair<string, List<string>> mtx in this.MatrixEmptyRows)
            {
                foreach(string col in mtx.Value)
                {
                    this.MatrixClearEmptyRows(mtx.Key, col);
                }
            }

            foreach(KeyValuePair<string, List<string>> mtx in this.MatrixUniqueRows)
            {
                res = this.CheckUniqueValueColumn(mtx.Key);
            }

            // Executa validações
            if(res && this.HasToValidate)
            {
                res = this.Validate();
            }

            this.ValidateError = !res;

            return res;
        }

        /// <summary>
        /// Chamado antes de executar um delete no UDO. Implementa a confirmação com o usuário.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnBeforeDelete(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this._executeFormEvent(this.FormId + "OnBeforeDelete", out BubbleEvent, new object[] { evObj, BubbleEvent });
            if(BubbleEvent)
            {
                BubbleEvent = (this.Addon.SBO_Application.MessageBox("Tem certeza de que deseja remover esse registro?", 2, "Sim", "Não") == 1);
            }
        }

        /// <summary>
        /// Chamado após uma remoção de registro UDO.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnAfterDelete(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this._executeFormEvent(this.FormId + "OnDelete", out BubbleEvent, new object[] { evObj, BubbleEvent });
        }

        /// <summary>
        /// Chamado antes de fechar o form
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnFormClose(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this._executeFormEvent(this.FormId + "OnFormClose", out BubbleEvent, new object[] { evObj, BubbleEvent });
        }

        /// <summary>
        /// Chamado após o form ter sido fechado
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnAfterFormClose(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this._executeFormEvent(this.FormId + "OnAfterFormClose", out BubbleEvent, new object[] { evObj, BubbleEvent });
        }

        #endregion


        #region :: Eventos Opcionais

        /// <summary>
        /// Chamado quando o form aberto que estava em segundo plano é selecionado novamente
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnFormFocus(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this._executeFormEvent(this.FormId + "OnFormFocus", out BubbleEvent, new object[] { evObj, BubbleEvent });
        }

        /// <summary>
        /// Chamado quando outro form recebe o foco
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnFormLostFocus(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this._executeFormEvent(this.FormId + "OnFormLostFocus", out BubbleEvent, new object[] { evObj, BubbleEvent });
        }

        #endregion


        #region :: Eventos UDO

        /// <summary>
        /// Handler para ajustar o campo Code do datasource de um form UDO e executa o OnInsertUDO no form.
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void OnInsertUDO(ref SAPbouiCOM.MenuEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if(!String.IsNullOrEmpty(this.FormParams.BusinessObjectId))
            {
                try
                {
                    if(!this.CodeAfter)
                    {
                        string code = this.getNextCode(this.FormParams.MainDatasource, 5);
                        //this.Addon.StatusAlerta(code);
                    
                        this.UpdateOnClient(new Dictionary<string, dynamic>() { 
                            {"Code", code},
                            {"DocEntry", code},
                        }, this.FormParams.MainDatasource);

                        if(!String.IsNullOrEmpty(this.FormParams.Focus))
                        {
                            this.GetItem(this.FormParams.Focus).Click();
                        }

                        this.UDOCode = code;
                    }

                    FastOneItemEvent ev = new FastOneItemEvent()
                    {
                        FormTypeEx = this.FormId,
                        FormUID = this.FormId,
                        BeforeAction = false
                    };

                    // Executa o handler no form
                    this._executeFormEvent(this.FormId + "OnInsertUDO", out BubbleEvent, new object[] { ev, BubbleEvent });

                    if(BubbleEvent)
                    {
                        this.InInsertMode = true;
                    }

                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, " em OnInsertUDO de " + this.FormId);
                }
            }
        }

        #endregion


        #endregion



        #region :: Tabs

        /// <summary>
        /// Cria Tabs.
        /// By Labs - 12/2012
        /// </summary>
        private void createTabs(ref tabParams Tabs)
        {
            short i = 1;
            string lastTab = "";
            string tabId = "";
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Folder oTab = null;
            SAPbouiCOM.Form frm = this.SapForm;

            frm.DataSources.UserDataSources.Add("TabsDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);

            // Cria a área do tab:
            this.FormParams.TabArea = frm.Items.Add("TabArea", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            foreach(KeyValuePair<string, Dictionary<string, int>> tab in Tabs.Tabs)
            {
                try
                {
                    tabId = "tab" + i;
                    oItem = frm.Items.Add(tabId, SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                    oItem.Width = 100;
                    oItem.Height = 19;
                    oItem.AffectsFormMode = false;

                    oTab = ((SAPbouiCOM.Folder)(oItem.Specific));
                    oTab.Pane = i;
                    oTab.Caption = tab.Key;
                    oTab.DataBind.SetBound(true, "", "TabsDS");

                    // Distribui os componentes pela tab
                    foreach(KeyValuePair<string, int> comp in tab.Value)
                    {

                        if(this.spaces.IndexOf(comp.Key) != -1)
                        {
                            continue;
                        }

                        // ctrl
                        oItem = frm.Items.Item(comp.Key);
                        oItem.FromPane = i;
                        oItem.ToPane = i;

                        // Label
                        if(oItem.LinkTo != "")
                        {
                            oItem = frm.Items.Item(oItem.LinkTo);
                            oItem.FromPane = i;
                            oItem.ToPane = i;
                        }

                    }

                    // Registra swap:
                    if(i == 1)
                    {
                        oTab.Select();
                        this.FirstTab = oTab.Item;
                    } else
                    {
                        oTab.GroupWith(lastTab);
                    }
                    this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, this.FormId, tabId, "swapTabs");

                    // Registra actions:
                    if(Tabs.actions != null)
                    {
                        foreach(action action in Tabs.actions)
                        {
                            this.Addon.RegisterEvent(action.EventType, this.FormId, tabId, action.EventHandler);
                        }
                    }
                } catch
                {

                }

                lastTab = tabId;
                i++;
            }
            frm.PaneLevel = 1;
            GC.Collect();
        }

        /// <summary>
        /// Implementa onClick para Tabs.
        /// By Labs - 12/2012
        /// </summary>
        public bool swapTabs(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            SAPbouiCOM.Form frm = null;
            try
            {
                frm = this.Addon.SBO_Application.Forms.Item(evObj.FormUID);
                SAPbouiCOM.Folder tab = frm.Items.Item(evObj.ItemUID).Specific;
                frm.Freeze(true);
                frm.PaneLevel = tab.Pane;

            } catch(Exception e)
            {

            } finally
            {
                frm.Freeze(false);
                GC.Collect();
            }
            return BubbleEvent = true;
        }

        #endregion


        #region :: Criação de componentes

        /// <summary>
        /// Criação de componentes.
        /// By Labs - 12/2012
        /// </summary>
        /// <param name="compId"></param>
        /// <param name="frm"></param>
        /// <param name="comp"></param>
        /// <returns>SAPbouiCOM.Item</returns>
        public SAPbouiCOM.Item makeComp(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            if(!String.IsNullOrEmpty(comp.Id))
            {
                compId = comp.Id;
            }

            if(compId.Length > 10)
            {
                //compId = compId.Substring(0, 9);
                this.Addon.StatusErro("Nome de componente acima de 10 chars: " + compId);
                return null;
            }

            if(this.spaces.IndexOf(compId) != -1)
            {
                return null;
            }

            string formId = frm.UniqueID;
            bool isButton = (comp.Type == SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            bool isLabel = (comp.Type == SAPbouiCOM.BoFormItemTypes.it_STATIC);

            // Label
            this.lblIdCount++;
            string lblId = "lbl_" + this.lblIdCount;
            SAPbouiCOM.Item lbl = null;
            if((comp.Label != null) && (!isButton && !isLabel))
            {

                lbl = frm.Items.Add(lblId, SAPbouiCOM.BoFormItemTypes.it_STATIC);
                lbl.FromPane = comp.FromPane;
                lbl.ToPane = comp.ToPane;
                lbl.Visible = comp.Visible;

                SAPbouiCOM.StaticText st = ((SAPbouiCOM.StaticText)(lbl.Specific));
                st.Caption = comp.Label;

            }

            // Chama o metodo de criação adequado:
            comp.Id = compId;
            SAPbouiCOM.Item ctrl = this.compFactory(compId, ref frm, ref comp);

            // Configura o componente: 
            ctrl.FromPane = comp.FromPane; // 0;
            ctrl.ToPane = comp.ToPane; //0;

            // Bind e eventos
            this.setupComp(frm, comp, ctrl);

            if((comp.Label != null) && (!isButton && !isLabel))
            {
                ctrl.LinkTo = lblId;
                if(lbl != null)
                {
                    lbl.LinkTo = comp.Id;
                }
            }

            ctrl.RightJustified = comp.RightJustified;

            // Permite a customização do item criado por handler especifico:
            bool BubbleEvent = true;
            FastOneItemEvent evObj = new FastOneItemEvent()
            {
                FormUID = frm.UniqueID,
                ItemUID = ctrl.UniqueID
            };
            if(frm.IsSystem)
            {
                evObj.userFieldsHandler = (this.Addon.UserFields.ContainsKey(frm.TypeEx) ? this.Addon.UserFields[frm.TypeEx] : this.Addon.UserFields["UserFields"]);
                evObj.FormUID = frm.TypeEx;
            }
            evObj.FormMode = (int)frm.Mode;
            evObj.FormType = frm.Type;
            evObj.FormTypeEx = frm.TypeEx;
            evObj.FormTypeCount = frm.TypeCount;


            if(comp.onCreateHandler != null)
            {
                this.Addon.ExecEvent(comp.onCreateHandler, out BubbleEvent, new Object[] { evObj, BubbleEvent });

            // ou padrao
            } else
            {
              //  if(this.EventMethods.ContainsKey(ctrl.UniqueID + "OnCreate"))
              //  {
                    this.Addon.ExecEvent(ctrl.UniqueID + "OnCreate", out BubbleEvent, new Object[] { evObj, BubbleEvent });
              //  }
            }

            switch(comp.TipoEspecial)
            {
                case compSpecialType.cspComboSimNao:
                    this.comboSimNao((SAPbouiCOM.ComboBox)(ctrl.Specific));
                    break;

                case compSpecialType.cspComboMeses:
                    this.comboMeses((SAPbouiCOM.ComboBox)(ctrl.Specific));
                    break;

                case compSpecialType.cspComboDias:
                    this.comboDias((SAPbouiCOM.ComboBox)(ctrl.Specific));
                    break;

                case compSpecialType.cspComboHoras:
                    this.comboHoras((SAPbouiCOM.ComboBox)(ctrl.Specific));
                    break;

                case compSpecialType.cspComboMinutos:
                    this.comboMinutos((SAPbouiCOM.ComboBox)(ctrl.Specific));
                    break;
            }

            // Executa refresh:
            if(this.EventMethods.ContainsKey(ctrl.UniqueID + "OnRefresh"))
            {
                this.Addon.ExecEvent(ctrl.UniqueID + "OnRefresh", out BubbleEvent, new Object[] { evObj, BubbleEvent });
            }

            // Retorna:
            GC.Collect();
            return ctrl;
        }
        
        private void setupComp(string frmId, CompDefinition comp, SAPbouiCOM.Item ctrl)
        {
            SAPbouiCOM.Form frm = this.getForm(frmId);
            this.setupComp(frm, comp, ctrl);
        }

        /// <summary>
        /// Cria o bind de dados e de eventos do componente
        /// By Labs - 05/2013
        /// </summary>
        /// <param name="frm"></param>
        /// <param name="comp"></param>
        /// <param name="ctrl"></param>
        private void setupComp(SAPbouiCOM.Form frm, CompDefinition comp, SAPbouiCOM.Item ctrl)
        {
            bool isButton = (comp.Type == SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            bool isLabel = (comp.Type == SAPbouiCOM.BoFormItemTypes.it_STATIC);
            bool isMatrix = (comp.Type == SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            bool isGrid = (comp.Type == SAPbouiCOM.BoFormItemTypes.it_GRID);
            bool isHeader = (comp.Type == SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            

            // Bind de dados
            if(!isHeader && comp.BindTo != "_no_bind_")
            {
                if(!String.IsNullOrEmpty(comp.BindTo))
                {
                    try
                    {
                        if(!String.IsNullOrEmpty(comp.BindTable))
                        {
                            (ctrl.Specific).DataBind.SetBound(true, comp.BindTable, comp.BindTo);
                        } else
                        {
                            (ctrl.Specific).DataBind.SetBound(true, this.FormParams.MainDatasource, comp.BindTo);
                        }
                        ctrl.AffectsFormMode = true;
                    } catch(Exception e)
                    {
                        //this.Addon.StatusErro(e.Message);
                    }

                //} else if(!isButton && !isLabel && !isMatrix)
                } else if(!isButton && !isLabel && !isMatrix && !isGrid)
                {
                    try
                    {
                        frm.DataSources.UserDataSources.Add(ctrl.UniqueID, comp.UserDataType, comp.UserDataSize);
                    } catch {}
                    (ctrl.Specific).DataBind.SetBound(true, "", ctrl.UniqueID);
                }

                // ChooseFromList
                if(comp.Type == SAPbouiCOM.BoFormItemTypes.it_EDIT)
                {
                    if(comp.ChooseFromList != CFLType.cflNone)
                    {
                        this.AddDefChooseFromList(comp.ChooseFromList, comp, comp.Id, frm, comp.ChooseFromListUID, comp.ChooseFromListAlias);

                    } else
                    {
                        if(!String.IsNullOrEmpty(comp.ChooseFromListUID))
                        {
                            ((SAPbouiCOM.EditText)ctrl.Specific).ChooseFromListUID = comp.ChooseFromListUID;
                            ((SAPbouiCOM.EditText)ctrl.Specific).ChooseFromListAlias = comp.ChooseFromListAlias;
                        }
                    }

                    if(comp.LinkedObject != SAPbouiCOM.BoLinkedObject.lf_None || !String.IsNullOrEmpty(comp.LinkedObjectForm))
                    {
                        try
                        {
                            string lkid = comp.Id.Remove(0, 2);
                            lkid = "lk" + lkid;
                            SAPbouiCOM.Item link = null;
                            try
                            {
                                link = frm.Items.Add(lkid, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            } catch {
                                link = frm.Items.Item(lkid);
                            }
                            if(!String.IsNullOrEmpty(comp.LinkedObjectForm))
                            {
                                link.Description = comp.LinkedObjectForm;
                                this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_CLICK, frm.UniqueID, lkid, "CompOnLinkPressed");
                            } else
                            {
                                ((SAPbouiCOM.LinkedButton)link.Specific).LinkedObject = comp.LinkedObject;
                            }
                            link.LinkTo = comp.Id;
                            //link.Visible = true;

                        } catch(Exception e)
                        {
                            this.Addon.StatusErro(e.Message);
                        }
                    }
                }
            }

            // Validação
            if(comp.NonEmpty || comp.Validate != null) 
            //if(comp.NonEmpty || !String.IsNullOrEmpty(comp.Validate.OnEmptyError) ||
            //    !String.IsNullOrEmpty(comp.Validate.RangeIntMin) || !String.IsNullOrEmpty(comp.Validate.RangeIntMax) ||
            //    !String.IsNullOrEmpty(comp.Validate.RangeDateMin) || !String.IsNullOrEmpty(comp.Validate.RangeDateMax))
            {
                this.HasToValidate = true;
                this.ToValidate.Add(comp.Id);
            }

            // Actions
            this.setCompActions(ctrl.UniqueID, frm.TypeEx, comp);
        }

        /// <summary>
        /// Seta actions de componentes
        /// </summary>
        /// <param name="compId"></param>
        /// <param name="formId"></param>
        /// <param name="comp"></param>
        private void setCompActions(string compId, string frmId, CompDefinition comp)
        {

            // Registra actions definidos no form:
            if(comp.actions != null)
            {
                foreach(action action in comp.actions)
                {
                    this.Addon.RegisterEvent(action.EventType, frmId, compId, action.EventHandler);
                }
            }

            // Registra evento pra onClick:
            if(!String.IsNullOrEmpty(comp.onClickHandler))
            {
                this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, frmId, compId, comp.onClickHandler);

                // Se não houver, registra "onClicks" padrão por tipo de componente:
            } else
            {

                // Botão ganha 'onClick' de presente:
                string ev = compId;
                bool after = true;
                if(comp.Type == SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                {
                    if(compId == "1" || compId == "2")
                    {
                        ev = FormId + "Btn" + compId;
                        after = false;
                    }
                    this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, frmId, compId, ev + "OnClick", after);
                } else
                {
                    if(this.EventMethods.ContainsKey(ev + "OnClick"))
                    {
                        this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, frmId, compId, ev + "OnClick", after);
                    }
                }

                // Matrix recebe a capacidade de selecionar rows
                if(comp.Type == SAPbouiCOM.BoFormItemTypes.it_MATRIX)
                {

                    // Seta automaticamente um row como ativo ao se clicar em uma celula:
                    this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, frmId, compId, "selectMatrixRowOnClick", false);

                    this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS, frmId, compId, "selectMatrixRowOnFocus", false);

                    // Garante evento automatico de onRowClick: TRANSFERIDO PARA ATIVAR EM selectMatrixRowOnClick
                    // this.Addon.registerEvent(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, frmId, compId, compId + "OnRowClick");
                }

            }

            // Registra evento pra onKeydown:
            if(!String.IsNullOrEmpty(comp.onKeyDownHandler))
            {
                this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_KEY_DOWN, frmId, compId, comp.onKeyDownHandler);
            }

            // Evento no change:
            if(!String.IsNullOrEmpty(comp.onChangeHandler))
            {

                // OnChange pra combos
                if(comp.Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                {
                    this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT, frmId, compId, comp.onChangeHandler);

                    // OnChange para edits
                } else if(comp.Type == SAPbouiCOM.BoFormItemTypes.it_EDIT)
                {
                    this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS, frmId, compId, comp.onChangeHandler);
                }

            } else
            {
                // OnChange padrão pra combos
                if(comp.Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX && this.EventMethods.ContainsKey(compId + "onChange"))
                {
                    this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT, frmId, compId, compId + "onChange");
                }
            }

            // Evento onExit
            if((comp.onExitHandler != null) && (comp.onExitHandler != ""))
            {
                this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS, frmId, compId, comp.onExitHandler);
            } else
            {
                if(this.EventMethods.ContainsKey(compId + "OnExit"))
                {
                    this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS, frmId, compId, compId + "OnExit");
                }
            }
        }

        /// <summary>
        /// Abre um form de UDO no click de um link no componente
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void CompOnLinkPressed(SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string formClass = "";
            try
            {
                SAPbouiCOM.Item link = this.GetItem(evObj.ItemUID);
                SAPbouiCOM.Item comp = this.GetItem(link.LinkTo);
                formClass = link.Description;

                this.Addon.OpenFormUDOFind(formClass, ((SAPbouiCOM.EditText)comp.Specific).Value);
            
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " abrindo form '" + formClass + "'");
            }
        }


        #endregion


        #region :: Comp Factory

        /// <summary>
        /// Factory para criação de componentes.
        /// By Labs - 12/2012
        /// </summary>
        /// <param name="compId"></param>
        /// <param name="formId"></param>
        /// <param name="comp"></param>
        /// <returns>SAPbouiCOM.Item</returns>
        public SAPbouiCOM.Item compFactory(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            if(String.IsNullOrEmpty(compId))
            {
                this.Addon.StatusErro("compFactory: compId não pode ser nulo.");
                return null;
            }

            // Chama o metodo de criação adequado:
            SAPbouiCOM.Item ctrl = null;
            bool botao = (comp.Type == SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            switch(comp.Type)
            {
                case SAPbouiCOM.BoFormItemTypes.it_ACTIVE_X: ctrl = this.makeActivex(compId, ref frm, ref comp); break;
                case SAPbouiCOM.BoFormItemTypes.it_BUTTON: ctrl = this.makeButton(compId, ref frm, ref comp); break;
                case SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO: ctrl = this.makeButtonCombo(compId, ref frm, ref comp); break;
                case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX: ctrl = this.makeCheckBox(compId, ref frm, ref comp); break;
                case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX: ctrl = this.makeComboBox(compId, ref frm, ref comp); break;
                case SAPbouiCOM.BoFormItemTypes.it_EDIT: ctrl = this.makeEdit(compId, ref frm, ref comp); break;
                case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT: ctrl = this.makeExtEdit(compId, ref frm, ref comp); break;
                case SAPbouiCOM.BoFormItemTypes.it_FOLDER: ctrl = this.makeFolder(compId, ref frm, ref comp); break;
                case SAPbouiCOM.BoFormItemTypes.it_GRID: ctrl = this.makeGrid(compId, ref frm, ref comp); break;
                case SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON: ctrl = this.makeLinkedButton(compId, ref frm, ref comp); break;
                case SAPbouiCOM.BoFormItemTypes.it_MATRIX: ctrl = this.makeMatrix(compId, ref frm, ref comp); break;
                case SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON: ctrl = this.makeOptionButton(compId, ref frm, ref comp); break;
                case SAPbouiCOM.BoFormItemTypes.it_PANE_COMBO_BOX: ctrl = this.makePaneComboBox(compId, ref frm, ref comp); break;
                case SAPbouiCOM.BoFormItemTypes.it_PICTURE: ctrl = this.makePicture(compId, ref frm, ref comp); break;
                case SAPbouiCOM.BoFormItemTypes.it_RECTANGLE: ctrl = this.makeRectangle(compId, ref frm, ref comp); break;
                case SAPbouiCOM.BoFormItemTypes.it_STATIC: ctrl = this.makeStatic(compId, ref frm, ref comp); break;
            }

            ctrl.Visible = comp.Visible;
            ctrl.Enabled = comp.Enabled;

            //ctrl.Width = comp.force_width;

            return ctrl;
        }

        private void noBind(string frmId, string compId)
        {
            if(this.FormId == frmId)
            {
                this.FormParams.Controls[compId].BindTo = "_no_bind_";
            }
        }

        private SAPbouiCOM.Item makeActivex(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            SAPbouiCOM.Item ctrl = frm.Items.Add(compId, SAPbouiCOM.BoFormItemTypes.it_ACTIVE_X);
            return ctrl;
        }

        private SAPbouiCOM.Item makeButton(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            SAPbouiCOM.Item ctrl = frm.Items.Add(compId, SAPbouiCOM.BoFormItemTypes.it_BUTTON);

            // Sem bind
            this.noBind(frm.UniqueID, compId);

            // Em botões, o Label vira o Caption
            (ctrl.Specific).Caption = comp.Caption;

            return ctrl;
        }

        private SAPbouiCOM.Item makeButtonCombo(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            SAPbouiCOM.Item ctrl = frm.Items.Add(compId, SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO);

            // Sem bind
            this.noBind(frm.UniqueID, compId);

            return ctrl;
        }

        private SAPbouiCOM.Item makeCheckBox(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            SAPbouiCOM.Item ctrl = frm.Items.Add(compId, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            return ctrl;
        }

        private SAPbouiCOM.Item makeComboBox(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            SAPbouiCOM.Item ctrl = frm.Items.Add(compId, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            ctrl.DisplayDesc = comp.DisplayDesc;

            if(!String.IsNullOrEmpty(comp.PopulateSQL))
            {
                this.populateCombo(ctrl.Specific, comp.PopulateSQL, comp.FirstKey, comp.FirstValue, comp.DefValue);
            } else if(null != comp.PopulateItens)
            {
                this.populateCombo(ctrl.Specific, comp.PopulateItens, comp.DefValue);
            }

            return ctrl;
        }

        private SAPbouiCOM.Item makeEdit(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            SAPbouiCOM.Item ctrl = frm.Items.Add(compId, SAPbouiCOM.BoFormItemTypes.it_EDIT);
            return ctrl;
        }

        private SAPbouiCOM.Item makeExtEdit(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            SAPbouiCOM.Item ctrl = frm.Items.Add(compId, SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
            return ctrl;
        }

        private SAPbouiCOM.Item makeFolder(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            SAPbouiCOM.Item ctrl = frm.Items.Add(compId, SAPbouiCOM.BoFormItemTypes.it_FOLDER);

            // Sem bind
            this.noBind(frm.UniqueID, compId);
            return ctrl;
        }

        private SAPbouiCOM.Item makeGrid(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            SAPbouiCOM.Item ctrl = frm.Items.Add(compId, SAPbouiCOM.BoFormItemTypes.it_GRID);
            return ctrl;
        }

        private SAPbouiCOM.Item makeLinkedButton(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            SAPbouiCOM.Item ctrl = frm.Items.Add(compId, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);

            // Sem bind
            this.noBind(frm.UniqueID, compId);
            return ctrl;
        }

        private SAPbouiCOM.Item makeMatrix(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            SAPbouiCOM.Item ctrl = frm.Items.Add(compId, SAPbouiCOM.BoFormItemTypes.it_MATRIX);

            // Sem bind
            this.noBind(frm.UniqueID, compId);

            return ctrl;
        }

        private SAPbouiCOM.Item makeOptionButton(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            SAPbouiCOM.Item ctrl = frm.Items.Add(compId, SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
            return ctrl;
        }

        private SAPbouiCOM.Item makePaneComboBox(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            SAPbouiCOM.Item ctrl = frm.Items.Add(compId, SAPbouiCOM.BoFormItemTypes.it_PANE_COMBO_BOX);
            return ctrl;
        }

        private SAPbouiCOM.Item makePicture(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            SAPbouiCOM.Item ctrl = frm.Items.Add(compId, SAPbouiCOM.BoFormItemTypes.it_PICTURE);
            return ctrl;
        }

        private SAPbouiCOM.Item makeRectangle(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            SAPbouiCOM.Item ctrl = frm.Items.Add(compId, SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);

            // Sem bind
            this.noBind(frm.UniqueID, compId);

            return ctrl;
        }

        private SAPbouiCOM.Item makeStatic(string compId, ref SAPbouiCOM.Form frm, ref CompDefinition comp)
        {
            SAPbouiCOM.Item ctrl = frm.Items.Add(compId, SAPbouiCOM.BoFormItemTypes.it_STATIC);

            // Sem bind
            this.noBind(frm.UniqueID, compId);

            // Em Static, o Label vira o Caption
            (ctrl.Specific).Caption = comp.Caption;

            return ctrl;
        }

        #endregion


        #region :: Componentes Especiais


        #region :: Matrix


        #region :: Eventos

        /// <summary>
        /// Seleciona automaticamente o row de uma matriz ao se clicar em qualquer de suas células
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void selectMatrixRowOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Matrix matrix = this.GetItem(evObj.ItemUID).Specific; // this.Addon.SBO_Application.Forms.ActiveForm.Items.Item(evObj.ItemUID).Specific;
                if(evObj.Row > 0 && evObj.Row <= matrix.RowCount)
                {
                    bool multi = (evObj.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_SHIFT || evObj.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_CTRL);
                    matrix.SelectRow(evObj.Row, true, multi);


                    // Garante evento automatico de onRowClick: 
                    // SE SE SENTIR TENTADO A MUDAR DAQUI, É PQ VC PRECISA NA VERDADE DE UM onClick NO "COMPONENTE" TODO, E NÃO NO ROW.
                    this.Addon.ExecEvent(evObj.ItemUID + "OnRowClick", out BubbleEvent, new Object[] { evObj, BubbleEvent });
                }
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - selectMatrixRowOnClick de " + evObj.ItemUID);
            } finally
            {
                GC.Collect();
            }
        }

        public void selectMatrixRowOnFocus(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Matrix matrix = this.GetItem(evObj.ItemUID).Specific; // this.Addon.SBO_Application.Forms.ActiveForm.Items.Item(evObj.ItemUID).Specific;
                if(evObj.Row > 0 && evObj.Row <= matrix.RowCount)
                {
                    bool multi = (evObj.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_SHIFT || evObj.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_CTRL);
                    matrix.SelectRow(evObj.Row, true, multi);

                    // Garante evento automatico de onRowClick: 
                    // SE SE SENTIR TENTADO A MUDAR DAQUI, É PQ VC PRECISA NA VERDADE DE UM onClick NO "COMPONENTE" TODO, E NÃO NO ROW.
                    //this.Addon.execEvent(evObj.ItemUID + "OnRowClick", out BubbleEvent, new Object[] { evObj, BubbleEvent });

                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - selectMatrixRowOnFocus de " + evObj.ItemUID);
            } finally
            {
                GC.Collect();
            }
        }

        /// <summary>
        /// Garante valores únicos emcolunas da matrix
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void MatrixOnCheckUnique(SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if(((List<string>)this.FormParams.Controls[evObj.ItemUID].ExtraData["Unique"]).Contains(evObj.ColUID))
            {
                if(!this.CheckUniqueValueColumn(evObj.ItemUID, evObj.ColUID, evObj.Row))
                {
                    this.Addon.ShowMessage("O valor informado está duplicado. Favor selecionar outro.");
                    BubbleEvent = false;
                }
            }
        }

        /// <summary>
        /// Garante valores únicos em colunas da matrix
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public bool CheckUniqueValueColumn(string mtxId, string colId, int row, string valor = "")
        {
            bool res = true;

            // Verifica se já tem o mesmo valor na matriz
            SAPbouiCOM.Matrix mtx = this.GetItem(mtxId).Specific;
            if(String.IsNullOrEmpty(valor))
            {
                valor = mtx.GetCellSpecific(colId, row).Value;
            }

            for(int r = 1; r <= mtx.RowCount; r++)
            {
                if(r != row)
                {
                    if(mtx.GetCellSpecific(colId, r).Value == valor)
                    {
                        res = false;
                        try
                        {
                            mtx.SetLineData(row);
                        } catch (Exception e) {
                            this.Addon.DesenvTimeError(e, "CheckUniqueValueColumn");
                        }
                        break;
                    }
                }
            }

            return res;
        }

        /// <summary>
        /// Garante valores únicos em colunas da matrix
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public bool CheckUniqueValueColumn(string mtxId)
        {
            bool res = true;

            if(this.FormParams.Controls[mtxId].ExtraData.ContainsKey("Unique") && ((List<string>)this.FormParams.Controls[mtxId].ExtraData["Unique"]).Count > 0)
            {

                Dictionary<string, List<dynamic>> uniques = new Dictionary<string, List<dynamic>>();

                // Verifica se já tem o mesmo valor na matriz
                SAPbouiCOM.Matrix mtx = this.GetItem(mtxId).Specific;
                for(int r = 1; r <= mtx.RowCount; r++)
                {
                    foreach(string col in ((List<string>)this.FormParams.Controls[mtxId].ExtraData["Unique"]))
                    {
                        string val = mtx.GetCellSpecific(col, r).Value;
                        if(!uniques.ContainsKey(col))
                        {
                            uniques.Add(col, new List<dynamic>());
                        }

                        if(!uniques[col].Contains(val))
                        {
                            uniques[col].Add(val);
                        } else
                        {
                            res = false;
                            this.ShowMessage("Existem valores duplicados na coluna '" + mtx.Columns.Item(col).Title + "'");
                            break;
                        }
                    }
                }
            }

            return res;
        }

        /// <summary>
        /// Garante valores únicos em colunas da matrix
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        private bool CheckUniqueValueDataTable(SAPbouiCOM.DataTable dts, string colId, int row, string valor)
        {
            bool res = true;

            // Verifica se já tem o mesmo valor no dataset
            for(int r = 0; r < dts.Rows.Count; r++)
            {
                if(r != row)
                {
                    if(dts.GetValue(colId, r) == valor)
                    {
                        res = false;
                        break;
                    }
                }
            }

            return res;
        }

        /// <summary>
        /// Abre um form de UDO no click de um link na coluna da matriz
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void MatrixOnLinkPressed(SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string formClass = "";
            try
            {
                if(((Dictionary<string, string>)this.FormParams.Controls[evObj.ItemUID].ExtraData["LinkedUDOs"]).ContainsKey(evObj.ColUID))
                {
                    string code = ((SAPbouiCOM.Matrix)this.GetItem(evObj.ItemUID).Specific).GetCellSpecific(evObj.ColUID, evObj.Row).Value;
                    if(!String.IsNullOrEmpty(code))
                    {
                        formClass = ((Dictionary<string, string>)this.FormParams.Controls[evObj.ItemUID].ExtraData["LinkedUDOs"])[evObj.ColUID];
                        this.Addon.OpenFormUDOFind(formClass, code);
                    }
                }
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " abrindo form '" + formClass + "'");
            }
        }

        #endregion


        #region :: Criação
        
        public void SetupMatrixOnTab(string tab, string mtxId, string tbId, List<ColumnDefinition> columns, bool UsingDataTable = false, string DataTableSQL = "")
        {
            // Registra o click no tab
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_CLICK, this.FormId, tab, "SetupMatrixOnTabClick", false);

            // Armazena
            this.MatrixParams.Add(tab, new SetupMatrixParams()
            {
                mtxId = mtxId,
                tbId = tbId,
                columns = columns,
                UsingDataTable = UsingDataTable,
                DataTableSQL = DataTableSQL
            });

        }

        /// <summary>
        /// Cria dinamicamente uma matrix no primeiro click de uma tab
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void SetupMatrixOnTabClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (this.MatrixParams.ContainsKey(evObj.ItemUID)) 
            {
                try
                {
                    this.SapForm.Freeze(true);
                    SetupMatrixParams mtxParams = this.MatrixParams[evObj.ItemUID];
                    this.SetupMatrix(mtxParams.mtxId, mtxParams.tbId, mtxParams.columns, mtxParams.UsingDataTable, mtxParams.DataTableSQL);
                    this.formResize(this.FormId);

                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, "Criando a matrix em SetupMatrixOnTabClick");
                } finally
                {
                    this.SapForm.Freeze(false);
                    this.MatrixParams.Remove(evObj.ItemUID);
                }
            }

        }



        /// <summary>
        /// Configura uma matriz passando seu id e columns 
        /// - ESSA VERSÃO NÃO DEVE SER USADO EM USERFIELDS - 
        /// By Labs - 04/2014
        /// </summary>
        /// <param name="mtxId">Id da matrix - NÃO DEVE SER USADO EM USERFIELDS</param>
        /// <param name="tbId"></param>
        /// <param name="columns"></param>
        /// <param name="UsingDataTable">TRUE se os dados dessa matriz vem de um DataTables</param>
        public SAPbouiCOM.Matrix SetupMatrix(string mtxId, string tbId, List<ColumnDefinition> columns, bool UsingDataTable = false, string DataTableSQL = "", SAPbouiCOM.Form frm = null)
        {
            SAPbouiCOM.Matrix matrix = this.GetItem(mtxId).Specific;
            SAPbouiCOM.Column column = null;
            
          //  this.Addon.DesenvTimeInfo(this.FormId + " Start: SetupMatrix de " + mtxId);

            if(frm == null)
            {
                frm = this.SapForm;
            }

            try
            {
                bool addWidth = false;
                if(!this.FormParams.Controls.ContainsKey(mtxId))
                {
                    this.FormParams.Controls[mtxId] = new CompDefinition();
                }

                // Ajusta datatable
                if(!String.IsNullOrEmpty(tbId))
                {
                    if(UsingDataTable)
                    {
                        try
                        {
                            SAPbouiCOM.DataTable dt = frm.DataSources.DataTables.Add(tbId);
                            if(!String.IsNullOrEmpty(DataTableSQL))
                            {
                                dt.ExecuteQuery(DataTableSQL);
                            }
                            if(!this.FormParams.Controls[mtxId].ExtraData.ContainsKey("DataTable"))
                            {
                                this.FormParams.Controls[mtxId].ExtraData.Add("DataTable", tbId);
                            }

                        } catch(Exception e)
                        {
                            this.Addon.DesenvTimeError(e, "Criando datatable em SetupMatrix de " + mtxId + "\nSQL: " + DataTableSQL);
                            return matrix;
                        }

                    } else
                    {
                       // if(!this.FormParams.ExtraDatasources.Contains(tbId))  Userfields precisa disso
                        //{
                            try
                            {
                                frm.DataSources.DBDataSources.Add(tbId);
                                //this.FormParams.ExtraDatasources.Add(tbId);
                            } catch(Exception e)
                            {
                        //        this.Addon.DesenvTimeError(e, " - Erro adicionando " + tbId + " ao DBDatasources em SetupMatrix " + this.FormId);
                            }
                        //}
                    }
                }


                if(this.FormParams.Controls[mtxId].Columns == null)
                {
                    this.FormParams.Controls[mtxId].Columns = new columnParams();
                    this.FormParams.Controls[mtxId].Columns.Widths = new List<int>();
                    addWidth = true;
                }

                matrix.Layout = SAPbouiCOM.BoMatrixLayoutType.mlt_Normal;
                matrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                foreach(ColumnDefinition colParams in columns)
                {
                    if(colParams.ChooseFromList != CFLType.cflNone)
                    {
                        colParams.LinkedObject = this.GetLinkedObjByCLF(colParams.ChooseFromList);
                    }

                    if(!String.IsNullOrEmpty(colParams.LinkedObjectType) || colParams.LinkedObject != SAPbouiCOM.BoLinkedObject.lf_None)
                    {
                        colParams.Type = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON;
                        
                        if(colParams.LinkedObject == SAPbouiCOM.BoLinkedObject.lf_UserDefinedObject && !String.IsNullOrEmpty(colParams.LinkedObjectForm))
                        {
                            colParams.LinkedObjectType = colParams.ChooseFromListUDOName;

                            if(!this.FormParams.Controls[mtxId].ExtraData.ContainsKey("LinkedUDOs"))
                            {
                                this.FormParams.Controls[mtxId].ExtraData.Add("LinkedUDOs", new Dictionary<string, string>());
                            }
                            if(!((Dictionary<string, string>)this.FormParams.Controls[mtxId].ExtraData["LinkedUDOs"]).ContainsKey(colParams.Id))
                            {
                                ((Dictionary<string, string>)this.FormParams.Controls[mtxId].ExtraData["LinkedUDOs"]).Add(colParams.Id, colParams.LinkedObjectForm);
                            }
                            this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED, frm.UniqueID /*this.FormId*/, mtxId, "MatrixOnLinkPressed");
                        }
                    }

                    column = matrix.Columns.Add(colParams.Id, colParams.Type);

                    this.AddColumn(frm, mtxId, column, colParams, (colParams.Bind ? tbId : ""), UsingDataTable);

                    // Registra colunas únicas
                    string unique = colParams.Id; // (!String.IsNullOrEmpty(colParams.BindTo) ? colParams.BindTo : colParams.Id);
                    if(colParams.Unique)
                    {
                        if(!this.FormParams.Controls[mtxId].ExtraData.ContainsKey("Unique"))
                        {
                            this.FormParams.Controls[mtxId].ExtraData.Add("Unique", new List<string>());
                        }
                        if(!((List<string>)this.FormParams.Controls[mtxId].ExtraData["Unique"]).Contains(unique))
                        {
                            ((List<string>)this.FormParams.Controls[mtxId].ExtraData["Unique"]).Add(unique);
                        }
                        
                        if(!this.MatrixUniqueRows.ContainsKey(mtxId))
                        {
                            this.MatrixUniqueRows.Add(mtxId, new List<string>());
                        }
                        if(!this.MatrixUniqueRows[mtxId].Contains(unique))
                        {
                            this.MatrixUniqueRows[mtxId].Add(unique);
                        }

                        this.Addon.RegisterEventHandler(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT, frm.UniqueID /*this.FormId*/, mtxId, "MatrixOnCheckUnique");
                    }

                    // Registra colunas mandatórias
                    if(colParams.NonEmpty)
                    {
                        if(!this.MatrixEmptyRows.ContainsKey(mtxId))
                        {
                            this.MatrixEmptyRows.Add(mtxId, new List<string>());
                        }
                        if(!this.MatrixEmptyRows[mtxId].Contains(unique))
                        {
                            this.MatrixEmptyRows[mtxId].Add(unique);
                        }
                    }

                    if(addWidth)
                    {
                        this.FormParams.Controls[mtxId].Columns.Widths.Add(colParams.Percent != 0 ? colParams.Percent : colParams.Width);
                    }
                }

                matrix.LoadFromDataSourceEx();

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Erro ao configurar a matriz ");
            }

           // this.Addon.DesenvTimeInfo(this.FormId + " End: SetupMatrix de " + mtxId);

            return matrix;
        }

        internal SAPbouiCOM.Column AddColumn(SAPbouiCOM.Form frm, string mtxid, SAPbouiCOM.Column column, ColumnDefinition colParams, string tbId = "", bool dynamic_bind = false)
        {
            try
            {
                // Configura o básico:
                column.Description      = colParams.Caption;
                column.DisplayDesc      = colParams.DisplayDesc;
                column.Editable         = colParams.Enabled;
                column.Visible          = colParams.Visible;
                column.TitleObject.Caption = colParams.Caption;
                column.AffectsFormMode  = colParams.AffectsFormMode;
                column.RightJustified   = colParams.RightJustified;
                column.ValOff           = colParams.ValOff;
                column.ValOn            = colParams.ValOn;

                if(colParams.Width > 0)
                {
                    column.Width = colParams.Width;
                }

                if(colParams.ForeColor > 0)
                {
                    column.ForeColor = colParams.ForeColor;
                }

                // Bound:
                if(colParams.Bind)
                {
                    if(!String.IsNullOrEmpty(tbId))
                    {
                        string fld = (String.IsNullOrEmpty(colParams.BindTo) ? colParams.Id : colParams.BindTo);

                        if(dynamic_bind)
                        {
                            column.DataBind.Bind(tbId, fld);

                        } else
                        {
                            try
                            {
                                column.DataBind.SetBound(true, tbId, fld);

                            } catch(Exception e)
                            {
                                column.DataBind.Bind(tbId, fld);
                            }
                        }
                    } else
                    {
                        frm.DataSources.UserDataSources.Add(colParams.Id, colParams.UserDataType, colParams.UserDataSize);
                        column.DataBind.SetBound(true, "", colParams.Id);
                    }
                }

                // Populate via SQL:
                if(!String.IsNullOrEmpty(colParams.PopulateSQL))
                {
                    this.populateColumn(mtxid, ref column, colParams.PopulateSQL, colParams.FirstKey, colParams.FirstValue);
                }

                // Populate via Values
                if(colParams.PopulateItens != null)
                {
                    this.populateColumn(ref column, colParams.PopulateItens);
                }

                // ChooseFromList
                if(colParams.ChooseFromList != CFLType.cflNone)
                {
                    this.AddDefChooseFromList(colParams.ChooseFromList, colParams, null, null, colParams.ChooseFromListUID, colParams.ChooseFromListAlias);
                    Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST, frm.TypeEx, mtxid, mtxid + "OnChooseFromList");
                }
                if(!String.IsNullOrEmpty(colParams.ChooseFromListUID))
                {
                    column.ChooseFromListUID = colParams.ChooseFromListUID;
                    column.ChooseFromListAlias = colParams.ChooseFromListAlias;
                }

                // Footer:
                column.ColumnSetting.SumType = colParams.SumType;
                column.ColumnSetting.SumValue = colParams.SumValue;

                if((!String.IsNullOrEmpty(colParams.LinkedObjectType) || colParams.LinkedObject != SAPbouiCOM.BoLinkedObject.lf_None) && column.ExtendedObject != null)
                {
                    
                    SAPbouiCOM.LinkedButton link = column.ExtendedObject;
                    if(String.IsNullOrEmpty(colParams.LinkedObjectType) && colParams.LinkedObject != SAPbouiCOM.BoLinkedObject.lf_UserDefinedObject)
                    {
                        link.LinkedObject = colParams.LinkedObject;
                    } else
                    {
                        link.LinkedObjectType = colParams.LinkedObjectType;
                    }
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, "Acrescentando coluna: " + mtxid + "-" + colParams.Id + " | " + tbId);
            }

            return column;
        }

        /// <summary>
        ///  Pupula os ítens de uma coluna via SQL e garante sua atualização
        ///  em refreshs.
        /// </summary>
        /// <param name="matrixId"></param>
        /// <param name="column"></param>
        /// <param name="sql"></param>
        /// <param name="first_key"></param>
        /// <param name="first_value"></param>
        public bool populateColumn(string matrixId, ref SAPbouiCOM.Column column, string sql, string first_key = "", string first_value = "")
        {
            return this._populateColumn(ref column, matrixId, sql, null, first_key, first_value);
        }

        /// <summary>
        /// Popula os itens de uma coluna via params.
        /// By Labs - 07/2013
        /// </summary>
        /// <param name="matrixId"></param>
        /// <param name="column">Coluna do tipo it_COMBO_BOX</param>
        /// <param name="itens">Objeto contendo os valores dos itens do combo</param>
        public bool populateColumn(ref SAPbouiCOM.Column column, Dictionary<string, string> itens)
        {
            return this._populateColumn(ref column, "", "", itens);
        }

        /// <summary>
        /// Executa o preenchimento de itens em uma coluna combobox de uma matrix.
        /// </summary>
        /// <param name="column"></param>
        /// <param name="mtxId"></param>
        /// <param name="sql"></param>
        /// <param name="itens"></param>
        /// <param name="first_key"></param>
        /// <param name="first_value"></param>
        internal bool _populateColumn(ref SAPbouiCOM.Column column, string mtxId = "", string sql = "", Dictionary<string, string> itens = null, string first_key = "", string first_value = "Selecione...")
        {
            bool res = false;

            // Valida se a coluna é combobox:
            if(column.Type != SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            {
                this.Addon.StatusErro("Não é possível popular coluna que não seja do tipo it_COMBO_BOX");
                return res;
            }

            // Limpa os itens pré-existentes:
            try
            {
                while(column.ValidValues.Count > 0)
                {
                    column.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }
            } catch(Exception e)
            {
            }

            // Acrescenta valor inicial:
            //if(!String.IsNullOrEmpty(first_key))
            //{
                column.ValidValues.Add(first_key, first_value);
            //}


            // Popula via SQL:
            if(!String.IsNullOrEmpty(sql))
            {
                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                try
                {
                    rec.DoQuery(sql);
                    rec.MoveFirst();
                    while(!rec.EoF)
                    {
                        try
                        {
                            column.ValidValues.Add(rec.Fields.Item(0).Value, rec.Fields.Item(1).Value);
                        } catch(Exception e)
                        {
                            this.Addon.DesenvTimeError(e, " - Acrescentando valores ao combo " + column.UniqueID + " da matriz " + mtxId);
                        }
                        rec.MoveNext();
                        res = true;
                    }
                    rec = null;

                    // Registra para refresh automático:
                    if(!this.MatrixRefreshColumns.ContainsKey(mtxId))
                    {
                        this.MatrixRefreshColumns[mtxId] = new Dictionary<string, List<string>>() { 
                            {column.UniqueID, new List<string>{ sql, first_key, first_value }}
                        };
                    }
                } catch(Exception e)
                {
                    this.Addon.StatusErro(e.Message);
                    res = false;
                }

                // Popula via itens:
            } else if(itens != null)
            {
                foreach(KeyValuePair<string, string> item in itens)
                {
                    column.ValidValues.Add(item.Key, item.Value);
                }
            }

            GC.Collect();
            return res;
        }


        /// <summary>
        /// Limpa valores de combo em uma coluna
        /// </summary>
        /// <param name="column"></param>
        public void clearColumns(SAPbouiCOM.Column column)
        {
            try
            {
                for(int i = column.ValidValues.Count - 1; i > -1; i--)
                {
                    column.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                }
            } catch(Exception e)
            {
                // this.Addon.DesenvTimeError(e, " - Limpndo ítens no combo " + combo.Item.UniqueID);
            }
        }

        /// <summary>
        /// Preenche os valores do combo de uma coluna com base nos valores em um DATATABLE
        /// </summary>
        /// <param name="mtxId"></param>
        /// <param name="columnId"></param>
        /// <param name="DataTable"></param>
        /// <param name="key"></param>
        /// <param name="val"></param>
        /// <param name="first_key"></param>
        /// <param name="first_value"></param>
        /// <param name="def_value"></param>
        public bool populateColumnDataTable(string mtxId, string columnId, string DataTable, string key, string val, string first_key = "", string first_value = "", string def_value = "")
        {
            SAPbouiCOM.Matrix matrix = this.GetItem(mtxId).Specific;
            return this.populateColumnDataTable(matrix, columnId, DataTable, key, val, first_key, first_value, def_value);
        }

        /// <summary>
        /// Preenche os valores do combo de uma coluna com base nos valores em um DATATABLE
        /// </summary>
        /// <param name="matrix"></param>
        /// <param name="columnId"></param>
        /// <param name="DataTable"></param>
        /// <param name="key"></param>
        /// <param name="val"></param>
        /// <param name="first_key"></param>
        /// <param name="first_value"></param>
        /// <param name="def_value"></param>
        public bool populateColumnDataTable(SAPbouiCOM.Matrix matrix, string columnId, string DataTable, string key, string val, string first_key = "", string first_value = "", string def_value = "")
        {
            bool res = false;
            SAPbouiCOM.Column column = matrix.Columns.Item(columnId);

            // Limpa            
            this.clearColumns(column);

            //if(!String.IsNullOrEmpty(first_value))
            //{
                try
                {
                    column.ValidValues.Add(first_key, first_value);
                } catch { }
            //}

            try
            {
                SAPbouiCOM.DataTable dt = this.SapForm.DataSources.DataTables.Item(DataTable);
                if(!dt.IsEmpty)
                {
                    for(int r = 0; r < dt.Rows.Count; r++)
                    {
                        try
                        {
                            dynamic key_val = dt.GetValue(key, r);
                            dynamic value = dt.GetValue(val, r);
                            
                            if (!String.IsNullOrEmpty(key_val)){
                                column.ValidValues.Add(key_val, value);
                            }

                        } catch(Exception e)
                        {
                            this.Addon.DesenvTimeError(e, "populateColumnDataTable");

                        }

                    }
                    res = true;
                }
            } catch { }


            GC.Collect();
            return res;
        }

        /// <summary>
        /// Preenche os valores do combo de uma coluna com base nos valores em um DBDataSource
        /// </summary>
        /// <param name="mtxId"></param>
        /// <param name="columnId"></param>
        /// <param name="DataSource"></param>
        /// <param name="key"></param>
        /// <param name="val"></param>
        /// <param name="flushMatrix"></param>
        /// <param name="first_key"></param>
        /// <param name="first_value"></param>
        /// <param name="def_value"></param>
        public bool populateColumnDataSource(string mtxId, string columnId, string DataSource, string key, string val, string flushMatrix = "", string first_key = "", string first_value = "", string def_value = "")
        {
            SAPbouiCOM.Matrix matrix = this.GetItem(mtxId).Specific;
            return this.populateColumnDataSource(matrix, columnId, DataSource, key, val, flushMatrix, first_key, first_value, def_value);
        }

        /// <summary>
        /// Preenche os valores do combo de uma coluna com base nos valores em um DBDataSource
        /// </summary>
        /// <param name="matrix"></param>
        /// <param name="columnId"></param>
        /// <param name="DataSource"></param>
        /// <param name="key"></param>
        /// <param name="val"></param>
        /// <param name="flushMatrix"></param>
        /// <param name="first_key"></param>
        /// <param name="first_value"></param>
        /// <param name="def_value"></param>
        public bool populateColumnDataSource(SAPbouiCOM.Matrix matrix, string columnId, string DataSource, string key, string val, string flushMatrix = "", string first_key = "", string first_value = "", string def_value = "")
        {
            bool res = false;
            SAPbouiCOM.Column column = matrix.Columns.Item(columnId);

            // Limpa            
            this.clearColumns(column);

           // if(!String.IsNullOrEmpty(first_value))
            //{
                try
                {
                    column.ValidValues.Add(first_key, first_value);
                } catch { }
            //}

            try
            {
                if(!String.IsNullOrEmpty(flushMatrix))
                {
                    SAPbouiCOM.Matrix mtx = this.GetItem(flushMatrix).Specific;
                    mtx.FlushToDataSource();
                }

                SAPbouiCOM.DBDataSource dt = this.SapForm.DataSources.DBDataSources.Item(DataSource);
                for(int r = 0; r < dt.Size; r++)
                {
                    try
                    {
                        column.ValidValues.Add(dt.GetValue(key, r), dt.GetValue(val, r));
                    } catch(Exception e)
                    {
                        this.Addon.DesenvTimeError(e, "populateColumnDataSource");
                    }
                }
                res = true;

            } catch { 
                
            }

            GC.Collect();

            return res;
        }



        #endregion


        #region :: Refresh

        /// <summary>
        /// Refresh em matrix
        /// </summary>
        /// <param name="matrixId">Id da Matrix</param>
        /// <param name="tbId">Id da tabela em DBDataSources</param>
        /// <param name="cond">Conditions para o refresh</param>
        public void RefreshMatrix(string matrixId, string tbId, SAPbouiCOM.Conditions cond)
        {
            this._refreshMatrix(ref this.SapForm, matrixId, tbId, cond);
        }

        /// <summary>
        /// Refresh em matrix
        /// </summary>
        /// <param name="matrixId">Id da Matrix</param>
        /// <param name="tbId">Id da tabela em DataTables</param>
        /// <param name="SQL">SQL para o refresh</param>
        public void RefreshMatrix(string matrixId, string tbId, string SQL = "")
        {
            this._refreshMatrix(ref this.SapForm, matrixId, tbId, SQL);
        }

        /// <summary>
        /// Refresh em matrix
        /// </summary>
        /// <param name="frmId"></param>
        /// <param name="matrixId"></param>
        /// <param name="tbId"></param>
        /// <param name="cond"></param>
        public void RefreshMatrix(string frmId, string matrixId, string tbId, SAPbouiCOM.Conditions cond = null)
        {
            SAPbouiCOM.Form form = this.getForm(frmId);
            this._refreshMatrix(ref form, matrixId, tbId, cond);
        }

        /// <summary>
        /// Refresh em matrix
        /// </summary>
        /// <param name="form"></param>
        /// <param name="matrixId"></param>
        /// <param name="tbId"></param>
        /// <param name="cond"></param>
        public void RefreshMatrix(ref SAPbouiCOM.Form form, string matrixId, string tbId, SAPbouiCOM.Conditions cond = null)
        {
            this._refreshMatrix(ref form, matrixId, tbId, cond);
        }

        
        /// <summary>
        /// Executa o refresh em Matrix baseada em DBDataSources
        /// </summary>
        /// <param name="form"></param>
        /// <param name="matrixId"></param>
        /// <param name="tbId"></param>
        /// <param name="cond"></param>
        /// <param name="clear"></param>
        private void _refreshMatrix(ref SAPbouiCOM.Form form, string matrixId, string tbId, SAPbouiCOM.Conditions cond = null, bool clear = true)
        {
            try
            {
                // Refresh em DBDataSources
                SAPbouiCOM.DBDataSource dts = form.DataSources.DBDataSources.Item(tbId);
                if(cond == null)
                {
                    cond = this.Addon.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                    SAPbouiCOM.Condition oCond = cond.Add();
                    oCond.Alias = "Code";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_IS_NULL;
                }

                dts.Query(cond);

                // Atualiza a matrix
                this._doRefreshMatrix(ref form, matrixId, dts.Size, clear);
                
                if(dts != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(dts);
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Em RefreshMatrix de " + matrixId);
            }
        }

        /// <summary>
        /// Executa o refresh em Matrix baseada em DataTable
        /// </summary>
        /// <param name="form"></param>
        /// <param name="matrixId"></param>
        /// <param name="tbId"></param>
        /// <param name="SQL"></param>
        /// <param name="clear"></param>
        private void _refreshMatrix(ref SAPbouiCOM.Form form, string matrixId, string tbId, string SQL, bool clear = true)
        {
            try
            {
                // Refresh em DataTable
                SAPbouiCOM.DataTable dts = form.DataSources.DataTables.Item(tbId);
                dts.ExecuteQuery(SQL);

                // Atualiza a matrix
                this._doRefreshMatrix(ref form, matrixId, dts.Rows.Count, clear);

                if(dts != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(dts);
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Em RefreshMatrix de " + matrixId);
            }
        }

        /// <summary>
        /// Executa o refresh final
        /// </summary>
        /// <param name="form"></param>
        /// <param name="matrixId"></param>
        /// <param name="size"></param>
        /// <param name="clear"></param>
        private void _doRefreshMatrix(ref SAPbouiCOM.Form form, string matrixId, int size, bool clear = true)
        {
            
            SAPbouiCOM.Matrix matrix = null;
            try
            {
                matrix = form.Items.Item(matrixId).Specific;
            } catch(Exception e)
            {
                //this.Addon.DesenvTimeError(e, " - Matrix '" + matrixId + "' não encontrado em RefreshMatrix de " + matrixId);
            }

            if(matrix != null)
            {
                try
                {
                    form.Freeze(true);  // ATENÇÃO AQUI!!!!

                    if(clear)
                    {
                        matrix.Clear();
                        matrix.LoadFromDataSource();
                    } else
                    {
                        int s = matrix.RowCount;
                        for(int i = 0; i < size; i++)
                        {
                            matrix.AddRow();
                            matrix.SetLineData(i + s);
                        }
                    }

                    // Popula campos da matriz que usam SQL:
                    if(this.MatrixRefreshColumns.ContainsKey(matrixId))
                    {
                        SAPbouiCOM.Column column = null;
                        foreach(KeyValuePair<string, List<string>> refresh in this.MatrixRefreshColumns[matrixId])
                        {
                            column = matrix.Columns.Item(refresh.Key);
                            if(column != null)
                            {
                                this.populateColumn(matrixId, ref column, refresh.Value[0], refresh.Value[1], refresh.Value[2]);
                            } else
                            {
                                this.Addon.StatusErro("Não foi encontrado a coluna '" + refresh.Key + "' na matriz '" + matrixId + "' para atualização.");
                            }
                        }
                    }

                    // Permite a customização do refresh:
                    bool BubbleEvent = true;
                    FastOneItemEvent evObj = new FastOneItemEvent()
                    {
                        FormUID = form.UniqueID,
                        ItemUID = matrixId
                    };
                    if(form.IsSystem)
                    {
                        evObj.userFieldsHandler = (this.Addon.UserFields.ContainsKey(form.TypeEx) ? this.Addon.UserFields[form.TypeEx] : this.Addon.UserFields["UserFields"]);
                        evObj.FormUID = form.TypeEx;
                    }
                    if(this.EventMethods.ContainsKey(matrixId + "OnRefresh"))
                    {
                        this.Addon.ExecEvent(matrixId + "OnRefresh", out BubbleEvent, new Object[] { evObj, BubbleEvent });
                    }

                    matrix.AutoResizeColumns();

                // dts.Clear();
                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, " - RefreshMatrix de " + matrixId);
                } finally
                {
                    form.Freeze(false);
                    GC.Collect();
                }
            }
        }

        /// <summary>
        /// Faz um refresh em um matrix de detalhe
        /// </summary>
        /// <param name="formId"></param>
        /// <param name="mtxId"></param>
        /// <param name="tbId"></param>
        /// <param name="field"></param>
        /// <param name="cond"></param>
        public void refreshDetail(string frmId, string mtxId, string tbId, string field, string cond = null)
        {
            SAPbouiCOM.DBDataSource dts = null;
            SAPbouiCOM.Form form = null;
            try
            {
                form = this.getForm(frmId);
                if(this.Status > FormStatus.frmControlsCreated || form.IsSystem)
                {
                    SAPbouiCOM.Matrix matrix = this.GetItem(mtxId, frmId).Specific;
                    matrix.Clear(); // Tem que ser primeiro, senão afeta o dts.Size !!!

                    dts = form.DataSources.DBDataSources.Item(tbId);
                    dts.Query();

                    for(int s = 0; s < dts.Size; s++)
                    {
                        dts.Offset = s;
                        string c = dts.GetValue(field, s);
                        if(cond == null || cond.Trim() == c.Trim())
                        {
                            matrix.AddRow();
                            matrix.GetLineData(matrix.RowCount);
                        }
                    }
                }
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - refreshDetail de " + mtxId);

            } finally
            {
                if(dts != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(dts);
                }
                GC.Collect();
            }
        }

        #endregion


        #region :: Util

        /// <summary>
        /// Verifica se o último row da matrix está vazio com base em um field de referencia
        /// </summary>
        /// <param name="Matrix"></param>
        /// <param name="EmptyField"></param>
        /// <returns></returns>
        public bool MatrixEmptyLastRow(SAPbouiCOM.Matrix Matrix, string EmptyField)
        {
            bool res = false;
            if(Matrix.RowCount > 0)
            {
                string v = Matrix.GetCellSpecific(EmptyField, Matrix.RowCount).Value;
                res = String.IsNullOrEmpty(v);
            }
            return res;
        }
        
        public bool MatrixEmptyLastRow(string mtxId, string EmptyField)
        {
            SAPbouiCOM.Matrix Matrix = this.GetItem(mtxId).Specific;
            return this.MatrixEmptyLastRow(Matrix, EmptyField);
        }

        public void MatrixClearEmptyRows(SAPbouiCOM.Matrix Matrix, string no_empty_field)
        {
            try
            {
                if(!String.IsNullOrEmpty(no_empty_field))
                {
                    for(int r = Matrix.RowCount; r > 0; r--)
                    {
                        string teste = Matrix.GetCellSpecific(no_empty_field, r).Value;
                        if(String.IsNullOrEmpty(teste))
                        {
                            Matrix.DeleteRow(r);
                        }
                    }
                    Matrix.FlushToDataSource();
                }
            } catch (Exception e)
            {
                this.Addon.DesenvTimeError(e, " em MatrixClearEmptyRows usando campo " + no_empty_field);
            }
        }

        public void MatrixClearEmptyRows(string mtxId, string no_empty_field)
        {
            SAPbouiCOM.Matrix Matrix = this.GetItem(mtxId).Specific;
            this.MatrixClearEmptyRows(Matrix, no_empty_field);
        }

        #endregion


        #region :: Manipulação de dados

        /// <summary>
        /// Insere um row na matrix que reflete no datasource do form atual, ou do especificado.
        /// </summary>
        /// <param name="mtxId">Id da matrix</param>
        /// <param name="values">Row a ser inserido ({field, value})</param>
        /// <param name="formId">Se informado, identifica o form via "Forms.Item()"</param>
        /// <param name="formCount">Se informado, identifica o form via "Forms.GetForm()"</param>
        public void InsertOnMatrix(string mtxId, Dictionary<string, dynamic> values, string formId = "", int formCount = 0)
        {
            SAPbouiCOM.Form form = this.getForm(formId, formCount);
            this.InsertOnMatrix(mtxId, values, form);
        }

        public void InsertOnMatrix(string mtxId, Dictionary<string, dynamic> values, bool CheckLastLine, string formId = "", int formCount = 0)
        {
            SAPbouiCOM.Form form = this.getForm(formId, formCount);
            SAPbouiCOM.Matrix mtx = this.GetItem(mtxId, form).Specific;

            bool ok = true;
            if(CheckLastLine && this.MatrixEmptyRows.ContainsKey(mtxId))
            {
                if(mtx.RowCount > 0)
                {
                    foreach(string col in this.MatrixEmptyRows[mtxId])
                    {
                        string v = mtx.GetCellSpecific(col, mtx.RowCount).Value;
                        ok = !String.IsNullOrEmpty(v);
                        if(!ok)
                        {
                            break;
                        }
                    }
                }
            }

            if(ok)
            {
                this._insertOnMatrix(mtx, values, form, CheckLastLine);
            } else
            {
                mtx.SelectRow(mtx.RowCount, true, false);
                this.Addon.StatusErro("Preencha corretamente os campos na linha atual antes de inserir nova linha.");
            }
        }

        /// <summary>
        /// Insere um row na matrix que reflete no datasource do form especificado.
        /// </summary>
        /// <param name="mtxId">Id da matrix</param>
        /// <param name="values">Row a ser inserido ({field, value})</param>
        /// <param name="form">Instancia do form a ser utilizado</param>
        public void InsertOnMatrix(string mtxId, Dictionary<string, dynamic> values, SAPbouiCOM.Form form)
        {
            SAPbouiCOM.Matrix mtx = this.GetItem(mtxId, form).Specific;
            this._insertOnMatrix(mtx, values, form);
        }

        internal void _insertOnMatrix(SAPbouiCOM.Matrix mtx, Dictionary<string, dynamic> values, SAPbouiCOM.Form form, bool CheckEmptyLine = false)
        {
            try
            {
                //form.Freeze(true); Isso deve ser feito no form e não aqui
                if(mtx != null)
                {
                  
                    // Pega primeira coluna editável para receber foco
                    int toFocus = -1;
                    SAPbouiCOM.Column col = null;
                    for(int c = 0; c < mtx.Columns.Count; c++)
                    {
                        col = mtx.Columns.Item(c);
                        if(col.Editable && toFocus < 0)
                        {
                            toFocus = c;
                        }
                    }

                    // Verifica se ultima linha em branco
                    bool ok = true;
                    /**if(CheckEmptyLine)
                    {
                        if(mtx.RowCount > 0)
                        {
                            if (this.MatrixEmptyRows.ContainsKey(mtx.)
                            string EmptyField = mtx.Columns.Item(1).UniqueID;
                            string v = mtx.GetCellSpecific(EmptyField, mtx.RowCount).Value;
                            ok = !String.IsNullOrEmpty(v);
                            if(!ok)
                            {
                                try
                                {
                                    mtx.SetCellFocus(mtx.RowCount, toFocus);
                                } catch {
                                    mtx.SelectRow(mtx.RowCount, true, false);
                                }
                            }
                        }
                    }*/

                    if(ok)
                    {
                        mtx.FlushToDataSource();
                        int RecNo = mtx.RowCount + 1;
                        mtx.AddRow(1, RecNo);
                        mtx.ClearRowData(RecNo);
                        mtx.FlushToDataSource();

                        // Alimenta:
                        bool Enabled = true;
                        List<dynamic> disabled = new List<dynamic>();
                        foreach(KeyValuePair<string, dynamic> v in values)
                        {
                            col = mtx.Columns.Item(v.Key);
                            Enabled = col.Editable;

                            if(!col.Editable)
                            {
                                col.Editable = true;
                                disabled.Add(v.Key);
                            }
                            try
                            {
                                if(col.Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                                {
                                    col.Cells.Item(RecNo).Specific.Select(v.Value);
                                } else
                                {
                                    //col.Cells.Item(RecNo).Specific.Value = v.Value;

                                    //mtx.SetCellWithoutValidation(RecNo, v.Key, v.Value);

                                    try
                                    {
                                        if(v.Value is Double || v.Value is float || v.Value is Decimal)
                                        {
                                            string tmp = Convert.ToString(v.Value);
                                            mtx.SetCellWithoutValidation(RecNo, v.Key, tmp.Replace(',', '.'));
                                        } else
                                        {
                                            mtx.SetCellWithoutValidation(RecNo, v.Key, v.Value);
                                        }
                                        
                                    } catch
                                    {
                                        mtx.SetCellWithoutValidation(RecNo, v.Key, v.Value);
                                    }
                                }
                            } catch(Exception e) {

                            }
                        }

                        // Tenta zerar "Code":
                        try
                        {
                            mtx.Columns.Item("Code").Cells.Item(RecNo).Specific.Value = "";
                        } catch(Exception e) { }

                        // Posiciona
                        mtx.SelectRow(RecNo, true, false);
                        try
                        {
                            mtx.SetCellFocus(RecNo, toFocus);
                        } catch(Exception e)
                        {
                            //this.Addon.StatusErro(e.Message);
                        }

                        // Desativa quem de direito
                        foreach(dynamic c in disabled)
                        {
                            col = mtx.Columns.Item(c);
                            col.Editable = false;
                        }


                        if(this.SapForm.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && this.SapForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            this.SapForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Inserindo valor na matriz ");
            } finally
            {
                //form.Freeze(false);Isso deve ser feito no form e não aqui
            }
        }

        /// <summary>
        /// Atualiza campos em um row de uma matriz e reflete no client dataset com base no RecNo informado,
        /// ou no row selecionado na matriz, do form atual ou especificado.
        /// </summary>
        /// <param name="mtxId">Identificador da matriz</param>
        /// <param name="values">Valores a serem atualizados ({field, value})</param>
        /// <param name="RecNo">Se informado, é o número do row da matriz, se não é atualizado o row selecionado.</param>
        /// <param name="formId">Se informado, identifica o form via "Forms.Item()"</param>
        /// <param name="formCount">Se informado, identifica o form via "Forms.GetForm()"</param>
        public void UpdateOnMatrix(string mtxId, Dictionary<string, dynamic> values, int RecNo = -1, string formId = "", int formCount = 0)
        {
            try
            {
                SAPbouiCOM.Matrix mtx = this.GetItem(mtxId, formId, formCount).Specific;
                if(mtx != null)
                {
                    if(RecNo < 0)
                    {
                        RecNo = mtx.GetNextSelectedRow();
                    }

                    if(RecNo > -1)
                    {
                        bool checkUnique = false;
                        if(this.FormParams.Controls.ContainsKey(mtxId))
                        {
                            checkUnique = this.FormParams.Controls[mtxId].ExtraData.ContainsKey("Unique");
                        }

                        foreach(KeyValuePair<string, dynamic> v in values)
                        {

                            if(checkUnique && ((List<string>)this.FormParams.Controls[mtxId].ExtraData["Unique"]).Contains(v.Key))
                            {
                                if(!this.CheckUniqueValueColumn(mtxId, v.Key, RecNo, v.Value))
                                {
                                    this.Addon.ShowMessage("O valor informado está duplicado. Favor selecionar outro.");
                                    foreach(KeyValuePair<string, dynamic> z in values)
                                    {
                                        mtx.SetCellWithoutValidation(RecNo, z.Key, "");
                                    }

                                    break;
                                }
                            }

                            mtx.SetCellWithoutValidation(RecNo, v.Key, v.Value);
                        }
                        mtx.FlushToDataSource();

                        if(this.SapForm.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && this.SapForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            this.SapForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Alterando valor na matirz (A matriz usa DBDataSource?) " + mtxId);
            }
        }

        public void UpdateOnMatrix(string mtxId, string datatable, Dictionary<string, dynamic> values, int RecNo = -1, string formId = "", int formCount = 0)
        {
            try
            {
                SAPbouiCOM.Form frm = this.getForm(formId, formCount);
                SAPbouiCOM.Matrix mtx = this.GetItem(mtxId, frm).Specific;
                SAPbouiCOM.DataTable dts = frm.DataSources.DataTables.Item(datatable);
                if(mtx != null)
                {
                    mtx.FlushToDataSource();
                    if(RecNo < 0)
                    {
                        RecNo = mtx.GetNextSelectedRow();
                    } 
                    
                    bool checkUnique = this.FormParams.Controls[mtxId].ExtraData.ContainsKey("Unique");

                    foreach(KeyValuePair<string, dynamic> v in values)
                    {
                        if(checkUnique && ((List<string>)this.FormParams.Controls[mtxId].ExtraData["Unique"]).Contains(v.Key))
                        {
                            if(!this.CheckUniqueValueDataTable(dts, v.Key, RecNo-1, v.Value))
                            {
                                this.Addon.ShowMessage("O valor informado está duplicado. Favor selecionar outro.");
                                foreach(KeyValuePair<string, dynamic> z in values){
                                    dts.SetValue(z.Key, RecNo - 1, "");
                                }
                                break;
                            }
                        }

                        dts.SetValue(v.Key, RecNo-1, v.Value);
                    }
                    mtx.LoadFromDataSourceEx();

                    if(this.SapForm.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && this.SapForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        this.SapForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Alterando valor na matirz " + mtxId);
            }
        }


        /// <summary>
        /// Salva os registros de um client dataset de uma matrix no banco de dados.
        /// </summary>
        /// <param name="table">Nome do dbDatasource</param>
        /// <param name="matrixId">Se informado, cuida do refresh da matriz</param>
        /// <param name="no_empty_field">Se informado, exclui da operação rows com esse campo em branco</param>
        /// <returns></returns>
        public bool SaveMatrixToServer(string matrixId, string mtxDatasource, string no_empty_field = "")
        {
            bool res = false;
            SAPbouiCOM.Matrix matrix = ((SAPbouiCOM.Matrix)this.GetItem(matrixId).Specific);

            if(this.CheckUniqueValueColumn(matrixId))
            {
                // Limpa a matriz
                if(!String.IsNullOrEmpty(no_empty_field))
                {
                    this.MatrixClearEmptyRows(matrix, no_empty_field);
                }

                matrix.FlushToDataSource();
                res = this.SaveToServer(mtxDatasource);
            }
            return res;
        }

        /// <summary>
        /// Executa um UPDATE em uma tabela padrão SQL "table_to" (não SAP) no server em uma matriz, com base 
        /// nos dados no DataTable "data_table_from", utilizando "map_fields" para mapear campos
        /// do DataTable para a tabela no banco.
        /// </summary>
        /// <param name="mtxId">Matriz </param>
        /// <param name="table_to">Informar com o "@", caso houver</param>
        /// <param name="data_table_from">Nome do DataTable de onde buscar os valores</param>
        /// <param name="map_fields">A Chave é a coluna no DataTable e o Valor é a coluna da tabela SQL</param>
        /// <param name="sqlRefresh">SQL utilizado para o refresh da matriz</param>
        /// <returns>True, se correr tudo bem</returns>
        public bool SaveMatrixToServer(string matrixId, string table_to, string data_table_from, Dictionary<string, string> map_fields, string sqlRefresh)
        {
            SAPbouiCOM.Matrix matrix = ((SAPbouiCOM.Matrix)this.GetItem(matrixId).Specific);
            matrix.FlushToDataSource();

            bool ok = this.SaveToServer(table_to, data_table_from, map_fields);
            if(ok)
            {
                this.RefreshMatrix(matrixId, data_table_from, sqlRefresh);
            }

            return ok;
        }


        /// <summary>
        /// Deleta um row em matrix APENAS. NÃO afeta o dataset.
        /// </summary>
        /// <param name="mtxId">Identificador da matrix, que terá o row selecionado removido.</param>
        public SAPbouiCOM.Matrix DeleteOnMatrix(string mtxId, bool quiet = false)
        {
            SAPbouiCOM.Matrix matrix = this.GetItem(mtxId).Specific;
            bool del = true;
            if(!quiet)
            {
                del = (this.Addon.SBO_Application.MessageBox("Tem certeza de que deseja remover esse registro?", 2, "Sim", "Não") == 1);
            }

            if(del)
            {
                this.DeleteOnMatrix(matrix, true);
            }
            return matrix;
        }

        /// <summary>
        /// Deleta um row em matrix APENAS. NÃO afeta o dataset.
        /// </summary>
        /// <param name="matrix">Instancia da matrix, que terá o row selecionado removido.</param>
        public void DeleteOnMatrix(SAPbouiCOM.Matrix matrix, bool quiet = false)
        {
            bool del = true;
            if(!quiet)
            {
                del = (this.Addon.SBO_Application.MessageBox("Tem certeza de que deseja remover esse registro?", 2, "Sim", "Não") == 1);
            }

            if (del){
                try
                {
                    for(int i = matrix.RowCount; i >= 1; i--)
                    {
                        if(matrix.IsRowSelected(i))
                        {
                            matrix.DeleteRow(i);
                            matrix.FlushToDataSource();
                        }
                    }
                    matrix.FlushToDataSource();

                    if(this.SapForm.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && this.SapForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        this.SapForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }

                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, " - Removendo row em matriz");
                }
            }
        }

        /// <summary>
        /// Remove um row em um usertable diretamente no banco com base no row da matrix.
        /// NÃO use em forms UDOs.
        /// </summary>
        /// <param name="mtxId">A matriz DEVERÁ ter um campo "Code" que será utilizado na operação</param>
        /// <param name="table">Não esquecer o arroba.</param>
        /// <param name="formId">Se informado, identifica o form via "Forms.Item()"</param>
        /// <param name="formCount">Se informado, identifica o form via "Forms.GetForm()"</param>
        public void DeleteMatrixOnServer(string mtxId, string table, string formId = "", int formCount = 0)
        {
            if(this.Addon.SBO_Application.MessageBox("Tem certeza de que deseja remover esse registro?", 2, "Sim", "Não") == 1)
            {
                try
                {
                    SAPbouiCOM.Form form = this.getForm(formId, formCount);
                    SAPbouiCOM.Matrix matrix = this.GetItem(mtxId, form).Specific;

                    for(int i = matrix.RowCount; i >= 1; i--)
                    {
                        if(matrix.IsRowSelected(i))
                        {
                            string code = matrix.GetCellSpecific("Code", i).value;
                            bool ok = this.DeleteOnServer(table, new Dictionary<string, dynamic>(){ 
                            {"Code", code}
                        }, true);

                            if(ok)
                            {
                                matrix.DeleteRow(i);
                                matrix.FlushToDataSource();
                            }
                        }
                    }
                    form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, " - Removendo row em " + mtxId);
                }
            }
        }

        #endregion


        #endregion


        #region :: Grid


        /// <summary>
        ///  Pupula os ítens de combo de uma coluna 
        /// </summary>
        /// <param name="matrixId"></param>
        /// <param name="column"></param>
        /// <param name="sql"></param>
        /// <param name="first_key"></param>
        /// <param name="first_value"></param>
        public void populateColumn(ref SAPbouiCOM.GridColumn column, string sql, string first_key = "", string first_value = "")
        {
            SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            if(column.Type != SAPbouiCOM.BoGridColumnType.gct_ComboBox)
            {
                this.Addon.StatusErro("Não é possível popular coluna que não seja do tipo it_COMBO_BOX");
                return;
            }

            // Limpa
            while(((SAPbouiCOM.ComboBox)column).ValidValues.Count > 0)
            {
                ((SAPbouiCOM.ComboBox)column).ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }

            rec.DoQuery(sql);
            rec.MoveFirst();

            //if(first_key != "")
            //{
                ((SAPbouiCOM.ComboBox)column).ValidValues.Add(first_key, first_value);
            //}

            while(!rec.EoF)
            {
                try
                {
                    ((SAPbouiCOM.ComboBox)column).ValidValues.Add(rec.Fields.Item(0).Value, rec.Fields.Item(1).Value);
                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, " - Acrescentando valores ao combo " + column.UniqueID + " do grid");
                }
                rec.MoveNext();
            }
            rec = null;

            GC.Collect();
        }

        /// <summary>
        /// Popula os itens de uma coluna via params
        /// By Labs - 07/2013
        /// </summary>
        /// <param name="column"></param>
        /// <param name="itens">Objetto contendo os valores dos itens do combo</param>
        public void populateColumn(ref SAPbouiCOM.GridColumn column, Dictionary<string, string> itens)
        {
            if(column.Type != SAPbouiCOM.BoGridColumnType.gct_ComboBox)
            {
                this.Addon.StatusErro("Não é possível popular coluna que não seja do tipo it_COMBO_BOX");
                return;
            }

            // Limpa
            while(((SAPbouiCOM.ComboBox)column).ValidValues.Count > 0)
            {
                ((SAPbouiCOM.ComboBox)column).ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }

            foreach(KeyValuePair<string, string> item in itens)
            {
                ((SAPbouiCOM.ComboBox)column).ValidValues.Add(item.Key, item.Value);
            }

            GC.Collect();
        }


        #endregion


        #region :: Combos

        /// <summary>
        /// Limpa os valores de um combo
        /// </summary>
        /// <param name="combo"></param>
        public void clearCombo(SAPbouiCOM.ComboBox combo)
        {
            try
            {
                for(int i = combo.ValidValues.Count - 1; i > -1; i--)
                {
                    combo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                }
            } catch(Exception e)
            {
               // this.Addon.DesenvTimeError(e, " - Limpndo ítens no combo " + combo.Item.UniqueID);
            }
        }

        /// <summary>
        /// Popula um combo via SQL
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="ctrlId"></param>
        /// <param name="frmId"></param>
        /// <param name="sql"></param>
        /// <param name="first_key">Use NO_FIRST_VALUE para evitar o primeiro valor default</param>
        /// <param name="first_value"></param>
        /// <param name="def_value"></param>
        public bool populateCombo(string ctrlId, string frmId, string sql, string first_key = "", string first_value = "", string def_value = "")
        {
            SAPbouiCOM.ComboBox combo = this.GetItem(ctrlId, frmId).Specific;
            return this.populateCombo(combo, sql, first_key, first_value, def_value);
        }

        public bool populateCombo(SAPbouiCOM.ComboBox combo, string sql, string first_key = "", string first_value = "", string def_value = "", bool first_if_only = false)
        {
            bool res = false;
            SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // Limpa
            this.clearCombo(combo);
            
            try
            {
                rec.DoQuery(sql);
                rec.MoveFirst();

                if(first_key != "NO_FIRST_VALUE")
                {
                    try
                    {
                        combo.ValidValues.Add(first_key, first_value);
                    } catch { }
                }

                while(!rec.EoF)
                {
                    try
                    {
                        combo.ValidValues.Add(rec.Fields.Item(0).Value, rec.Fields.Item(1).Value);
                    } catch { }
                    rec.MoveNext();
                }

                if(!String.IsNullOrEmpty(def_value))
                {
                    try
                    {
                        combo.Select(def_value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    } catch { }
                } else
                {
                    if(rec.RecordCount == 1) // && first_if_only)
                    {
                        combo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                }
                res = (rec.RecordCount > 0);
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Carregando ítens para o combo " + combo.Item.UniqueID);
            }

            rec = null;
            GC.Collect();
            return res;
        }

        /// <summary>
        /// Popula um combo via params
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="ctrlId">O combo</param>
        /// <param name="frmId"></param>
        /// <param name="itens">Objetto contendo os valores dos itens do combo</param>
        /// <param name="def_value"></param>
        public bool populateCombo(string ctrlId, string frmId, Dictionary<string, string> itens, string def_value = "")
        {
            SAPbouiCOM.ComboBox combo = this.GetItem(ctrlId, frmId).Specific;
            return this.populateCombo(combo, itens, def_value);
        }

        public bool populateCombo(SAPbouiCOM.ComboBox combo, Dictionary<string, string> itens, string def_value = "")
        {
            bool res = true;

            // Limpa            
            this.clearCombo(combo);


            foreach(KeyValuePair<string, string> item in itens)
            {
                try
                {
                    combo.ValidValues.Add(item.Key, item.Value);
                } catch {
                    res = false;
                }
            }

            if(!String.IsNullOrEmpty(def_value))
            {
                try
                {
                    combo.Select(def_value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                } catch { }
            }

            GC.Collect();
            return res;
        }

        /// <summary>
        /// Popula um combo com base em valores de um DataTable
        /// </summary>
        /// <param name="ctrlId"></param>
        /// <param name="DataTable"></param>
        /// <param name="key"></param>
        /// <param name="val"></param>
        /// <param name="first_key"></param>
        /// <param name="first_value"></param>
        /// <param name="def_value"></param>
        public bool populateComboDataTable(string ctrlId, string DataTable, string key, string val, string first_key = "", string first_value = "", string def_value = "")
        {
            SAPbouiCOM.ComboBox combo = this.GetItem(ctrlId).Specific;
            return this.populateComboDataTable(combo, DataTable, key, val, first_key, first_value, def_value);
        }

        /// <summary>
        /// Popula um combo com base em valores de um DataTable
        /// </summary>
        /// <param name="combo"></param>
        /// <param name="DataTable"></param>
        /// <param name="key"></param>
        /// <param name="val"></param>
        /// <param name="first_key"></param>
        /// <param name="first_value"></param>
        /// <param name="def_value"></param>
        public bool populateComboDataTable(SAPbouiCOM.ComboBox combo, string DataTable, string key, string val, string first_key = "", string first_value = "", string def_value = "")
        {
            bool res = false;

            // Limpa            
            this.clearCombo(combo);

            //if(!String.IsNullOrEmpty(first_value))
            //{
                try
                {
                    combo.ValidValues.Add(first_key, first_value);
                } catch { }
            //}

            try{
                SAPbouiCOM.DataTable dt = this.SapForm.DataSources.DataTables.Item(DataTable);
                if (!dt.IsEmpty){
                    for (int r = 0; r < dt.Rows.Count; r++){
                        try
                        {
                            combo.ValidValues.Add(dt.GetValue(key, r), dt.GetValue(val, r));
                        } catch (Exception e) {
                            this.Addon.DesenvTimeError(e, "populateComboDataTable");

                        }
                    }
                    res = true;
                }
            } catch {}

            if(!String.IsNullOrEmpty(def_value))
            {
                try
                {
                    combo.Select(def_value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                } catch { }
            }

            GC.Collect();
            return res;
        }

        /*
        public void populateComboDbDataSource(string ctrlId, string DataSource, string key, string val, string first_key = "", string first_value = "", string def_value = "")
        {
            SAPbouiCOM.ComboBox combo = this.GetItem(ctrlId).Specific;
            this.populateComboDbDataSource(combo, DataSource, key, val, first_key, first_value, def_value);
        }

        public void populateComboDbDataSource(SAPbouiCOM.ComboBox combo, string DataSource, string key, string val, string first_key = "", string first_value = "", string def_value = "")
        {
            // Limpa            
            this.clearCombo(combo);

            if(!String.IsNullOrEmpty(first_value))
            {
                try
                {
                    combo.ValidValues.Add(first_key, first_value);
                } catch { }
            }

            try
            {
                SAPbouiCOM.DataTable dt = this.SapForm.DataSources.DataTables.Item(DataTable);
                if(!dt.IsEmpty)
                {
                    for(int r = 0; r < dt.Rows.Count; r++)
                    {
                        try
                        {
                            combo.ValidValues.Add(dt.GetValue(key, r), dt.GetValue(val, r));
                        } catch(Exception e)
                        {
                            this.Addon.DesenvTimeError(e, "populateComboDataTable");

                        }

                    }
                }
            } catch { }

            if(!String.IsNullOrEmpty(def_value))
            {
                try
                {
                    combo.Select(def_value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                } catch { }
            }

            GC.Collect();
        }
        */

        /// <summary>
        /// Popula um combo via com valores definidos
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="formID"></param>
        public void comboSimNao(ref SAPbouiCOM.Item ctrl, string formID)
        {
            SAPbouiCOM.ComboBox combo = ctrl.Specific;
            combo.ValidValues.Add("S", "Sim");
            combo.ValidValues.Add("N", "Não");
            GC.Collect();
        }

        private void comboSimNao(SAPbouiCOM.ComboBox combo)
        {
            combo.ValidValues.Add("S", "Sim");
            combo.ValidValues.Add("N", "Não");
        }

        private void comboDias(SAPbouiCOM.ComboBox combo)
        {
            for(int x = 1; x < 32; x++)
            {
                string dia = (x < 10 ? "0" + x : x.ToString());
                combo.ValidValues.Add(dia, dia);
            }
        }

        private void comboMeses(SAPbouiCOM.ComboBox combo)
        {
            string[] meses = { "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez" };
            for(int x = 0; x < 12; x++)
            {
                combo.ValidValues.Add(x.ToString(), meses[x]);
            }
        }

        private void comboHoras(SAPbouiCOM.ComboBox combo)
        {
            for(int x = 0; x < 24; x++)
            {
                combo.ValidValues.Add(x.ToString(), x.ToString());
            }
        }

        private void comboMinutos(SAPbouiCOM.ComboBox combo, int step = 5)
        {
            for(int x = 0; x < 61; x +=step)
            {
                combo.ValidValues.Add(x.ToString(), x.ToString());
            }
        }

        #endregion


        #endregion


        #region :: Eventos em Componentes

        /// <summary>
        /// Cria evento de OnClick para um componente.
        /// </summary>
        /// <param name="CompId">ID do componente</param>
        /// <param name="FunctionName">Por padrão: CompId + "OnClick" </param>
        /// <param name="AfterSAP">Executar antes ou depois do SAP</param>
        public void SetOnClick(string CompId, bool AfterSAP = true, string FunctionName = "")
        {
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_CLICK, this.FormId, CompId,
                (String.IsNullOrEmpty(FunctionName) ? CompId + "OnClick" : FunctionName), AfterSAP
            );
        }

        /// <summary>
        /// Cria evento de OnDblClick para um componente.
        /// </summary>
        /// <param name="CompId">ID do componente</param>
        /// <param name="FunctionName">Por padrão: CompId + "OnDblClick" </param>
        /// <param name="AfterSAP">Executar antes ou depois do SAP</param>
        public void SetOnDblClick(string CompId, bool AfterSAP = true, string FunctionName = "")
        {
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK, this.FormId, CompId,
                (String.IsNullOrEmpty(FunctionName) ? CompId + "OnDblClick" : FunctionName), AfterSAP
            );
        }

        /// <summary>
        /// Cria evento de OnChange para um componente.
        /// </summary>
        /// <param name="CompId">ID do componente</param>
        /// <param name="FunctionName">Por padrão: CompId + "OnClick" </param>
        /// <param name="AfterSAP">Executar antes ou depois do SAP</param>
        public void SetOnComboChange(string CompId, bool AfterSAP = true, string FunctionName = "")
        {
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT, this.FormId, CompId,
                (String.IsNullOrEmpty(FunctionName) ? CompId + "OnComboChange" : FunctionName), AfterSAP
            );
        }

        /// <summary>
        /// Cria evento de OnRightClick para um componente.
        /// </summary>
        /// <param name="CompId">ID do componente</param>
        /// <param name="FunctionName">Por padrão: CompId + "OnRightClick" </param>
        /// <param name="AfterSAP">Executar antes ou depois do SAP</param>
        public void SetOnRightClick(string CompId, bool AfterSAP = true, string FunctionName = "")
        {
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK, this.FormId, CompId,
                (String.IsNullOrEmpty(FunctionName) ? CompId + "OnRightClick" : FunctionName), AfterSAP
            );
        }

        /// <summary>
        /// Cria evento de OnPickerClick para um componente.
        /// </summary>
        /// <param name="CompId">ID do componente</param>
        /// <param name="FunctionName">Por padrão: CompId + "OnPickerClick" </param>
        /// <param name="AfterSAP">Executar antes ou depois do SAP</param>
        public void SetOnPickerClick(string CompId, bool AfterSAP = true, string FunctionName = "")
        {
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED, this.FormId, CompId,
                (String.IsNullOrEmpty(FunctionName) ? CompId + "OnPickerClick" : FunctionName), AfterSAP
            );
        }

        /// <summary>
        /// Cria evento de OnFocus para um componente.
        /// </summary>
        /// <param name="CompId">ID do componente</param>
        /// <param name="FunctionName">Por padrão: CompId + "OnFocus" </param>
        /// <param name="AfterSAP">Executar antes ou depois do SAP</param>
        public void SetOnFocus(string CompId, bool AfterSAP = true, string FunctionName = "")
        {
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS, this.FormId, CompId,
                (String.IsNullOrEmpty(FunctionName) ? CompId + "OnFocus" : FunctionName), AfterSAP
            );
        }

        /// <summary>
        /// Cria evento de OnLostFocus para um componente.
        /// </summary>
        /// <param name="CompId">ID do componente</param>
        /// <param name="FunctionName">Por padrão: CompId + "OnLostFocus" </param>
        /// <param name="AfterSAP">Executar antes ou depois do SAP</param>
        public void SetOnLostFocus(string CompId, bool AfterSAP = true, string FunctionName = "")
        {
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS, this.FormId, CompId,
                (String.IsNullOrEmpty(FunctionName) ? CompId + "OnLostFocus" : FunctionName), AfterSAP
            );
        }

        /// <summary>
        /// Cria evento de OnItemPressed para um componente.
        /// </summary>
        /// <param name="CompId">ID do componente</param>
        /// <param name="FunctionName">Por padrão: CompId + "OnItemPressed" </param>
        /// <param name="AfterSAP">Executar antes ou depois do SAP</param>
        public void SetOnItemPressed(string CompId, bool AfterSAP = true, string FunctionName = "")
        {
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, this.FormId, CompId,
                (String.IsNullOrEmpty(FunctionName) ? CompId + "OnItemPressed" : FunctionName), AfterSAP
            );
        }

        /// <summary>
        /// Cria evento de OnKeyDown para um componente.
        /// </summary>
        /// <param name="CompId">ID do componente</param>
        /// <param name="FunctionName">Por padrão: CompId + "OnKeyDown" </param>
        /// <param name="AfterSAP">Executar antes ou depois do SAP</param>
        public void SetOnKeyDown(string CompId, bool AfterSAP = true, string FunctionName = "")
        {
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_KEY_DOWN, this.FormId, CompId,
                (String.IsNullOrEmpty(FunctionName) ? CompId + "OnKeyDown" : FunctionName), AfterSAP
            );
        }


        #endregion



        #region :: Utils

        /// <summary>
        /// Navega para o primeiro registro "clicando" no botão de navegação.
        /// </summary>
        public void GotoFirst()
        {
            this._navTo("1290");
        }

        /// <summary>
        /// Volta um registro "clicando" no botão de navegação.
        /// </summary>
        public void GotoBack()
        {
            this._navTo("1289");
        }

        /// <summary>
        /// Avança um registro "clicando" no botão de navegação.
        /// </summary>
        public void GotoNext()
        {
            this._navTo("1288");
        }

        /// <summary>
        /// Navega para o último registro "clicando" no botão de navegação.
        /// </summary>
        public void GotoLast()
        {
            this._navTo("1291");
        }

        public void GotoFormCode(string UDOFindCode = "")
        {
            this.timerUDOFind.Stop();

            if(String.IsNullOrEmpty(UDOFindCode))
            {
                UDOFindCode = this.UDOCode;
            }

            if(String.IsNullOrEmpty(this.FormParams.BrowseByComp))
            {
                this.Addon.DesenvTimeAlert("Não foi definido 'FormParams.BrowseByComp' no formulário. O refresh não será possível.");
                return;
            }

            if(this.SapForm != null)
            {
                try
                {

                    this.SapForm.Freeze(true);

                    SAPbouiCOM.Item comp = this.GetItem(this.FormParams.BrowseByComp);

                    this.SapForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                    bool oldEnabled = comp.Enabled;
                    comp.Enabled = true;
                    comp.Specific.Value = UDOFindCode;

                    this.GetItem("1").Click();

                    comp.Enabled = oldEnabled;
                } catch (Exception e){

                } finally {
                    this.SapForm.Freeze(false);
                }
            }
        }

        internal void _navTo(string menuId)
        {
            try
            {
                this.Addon.SBO_Application.ActivateMenuItem(menuId);
            } catch(Exception e)
            {

            }
        }


        /// <summary>
        /// Recupera o DocEntry com base no evento.
        /// </summary>
        /// <param name="evObj"></param>
        /// <returns></returns>
        public string GetFormCode(SAPbouiCOM.BusinessObjectInfo evObj)
        {
            this.UDOCode = "";
            try
            {
                if(!String.IsNullOrEmpty(evObj.ObjectKey))
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(evObj.ObjectKey);
                    this.UDOCode = xmlDoc.DocumentElement.FirstChild.ChildNodes[0].Value;
                }

            } catch(Exception e)
            {
               // this.Addon.DesenvTimeError(e, "em GetFormCode " + evObj.FormUID);
            }

            return this.UDOCode;
        }

        /// <summary>
        /// Recupera o próximo valor utilizavel de "Code" em uma tabela.
        /// NÃO É UM PROCEDIMENTO INTEIRAMENTE SEGURO!!!
        /// </summary>
        /// <param name="table">Com o "@" se houver</param>
        /// <param name="padding">Número de padding</param>
        /// <param name="padchar">char para padding</param>
        /// <returns></returns>
        public string getNextCode(string table, int padding = 0, char padchar = '0')
        {
            return this.DtSources.getNextCode(table, padding, padchar);
        }

        /// <summary>
        /// Recupera o próximo valor utilizavel de "Code" em uma tabela.
        /// NÃO É UM PROCEDIMENTO INTEIRAMENTE SEGURO!!!
        /// </summary>
        /// <param name="table"></param>
        /// <param name="padding"></param>
        /// <param name="padchar"></param>
        /// <returns></returns>
        public string getNextIdentityCode(string table, int padding = 0, char padchar = '0')
        {
            return this.DtSources.getNextIdentityCode(table, padding, padchar);
        }

        /// <summary>
        /// Recupera o próximo valor utilizavel de "Code" em uma tabela.
        /// NÃO É UM PROCEDIMENTO INTEIRAMENTE SEGURO!!!
        /// </summary>
        /// <param name="table"></param>
        /// <param name="padding"></param>
        /// <param name="padchar"></param>
        /// <returns></returns>
        public int getNextIdentityCodeInt(string table, int padding = 0, char padchar = '0')
        {
            int res = 0;
            Int32.TryParse(this.DtSources.getNextIdentityCode(table, padding, padchar), out res);
            return res;
        }

        public int getNextCodeInt(string table, int padding = 0, char padchar = '0')
        {
            int res = 0;
            Int32.TryParse(this.DtSources.getNextCode(table, padding, padchar), out res);
            return res;
        }

        /// <summary>
        /// Recupera o próximo valor utilizavel de "Code" em uma tabela.
        /// NÃO É UM PROCEDIMENTO INTEIRAMENTE SEGURO!!!
        /// </summary>
        /// <param name="table">Com o "@" se houver</param>
        /// <param name="padding">Número de padding</param>
        /// <param name="padchar">char para padding</param>
        /// <returns></returns>
        public string getMaxCode(string table, int padding = 0, char padchar = '0')
        {
            return this.DtSources.getMaxCode(table, padding, padchar);
        }

        public int getMaxCodeInt(string table, int padding = 0, char padchar = '0')
        {
            int res = 0;
            Int32.TryParse(this.DtSources.getMaxCode(table, padding, padchar), out res);
            return res;
        }

        /// <summary>
        /// Recupera o próximo LineId em um DBDataSource.
        /// SÓ DEVE SER USADO EM MATRIZ UDOCHILD!!!!
        /// </summary>
        /// <param name="mtxId">Id da Matriz</param>
        /// <param name="dbdatasource">Nome da tabela DBDataSource</param>
        /// <returns>retorna o próximo LineId que deve ser inserido, não precisa ser incrementado</returns>
        public string getNextLineId(string mtxId, string dbdatasource)
        {
            SAPbouiCOM.Matrix mtx = this.GetItem(mtxId).Specific;
            mtx.FlushToDataSource();
            
            SAPbouiCOM.DBDataSource dt = this.SapForm.DataSources.DBDataSources.Item(dbdatasource);
            string line_id = dt.GetValue("LineId", dt.Size - 1);
            int int_line_id = 0;

            if(String.IsNullOrEmpty(line_id))
            {
                line_id = "1";
            } else
            {
                int_line_id = Int32.Parse(line_id) + 1;
                line_id = int_line_id.ToString();
            }

            return line_id;
        }


        public bool RunOnOppener(string func)
        {
            bool res = false;
            
            if(this.Oppener != null && this.Oppener.GetType().Name == "Forms")
            {
              //  ((FrmListaVeiculos)this.Oppener).RefreshListagem();
            }

            return res;
        }

        /// <summary>
        /// Recupera ou um form SAP padrão ou o proprio form instanciado
        /// </summary>
        /// <param name="frmId">Id do Form</param>
        /// <param name="frmCount"></param>
        /// <returns>Retorna um form sap</returns>
        public SAPbouiCOM.Form getForm(string frmId = "", int frmCount = 0)
        {
            SAPbouiCOM.Form frm = null;

            if(String.IsNullOrEmpty(frmId) && this.SapForm != null)
            {
                frm = this.SapForm;
            } else
            {
                // Pega em FastOne:
                frm = this.Addon.getForm(frmId, frmCount);

                // Se não rolou o form (ainda):
                if(frm == null)
                {
                    frm = this.SapForm;
                }
            }

            // Retorna o form:
            return frm;
        }

        /// <summary>
        /// Recupera um item em um form
        /// </summary>
        /// <param name="ctrlId">Id do componente</param>
        /// <param name="frmId">Id do form onde ele esta</param>
        /// <param name="frmCount"></param>
        /// <returns>SAPbouiCOM.Item</returns>
        public SAPbouiCOM.Item GetItem(string ctrlId, string frmId = "", int frmCount = 0)
        {
            SAPbouiCOM.Item ctrl = null;
            try
            {
                SAPbouiCOM.Form frm = this.getForm(frmId, frmCount);
                ctrl = this.GetItem(ctrlId, frm);
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Erro ao tentar encontrar " + ctrlId + " no form " + FormId);
            }
            return ctrl;
        }

        /// <summary>
        /// Recupera um item em um form
        /// </summary>
        /// <param name="ctrlId"></param>
        /// <param name="form"></param>
        /// <returns></returns>
        public SAPbouiCOM.Item GetItem(string ctrlId, SAPbouiCOM.Form form)
        {
            SAPbouiCOM.Item ctrl = null;
            try
            {
                ctrl = form.Items.Item(ctrlId);
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Erro ao tentar encontrar " + ctrlId + " no form " + FormId);
            }
            return ctrl;
        }


        /// <summary>
        /// Recupera o valor de uma coluna no row selecionado de uma matriz
        /// </summary>
        /// <param name="matrixId"></param>
        /// <param name="frmId"></param>
        /// <param name="colName"></param>
        /// <param name="force_first_line"></param>
        /// <returns></returns>
        public dynamic getCellValue(string matrixId, string frmId, string colName, bool force_first_line = false, int r = -1)
        {
            try
            {
                SAPbouiCOM.Form frm = this.getForm(frmId);
                SAPbouiCOM.Matrix matrix = frm.Items.Item(matrixId).Specific;
                return this.getCellValue(ref matrix, colName, force_first_line, r);
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - getCellValue de " + matrixId + ", coluna " + colName);
                return null;
            }
        }

        /// <summary>
        /// Recupera o valor de uma coluna no row selecionado de uma matriz
        /// </summary>
        /// <param name="matrix"></param>
        /// <param name="colName"></param>
        /// <returns></returns>
        public dynamic getCellValue(ref SAPbouiCOM.Matrix matrix, string colName, bool force_first_line = false, int r = -1)
        {
            dynamic res = null;
            try
            {
                if(r == -1)
                {
                    r = (force_first_line ? 1 : matrix.GetNextSelectedRow());
                }

                if(r > -1)
                {
                    res = matrix.GetCellSpecific(colName, r).Value; // .Columns.Item(colName).Cells.Item(r).Specific.Value;
                } else
                {
                    this.Addon.StatusAlerta("Nenhuma linha está selecionada. Clique na coluna '#' para selecionar a linha desejada.");
                }
            } catch(Exception e)
            {
               // this.Addon.DesenvTimeError(e, "getCellValue: " + colName);
            }
            return res;
        }

        /// <summary>
        /// Coloca um valor em uma celula da matriz
        /// </summary>
        /// <param name="matrixId"></param>
        /// <param name="formId"></param>
        /// <param name="colName"></param>
        /// <param name="value"></param>
        /// <param name="row">Se não informado, pega o último row por defaul</param>
        /// <returns></returns>
        public bool putCellValue(string matrixId, string frmId, string colName, dynamic value, int row = -1)
        {
            try
            {
                SAPbouiCOM.Form frm = this.getForm(frmId);
                SAPbouiCOM.Matrix matrix = frm.Items.Item(matrixId).Specific;
                return this._putCellValue(ref matrix, colName, value, row);
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Inserindo valor na " + matrixId + ", coluna " + colName);
                return false;
            }
        }

        /// <summary>
        /// Coloca um valor em uma celula da matriz
        /// </summary>
        /// <param name="matrix">SAPbouiCOM.Matrix</param>
        /// <param name="colName"></param>
        /// <param name="value"></param>
        /// <param name="row">Se não informado, pega o último row por default</param>
        /// <returns></returns>
        public bool putCellValue(ref SAPbouiCOM.Matrix matrix, string colName, dynamic value, int row = -1)
        {
            return this._putCellValue(ref matrix, colName, value, row);
        }

        /// <summary>
        /// Coloca um valor em uma celula da matriz
        /// </summary>
        /// <param name="matrixId"></param>
        /// <param name="formId"></param>
        /// <param name="colName"></param>
        /// <param name="value"></param>
        /// <param name="row">Se não informado, pega o último row por defaul</param>
        /// <returns></returns>
        public bool putComboCellValue(string matrixId, string frmId, string colName, dynamic value, int row = -1)
        {
            try
            {
                SAPbouiCOM.Form frm = this.getForm(frmId);
                SAPbouiCOM.Matrix matrix = frm.Items.Item(matrixId).Specific;
                return this._putCellValue(ref matrix, colName, value, row, true);
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Inserindo valor em " + matrixId + ", na coluna " + colName);
                return false;
            }
        }

        /// <summary>
        /// Coloca um valor em uma celula da matriz
        /// </summary>
        /// <param name="matrix">SAPbouiCOM.Matrix</param>
        /// <param name="colName"></param>
        /// <param name="value"></param>
        /// <param name="row">Se não informado, pega o último row por default</param>
        /// <returns></returns>
        public bool putComboCellValue(ref SAPbouiCOM.Matrix matrix, string colName, dynamic value, int row = -1)
        {
            return this._putCellValue(ref matrix, colName, value, row, true);
        }

        /// <summary>
        /// Coloca um valor em uma celula da matriz
        /// </summary>
        /// <param name="matrix"></param>
        /// <param name="colName"></param>
        /// <param name="value"></param>
        /// <param name="row"></param>
        /// <param name="combo"></param>
        /// <returns></returns>
        private bool _putCellValue(ref SAPbouiCOM.Matrix matrix, string colName, dynamic value, int row = -1, bool combo = false)
        {
            bool res = false;
            try
            {
                int r = matrix.GetNextSelectedRow();
                if(row == -1)
                {
                    row = matrix.RowCount;
                }
                SAPbouiCOM.Cell c = matrix.Columns.Item(colName).Cells.Item(row);
                if(combo)
                {
                    c.Specific.Select(value);
                } else
                {
                    bool travado = (!matrix.Columns.Item(colName).Editable);
                    if(travado)
                    {
                        matrix.Columns.Item(colName).Editable = true;
                    }
                    try
                    {
                        c.Specific.Value = value;
                    } catch(Exception e)
                    {
                        this.Addon.StatusAlerta(e.Message);
                    }

                    if(travado)
                    {
                        matrix.Columns.Item(colName).Editable = false;
                        matrix.Columns.Item(colName).ForeColor = 1;
                    }
                }
                res = true;

            } catch(Exception e)
            {
                //this.Addon.DesenvTimeError(e, " - Inserindo valor na coluna " + colName);
            }
            return res;
        }



        #region :: "onMatrix" DEPRECATED

        /// <summary>
        /// DEPRECATED :: Salva alterações feitas em um matrix com dataset não UDO.
        /// </summary>
        /// <param name="formId"></param>
        /// <param name="matrixId"></param>
        /// <param name="table"></param>
        public void UpdateUserMatrix(string frmId, string matrixId, string table)
        {
            SAPbouiCOM.Form frm = this.getForm(frmId);
            SAPbouiCOM.Matrix matrix = frm.Items.Item(matrixId).Specific;
            matrix.FlushToDataSource();
            this.Addon.DtSources.saveUserDataSource(table, frmId);
        }

        /// <summary>
        /// DEPRECATED :: Remove um row do matrix
        /// </summary>
        /// <param name="formId"></param>
        /// <param name="matrixId"></param>
        public void delOnMatrix(string frmId, string matrixId, string table = "")
        {
            try
            {
                SAPbouiCOM.Form frm = this.getForm(frmId);
                SAPbouiCOM.Matrix matrix = frm.Items.Item(matrixId).Specific;
                int r = matrix.GetNextSelectedRow();
                if(r > -1)
                {

                    // Se informado, remove na tabela
                    if(!String.IsNullOrEmpty(table))
                    {
                        this.Addon.DtSources.dtsDelete(table, new Dictionary<string, dynamic>() { 
                            {"Code", this.getCellValue(ref matrix, "Code")}
                        });
                    }

                    matrix.DeleteRow(r);
                    matrix.FlushToDataSource();
                    frm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                } else
                {
                    this.Addon.StatusAlerta("Nenhuma linha está selecionada. Clique na coluna '#' para selecionar a linha desejada.");
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Removendo row em " + matrixId);
            }
        }

        /// <summary>
        /// DEPRECATED :: Limpa toda a matriz - NAO IMPLEMENTADA AINDA - NÃO USAR
        /// </summary>
        /// <param name="formId"></param>
        /// <param name="matrixId"></param>
        private void clearMatrix(string frmId, string matrixId)
        {
            try
            {
                SAPbouiCOM.Form frm = this.getForm(frmId);
                SAPbouiCOM.Matrix matrix = frm.Items.Item(matrixId).Specific;
                matrix.Clear();
                matrix.FlushToDataSource();
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Limpando " + matrixId);
            }
        }

        #endregion



        public int GetMatrixRowIndex(string mtxId, string frmId = "")
        {
            SAPbouiCOM.Matrix mtx = GetItem(mtxId, frmId).Specific;
            return this.GetMatrixRowIndex(mtx);
        }

        public int GetMatrixRowIndex(SAPbouiCOM.Matrix mtx)
        {
            int res = 0;
            SAPbouiCOM.CellPosition c = mtx.GetCellFocus();
            if(null != c)
            {
                res = c.rowIndex;
            } else
            {
                res = mtx.GetNextSelectedRow();
            }
            return res;
        }

        /// <summary>
        /// Handler para o evento form_update para salvar os datasources definidos
        /// em saveDatasources.
        /// </summary>
        public void saveUserDataSources()
        {
            foreach(string dtsId in this.FormParams.SaveDatasources)
            {
                this.Addon.DtSources.saveUserDataSource(dtsId, this.FormId);
            }
        }

        /// <summary>
        /// Handler de atalho para botões fechar
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnFecharOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if(this.SapForm != null)
                {
                    this.SapForm.Close();
                }
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Ao tentar fechar o form");
            }
        }
        public void btnCloseOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            this.btnFecharOnClick(ref evObj, out BubbleEvent);
        }

        public void ShowMessage(string msg)
        {
            this.Addon.ShowMessage(msg);
        }

        #endregion


        #region :: Validador de campos

        /// <summary>
        /// Validação de campos no formulário
        /// </summary>
        /// <returns></returns>
        public bool Validate()
        {
            bool res = true;
            this.ValidateError = false;
            
            foreach(string compId in this.ToValidate)
            {
                try
                {
                    // valor do componente
                    string val = Convert.ToString(this.GetItem(compId).Specific.Value).Trim();
                    string msg = "";
                    string emptyMsg = "";

                    // Validação de Range
                    if(this.FormParams.Controls[compId].Validate != null)
                    {

                        #region :: Inteiros

                        // Valor em componente de referncia
                        string ref_min = "";
                        string compRef = this.FormParams.Controls[compId].Validate.IntMinByComp;
                        if(!String.IsNullOrEmpty(compRef))
                        {
                            ref_min = this.GetItem(compRef).Specific.Value;
                        }

                        string ref_max = "";
                        compRef = this.FormParams.Controls[compId].Validate.IntMaxByComp;
                        if(!String.IsNullOrEmpty(compRef))
                        {
                            ref_max = this.GetItem(compRef).Specific.Value;
                        }

                        // Range de int
                        string r_min = this.FormParams.Controls[compId].Validate.RangeIntMin;
                        string r_max = this.FormParams.Controls[compId].Validate.RangeIntMax;
                        if(!String.IsNullOrEmpty(r_min) || !String.IsNullOrEmpty(r_max) ||
                            !String.IsNullOrEmpty(ref_min) || !String.IsNullOrEmpty(ref_max))
                        {
                            this.FormParams.Controls[compId].NonEmpty = true;
                            msg += this.ValidateIntRange(val, compId, r_min, r_max, ref_min, ref_max);
                        }

                        #endregion

                        #region :: Datas

                        compRef = this.FormParams.Controls[compId].Validate.DateMinByComp;
                        if(!String.IsNullOrEmpty(compRef))
                        {
                            ref_min = this.GetItem(compRef).Specific.Value;
                        }

                        compRef = this.FormParams.Controls[compId].Validate.DateMaxByComp;
                        if(!String.IsNullOrEmpty(compRef))
                        {
                            ref_max = this.GetItem(compRef).Specific.Value;
                        }

                        // Range de date
                        r_min = this.FormParams.Controls[compId].Validate.RangeDateMin;
                        r_max = this.FormParams.Controls[compId].Validate.RangeDateMax;
                        if(!String.IsNullOrEmpty(r_min) || !String.IsNullOrEmpty(r_max) ||
                            !String.IsNullOrEmpty(ref_min) || !String.IsNullOrEmpty(ref_max))
                        {
                            //this.FormParams.Controls[compId].NonEmpty = true;
                            msg += this.ValidateDateRange(
                                this.Addon.fromSAPToDateStr(val), compId, 
                                this.Addon.fromSAPToDateStr(r_min), 
                                this.Addon.fromSAPToDateStr(r_max), 
                                this.Addon.fromSAPToDateStr(ref_min), 
                                this.Addon.fromSAPToDateStr(ref_max)
                            );
                        }

                        #endregion

                        // Campos obrigatórios
                        //emptyMsg = this.FormParams.Controls[compId].Validate.OnEmptyError;
                    }

                    if(this.FormParams.Controls[compId].NonEmpty || !String.IsNullOrEmpty(emptyMsg))
                    {
                        try
                        {
                            if(String.IsNullOrEmpty(val))
                            {
                                msg += !String.IsNullOrEmpty(emptyMsg)
                                    ? emptyMsg
                                    : "\nAtenção! O campo '" + this.FormParams.Controls[compId].Label + "' é de preenchimento obrigatório!"
                                ;
                            }
                        } catch(Exception e)
                        {
                            this.Addon.DesenvTimeError(e, " validando o valor do campo " + compId);
                        }
                    }

                    if(!String.IsNullOrEmpty(msg))
                    {
                        this.ValidateError = true;
                        this.ShowMessage(msg);
                        this.Addon.StatusInfo(msg);
                        this.GetItem(compId).Click();
                        return false;
                    }
                   
                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, " validando o valor do campo " + compId);
                }
            }

            return res;
        }

        /// <summary>
        /// Valida o valor int de "val" entre (r_min | ref_min) e (r_max_ref_max)
        /// </summary>
        /// <param name="val"></param>
        /// <param name="compId"></param>
        /// <param name="r_min"></param>
        /// <param name="r_max"></param>
        /// <param name="ref_min"></param>
        /// <param name="ref_max"></param>
        /// <returns></returns>
        public string ValidateIntRange(string val, string compId, string r_min = "", string r_max = "", string ref_min = "", string ref_max = "")
        {
            string msg = "";
            bool erro = false;
            if(!String.IsNullOrEmpty(val) && (!String.IsNullOrEmpty(r_min) || !String.IsNullOrEmpty(r_max) || !String.IsNullOrEmpty(ref_min) || !String.IsNullOrEmpty(ref_max)))
            {
                float v, n;
                this.FormParams.Controls[compId].NonEmpty = true;
                msg = "\nAtenção! O valor do campo '" + this.FormParams.Controls[compId].Label + "' ";

                try
                {
                    v = float.Parse(val, System.Globalization.CultureInfo.InvariantCulture.NumberFormat);
                } catch
                {
                    v = 0;
                }
                if(!String.IsNullOrEmpty(r_min) || !String.IsNullOrEmpty(ref_min))
                {
                    try
                    {
                        n = float.Parse((!String.IsNullOrEmpty(ref_min) ? ref_min : r_min), System.Globalization.CultureInfo.InvariantCulture.NumberFormat);
                    } catch
                    {
                        n = 0;
                    }
                    if(v < n)
                    {
                        msg += "deverá ser superior à '" + (!String.IsNullOrEmpty(ref_min) ? ref_min : r_min) + "'";
                        erro = true;
                    }
                }
                if(!String.IsNullOrEmpty(r_max) || !String.IsNullOrEmpty(ref_max))
                {
                    try
                    {
                        n = float.Parse((!String.IsNullOrEmpty(ref_max) ? ref_max : r_max), System.Globalization.CultureInfo.InvariantCulture.NumberFormat);
                    } catch
                    {
                        n = 0;
                    }
                    if(v > n)
                    {
                        msg += (erro ? "e" : "deverá ser") + " inferior à '" + (!String.IsNullOrEmpty(ref_max) ? ref_max : r_max) + "'";
                        erro = true;
                    }
                }
            }
            return (erro ? msg : "");
        }

        /// <summary>
        /// Valida o valor date de "val" entre (r_min | ref_min) e (r_max_ref_max)
        /// </summary>
        /// <param name="val"></param>
        /// <param name="compId"></param>
        /// <param name="r_min">Passar a data no formato dd/mm/yyyy (use this.Addon.fromSAPToDate se necessario)</param>
        /// <param name="r_max"></param>
        /// <param name="ref_min"></param>
        /// <param name="ref_max"></param>
        /// <returns></returns>
        public string ValidateDateRange(string val, string compId, string r_min = "", string r_max = "", string ref_min = "", string ref_max = "")
        {
            string msg = "";
            bool erro = false;
            if(!String.IsNullOrEmpty(val) && (!String.IsNullOrEmpty(r_min) || !String.IsNullOrEmpty(r_max) || !String.IsNullOrEmpty(ref_min) || !String.IsNullOrEmpty(ref_max)))
            {
                this.FormParams.Controls[compId].NonEmpty = true;
                msg = "\nAtenção! A data no campo '" + this.FormParams.Controls[compId].Label + "' ";
                
                DateTime n;
                DateTime v = Convert.ToDateTime(val); // DateTime.ParseExact(val, "dd/mm/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                if(!String.IsNullOrEmpty(r_min) || !String.IsNullOrEmpty(ref_min))
                {
                    n = Convert.ToDateTime((!String.IsNullOrEmpty(ref_min) ? ref_min : r_min)); //DateTime.ParseExact((!String.IsNullOrEmpty(ref_min) ? ref_min : r_min), "dd/mm/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    if(v < n)
                    {
                        msg += "deverá ser superior à '" + (!String.IsNullOrEmpty(ref_min) ? ref_min : r_min) + "'";
                        erro = true;
                    }
                }
                if(!String.IsNullOrEmpty(r_max) || !String.IsNullOrEmpty(ref_max))
                {
                    n = Convert.ToDateTime((!String.IsNullOrEmpty(ref_max) ? ref_max : r_max)); //, "dd/mm/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    if(v > n)
                    {
                        msg += (erro ? "e" : "deverá ser") + " inferior à '" + (!String.IsNullOrEmpty(ref_max) ? ref_max : r_max) + "'";
                        erro = true;
                    }
                }
            }
            return (erro ? msg : "");
        }


        /// <summary>
        /// Verifica o preenchimento de campos obrigatórios no form.
        /// </summary>
        /// <param name="campos">Nome do campo / Label a ser exibido na mensagem</param>
        /// <param name="form">Opcional</param>
        /// <returns></returns>
        public String CheckEmpty(Dictionary<string, string> campos, SAPbouiCOM.Form form = null)
        {
            String msg = "";

            if(form == null)
            {
                form = this.getForm();
            }

            string val;
            foreach(KeyValuePair<string, string> campo in campos)
            {
                try
                {
                    val = this.GetItem(campo.Key, form).Specific.Value;
                    if(string.IsNullOrEmpty(val))
                    {
                        msg += "\n             - " + campo.Value;
                    }
                } catch
                {

                }
            }

            if(!String.IsNullOrEmpty(msg))
            {
                msg = "Atenção! \nOs campos abaixo são de preenchimento obrigatório:" + msg + "\nPor favor, verifique os valores antes de continuar.";
            }

            return msg;
        }

        /// <summary>
        /// Verifica se as datas de um período estão ok.
        /// </summary>
        /// <param name="dtInicio"></param>
        /// <param name="dtFinal"></param>
        /// <returns></returns>
        public String CheckPeriodo(string dtInicio, string dtFinal)
        {
            string msg = "";
            try
            {
                DateTime dtIni = this.Addon.fromSAPToDate(dtInicio);
                DateTime dtFim = this.Addon.fromSAPToDate(dtFinal);

                if(dtFim < dtIni)
                {
                    msg = "Atenção!\nA data final não pode ser menor que a data inicial.";
                }
            } catch(Exception e)
            {
                msg = e.Message;
            }

            return msg;
        }


        /// <summary>
        /// Executa a verificacao de valores em um datasource.
        /// </summary>
        /// <param name="formId"></param>
        public StringBuilder ValidateXml(string formId = "")
        {
            StringBuilder msg = new StringBuilder();
            if(String.IsNullOrEmpty(formId))
            {
                formId = this.FormId;
            }

            if(this.Addon.Xml.Validacao != null)
            {
                foreach(XmlNode form in this.Addon.Xml.Validacao)
                {
                    foreach(XmlNode table in form.ChildNodes)
                    {
                        string tbl = table.Attributes["id"].Value;
                        int total = this.GetCount(tbl);
                        for(int x = 0; x < total; x++)
                        {
                            foreach(XmlNode field in table.ChildNodes)
                            {
                                try
                                {
                                    string tipo = field.Attributes["tipo"].Value;
                                    string fld = field.Attributes["id"].Value;
                                    string erro = field.InnerText;

                                    switch(tipo)
                                    {
                                        // Valores obrigatorios:
                                        case "not_empty":
                                            if(String.IsNullOrEmpty(this.GetValue(tbl, fld, x)))
                                            {
                                                msg.AppendLine(fld + ": " + erro);
                                                this.Addon.StatusErro("Campo de preenchimento obrigatório: " + fld);
                                            }
                                            break;

                                    }
                                } catch(Exception e)
                                {
                                    this.Addon.DesenvTimeError(e, "Erro em addon.xml");
                                }
                            }
                        }
                    }
                }
            }

            // Retorna msgs se houver erro
            return msg;
        }



        #endregion


        #region :: ChooseFromList

        /// <summary>
        /// Acrescenta um ChooseFromList pré-definido em CFLType.
        /// </summary>
        /// <todo>Implementar cache de Uid, de forma que mais de um CFL do mesmo tipo possa ser usado sem que 
        /// o desenvolvedor tenha que se definir UIDs no código
        /// </todo>
        /// <param name="CFLType">Tipo do choose</param>
        /// <param name="Alias"></param>
        /// <param name="ctrl"></param>
        /// <param name="Field"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddChooseFromList(CFLType CFLType, String ctrl = null, SAPbouiCOM.Form form = null, String Uid = null, String Alias = null)
        {
            return this.AddDefChooseFromList(CFLType, null, ctrl, form, Uid, Alias);
        }

        internal SAPbouiCOM.ChooseFromList AddDefChooseFromList(CFLType CFLType, ColCompDefinitions def = null, String ctrl = null, SAPbouiCOM.Form form = null, String Uid = null, String Alias = null)
        {
            SAPbouiCOM.ChooseFromList cfl = null;
            switch(CFLType)
            {
                case TShark.CFLType.cflUDO:
                    Alias = (String.IsNullOrEmpty(Alias) ? "UDOCode" : Alias);
                    Uid = (String.IsNullOrEmpty(Uid) ? "CFLUDO" : Uid);
                    cfl = this.AddCFLUDO(ctrl, Alias, Uid, def.ChooseFromListUDOName, form);
                    break;

                case TShark.CFLType.cflItens:
                    Alias = (String.IsNullOrEmpty(Alias) ? "ItemCode" : Alias);
                    Uid = (String.IsNullOrEmpty(Uid) ? "CFLItens" : Uid);
                    cfl = this.AddCFLItens(ctrl, Alias, Uid, form);
                    break;

                case TShark.CFLType.cflServicos:
                    Alias = (String.IsNullOrEmpty(Alias) ? "ItemCode" : Alias);
                    Uid = (String.IsNullOrEmpty(Uid) ? "CFLServ" : Uid);
                    cfl = this.AddCFLServicos(ctrl, Alias, Uid, form);
                    break;

                case TShark.CFLType.cflCartaoEquip:
                    Alias = (String.IsNullOrEmpty(Alias) ? "internalSN" : Alias);
                    Uid = (String.IsNullOrEmpty(Uid) ? "CFLCartao" : Uid);
                    cfl = this.AddCFLCartaoEquipamento(ctrl, Alias, Uid, form);
                    break;

                case TShark.CFLType.cflNumSerie:
                    Alias = (String.IsNullOrEmpty(Alias) ? "ItemCode" : Alias);
                    Uid = (String.IsNullOrEmpty(Uid) ? "CFLSerial" : Uid);
                    cfl = this.AddCFLNumSerie(ctrl, Alias, Uid, form);
                    break;

                case TShark.CFLType.cflDepositos:
                    Alias = (String.IsNullOrEmpty(Alias) ? "WhsCode" : Alias);
                    Uid = (String.IsNullOrEmpty(Uid) ? "CFLDepositos" : Uid);
                    cfl = this.AddCFLDepositos(ctrl, Alias, Uid, form);
                    break;

                case TShark.CFLType.cflBusinessPartners:
                    Alias = (String.IsNullOrEmpty(Alias) ? "CardCode" : Alias);
                    Uid = (String.IsNullOrEmpty(Uid) ? "CFLBp" : Uid);
                    cfl = this.AddCFLBusinessPartners(ctrl, Alias, Uid, form);
                    break;

                case TShark.CFLType.cflClientes:
                    Alias = (String.IsNullOrEmpty(Alias) ? "CardCode" : Alias);
                    Uid = (String.IsNullOrEmpty(Uid) ? "CFLCli" : Uid);
                    cfl = this.AddCFLClientes(ctrl, Alias, Uid, form);
                    break;

                case TShark.CFLType.cflFornecedores:
                    Alias = (String.IsNullOrEmpty(Alias) ? "CardCode" : Alias);
                    Uid = (String.IsNullOrEmpty(Uid) ? "CFLForn" : Uid);
                    cfl = this.AddCFLFornecedores(ctrl, Alias, Uid, form);
                    break;

                case TShark.CFLType.cflLeads:
                    Alias = (String.IsNullOrEmpty(Alias) ? "CardCode" : Alias);
                    Uid = (String.IsNullOrEmpty(Uid) ? "CFLLead" : Uid);
                    cfl = this.AddCFLLeads(ctrl, Alias, Uid, form);
                    break;

                case TShark.CFLType.cflUsuarios:
                    Alias = (String.IsNullOrEmpty(Alias) ? "USERID" : Alias);
                    Uid = (String.IsNullOrEmpty(Uid) ? "CFLUsuarios" : Uid);
                    cfl = this.AddCFLUsuarios(ctrl, Alias, Uid, form);
                    break;

                case TShark.CFLType.cflFuncionarios:
                    Alias = (String.IsNullOrEmpty(Alias) ? "FirstName" : Alias);
                    Uid = (String.IsNullOrEmpty(Uid) ? "CFLFuncionarios" : Uid);
                    cfl = this.AddCFLFuncionarios(ctrl, Alias, Uid, form);
                    break;

                case TShark.CFLType.cflAtividades:
                    Alias = (String.IsNullOrEmpty(Alias) ? "ClgCode" : Alias);
                    Uid = (String.IsNullOrEmpty(Uid) ? "CFLAtividades" : Uid);
                    cfl = this.AddCFLAtividades(ctrl, Alias, Uid, form);
                    break;

                case TShark.CFLType.cflOportunidades:
                    Alias = (String.IsNullOrEmpty(Alias) ? "DocEntry" : Alias);
                    Uid = (String.IsNullOrEmpty(Uid) ? "CFLOportunidades" : Uid);
                    cfl = this.AddCFLOportunidades(ctrl, Alias, Uid, form);
                    break;

                case TShark.CFLType.cflContaContabil:
                    Alias = (String.IsNullOrEmpty(Alias) ? "AcctCode" : Alias);
                    Uid = (String.IsNullOrEmpty(Uid) ? "CFLContaContabil" : Uid);
                    cfl = this.AddCFLContaContabil(ctrl, Alias, Uid, form);
                    break;
            }


            if(def != null)
            {
                def.ChooseFromListAlias = Alias;
                def.ChooseFromListUID = Uid;

                if(def.ChooseFromListConds != null)
                {
                    cfl.SetConditions(def.ChooseFromListConds);
                }
            }

            return cfl;
        }

        /// <summary>
        /// Retorna um LinkedObject com base no tipo de
        /// ChooseFromList
        /// </summary>
        /// <param name="CFLType"></param>
        /// <returns></returns>
        internal SAPbouiCOM.BoLinkedObject GetLinkedObjByCLF(CFLType CFLType)
        {
            SAPbouiCOM.BoLinkedObject lko = SAPbouiCOM.BoLinkedObject.lf_None;
            switch(CFLType)
            {
                case TShark.CFLType.cflUDO:
                    lko = SAPbouiCOM.BoLinkedObject.lf_UserDefinedObject;
                    break;

                case TShark.CFLType.cflItens:
                    lko = SAPbouiCOM.BoLinkedObject.lf_Items;
                    break;

                case TShark.CFLType.cflServicos:
                    lko = SAPbouiCOM.BoLinkedObject.lf_Items;
                    break;

                case TShark.CFLType.cflCartaoEquip:
                    lko = SAPbouiCOM.BoLinkedObject.lf_InstallBase;
                    break;

                case TShark.CFLType.cflNumSerie:
                    lko = SAPbouiCOM.BoLinkedObject.lf_SerialNumbersForItems;
                    break;

                case TShark.CFLType.cflDepositos:
                    lko = SAPbouiCOM.BoLinkedObject.lf_Warehouses;
                    break;

                case TShark.CFLType.cflClientes:
                    lko = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
                    break;

                case TShark.CFLType.cflFornecedores:
                    lko = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
                    break;

                case TShark.CFLType.cflLeads:
                    lko = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
                    break;

                case TShark.CFLType.cflUsuarios:
                    lko = SAPbouiCOM.BoLinkedObject.lf_User;
                    break;

                case TShark.CFLType.cflFuncionarios:
                    lko = SAPbouiCOM.BoLinkedObject.lf_Employee;
                    break;

                case TShark.CFLType.cflAtividades:
                    lko = SAPbouiCOM.BoLinkedObject.lf_ServiceCall;
                    break;

                case TShark.CFLType.cflOportunidades:
                    lko = SAPbouiCOM.BoLinkedObject.lf_SalesOpportunity;
                    break;

                case TShark.CFLType.cflContaContabil:
                    lko = SAPbouiCOM.BoLinkedObject.lf_GLAccounts;
                    break;
            }

            return lko;
        }

        /// <summary>
        /// Acrescenta um choose from list para um UDO
        /// </summary>
        /// <param name="UDOName">Nome do UDO</param>
        /// <param name="Uid"></param>
        /// <param name="Alias"></param>
        /// <param name="ctrl"></param>
        /// <param name="form"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddChooseFromList(string UDOName, string Uid, String Alias, String ctrl = null, SAPbouiCOM.Form form = null)
        {
            return this.AddChooseFromList(SAPbouiCOM.BoLinkedObject.lf_UserDefinedObject, Uid, ctrl, form, Alias, UDOName);
        }

        /// <summary>
        /// Acrescenta um choose from list
        /// </summary>
        /// <param name="ObjectType"></param>
        /// <param name="Uid"></param>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddChooseFromList(SAPbouiCOM.BoLinkedObject ObjectType, string Uid, String ctrl = null, SAPbouiCOM.Form form = null, String Alias = null, string UDO = "")
        {
            SAPbouiCOM.ChooseFromList cfl = null;

            if(null == form)
            {
                form = this.getForm();
            }

            if(String.IsNullOrEmpty(Alias))
            {
                this.Addon.StatusErro("Alias não pode ser null - Acrescentando ChooseFromList " + Uid);
            }

            try
            {

                // Cria CFL:
                SAPbouiCOM.ChooseFromListCreationParams cflParams = Addon.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                cflParams.MultiSelection = false;
                cflParams.ObjectType = (String.IsNullOrEmpty(UDO) ? ObjectType.GetHashCode().ToString() : UDO);
                cflParams.UniqueID = Uid; 
                try // Porque o cfl já existe em um form xml
                {
                    cfl = form.ChooseFromLists.Add(cflParams);
                } catch {
                    cfl = form.ChooseFromLists.Item(Uid);
                }

                // Seta componente:
                if(!String.IsNullOrEmpty(ctrl))
                {
                    SAPbouiCOM.Item item = this.GetItem(ctrl);
                    item.Specific.ChooseFromListUID = Uid;
                    item.Specific.ChooseFromListAlias = Alias;

                    if(String.IsNullOrEmpty(UDO))
                    {
                        try
                        {
                            string lkid = ctrl.Remove(0, 2);
                            lkid = "lk" + lkid;
                            SAPbouiCOM.Item link = null;
                            try
                            {
                                link = form.Items.Add(lkid, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            } catch
                            {
                                link = form.Items.Item(lkid);
                            }
                            if(link != null)
                            {
                                ((SAPbouiCOM.LinkedButton)link.Specific).LinkedObject = ObjectType;
                                link.LinkTo = ctrl;
                                this.CalcLinkButtonPos(item, form);
                            }

                        } catch(Exception e)
                        {
                            this.Addon.StatusErro(e.Message);
                        }
                    }

                    // Ajusta evento para selecao do item
                    if(this.EventMethods.ContainsKey(ctrl + "OnChooseFromList"))
                    {
                        Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST, form.TypeEx, ctrl, ctrl + "OnChooseFromList");
                    }
                    if(this.EventMethods.ContainsKey(form.UniqueID + "OnChooseFromList"))
                    {
                        Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST, form.TypeEx, ctrl, form.UniqueID + "OnChooseFromList");
                    }
                }


            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Acrescentando ChooseFromList " + Uid);
            }

            return cfl;
        }

        /// <summary>
        /// ChooseFromList de Itens
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <param name="Uid"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddCFLUDO(String ctrl, String Alias, string Uid, string UDO, SAPbouiCOM.Form form = null)
        {
            return this.AddChooseFromList(UDO, Uid, Alias, ctrl, form);
        }

        /// <summary>
        /// ChooseFromList de Itens
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <param name="Uid"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddCFLItens(String ctrl, String Alias, string Uid, SAPbouiCOM.Form form = null)
        {
            return this.AddChooseFromList(SAPbouiCOM.BoLinkedObject.lf_Items, Uid, ctrl, form, Alias);
        }

        /// <summary>
        /// ChooseFromList de Usuários
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <param name="Uid"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddCFLUsuarios(String ctrl, String Alias, string Uid, SAPbouiCOM.Form form = null)
        {
            return this.AddChooseFromList(SAPbouiCOM.BoLinkedObject.lf_User, Uid, ctrl, form, Alias);
        }

        /// <summary>
        /// ChooseFromList de Funcionarios
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <param name="Uid"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddCFLFuncionarios(String ctrl, String Alias, string Uid, SAPbouiCOM.Form form = null)
        {
            return this.AddChooseFromList(SAPbouiCOM.BoLinkedObject.lf_Employee, Uid, ctrl, form, Alias);
        }

        /// <summary>
        /// ChooseFromList de Atividades
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <param name="Uid"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddCFLAtividades(String ctrl, String Alias, string Uid, SAPbouiCOM.Form form = null)
        {
            return this.AddChooseFromList(SAPbouiCOM.BoLinkedObject.lf_ContactWithCustAndVend, Uid, ctrl, form, Alias);
        }


        /// <summary>
        /// ChooseFromList de Oportunidades de Venda
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <param name="Uid"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddCFLOportunidades(String ctrl, String Alias, string Uid, SAPbouiCOM.Form form = null)
        {
            return this.AddChooseFromList(SAPbouiCOM.BoLinkedObject.lf_SalesOpportunity, Uid, ctrl, form, Alias);
        }


        /// <summary>
        /// ChooseFromList de Contas Cobtábeis
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <param name="Uid"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddCFLContaContabil(String ctrl, String Alias, string Uid, SAPbouiCOM.Form form = null)
        {
            return this.AddChooseFromList(SAPbouiCOM.BoLinkedObject.lf_GLAccounts, Uid, ctrl, form, Alias);
        }



        /// <summary>
        /// ChooseFromList de Servicos
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <param name="Uid"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddCFLServicos(String ctrl, String Alias, string Uid, SAPbouiCOM.Form form = null)
        {
            SAPbouiCOM.ChooseFromList cfl = this.AddChooseFromList(SAPbouiCOM.BoLinkedObject.lf_Items, Uid, ctrl, form, Alias);

            try
            {
                SAPbouiCOM.Conditions conds = cfl.GetConditions();
                SAPbouiCOM.Condition con = conds.Add();
                con.Alias = "ItemClass";
                con.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                con.CondVal = "1";
                cfl.SetConditions(conds);

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Definindo ChooseFromList Serviços");
            }

            return cfl;
        }

        /// <summary>
        /// ChooseFromList de Cartão de equipamentos
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <param name="Uid"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddCFLCartaoEquipamento(String ctrl, String Alias, string Uid, SAPbouiCOM.Form form = null)
        {
            return this.AddChooseFromList(SAPbouiCOM.BoLinkedObject.lf_InstallBase, Uid, ctrl, form, Alias);
        }

        /// <summary>
        /// ChooseFromList de Lotes
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <param name="Uid"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddCFLLotes(String ctrl, String Alias, string Uid, SAPbouiCOM.Form form = null)
        {
            return this.AddChooseFromList(SAPbouiCOM.BoLinkedObject.lf_ItemBatchNumbers, Uid, ctrl, form, Alias);
        }

        /// <summary>
        /// ChooseFromList de Numero de serie
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <param name="Uid"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddCFLNumSerie(String ctrl, String Alias, string Uid, SAPbouiCOM.Form form = null)
        {
            return this.AddChooseFromList(SAPbouiCOM.BoLinkedObject.lf_SerialNumbersForItems, Uid, ctrl, form, Alias);
        }

        /// <summary>
        /// ChooseFromList de depósitos
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <param name="Uid"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddCFLDepositos(String ctrl, String Alias, string Uid, SAPbouiCOM.Form form = null)
        {
            return this.AddChooseFromList(SAPbouiCOM.BoLinkedObject.lf_Warehouses, Uid, ctrl, form, Alias);
        }

        /// <summary>
        /// ChooseFromList de parceiros
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <param name="Uid"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddCFLBusinessPartners(String ctrl, String Alias, string Uid, SAPbouiCOM.Form form = null)
        {
            return this.AddChooseFromList(SAPbouiCOM.BoLinkedObject.lf_BusinessPartner, Uid, ctrl, form, Alias);
        }

        /// <summary>
        /// ChooseFromList de clientes
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <param name="Uid"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddCFLClientes(String ctrl, String Alias, string Uid, SAPbouiCOM.Form form = null)
        {
            return this._setCFLBusinessPartnerType(ctrl, form, Alias, "C", Uid);
        }

        /// <summary>
        /// ChooseFromList de Fornecedores
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <param name="Uid"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddCFLFornecedores(String ctrl, String Alias, string Uid, SAPbouiCOM.Form form = null)
        {
            return this._setCFLBusinessPartnerType(ctrl, form, Alias, "S", Uid);
        }

        /// <summary>
        /// ChooseFromList de leads
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="Alias"></param>
        /// <param name="Uid"></param>
        /// <returns></returns>
        public SAPbouiCOM.ChooseFromList AddCFLLeads(String ctrl, String Alias, string Uid, SAPbouiCOM.Form form = null)
        {
            return this._setCFLBusinessPartnerType(ctrl, form, Alias, "L", Uid);
        }

        internal SAPbouiCOM.ChooseFromList _setCFLBusinessPartnerType(String ctrl, SAPbouiCOM.Form form, String Alias, String type, String Uid)
        {
            SAPbouiCOM.ChooseFromList cfl = this.AddChooseFromList(SAPbouiCOM.BoLinkedObject.lf_BusinessPartner, Uid, ctrl, form, Alias);

            try
            {
                SAPbouiCOM.Conditions conds = cfl.GetConditions();
                SAPbouiCOM.Condition con = conds.Add();
                con.Alias = "CardType";
                con.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                con.CondVal = type;
                cfl.SetConditions(conds);

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Definindo ChooseFromList PN type: " + type);
            }

            return cfl;
        }


        #endregion


        #region :: Manipulação de Dados

        /// <summary>
        /// Coloca um form UDO em modo de inserção.
        /// </summary>
        public void FormUDOSetAddMode()
        {
            try
            {
                this.timerUDOAdd.Stop();
                this.Addon.SBO_Application.ActivateMenuItem("1282");
            } catch(Exception e)
            {

            }
        }

        #region :: Inserts

        /// <summary>
        /// Insere um row diretamente em um DBDataSources do form atual, ou do especificado.
        /// </summary>
        /// <param name="values">Row a ser inserido ({field, value})</param>
        /// <param name="table">Não esquecer o arroba.</param>
        /// <param name="formId">Se informado, identifica o form via "Forms.Item()"</param>
        /// <param name="formCount">Se informado, identifica o form via "Forms.GetForm()"</param>
        public int InsertOnClient(Dictionary<string, dynamic> values, string table, string formId = "", int formCount = 0)
        {
            SAPbouiCOM.Form form = this.getForm(formId, formCount);
            return this.InsertOnClient(values, table, form);
        }

        /// <summary>
        /// Insere um row diretamente em um DBDataSources do form especificado.
        /// </summary>
        /// <param name="values">Row a ser inserido ({field, value})</param>
        /// <param name="table">Não esquecer o arroba.</param>
        /// <param name="form">Instancia do form a ser utilizado</param>
        public int InsertOnClient(Dictionary<string, dynamic> values, string table, SAPbouiCOM.Form form)
        {
            int count = 0;
            string fld = "";
            string val = "";
            try
            {
                string arroba = table.Substring(0, 1);

                // DBDataSources
                if(arroba == "@" || table.Length == 4)
                {
                    SAPbouiCOM.DBDataSource tb = form.DataSources.DBDataSources.Item(table);
                    tb.InsertRecord(tb.Size);
                    foreach(KeyValuePair<string, dynamic> item in values)
                    {
                        fld = item.Key;
                        val = Convert.ToString(item.Value);
                        tb.SetValue(fld, tb.Size - 1, val);
                    }
                    count = tb.Size;

                // DataTables
                } else
                {
                    SAPbouiCOM.DataTable tb = form.DataSources.DataTables.Item(table);
                    tb.Rows.Add();
                    tb.Rows.Offset = tb.Rows.Count - 1;
                    foreach(KeyValuePair<string, dynamic> item in values)
                    {
                        try
                        {
                            tb.SetValue(item.Key, tb.Rows.Offset, item.Value);
                        } catch { }
                    }
                    
                }

                

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Inserindo valor na tabela " + table + " - " + fld);
            }
            return count;
        }

        /// <summary>
        /// Insere "onClient" em um DBDatasource que representa uma tabela M:N EM UM UDO
        /// </summary>
        /// <param name="tbMN">Não esquecer o @</param>
        /// <param name="tbDados">Não esquecer o @</param>
        /// <param name="fieldMap">Mapeamento de campos da tabela MN e de Dados</param>
        /// <param name="mtxId">Matriz opcional a ser atualizado</param>
        /// <param name="sql">SQL Opcional</param>
        /// <param name="formId"></param>
        /// <param name="formCount"></param>
        /// <returns></returns>
        public bool InsertOnClientMN(string tbMN, string tbDados, Dictionary<string, string> fieldMap, string mtxId = "", string sql = "", string formId = "", int formCount = 0)
        {
            // Recupera Form e matriz
            SAPbouiCOM.Form form = this.getForm(formId, formCount);

            // Reseta dataset e matriz
            this.ClearValues(tbMN);

            // Recupera dados para inserção na tabela "left" do M:N
            SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                rs.DoQuery(String.IsNullOrEmpty(sql) ? "SELECT * FROM [" + tbDados + "]" : sql);
                rs.MoveFirst();
            } catch(Exception e)
            {
                if(this.Addon.showDesenvTimeMsgs)
                {
                    this.Addon.StatusErro("Erro em InsertOnClientMN: " + e.Message);
                    return false;
                }
            }

            // Alimenta o tbMN
            try
            {
                SAPbouiCOM.DBDataSource tb = form.DataSources.DBDataSources.Item(tbMN);
                while(!rs.EoF)
                {
                    tb.InsertRecord(tb.Size);
                    foreach(KeyValuePair<string, string> item in fieldMap)
                    {
                        tb.SetValue(item.Key, tb.Size - 1, rs.Fields.Item(item.Value).Value);
                    }
                    rs.MoveNext();
                }

                if(!String.IsNullOrEmpty(mtxId))
                {
                    SAPbouiCOM.Matrix matrix = this.GetItem(mtxId).Specific;
                    matrix.LoadFromDataSourceEx(true);
                }

            } catch(Exception e)
            {
                if(this.Addon.showDesenvTimeMsgs)
                {
                    this.Addon.StatusErro("Erro em InsertOnClientMN: " + e.Message);
                    return false;
                }
            }

            return true;
        }


        /// <summary>
        /// Insere um row em um usertable não UDO diretamente no banco.
        /// </summary>
        /// <param name="table">Não esquecer o arroba.</param>
        /// <param name="values">Row a ser inserido ({field, value})</param>
        /// <param name="mtxId">Se informado, a matriz é atualizada na operação</param>
        /// <param name="formId">Se informado, identifica o form via "Forms.Item()"</param>
        /// <param name="formCount">Se informado, identifica o form via "Forms.GetForm()"</param>
        public bool InsertOnServer(string table, Dictionary<string, dynamic> values, string mtxId = "", string formId = "", int formCount = 0)
        {
            bool res = false;
            try
            {
                SAPbouiCOM.Matrix matrix = null;
                if(!String.IsNullOrEmpty(mtxId))
                {
                    matrix = this.GetItem(mtxId, formId, formCount).Specific;

                    // Atualiza e salva alguma alteração:
                    if(matrix != null && matrix.RowCount > 0)
                    {
                        matrix.FlushToDataSource();
                        this.Addon.DtSources.saveUserDataSource(table, formId, formCount);
                    }
                }

                // Insere o novo item:
                if(values == null)
                {
                    values = new Dictionary<string, dynamic>() { };
                }
                res = this.Addon.DtSources.dtsInsert(table, values);

                if(matrix != null)
                {
                    try
                    {
                        this.RefreshMatrix(formId, mtxId, table);
                        //matrix.LoadFromDataSourceEx();
                        matrix.SetCellFocus(matrix.RowCount, 2);
                    } catch { }
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Inserindo row em " + table);
            }

            return res;
        }

        #endregion


        #region :: Updates

        /// <summary>
        /// Atualiza campos diretamente de em um row de um DBDataSource em um form específico com base no RecNo informado.
        /// </summary>
        /// <param name="values">Valores a serem atualizados ({field, value})</param>
        /// <param name="table">Se informado o arroba, trata como DBDataSources, senão como DataTables.</param>
        /// <param name="RecNo">Se informado, número do row do DBDataSource, se não atualiza o atual</param>
        /// <param name="formId">Se informado, identifica o form via "Forms.Item()"</param>
        /// <param name="formCount">Se informado, identifica o form via "Forms.GetForm()"</param>
        public void UpdateOnClient(Dictionary<string, dynamic> values, string table, int RecNo = -1, string formId = "", int formCount = 0)
        {
            SAPbouiCOM.Form form = this.getForm(formId, formCount);
            this.UpdateOnClient(values, table, form, RecNo);
        }

        /// <summary>
        /// Atualiza campos diretamente de em um row de um DBDataSource em um form específico com base no RecNo informado.
        /// </summary>
        /// <param name="values">Valores a serem atualizados ({field, value})</param>
        /// <param name="table">Se informado o arroba, trata como DBDataSources, senão como DataTables.</param>
        /// <param name="form">Instância do form onde mora o DBDataSource</param>
        /// <param name="RecNo">Se informado, número do row do DBDataSource, se não atualiza o atual</param>
        public void UpdateOnClient(Dictionary<string, dynamic> values, string table, SAPbouiCOM.Form form, int RecNo = -1)
        {
            try
            {
                string arroba = table.Substring(0,1);

                // DBDataSources
                if(arroba == "@" || table.Length == 4)
                {
                    SAPbouiCOM.DBDataSource tb = form.DataSources.DBDataSources.Item(table);
                    RecNo = (RecNo != -1 ? RecNo : tb.Offset);
                    if(RecNo < tb.Size)
                    {
                        foreach(KeyValuePair<string, dynamic> item in values)
                        {
                            try
                            {
                                tb.SetValue(item.Key, RecNo, item.Value);
                            } catch (Exception e) {
                                this.Addon.DesenvTimeError(e);
                            }
                        }
                    }

                // DataTables
                } else
                {
                    SAPbouiCOM.DataTable tb = form.DataSources.DataTables.Item(table);
                    RecNo = (RecNo != -1 ? RecNo : tb.Rows.Offset);
                    if(RecNo < tb.Rows.Count)
                    {
                        foreach(KeyValuePair<string, dynamic> item in values)
                        {
                            try
                            {
                                tb.SetValue(item.Key, RecNo, item.Value);
                            } catch { }
                        }
                    }
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Alterando valor na tabela " + table);
            }
        }

        /// <summary>
        /// Atualiza um campo específico em um row de um DBDataSource em um form específico com base no RecNo informado.
        /// </summary>
        /// <param name="field">Campo a ser alterado</param>
        /// <param name="value">Valor a ser alterado</param>
        /// <param name="table">Se informado o arroba, trata como DBDataSources, senão como DataTables.</param>
        /// <param name="form">Form, se não informado usa o atual</param>
        /// <param name="RecNo">Indice do registro no DBDataSource</param>
        public void UpdateOnClient(string field, dynamic value, string table, SAPbouiCOM.Form form = null, int RecNo = -1)
        {
            try
            {
                if(form == null)
                {
                    form = this.getForm();
                }

                string arroba = table.Substring(0, 1);

                // DBDataSources
                if(arroba == "@" || table.Length == 4)
                {
                    SAPbouiCOM.DBDataSource tb = form.DataSources.DBDataSources.Item(table);
                    RecNo = (RecNo != -1 ? RecNo : tb.Offset);
                    if(RecNo < tb.Size)
                    {
                        try
                        {
                            tb.SetValue(field, RecNo, value);
                        } catch { }
                    }

                // DataTables
                } else
                {
                    SAPbouiCOM.DataTable tb = form.DataSources.DataTables.Item(table);
                    RecNo = (RecNo != -1 ? RecNo : tb.Rows.Offset);
                    if(RecNo < tb.Rows.Count)
                    {
                        try
                        {
                            tb.SetValue(field, RecNo, value);
                        } catch { }
                    }
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Alterando valor na tabela " + table);
            }
        }

        /// <summary>
        /// Atualiza campos de UserDataSources, criados explicitamente ou por componentes sem BindTo. 
        /// </summary>
        /// <param name="values">Valores a serem atualizados ({field, value})</param>
        /// <param name="form">Instância do form, se não informado atualiza o atual</param>
        public void UpdateUserDataSource(Dictionary<string, dynamic> values, SAPbouiCOM.Form form = null)
        {
            try
            {
                if(form == null)
                {
                    form = this.getForm();
                }

                foreach(KeyValuePair<string, dynamic> item in values)
                {
                    this.UpdateUserDataSource(item.Key, item.Value, form);
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Alterando valor em UserDataSources");
            }
        }

        /// <summary>
        /// Atualiza um campo específico em UserDataSources, criados explicitamente ou por componentes sem BindTo. 
        /// </summary>
        /// <param name="key">Nome do campo</param>
        /// <param name="value">Valor a ser salvo</param>
        /// <param name="form">Form, se não informado usa o atual</param>
        public void UpdateUserDataSource(string key, dynamic value, SAPbouiCOM.Form form = null)
        {
            
            if(form == null)
            {
                form = this.getForm();
            }
            try
            {
                form.DataSources.UserDataSources.Item(key).Value = value;
            } catch(Exception e)
            {
                // this.Addon.StatusErro(e.Message);
            }
        }

        /// <summary>
        /// Atualiza um usertable não UDO diretamente no banco, com os valores passados em values.
        /// </summary>
        /// <param name="table">Não esquecer o arroba.</param>
        /// <param name="values">Valores a serem atualizados ({field, value})</param>
        /// <param name=""mtxId">Se informado, a matriz é atualizada na operação</param>
        /// <param name="formId">Se informado, identifica o form via "Forms.Item()"</param>
        /// <param name="formCount">Se informado, identifica o form via "Forms.GetForm()"</param>
        public void UpdateOnServer(string table, Dictionary<string, dynamic> values, string mtxId = "", string formId = "", int formCount = 0)
        {
            try
            {
                // Atualiza row:
                this.Addon.DtSources.dtsUpdate(table, values);

                if(!String.IsNullOrEmpty(mtxId))
                {
                    SAPbouiCOM.Matrix matrix = this.GetItem(mtxId, formId, formCount).Specific;

                    // Atualiza:
                    if(matrix != null && matrix.RowCount > 0)
                    {
                        matrix.LoadFromDataSourceEx();
                    }
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Atualizando row em " + table);
            }
        }


        public void UpdateFormCodeOn(string table, string field_to)
        {
            try
            {
                string arroba = table.Substring(0, 1);

                // DBDataSources
                if(arroba == "@" || table.Length == 4)
                {
                    SAPbouiCOM.DBDataSource tb = this.SapForm.DataSources.DBDataSources.Item(table);
                    for(int r = 0; r < tb.Size; r++)
                    {
                        tb.SetValue(field_to, r, this.UDOCode);
                    }
                    
                // DataTables
                } else
                {
                    SAPbouiCOM.DataTable tb = this.SapForm.DataSources.DataTables.Item(table);
                    for(int r = 0; r < tb.Rows.Count; r++)
                    {
                        tb.SetValue(field_to, r, this.UDOCode);
                    }
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Alterando Code na tabela " + table);
            }
        }


        /// <summary>
        /// Salva os registros de um DBDataSources no banco de dados.
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public bool SaveToServer(string table)
        {
            bool res = this.DtSources.saveUserDataSource(table);
            if(res && table == this.FormParams.MainDatasource)
            {
                SAPbouiCOM.DBDataSource dts = this.SapForm.DataSources.DBDataSources.Item(table);
                this.UDOCode = dts.GetValue("Code", dts.Offset);
            }
            return res;
        }
        
        /// <summary>
        /// Executa um UPDATE em uma tabela padrão SQL "table_to" (não SAP) no server, com base 
        /// nos dados no DataTable "data_table_from", utilizando "map_fields" para mapear campos
        /// do DataTable para a tabela no banco.
        /// </summary>
        /// <param name="table_to">Informar com o "@", caso houver</param>
        /// <param name="data_table_from">Nome do DataTable de onde buscar os valores</param>
        /// <param name="map_fields">A Chave é a coluna no DataTable e o Valor é a coluna da tabela SQL</param>
        /// <param name="no_empty_field">Se informado, caso essa campo esteja vazio, o row é excluido da operação.</param>
        /// <param name="field_key">Se não informado, utilizará "code" como chave default</param>
        /// <param name="field_key_map">Se não informado, procurará "code" para o valor no DataTable</param>
        /// <returns>True, se correr tudo bem</returns>
        public bool SaveToServer(string table_to, string data_table_from, Dictionary<string, string> map_fields = null, string no_empty_field = "", string field_key = "", string field_key_map = "")
        {
            bool res = false;
            string erro = "";

            SAPbobsCOM.Recordset rec = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string v1 = "";
                string sql = "";
                

                SAPbouiCOM.DataTable dts = this.SapForm.DataSources.DataTables.Item(data_table_from);
                if(!dts.IsEmpty)
                {

                    if(map_fields == null)
                    {
                        try
                        {
                            SAPbouiCOM.DBDataSource dtSrc;
                            try
                            {
                                dtSrc = this.SapForm.DataSources.DBDataSources.Item(table_to);
                            } catch
                            {
                                dtSrc = this.SapForm.DataSources.DBDataSources.Add(table_to);
                            }
                            map_fields = this.MapFields(dtSrc, dts);
                            
                        } catch(Exception e)
                        {
                            return this.Addon.DesenvTimeError(e, "Não foi possível mapear automaticamente as tabelas. Existe '" + table_to + "' em DBDataSources no form?");
                        }
                    }

                    for(int r = 0; r < dts.Rows.Count; r++)
                    {
                        string code = Convert.ToString(dts.GetValue((String.IsNullOrEmpty(field_key_map) ? "code" : field_key_map), r));

                        // Abre SQL
                        string v2 = "";
                        string fields = "";
                        string ins = "";
                        string upd = "";
                        string val = "";

                        bool ok = true;
                        foreach(KeyValuePair<string, string> field in map_fields)
                        {
                            erro = " - verifique o campo: " + field.Key;
                            
                            try
                            {
                                val = Convert.ToString(dts.GetValue(field.Key, r));
                            } catch(Exception e)
                            {
                                if(e.Message.Contains("DateTime"))
                                {
                                    DateTime date = dts.GetValue(field.Key, r);
                                    val = date.Day + "/" + date.Month + "/" + date.Year; //this.Addon.ToSAPDate(dts.GetValue(field.Key, r));
                                }
                            } 


                            fields += v2 + " " + field.Value;
                            ins    += v2 + " '" + val + "' ";
                            upd    += v2 + field.Value + " = '" + val + "' ";
                            v2 = ", ";
                            
                            // Valida empty_field
                            if(!String.IsNullOrEmpty(no_empty_field) && no_empty_field == field.Key && String.IsNullOrEmpty(val))
                            {
                                ok = false;
                            }
                        }

                        if(ok)
                        {

                            // Se não tem code, monta um insert
                            if(String.IsNullOrEmpty(code) || code == "0")
                            {
                                sql += v1 + " INSERT INTO [" + table_to + "] (" + fields + ")  VALUES ( " + ins + " ) ";

                            // Se tem, monta um update
                            } else
                            {
                                sql += v1 + " UPDATE [" + table_to + "] SET " + upd;
                                sql += " WHERE " + (String.IsNullOrEmpty(field_key) ? "code" : field_key) + " = '";
                                sql += code + "' ";
                            }

                            // Fecha
                            v1 = "; ";
                        }
                    }

                    // Executa:
                    erro = " - verifique o SQL " + sql;
                    if(!String.IsNullOrEmpty(sql))
                    {
                        rec.DoQuery(sql);
                    }
                    res = true;
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " em UpdateToServer " + table_to + erro);

            } finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
                rec = null;
            }

            return res;
        }



        /// <summary>
        /// Executa inserções em tabelas filhas em UDOs.
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="masterTable"></param>
        /// <param name="childTable"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public bool SaveToUDOChild(string UDOTable, string UDOCode, string UDOChild, string TableFrom, SAPbouiCOM.Form form = null, Dictionary<string, string> map_fields = null)
        {
            try
            {
                if(form == null)
                {
                    form = this.getForm();
                }
                
                DateTime date;
                string[] format = new string[] { "dd/MM/yyyy HH:mm:ss" };
                SAPbouiCOM.DataTable tbl = form.DataSources.DataTables.Item(TableFrom);
                
                if (map_fields == null){
                    try{
                        SAPbouiCOM.DBDataSource dts = form.DataSources.DBDataSources.Item(UDOChild);
                        map_fields = this.MapFields(dts, tbl);
                    } catch (Exception e) {
                        this.Addon.DesenvTimeError(e, "Se não for informado o map_fields, a tabela UDOChild deverá estar declarada em ExtraDatasources.");
                        return false;
                    }
                }

                List<Dictionary<string, dynamic>> rows = new List<Dictionary<string, dynamic>>();
                for(int r = 0; r < tbl.Rows.Count; r++)
                {
                    rows.Add(new Dictionary<string, dynamic>());
                    foreach(KeyValuePair<string, string> item in map_fields)
                    {
                        string fld = item.Key.Trim();
                        string val = Convert.ToString(tbl.GetValue(fld, r));
                        if(DateTime.TryParseExact(val, format, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.NoCurrentDateDefault, out date))
                        {
                            val = date.ToString("yyyyMMdd");
                        }

                        rows[r].Add(item.Value, val);
                    }
                }

                this.Addon.DtSources.udoChildReplace(UDOTable, UDOCode, UDOChild, rows);

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, ", em SaveToUDOChild");
                return false;
            }

            // Retorna
            return true;
        }


        #endregion


        #region :: Deletes

        /// <summary>
        /// Deleta um row diretamente de um client dataset.
        /// </summary>
        /// <param name="table">Não esquecer o arroba.</param>
        /// <param name="RecNo">Indice do row a ser removido</param>
        /// <param name="formId">Se informado, identifica o form via "Forms.Item()"</param>
        /// <param name="formCount">Se informado, identifica o form via "Forms.GetForm()"</param>
        public void DeleteOnClient(string table, int RecNo, string formId = "", int formCount = 0)
        {
            if(this.Addon.SBO_Application.MessageBox("Tem certeza de que deseja remover esse registro?", 2, "Sim", "Não") == 1)
            {
                try
                {
                    SAPbouiCOM.Form form = this.getForm(formId, formCount);
                    SAPbouiCOM.DBDataSource tb = form.DataSources.DBDataSources.Item(table);
                    if(tb.Size > 0)
                    {
                        tb.RemoveRecord((RecNo != -1 ? RecNo : tb.Offset));
                        this.getForm(formId, formCount).Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }

                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, " - Removendo registro na tabela " + table);
                }
            }
        }

        /// <summary>
        /// Remove um row em um usertable diretamente no banco.
        /// Não use em UDOs. A remoção é no banco e não em client.
        /// </summary>
        /// <param name="table">Não esquecer o arroba.</param>
        /// <param name="values">Row ({field, value}) com o code a ser removido.</param>
        public bool DeleteOnServer(string table, Dictionary<string, dynamic> values, bool quiet = false)
        {
            if(!quiet)
            {
                if(this.Addon.SBO_Application.MessageBox("Tem certeza de que deseja remover esse registro?", 2, "Sim", "Não") == 2)
                {
                    return false;
                }
            }

            bool res = true;
            try
            {
                this.Addon.DtSources.dtsDelete(table, values);

            } catch(Exception e)
            {
                res = false;
                if(!quiet)
                {
                    this.Addon.DesenvTimeError(e, " - Removendo row em " + table);
                }
            }

            return res;
        }


        /// <summary>
        /// exclui o relacionamento em uma tabela SQL padrão, e exclui a linha da matriz.
        /// </summary>
        /// <param name="mtxId">Id da matriz a ser removido o registro</param>
        /// <param name="tabela_sqlserver">Nome da tabela SQLServer padrão sem @</param>
        public bool DeleteOnServer(string mtxId, string tabela_sqlserver, string frmId = "")
        {
            bool res = false;

            if(this.Addon.SBO_Application.MessageBox("Tem certeza de que deseja remover esse registro?", 2, "Sim", "Não") == 2)
            {
                return res;
            }

            string code = "";
            try
            {
                code = this.getCellValue(mtxId, (String.IsNullOrEmpty(frmId) ? this.FormId : frmId), "code");
                
                //se tiver o codigo na coluna, entao deleta do banco
                if(!String.IsNullOrEmpty(code))
                {
                    string sql = "DELETE FROM [@" + tabela_sqlserver + "]  WHERE code = '" + code + "'";

                    SAPbobsCOM.Recordset rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rs.DoQuery(sql);
                }

                //apaga a linha na matriz.
                this.DeleteOnMatrix(mtxId, true);
                res = true;

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " removendo um registro na tabela " + tabela_sqlserver + " code: '" + code + "'");
            }

            return res;
        }

        #endregion


        internal Dictionary<string, string> MapFields(SAPbouiCOM.DBDataSource dts, SAPbouiCOM.DataTable tbl)
        {
            Dictionary<string, string> map_fields = new Dictionary<string, string>();
            for(int f = 0; f < dts.Fields.Count; f++)
            {
                string fld_from = "";
                string fld_to = dts.Fields.Item(f).Name;
                if(fld_to.Substring(0, 2) == "U_")
                {
                    try
                    {
                        fld_from = tbl.Columns.Item(fld_to).Name;
                    } catch { }
                    if(!String.IsNullOrEmpty(fld_from))
                    {
                        map_fields[fld_from] = fld_to;
                    }
                }
            }
            return map_fields;
        }

        /// <summary>
        /// Copia os rows de um DataTable para dentro de um DBDataSource. O 
        /// DBDataSource é resetado na operação.
        /// </summary>
        /// <param name="table">DataTable de origem</param>
        /// <param name="datasource">DBDataSource de destino</param>
        /// <param name="map_fields">Se null, um mapeamento dinâmico entre as fontes será tentado.</param>
        /// <param name="formId"></param>
        /// <param name="formCount"></param>
        public void CopyTableToDatasource(string table, string datasource, Dictionary<string, string> map_fields = null, string formId = "", int formCount = 0)
        {
            this._CopyTableToDatasource(table, datasource, map_fields, formId, formCount, true);
        }

        /// <summary>
        /// Acrescenta os rows de um DataTable para dentro de um DBDataSource, somando aos registros já existentes.
        /// </summary>
        /// <param name="table"></param>
        /// <param name="datasource"></param>
        /// <param name="map_fields"></param>
        /// <param name="formId"></param>
        /// <param name="formCount"></param>
        public void AppendTableToDatasource(string table, string datasource, Dictionary<string, string> map_fields = null, string formId = "", int formCount = 0)
        {
            this._CopyTableToDatasource(table, datasource, map_fields, formId, formCount, false);
        }

        internal void _CopyTableToDatasource(string table, string datasource, Dictionary<string, string> map_fields = null, string formId = "", int formCount = 0, bool clear = true)
        {
            try
            {
                SAPbouiCOM.Form form = this.getForm(formId, formCount);
                SAPbouiCOM.DBDataSource dts;
                try
                {
                    dts = form.DataSources.DBDataSources.Item(datasource);
                } catch
                {
                    dts = form.DataSources.DBDataSources.Add(datasource);
                }

                SAPbouiCOM.DataTable tbl = form.DataSources.DataTables.Item(table);
                
                string []format = new string []{"dd/MM/yyyy HH:mm:ss"};
                DateTime date;

                if(map_fields == null)
                {
                    map_fields = this.MapFields(dts, tbl);
                }

                if(clear)
                {
                    dts.Clear();
                }

                for(int r = 0; r < tbl.Rows.Count; r++)
                {
                    dts.InsertRecord(dts.Size);
                    foreach(KeyValuePair<string, string> item in map_fields)
                    {
                        string fld = item.Key.Trim();
                        dynamic tmp = tbl.GetValue(fld, r);
                        string val = "";
                        
                        if (tmp is DateTime)
                        {
                            val = Convert.ToDateTime(tmp).ToString("yyyyMMdd");
                        }
                        else if (tmp is float || tmp is double)
                        {
                            val = Convert.ToString(tmp);
                            val = val.Replace(".", "");
                            val = val.Replace(",", ".");
                        }
                        else
                        {
                            val = Convert.ToString(tmp);
                        }
                        
                        dts.SetValue(item.Value, dts.Size - 1, val);
                    }
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " em CopyTableToDatasource de " + table + " para " + datasource);
            }
        }

        /// <summary>
        /// Salva os registros exibidos em uma Matriz de um Datatable 
        /// child de um MainDatasource noObject no server, ajustando Codes e afins  
        /// </summary>
        /// <param name="matrix"></param>
        /// <param name="table"></param>
        /// <param name="datasource"></param>
        /// <param name="CodeTo"></param>
        /// <param name="map_fields"></param>
        /// <param name="formId"></param>
        /// <param name="formCount"></param>
        public void SaveMatrixToDatasource(string matrix, string table, string datasource, string CodeTo = "", Dictionary<string, string> map_fields = null, string formId = "", int formCount = 0)
        {
            SAPbouiCOM.Matrix mtx = this.GetItem(matrix).Specific;
            mtx.FlushToDataSource();
            this.SaveTableToDatasource(table, datasource, CodeTo, map_fields, formId, formCount);
        }

        /// <summary>
        /// Salva os registros de um Datatable 
        /// child de um MainDatasource noObject no server, ajustando Codes e afins
        /// </summary>
        /// <param name="table"></param>
        /// <param name="datasource"></param>
        /// <param name="CodeTo"></param>
        /// <param name="map_fields"></param>
        /// <param name="formId"></param>
        /// <param name="formCount"></param>
        public void SaveTableToDatasource(string table, string datasource, string CodeTo = "", Dictionary<string, string> map_fields = null, string formId = "", int formCount = 0)
        {
            // Prepara mapfields
            SAPbouiCOM.DataTable tb = this.SapForm.DataSources.DataTables.Item(table);
            SAPbouiCOM.DBDataSource dts;
            try
            {
                dts = this.SapForm.DataSources.DBDataSources.Item(datasource);
            } catch
            {
                dts = this.SapForm.DataSources.DBDataSources.Add(datasource);
            }
            if(map_fields == null)
            {
                map_fields = this.MapFields(dts, tb);
            }
            map_fields["Code"] = "Code";

            // Copia o datatable para o datasource
            this.CopyTableToDatasource(table, datasource, map_fields, formId, formCount);

            // Aplica o UDOCode nos registros
            if(!String.IsNullOrEmpty(CodeTo))
            {
                this.UpdateFormCodeOn(datasource, CodeTo);
            }

            // Salva os registros
            this.SaveToServer(datasource);

            // Aplica os Codes gerados no datatable
            string code = "Code";
            for(int r = 0; r < dts.Size; r++)
            {
                try
                {
                    tb.SetValue(code, r, dts.GetValue("Code", r));
                } catch
                {
                    code = "code";
                    try
                    {
                        tb.SetValue(code, r, dts.GetValue("Code", r));
                    } catch(Exception e)
                    {
                        this.Addon.DesenvTimeError(e, "Atualizando Code no DataTable depois de salvar");
                    }
                }
            }
        }



        /// <summary>
        /// Executa um SQL e alimenta this.Addon.Browser.
        /// </summary>
        /// <param name="sql"></param>
        public bool ExecSql(string sql, bool quiet = false)
        {
            return this.DtSources.Select(sql, quiet);
        }

        /// <summary>
        /// Retorna o valor de um campo em um client dataset.
        /// </summary>
        /// <param name="table">Se informado o "@" assume que table é DBDataSources, senão DataTables.</param>
        /// <param name="field">Campo que se deseja recuperar o valor.</param>
        /// <param name="RecNo">Se nao informado, assume o ultimo registro.</param>
        public string GetValue(string table, string field, int RecNo = -1)
        {
            return this.GetValue(this.SapForm, table, field, RecNo);
        }

        public string GetValue(SAPbouiCOM.Form form, string table, string field, int RecNo = -1)
        {
            string res = "";
            try
            {
                string arroba = table.Substring(0,1);

                // DBDataSources
                if(arroba == "@" || table.Length == 4)
                {
                    SAPbouiCOM.DBDataSource tb = form.DataSources.DBDataSources.Item(table);
                    if(tb.Size > 0)
                    {
                        res = tb.GetValue(field, (RecNo != -1 ? RecNo : tb.Offset));
                    }

                // DataTables
                } else
                {
                    SAPbouiCOM.DataTable tb = form.DataSources.DataTables.Item(table);
                    if(tb.Rows.Count > 0)
                    {
                        res = Convert.ToString(tb.GetValue(field, (RecNo != -1 ? RecNo : tb.Rows.Offset)));
                    }
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
        public string GetValue(string field, SAPbouiCOM.Form form = null)
        {
            string res = "";
            try
            {
                if(form == null)
                {
                    form = this.SapForm;
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
        public void ClearValues(string table)
        {
            this.ClearValues(this.SapForm, table);
        }

        public void ClearValues(SAPbouiCOM.Form form, string table)
        {
            try
            {
                form.DataSources.DBDataSources.Item(table).Clear();
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Limpando a tabela " + table);
            }
        }

        /// <summary>
        /// Retorna a quantidade de registros em um client dataset.
        /// </summary>
        /// <param name="table">Não esquecer o arroba.</param>
        public int GetCount(string table)
        {
            return this.GetCount(this.SapForm, table);
        }
        public int GetCount(SAPbouiCOM.Form form, string table)
        {
            int res = 0;
            try
            {
                SAPbouiCOM.DBDataSource tb = form.DataSources.DBDataSources.Item(table);
                res = tb.Size;
            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " - Recuperando a quantidade de registros na tabela " + table);
            }

            return res;
        }


        /// <summary>
        /// Garante que a ultima linha da matriz esteja pronta para insercao e executa 
        /// um refresh na matriz
        /// </summary>
        /// <param name="table"></param>
        /// <param name="check_field"></param>
        /// <param name="matrixId"></param>
        /// <param name="new_line_values"></param>
        public void CheckLastLine(string table, string check_field, string matrixId, Dictionary<string, dynamic> new_line_values)
        {
            bool ok = false;

            SAPbouiCOM.Matrix mtx = ((SAPbouiCOM.Matrix)this.GetItem(matrixId).Specific);

            // Verifica se vai inserir:
            string check_value = "";
            int c = this.GetCount(table);
            if(c == 0)
            {
                ok = true;
            } else
            {
                check_value = mtx.GetCellSpecific(check_field, c).Value;
                ok = (!String.IsNullOrEmpty(check_value) && !String.IsNullOrWhiteSpace(check_value));
            }

            if(ok)
            {
                if(c > 0)
                {
                    //  mtx.FlushToDataSource();
                    if(!String.IsNullOrEmpty(check_value))
                    {
                        this.UpdateOnMatrix(matrixId, new Dictionary<string, dynamic>() { { check_field, check_value } });
                    }
                }
                this.InsertOnMatrix(matrixId, new_line_values);
                //  mtx.LoadFromDataSource();
            }
        }

        #endregion


    }
}
