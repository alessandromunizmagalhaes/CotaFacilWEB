using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TShark {

    /// <summary>
    /// Armazena dimensões.
    /// By Labs - 12/2012
    /// </summary>
    public struct Bounds
    {
        public int Top;
        public int Left;
        public int Width;
        public int Height;

        internal int RealWidth;
        internal int RealHeight;

        public bool PinRight;
        public bool PinBottom;
    }

    /// <summary>
    /// Parâmetros para registro de eventos.
    /// By Labs - 12/2012
    /// </summary>
    public struct action
    {
        public SAPbouiCOM.BoEventTypes EventType;
        public string EventHandler;
    }

    /// <summary>
    /// Parametros para Tabs.
    /// By Labs - 12/2012
    /// </summary>
    public struct tabParams
    {
        public int Height;
        public int force_top;
        public int force_left;
        public Dictionary<string, Dictionary<string, int>> Tabs;
        public List<action> actions;
    }

    public class labelParams
    {
        public String Caption;
        public int Space = 5;
        public int Width = -1;
        public labelAlign Align = labelAlign.lblTop;
    }

    /// <summary>
    /// Parametros de colunas
    /// </summary>
    public class columnParams
    {
        public List<int> Widths;
    }

    public class SetupMatrixParams
    {
        public string mtxId;
        public string tbId;
        public List<ColumnDefinition> columns;
        public bool UsingDataTable = false;
        public string DataTableSQL = "";
    }


    public class ValidateParams
    {
        public string OnEmptyError;
        public string RangeIntMin;
        public string RangeIntMax;
        public string IntMinByComp;
        public string IntMaxByComp;

        public string RangeDateMin;
        public string RangeDateMax;
        public string DateMinByComp;
        public string DateMaxByComp;
    }

    /// <summary>
    /// ALinhamento de componentes
    /// </summary>
    public enum compAlign { calLeft, calRight, calTop, calBottom };

    /// <summary>
    /// ALinhamento de Label
    /// </summary>
    public enum labelAlign { lblTop, lblLeft, lblRight };


    public enum CFLType
    {
        cflNone, cflUDO,
        cflItens, cflServicos, cflCartaoEquip, cflNumSerie,
        cflDepositos, cflClientes, cflFornecedores, cflLeads,
        cflBusinessPartners, cflUsuarios, cflFuncionarios,
        cflOportunidades, cflAtividades, cflContaContabil
    };

    public enum compSpecialType
    {
        cspNone, cspComboSimNao, cspComboDias, cspComboMeses, cspComboHoras, cspComboMinutos,
        cspBtnAdd, cspBtnDel, cspBtnClose
    };

    public class ColCompDefinitions
    {
        public string Id;
        public string Caption;

        public string BindTo;
        public bool AffectsFormMode = true;

        public bool Visible = true;
        public bool Enabled = true;
        public bool DisplayDesc = true;

        public int ForeColor;
        public int BackColor;

        public SAPbouiCOM.BoDataType UserDataType = SAPbouiCOM.BoDataType.dt_SHORT_TEXT;
        public int UserDataSize = 55;

        public SAPbouiCOM.BoFormItemTypes Type = SAPbouiCOM.BoFormItemTypes.it_EDIT;
        public compSpecialType TipoEspecial = compSpecialType.cspNone;

        public bool NonEmpty = false;
        public ValidateParams Validate;

        public string ChooseFromListUID;
        public string ChooseFromListAlias;
        public string ChooseFromListUDOName;
        public CFLType ChooseFromList = CFLType.cflNone;
        public SAPbouiCOM.Conditions ChooseFromListConds = null;

        public SAPbouiCOM.BoLinkedObject LinkedObject = SAPbouiCOM.BoLinkedObject.lf_None;
        public string LinkedObjectType;
        public string LinkedObjectForm;

        public string PopulateSQL;
        public Dictionary<string, string> PopulateItens;
        public string FirstKey;
        public string FirstValue;
        public string DefValue;
        public string ValOn = "1";
        public string ValOff = "0";
    }

    public class ColumnDefinition : ColCompDefinitions
    {
        public int Width;
        public int Percent;

        public bool Bind = true;
        public bool RightJustified = false;
        public bool Unique = false;
        public bool NonEmpty = false;

        public string SumValue = "";
        public SAPbouiCOM.BoColumnSumType SumType = SAPbouiCOM.BoColumnSumType.bst_None;
    }

    /// <summary>
    /// Parametros para criação de componentes.
    /// By Labs - 12/2012
    /// </summary>
    public class CompDefinition : ColCompDefinitions
    {
        public string Label;
        public int LblSpace = 5;
        public int LblWidth = -1;
        public labelAlign LblAlign = labelAlign.lblTop;
        public bool RightJustified = false;

        public string BindTable;

        public labelParams LabelParams;

        public SAPbouiCOM.BoPickerType Picker;
        public compAlign Align = compAlign.calLeft;

        public Int32 ModeMask = -1;
        public columnParams Columns;

        public string itemRef;
        public Bounds Bounds;

        public int marginTop;
        //public int marginBottom;

        public int Pane;
        public int FromPane;
        public int ToPane;

        public Dictionary<string, dynamic> ExtraData = new Dictionary<string, dynamic>();

        internal int _getTop()
        {
            return (this.Bounds.Top != 0
                ? this.Bounds.Top
                : (this.UserFieldTop != -1
                    ? this.UserFieldTop
                    : this.ForceTop
                )
            );
        }
        internal int _getLeft()
        {
            return (this.Bounds.Left != 0
                ? this.Bounds.Left
                : (this.UserFieldLeft != -1
                    ? this.UserFieldLeft
                    : this.ForceLeft
                )
            );
        }
        internal int _getHeight()
        {
            return (this.Bounds.Height > 0
                ? this.Bounds.Height
                : this.Height
            );
        }
        internal int _getWidth()
        {
            return (this.Bounds.Width > 0
                ? this.Bounds.Width
                : (this.UserFieldWidth > 0
                    ? this.UserFieldWidth
                    : this.ForceWidth
                )
            );
        }


        #region DEPRECATED

        public int Height;
        public int ForceWidth;
        public int ForceTop;
        public int ForceLeft;
        public int UserFieldWidth = -1;
        public int UserFieldTop = -1;
        public int UserFieldLeft = -1;

        /// <summary>
        /// Permite que se execute um método no momento da criação do componente para
        /// que se possa customiza-lo.
        /// O método deve ter a assinatura:
        /// public void NOME_DO_METODO(ref SAPbouiCOM.Item oComp){
        /// 
        /// };
        /// By Labs - 12/2012
        /// </summary>
        public string onCreateHandler;
        public string onClickHandler;
        public string onChangeHandler;
        public string onKeyDownHandler;
        public string onExitHandler;
        public List<action> actions;

        #endregion

    }

    /// <summary>
    /// Status de criação de form
    /// By Labs - 07/2013
    /// </summary>
    public enum FormStatus
    {
        frmNull, frmClosed, frmCreating, frmDtsCreated, frmDtsStarted, frmControlsCreated, frmCreated
    }

    /// <summary>
    /// Parametros para criação de formulários.
    /// By Labs - 12/2012
    /// </summary>
    public class FormParams
    {
        public string Title;
        public string Focus;
        public string BrowseByComp;
        public int LabelCount;
        public string MainDatasource;
        public List<string> ExtraDatasources;
        public List<string> SaveDatasources;
        public string BusinessObjectId;
        public SAPbouiCOM.BoFormBorderStyle BorderStyle;
        public Bounds Bounds;
        public tabParams Tabs;
        public SAPbouiCOM.Item ButtonArea;
        public SAPbouiCOM.Item TabArea;
        public SAPbouiCOM.Item Area;
        public Dictionary<string, int> Linhas;
        public Dictionary<string, int> Buttons;
        public Dictionary<string, CompDefinition> Controls;
        public List<action> FormActions;
        public bool UseXML = false;

        public FormParams()
        {
            this.ExtraDatasources = new List<string>();
            this.SaveDatasources = new List<string>();
        }

        public void Merge(FormParams fparams)
        {

            // Strings
            try
            {
                if(!String.IsNullOrEmpty(fparams.Title)) { this.Title = fparams.Title; }
                if(!String.IsNullOrEmpty(fparams.Focus)) { this.Title = fparams.Focus; }
                if(!String.IsNullOrEmpty(fparams.BrowseByComp)) { this.Title = fparams.BrowseByComp; }
                if(!String.IsNullOrEmpty(fparams.MainDatasource)) { this.Title = fparams.MainDatasource; }
                if(!String.IsNullOrEmpty(fparams.BusinessObjectId)) { this.Title = fparams.BusinessObjectId; }
                if(!String.IsNullOrEmpty(fparams.Title)) { this.Title = fparams.Title; }
            } catch(Exception e) { };

            // Objects
            //if(fparams.Bounds != null) { this.Bounds = fparams.Bounds; }

            // Lists
            try
            {
                fparams.ExtraDatasources.ForEach(d => this.ExtraDatasources.Add(d));
            } catch(Exception e) { }
            try
            {
                fparams.SaveDatasources.ForEach(d => this.SaveDatasources.Add(d));
            } catch(Exception e) { }
            try
            {
                fparams.FormActions.ForEach(d => this.FormActions.Add(d));
            } catch(Exception e) { }


            // Dictionary
            try
            {
                this.Linhas = this.Linhas.Concat(fparams.Linhas).GroupBy(d => d.Key)
                            .ToDictionary(d => d.Key, d => d.First().Value);
            } catch(Exception e) { }

            try
            {
                this.Buttons = this.Buttons.Concat(fparams.Buttons).GroupBy(d => d.Key)
                            .ToDictionary(d => d.Key, d => d.First().Value);
            } catch(Exception e) { }

            try
            {
                this.Controls = this.Controls.Concat(fparams.Controls).GroupBy(d => d.Key)
                            .ToDictionary(d => d.Key, d => d.First().Value);
            } catch(Exception e) { }

        }
    }

    /// <summary>
    /// Enumera os possíveis status de progressbar
    /// </summary>
    public enum progressBarStatus
    {
        pgb_null, pgb_created, pgb_stopped, pdb_released
    };

    public struct Color
    {
        public static int Red = (65536 * 130) + 130 * 256 + 255;
        public static int Yellow = (65536 * 130) + 250 * 256 + 255;
        public static int Green = (65536 * 130) + 255 * 256 + 160;
        public static int Gray = (65536 * 231) + 231 * 256 + 231;
    }



    #region :: Menus

    /// <summary>
    /// Enumerado com IDs dos menus padrão SAP
    /// </summary>
    public enum MenusSAP
    {
        menuNone = 0,
        menuPrincipal = 43520,
        menuAdministracao = 3328,
        menuConfiguracao = 43525,
        menuFinancas = 1536,
        menuOpVendas = 2560,
        menuVendas = 2048,
        menuCompras = 2304,
        menuPN = 43535,
        menuBanco = 43537,
        menuEstoque = 3072,
        menuRecursos = 13312,
        menuProducao = 4352,
        menuMRP = 43543,
        menuServico = 3584,
        menuRH = 43544,
        menuRelatorios = 43545
    }

    public enum MenuOpenType
    {
        mnOpNormal, mnOpUDOAdd //, mnOpUDOFind
    }


    /// <summary>
    /// Estrutura para criação de ítens de menu.
    /// By Labs - 10/2012
    /// </summary>
    public class menuStruct
    {
        public string parentUID;
        public string UID;
        public string Label;
        public SAPbouiCOM.BoMenuType Type = SAPbouiCOM.BoMenuType.mt_STRING;
        public MenusSAP SAPParentId = MenusSAP.menuNone;

        public string OpenForm;
        public MenuOpenType OpenFormType = MenuOpenType.mnOpNormal;

        public int Position = 99;
        public string Image = "";
        public string ImgPath = "";
        public bool Checked = false;
        public bool Enabled = true;
    }

    #endregion


}
