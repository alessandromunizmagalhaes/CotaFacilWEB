using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Reflection;
using Google.Maps.Geocoding;
using System.Net;

/// <summary>
/// Classe ancestral para implementação e criação de add-on 
/// para o SAP, definindo o padrão de boas práticas e
/// regras de desenvolvimento.
/// By Labs - 10/2012
/// </summary>
namespace TShark
{

    public static class addonConfig
    {

    }

    /// <summary>
    /// Dados de informação do add_on.
    /// By Labs - 12/2012
    /// </summary>
    public struct addon_info
    {
        public int Versao;
        public int Release;
        public int Revisao;
        public string AppName;
        public string Namespace;
        public string Descricao;
        public string Autor;
        public string ExeName;
        public string VersaoStr;

        public void setInfo(int versao, int release, int revisao, string app_name, string Namespace, string desc, string autor)
        {
            this.Versao = versao;
            this.Release = release;
            this.Revisao = revisao;
            this.Autor = autor;
            this.AppName = app_name;
            this.Namespace = Namespace;
            this.Descricao = desc;
            this.ExeName = (System.AppDomain.CurrentDomain.SetupInformation.ApplicationName.Split('.'))[0]; // Path.GetFileName(Application.ExecutablePath);
            this.VersaoStr = Versao + "." + Release + "." + Revisao;
        }
    }


    public class FastOneItemEvent : SAPbouiCOM.ItemEvent
    {
        public string FormUID { get; set; }
        public string ItemUID { get; set; }
        public bool BeforeAction { get; set; }
        public int CharPressed { get; set; }
        public string ColUID { get; set; }
        public int FormMode { get; set; }
        public int FormTypeCount { get; set; }
        public string FormTypeEx { get; set; }
        public bool InnerEvent { get; set; }
        public bool ItemChanged { get; set; }
        public int PopUpIndicator { get; set; }
        public int Row { get; set; }
        public int FormType { get; set; }
        public bool Before_Action { get; set; }
        public bool ActionSuccess { get; set; }
        public bool Action_Success { get; set; }
        public SAPbouiCOM.BoEventTypes EventType { get; set; }
        public SAPbouiCOM.BoModifiersEnum Modifiers { get; set; }

        public UserFields userFieldsHandler;
    }

    /// <summary>
    /// Classe para armazenamento de configurações diversas
    /// do addon em arquivo xml.
    /// </summary>
    public class AddonXML
    {
        private XmlDocument _doc = null;
        private string arq = "addon.xml";

        public XmlNode Info         { get { return this.setNode("/addon/info"); } }
        public XmlNode Config       { get { return this.setNode("/addon/config"); } }
        public XmlNode UserTables   { get { return this.setNode("/addon/usertables"); } }
        public XmlNode UserFields   { get { return this.setNode("/addon/userfields"); } }
        public XmlNode Validacao    { get { return this.setNode("/addon/validacao"); } }
        public XmlNode Data         { get { return this.setNode("/addon/data"); } }


        private XmlNode setNode(string node){
            XmlNode Node = this._doc.SelectSingleNode(node);

            if(Node == null)
            {
                this._doc.LoadXml(Properties.Resources.addon_xml);
                this.Save();
                Node = this._doc.SelectSingleNode(node);
            }

            return Node;
        }


        public AddonXML()
        {
            // Parametrizacao externa:
            this._doc = new XmlDocument(); 
            if(System.IO.File.Exists(this.arq))
            {
                this._doc.Load(this.arq);

            // Cria um novo arquivo com base no que está em resources
            } else
            {
                this._doc.LoadXml(Properties.Resources.addon_xml);
                this.Save();
            }
        } 

        public void Save(){
            this._doc.Save(this.arq);
        }
    }

    /// <summary>
    /// Encapsula uma janela windows para host de processos
    /// </summary>
    public class WindowWrapper : System.Windows.Forms.IWin32Window
    {
        private IntPtr _hwnd;

        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }
        
        public IntPtr Handle
        {
            get { return _hwnd; }
        }
    }

    /// <summary>
    /// Implementa configuração para acesso Soap a servidores remotos
    /// </summary>
    public class SoapConfig
    {
        public string user;
        public string pwd;
        public string host;
        public string port;
        public bool useCredentials = false;
    }

    /// <summary>
    /// Classe ancestral para implementação de add-ons para o 
    /// SAP Business One - Padrão.
    /// By Labs - 10/2012
    /// </summary>
    public class FastOne
    {
        
        /// <summary>
        /// Objetos 'Application' e 'Company'
        /// By Labs - 10/2012
        /// </summary>
        public SAPbouiCOM.Application SBO_Application;
        public SAPbobsCOM.Company oCompany;


        /// <summary>
        /// Filtros de eventos
        /// By Labs - 10/2012
        /// </summary>
        public SAPbouiCOM.EventFilters evFilters;
        public SAPbouiCOM.EventFilter evFilter;
        private Dictionary<SAPbouiCOM.BoEventTypes, int> evIndexer;

        /// <summary>
        /// Armazena o form com base no ultimo evento ocorrido 
        /// </summary>
        internal SAPbouiCOM.Form _lastFormByEvent = null;

        /// <summary>
        /// Informações do add_on
        /// By Labs - 12/2012
        /// </summary>
        public addon_info AddonInfo;

        /// <summary>
        /// Dicionário para registro de eventos.
        /// Evento::Before_or_After::OnForm::OnItem::ExecHandler
        /// By Labs - 12/2012
        /// </summary>
        internal Dictionary<SAPbouiCOM.BoEventTypes, Dictionary<string, Dictionary<string, Dictionary<string, List<string>>>>> _eventos_registrados;

        /// <summary>
        /// Armazena o path de execução para cargas de XML 
        /// By Labs - 12/2012
        /// </summary>
        public string ExecPath = "";

        /// <summary>
        /// Armazena o path da pasta de addons do var 
        /// By Labs - 12/2012
        /// </summary>
        public string VarPath = "";

        /// <summary>
        /// Indica se a configuração da empresa é de multi filiais.
        /// </summary>
        public bool UsaFiliais = false;

        /// <summary>
        /// Flag para uso em tempo de desenvolvimento. Se TRUE,
        /// ativa a exibição de mensagens de execução na barra do B1,
        /// como clicks em menus, etc...
        /// By Labs - 12/2012
        /// </summary>
        public bool VerboseMode = false;

        /// <summary>
        /// Flag para uso em tempo de desenvolvimento. Se TRUE,
        /// ativa a exibição de mensagens de erro de execução.
        /// By Labs - 08/2014
        /// </summary>
        public bool showDesenvTimeMsgs = false;

        /// <summary>
        /// DEBUG: exibe nome dos eventos que o SAP for gerando.
        /// By Labs - 01/2013
        /// </summary>
        public bool debugShowEvents = false;

        /// <summary>
        /// DEBUG: Força a recriação de todas as tabelas em UserTables. 
        /// </summary>
        public bool ForceTableReset = false;

        /// <summary>
        /// Manipulação de Menus.
        /// By Labs - 01/2013
        /// </summary>
        public Menus Menus;

        /// <summary>
        /// Armazena estruturas de criação de forms e os forms criados.
        /// By Labs - 12/2012
        /// </summary>
        public Dictionary<string, Forms> FormList;

        /// <summary>
        /// Armazena valores de chaves (ids) utilizadas pelas classes e forms 
        /// dentro do addon.
        /// By Labs - 11/2013
        /// </summary>
        public Dictionary<string, dynamic> AddonKeys = new Dictionary<string, dynamic>();

        /// <summary>
        /// Status de progressBar
        /// </summary>
        public progressBarStatus pgBarStatus = progressBarStatus.pgb_null;

        /// <summary>
        /// Informa se alguem clicou em "Parar" no pgBar
        /// </summary>
        public bool pgBarStopped = false;

        /// <summary>
        /// Manipulação de Datasources.
        /// By Labs - 01/2013
        /// </summary>
        public Datasources DtSources;

        /// <summary>
        /// Campos de usuário em forms padrão.
        /// By Labs - 08/2013
        /// </summary>
        public Dictionary<string, UserFields> UserFields;
        internal Dictionary<string, string> UserFieldClass;

        /// <summary>
        /// Recordset de uso genérico, alimentado pelo método Select.
        /// </summary>
        public SAPbobsCOM.Recordset Browser = null;

        /// <summary>
        /// Conexão Addon - server TShark Web Application.
        /// By Labs - 05/2013
        /// </summary>
        public WebDriver WebDriver;

        /// <summary>
        /// Acesso ao addon.xml
        /// </summary>
        public AddonXML Xml;

        /// <summary>
        /// Faz com que os forms sejam salvos em XML
        /// </summary>
        public bool UseFormXML;

        /// <summary>
        /// Implementa metodos de boBridge
        /// </summary>
        public SAPbobsCOM.SBObob SBOBob;

        /// <summary>
        /// Configurações para acesso Soap
        /// </summary>
        public SoapConfig SoapConfig;

        /// <summary>
        /// Inicialização da Classe.
        /// By Labs - 12/2012
        /// </summary>
        public FastOne(string appPath, bool show_msg = false)
        {

            // Define o path de execução:
            this.ExecPath = System.IO.Directory.GetParent(appPath).ToString();

            // Userfields
            this.UserFields = new Dictionary<string, UserFields>();
            this.UserFieldClass = new Dictionary<string, string>();

            this.SoapConfig = new SoapConfig();

            // Parametrizacao externa:
            this.Xml = new AddonXML();
             
            // Exibe ou não mensagens de desenvolvimento em status bar:
            this.VerboseMode = false;
            try
            {
                this.VerboseMode = (this.Xml.Config.SelectSingleNode("/addon/config/mensagens/verboseMode").Attributes["value"].Value == "true");
            } catch { }

            // Exibe ou não mensagens de desenvolvimento em status bar:
            this.showDesenvTimeMsgs = true;
            try
            {
                this.showDesenvTimeMsgs = (this.Xml.Config.SelectSingleNode("/addon/config/mensagens/showDesenvTimeMsgs").Attributes["value"].Value == "true");
            } catch { }
        }

        /// <summary>
        /// Exibe mensagens de erro internas de forma mais amigável ao desenvolvedor, 
        /// de preferência, exibindo um hint do problema.
        /// </summary>
        /// <author>Labs - 10/2013</author>
        /// <param name="e"></param>
        /// <param name="hint"></param>
        public bool DesenvTimeError(Exception e, string hint = "")
        {
            if(!this.showDesenvTimeMsgs) return false;

            string msg = "Erro de Desenvolvimento: ";
            msg += "\n  - " + e.Message;
            if(!String.IsNullOrEmpty(hint))
            {
                msg += "\n\nAjuda TShark FastOne:\n " + hint + "\n";
            }
            msg += "\n Stack: \n" + e.StackTrace;
            try
            {
                this.SBO_Application.MessageBox(msg);
            } catch (Exception e2) { }

            return false;
        }
        public void DesenvTimeInfo(string info)
        {
            if(!this.showDesenvTimeMsgs) return;
            this.StatusInfo(info);
        }
        public void DesenvTimeAlert(string alerta)
        {
            if(!this.showDesenvTimeMsgs) return;
            this.StatusAlerta(alerta);
        }


        #region :: Inicialização

        /// <summary>
        /// Seta as informações do addon.
        /// By Labs - 10/2012
        /// </summary>
        /// <param name="versao"></param>
        /// <param name="release"></param>
        /// <param name="revisao"></param>
        /// <param name="app_name"></param>
        /// <param name="Namespace"></param>
        /// <param name="desc"></param>
        /// <param name="autor"></param>
        public void setInfo(int versao, int release, int revisao, string app_name, string Namespace, string desc, string autor)
        {
            this.AddonInfo.setInfo(versao, release, revisao, app_name, Namespace, desc, autor);
        }

        /// <summary>
        /// Inicializa o add_on.
        /// By Labs - 10/2012
        /// </summary>
        public void AddOnInitialize()
        {

            // Define e inicializa SBO_Application:
            SetApplication();

            // Define o contexto de conexão:
            if(!(SetConnectionContext() == 0))
            {
                this.SBO_Application.MessageBox("Falha ao se estabelecer a conexão com a DI API", 1, "Ok", "", "");
                System.Environment.Exit(0);
            }

            // Estabelece conexão com o banco de dados:
            this.StatusInfo("Conectando-se ao banco de dados... Aguarde...");
            if(!(ConnectToCompany() == 0))
            {
                this.SBO_Application.MessageBox("Falha ao se estabelecer conexão com o banco de dados", 1, "Ok", "", "");
                System.Environment.Exit(0);
            }
            this.StatusInfo("Conectado!");


            // Manipuladores de Eventos
            this.evFilters = new SAPbouiCOM.EventFilters();
            this.evIndexer = new Dictionary<SAPbouiCOM.BoEventTypes, int>();
            this.evFilter = this.evFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
            this.evIndexer[SAPbouiCOM.BoEventTypes.et_MENU_CLICK] = this.evFilters.Count - 1;// .evIndex++;
            this.SBO_Application.SetFilter(this.evFilters);

            // Armazena os eventos registrados pelo add-on:
            this._eventos_registrados = new Dictionary<SAPbouiCOM.BoEventTypes, Dictionary<string, Dictionary<string, Dictionary<string, List<string>>>>>();

            // Factory de menus:
            this.Menus = new Menus(this);

            // Objeto de datasource:
            this.DtSources = new Datasources(this);

            // Lista de forms
            this.FormList = new Dictionary<string, Forms>();

            // Campos definidos por usuário:
            //this.userFieldsParams = new Dictionary<string, List<userFieldsParams>>();

            // WebDriver:
            this.WebDriver = new WebDriver(this);

            // Configura o webdriver
            if (this.WebDriver.SetupByXML())
            {
                this.StatusInfo("Acesso ao cloud configurado!");
            }

            //this.ts_data_mapping = new Dictionary<string, dynamic>();

            // Registra manipulador para eventos da aplicação:
            this.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(AppEvents);

            // Registra controlador para os eventos de menu:
            this.SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(MenuEvents);

            // Registra manipulador para eventos de itens:
            this.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(AppItemEvents);

            // Registra manipulador para eventos de dados de form:
            this.SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(FormDataEvents);

            // Registra manipulador para eventos para ProgressBar
            this.SBO_Application.ProgressBarEvent += new SAPbouiCOM._IApplicationEvents_ProgressBarEventEventHandler(ProgressBarEvents);

            // Registra manipulador para eventos na barra de status
            this.SBO_Application.StatusBarEvent += new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler(StatusBarEvents);

            // Recordset genérico, alimentado por Select
            this.Browser = this.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            this.SBOBob = this.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            // Manda um "hello world":
            if(this.VerboseMode)
            {
                this.StatusInfo("DI Connectada: " + this.AddonInfo.Descricao);
            }

            // Informa que estamos na área:
            this.StatusInfo("Addon " + this.AddonInfo.Descricao + " conectando. Aguarde...");
        }

        /// <summary>
        /// Exibe info quando finalizar a inicialização
        /// </summary>
        /// <param name="msg">Mensagem opcional a ser exibida</param>
        public void AddOnInitialized(string msg = "")
        {
            // Verifica multifilial:
            try
            {
                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)this.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rec.DoQuery("SELECT MltpBrnchs FROM OADM (nolock) ");
                this.UsaFiliais = (rec.Fields.Item("MltpBrnchs").Value == "Y");
            } catch { }
            
            // limpa flag fullreset
            this.Xml.UserTables.SelectSingleNode("fullReset").Attributes[0].Value = "false";
            this.Xml.Save();
            
            this.StatusInfo(this.AddonInfo.Descricao + (String.IsNullOrEmpty(msg) ? ": Conectado!" : msg));
        }

        /// <sumary>
        /// Estabelece a conexão com o SAP e inicializa o objeto 'Application'.
        /// By Labs - 10/2012
        /// </sumary>
        private void SetApplication()
        {

            // Recupera string de conexão:
            string sConnectionString = "";
            try
            {
                sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            } catch(Exception e)
            {
                throw new Exception("Atenção! String de conexão não configurada");
            }

            // Conecta:
            SAPbouiCOM.SboGuiApi SboGuiApi = new SAPbouiCOM.SboGuiApi();
            SboGuiApi.Connect(sConnectionString);

            // Inicializa:
            this.SBO_Application = SboGuiApi.GetApplication(-1);

            // Manda um "hello world":
            if(this.VerboseMode)
            {
                this.StatusInfo("Conexão SBO_Application estabelecida: " + this.AddonInfo.Descricao);
            }

        }

        /// <sumary>
        /// Estabelece a conexão com o SAP e inicializa o objeto 'Company'.
        /// By Labs - 10/2012
        /// </sumary>
        private int SetConnectionContext()
        {

            // Inicializa 'Company':
            this.oCompany = new SAPbobsCOM.Company();
            
            // Recupera a string de conexão do contexto:
            string sCookie = this.oCompany.GetContextCookie();
            string sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie);
            if(this.oCompany.Connected == true)
            {
                this.oCompany.Disconnect();
            }

            // Conecta:
            int res = this.oCompany.SetSboLoginContext(sConnectionContext);

            // Manda um "hello world":
            if(this.VerboseMode)
            {
                this.StatusInfo("Contexto de conexão estabelecido: " + this.AddonInfo.Descricao);
            }

            // Retorna:
            return res;
        }

        /// <summary>
        /// Conecta com a companhia.
        /// By Labs - 10/2012
        /// </summary>
        private int ConnectToCompany()
        {
            return this.oCompany.Connect();
        }

        #endregion


        #region :: Status

        private void _showStatus(string msg, SAPbouiCOM.BoStatusBarMessageType tipo, bool popMessage = false)
        {
            try
            {
                this.SBO_Application.StatusBar.SetText(msg, SAPbouiCOM.BoMessageTime.bmt_Medium, tipo);
                if(popMessage)
                {
                    this.SBO_Application.MessageBox(msg);
                }
            } catch(Exception e)
            {

            }
        }

        public bool StatusInfo(string msg, bool popMessage = false)
        {
            try
            {
                this._showStatus(msg, SAPbouiCOM.BoStatusBarMessageType.smt_Success, popMessage);

            } catch(Exception e)
            {

            }
            return true;
        }

        public bool StatusAlerta(string msg, bool popMessage = false)
        {
            try
            {
                this._showStatus(msg, SAPbouiCOM.BoStatusBarMessageType.smt_Warning, popMessage);
            } catch(Exception e)
            {

            }
            return true;
        }

        public bool StatusErro(string msg, bool popMessage = false)
        {
            try
            {
                this._showStatus(msg, SAPbouiCOM.BoStatusBarMessageType.smt_Error, popMessage);
            } catch(Exception e)
            {

            }
            return false;
        }

        public void ClearStatus()
        {
            try
            {
                this.SBO_Application.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_None);
            } catch(Exception e)
            {

            }
        }

        public void ShowMessage(string msg)
        {
            try { 
                this.SBO_Application.MessageBox(msg);
            } catch(Exception e)
            {

            }
        }

        #endregion


        #region :: Eventos

        /// <summary>
        /// Registra eventos.
        /// By Labs - 12/2012
        /// </summary>
        /// <remarks>
        /// Ao registrar o evento, deve-se informar o tipo do evento, e o componente em qual form que 
        /// irá ouví-lo, bem como o NOME (string) do método que foi implementado no add-on para
        /// tratar do evento.
        /// </remarks>
        /// <param name="EventType">Tipo do evento que será tratado</param>
        /// <param name="FormTypeEx">Id do form onde o evento tem interesse para o add-on</param>
        /// <param name="ItemID">Id do ítem no form que escutará o evento</param>
        /// <param name="EventHandler">Nome do método implementado no add-on que será chamado para tratar o evento</param>
        /// <param name="TriggerAfterSAPEvent">Momento do evento em que o método será chamado: before ou after</param>
        public void RegisterEvent(SAPbouiCOM.BoEventTypes EventType, string FormTypeEx, string ItemID, string EventHandler, bool TriggerAfterSAPEvent = true)
        {
            this._registerEvent(EventType, FormTypeEx, ItemID, EventHandler, TriggerAfterSAPEvent);
        }

        internal void RegisterEventHandler(SAPbouiCOM.BoEventTypes EventType, string FormTypeEx, string ItemID, string EventHandler, bool TriggerAfterSAPEvent = true)
        {
            this._registerEvent(EventType, FormTypeEx, ItemID, EventHandler, TriggerAfterSAPEvent, true);
        }
        
        internal void _registerEvent(SAPbouiCOM.BoEventTypes EventType, string FormTypeEx, string ItemID, string EventHandler, bool TriggerAfterSAPEvent = true, bool force = false)
        {

            // Registra o evento desejado no SAP:
            try
            {
                this.evFilter = this.evFilters.Add(EventType);
                this.evIndexer[EventType] = this.evFilters.Count - 1;// this.evIndex++;

            } catch(Exception e)
            {
                if(this.VerboseMode)
                {
                    this.StatusErro("Erro ao registrar evento no SAP: " + FormTypeEx + "::" + ItemID + "::" + EventHandler + "' - " + e.Message);
                }

                try
                {
                    this.evFilter = this.evFilters.Item(this.evIndexer[EventType]);
                } catch(Exception e2)
                {
                    this.StatusErro("Erro ao registrar evento no SAP: " + FormTypeEx + "::" + ItemID + "::" + EventHandler + "' - " + e2.Message);
                }
            }

            try
            {
                this.evFilter.AddEx(FormTypeEx);

            } catch(Exception e)
            {
                if(this.VerboseMode)
                {
                    this.StatusErro("Erro ao registrar evento no SAP: " + FormTypeEx + "::" + ItemID + "::" + EventHandler + "' - " + e.Message);
                }
            }


            // Registra o evento no framework:
            try
            {
                if(!this._eventos_registrados.ContainsKey(EventType))
                {
                    this._eventos_registrados.Add(EventType, new Dictionary<string, Dictionary<string, Dictionary<string, List<string>>>>());
                }

                // Momento do evento em que estamos interessados:
                string momento = (TriggerAfterSAPEvent ? "after" : "before");
                if(!this._eventos_registrados[EventType].ContainsKey(momento))
                {
                    this._eventos_registrados[EventType].Add(momento, new Dictionary<string, Dictionary<string, List<string>>>());
                }

                // Registra o form para o evento:
                if(!this._eventos_registrados[EventType][momento].ContainsKey(FormTypeEx))
                {
                    this._eventos_registrados[EventType][momento].Add(FormTypeEx, new Dictionary<string, List<string>>());
                }

                // Registra o item que escutará o evento:
                if(!this._eventos_registrados[EventType][momento][FormTypeEx].ContainsKey(ItemID))
                {
                    this._eventos_registrados[EventType][momento][FormTypeEx].Add(ItemID, new List<string>());
                }

                // Evita que o mesmo handler seja registrado mais de uma vez:
                if(!this._eventos_registrados[EventType][momento][FormTypeEx][ItemID].Contains(EventHandler))
                {
                    this._eventos_registrados[EventType][momento][FormTypeEx][ItemID].Add(EventHandler);
                }

                this.SBO_Application.SetFilter(this.evFilters);

            } catch(Exception e)
            {
                if(this.VerboseMode)
                {
                    this.DesenvTimeError(e, " - Erro ao registrar evento no Framework: " + FormTypeEx + "::" + ItemID + "::" + EventHandler);
                }
            }
            GC.Collect();
        }

        /// <summary>
        /// Handler global para controle de eventos da aplicação.
        /// By Labs - 12/2012
        /// </summary>
        /// <remarks>
        /// A implementação dos handlers globais de eventos permite que o código do add-on possa organizar 
        /// os handlers de cada evento em métodos com uma assinatura padrão, de forma que a classe TShark.BusinessOne
        /// se encarrege de sua chamada.
        /// No caso dos "Eventos de Aplicação", os seguintes handlers podem ser criados no código do add-on:
        /// <code>
        /// public void onShutDown() {
        ///    // este método será chamado automaticamente quando a aplicação estiver se encerrando. 
        /// } 
        /// 
        /// public void onCompanyChange() {
        ///    // este método será chamado automaticamente quando houver mudança na companhia,
        ///    // exceto quando o add-on implementar restrição de uso e a nova companhia selecionada
        ///    // não estiver registrada como autorizada.
        /// } 
        /// 
        /// public void onLanguageChange() {
        ///    // este método será chamado automaticamente quando a aplicação mudar a linguagem. 
        /// } 
        /// </code>
        /// </remarks>
        private void AppEvents(SAPbouiCOM.BoAppEventTypes EventType)
        {
            GC.Collect();
            bool BubbleEvent;
            switch(EventType)
            {

                // onShutDown - A aplicação está encerrando:
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    try
                    {
                        this.ExecEvent("onShutDown", out BubbleEvent, new object[] { });
                    } catch(Exception e)
                    {

                    } finally
                    {
                        //System.Environment.Exit(0);
                        System.Windows.Forms.Application.Exit();
                    }
                    break;

                // onServerTerminition: 
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    try
                    {
                        this.ExecEvent("onServerTerminition", out BubbleEvent, new object[] { });
                    } catch(Exception e)
                    {

                    } finally
                    {
                        //System.Environment.Exit(0);
                        System.Windows.Forms.Application.Exit();
                    }
                    break;

                // onCompanyChange - Mudança de companhia: 
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    try
                    {
                        this.ExecEvent("onCompanyChange", out BubbleEvent, new object[] { });
                    } catch(Exception e)
                    {

                    } finally
                    {
                        //System.Environment.Exit(0);
                        System.Windows.Forms.Application.Exit();
                    }
                    break;

                // onLanguageChange - Mudança de linguagem:
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    try
                    {
                        this.ExecEvent("onLanguageChange", out BubbleEvent, new object[] { });
                    } catch(Exception e)
                    {

                    } finally
                    {
                        //System.Environment.Exit(0);
                        System.Windows.Forms.Application.Exit();
                    }
                    break;
            }
        }

        /// <summary>
        /// Handler global para controle de eventos de menu.
        /// By Labs - 12/2012
        /// </summary>
        /// <remarks>
        /// A implementação do handler global de eventos de menu permite que o código do add-on possa organizar os handlers
        /// de cada menu em métodos com uma assinatura padrão, de forma que a classe se encarrege de sua chamada.
        /// <example>
        ///  Supondo que se criou um ítem de menu com o MenuID 'mnMeuTeste', tem-se a possibilidade de se implementar 
        ///  os métodos de evento na classe do add-on com a seguinte assinatura:
        ///  <code>
        ///  // Por padrão, [IDMENU]OnClick é executado no onAfter (evObj.BeforeAction == false)
        /// public bool mnMeuTesteOnClick(ref SAPbouiCOM.MenuEvent evObj, out bool BubbleEvent) {
        ///    
        ///    SBO_Application.MessageBox("O evento onClick do menu '" + evObj.MenuID + "' foi executado", 1, "Ok", "", "");
        ///    
        ///    // Retorna true se correu tudo bem:
        ///    return true;
        /// } 
        /// </code>
        /// caso se deseje tratar o evento antes do B1 ou seja, quando (evObj.BeforeAction == true), basta se criar o metodo 
        /// com a seguinte assinatura: [IDMENU]OnBeforeClick
        /// <code>
        /// public bool mnMeuTesteOnBeforeClick(ref SAPbouiCOM.MenuEvent evObj, out bool BubbleEvent) {
        ///    SBO_Application.MessageBox("O evento onBeforeClick do menu '" + evObj.MenuID + "' foi executado", 1, "Ok", "", "");
        ///    
        ///    // Retorna true se correu tudo bem:
        ///    return true;
        /// } 
        /// </code>
        /// </example>
        /// </remarks>
        internal void MenuEvents(ref SAPbouiCOM.MenuEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if(this.debugShowEvents)
            {
                this.StatusAlerta("Evento [" + (evObj.BeforeAction == true ? "BEFORE" : "AFTER") + "] et_MENU_CLICK: " + evObj.MenuUID);
            }

            // Só processamos o evento se for um menu nosso:
            if(this.Menus._menu_ids.IndexOf(evObj.MenuUID) > -1)
            {

                // Form automatico
                if(this.Menus.MenuForms.ContainsKey(evObj.MenuUID) && evObj.BeforeAction)
                {
                    KeyValuePair<string, MenuOpenType> exec = this.Menus.MenuForms[evObj.MenuUID].First();
                    switch(exec.Value)
                    {
                        case MenuOpenType.mnOpNormal:
                            this.OpenForm(exec.Key);
                            break;

                        case MenuOpenType.mnOpUDOAdd:
                            this.OpenFormUDOAdd(exec.Key);
                            break;
                    }

                // Evento cadastrado
                } else
                {
                    // Tipo do evento:
                    string tipo_evento = evObj.MenuUID + "On" + (evObj.BeforeAction == true ? "Before" : "") + "Click";

                    // Executa:
                    bool handler_existe = this.ExecEvent(tipo_evento, out BubbleEvent, new object[] { evObj, BubbleEvent });

                    // Se houver handler, cancela bubble:
                    BubbleEvent = !handler_existe;
                }
            }

            // Intercepta onFormAdd
            if(evObj.MenuUID == "1282" && evObj.BeforeAction == false)
            {
                string FormTypeEx = this.SBO_Application.Forms.ActiveForm.TypeEx;
                if(this.FormList.ContainsKey(FormTypeEx))
                {
                  //  if(this._processFormEvents(SAPbouiCOM.BoEventTypes.et_MENU_CLICK, "after", FormTypeEx))
                  //  {
                        this.FormList[FormTypeEx].OnInsertUDO(ref evObj, out BubbleEvent);
                  //  }
                }
            }

            GC.Collect();
        }


        /// <summary>
        /// Registra eventos de progressbar e mantem o status da mesma. Seta pgBarStopped 
        /// para false quando cria um novo pgbar e para true se alguem clicar no botão de 
        /// parar. 
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        private void ProgressBarEvents(ref SAPbouiCOM.ProgressBarEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if(evObj.BeforeAction)
            {

                // Avisa que foi liberado, no before action
                if(evObj.EventType == SAPbouiCOM.BoProgressBarEventTypes.pbet_ProgressBarReleased)
                {
                    this.pgBarStatus = progressBarStatus.pdb_released;
                }

            } else
            {
                switch(evObj.EventType)
                {

                    // Registra a criação do pgBar:
                    case SAPbouiCOM.BoProgressBarEventTypes.pbet_ProgressBarCreated:
                        this.pgBarStatus = progressBarStatus.pgb_created;
                        this.pgBarStopped = false;
                        break;

                    // Registra o clique no stop:
                    case SAPbouiCOM.BoProgressBarEventTypes.pbet_ProgressBarStopped:
                        this.pgBarStatus = progressBarStatus.pgb_stopped;
                        this.pgBarStopped = true;
                        break;

                    // Registra a remoção do pgBar:
                    case SAPbouiCOM.BoProgressBarEventTypes.pbet_ProgressBarReleased:
                        this.pgBarStatus = progressBarStatus.pgb_null;
                        break;
                }
            }
        }


        /// <summary>
        /// Handler global para interceptar mensagens padrão SAP.
        /// By Labs - 10/2015
        /// </summary>
        /// <remarks>
        /// </remarks>
        private void StatusBarEvents(string Message, SAPbouiCOM.BoStatusBarMessageType MessageType)
        {
            GC.Collect();

            if(MessageType == SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            {
                // Substitui msg de erro padrão SAP quando uma validação retorna BubbleEvent false
                if(Message.Contains("(UI_API -7780)"))
                {
                    this.StatusErro("Verifique e corrija os valores incorretos.");
                }
            }
        }


        /// <summary>
        /// Handler global para controle de eventos de ítens de formulários.
        /// By Labs - 12/2012
        /// </summary>
        /// <remarks>
        /// Será executado o método registrado para tratar exatamente o EVENTO ocorrido NO ÍTEM do FORM ESPECÍFICO,
        /// tal como registrado anteriormente via this.Addon.registerEvent
        /// <example>
        ///  Desejamos exibir uma mensagem quando um determinado botão do form for clicado. 
        ///  Primeiro registramos o evento:
        ///  <code>
        ///  this.Addon.registerEvent(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, "meuForm", "btnMsg", "onBtnMsgClick");
        ///  </code>
        ///  Então criamos o método:
        ///  <code>
        ///  public bool onBtnMsgClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent) {
        ///    SBO_Application.MessageBox("O botão '" + evObj.ItemUID + "' foi clicado", 1, "Ok", "", "");
        ///    
        ///    // Retorna ok, se tudo correu bem:
        ///    return true;
        ///  } 
        ///  </code>
        /// </example>
        /// </remarks>
        private void AppItemEvents(string FormUID, ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if(evObj.FormTypeEx == "-9876")
            {
                BubbleEvent = true;
            }
            
            // Registra o Form do evento
            this._lastFormByEvent = null;
            try
            {
                this._lastFormByEvent = this.SBO_Application.Forms.GetForm(evObj.FormTypeEx, evObj.FormTypeCount);
            } catch(Exception e) { }


            if(this.debugShowEvents)
            {
                this.StatusAlerta("Evento [" + (evObj.BeforeAction == true ? "BEFORE" : "AFTER") + "] " + evObj.EventType + ": " + evObj.ItemUID);
            
            } else if(this.VerboseMode)
            {
                this.StatusAlerta("Evento '" + evObj.EventType + "': " + System.DateTime.Now.ToString());
            }

            // Verifica se estamos interessados no evento em questão:
            if(this._eventos_registrados.ContainsKey(evObj.EventType))
            {

                // Verifica se o momento do evento nos interessa:
                string momento = (evObj.Before_Action ? "before" : "after");
                if(this._eventos_registrados[evObj.EventType].ContainsKey(momento))
                {

                    // Verifica se o form foi registrado para o evento:
                    if(this._eventos_registrados[evObj.EventType][momento].ContainsKey(evObj.FormTypeEx))
                    {

                        // Verifica se o evento ocorreu em um ítem registrado:
                        string id = (evObj.ItemUID != "" ? evObj.ItemUID : evObj.FormTypeEx);
                        if(this._eventos_registrados[evObj.EventType][momento][evObj.FormTypeEx].ContainsKey(id))
                        {

                            // Bão, aí num mais tem jeito, temos que trabalhar:
                            bool handler_existe = false;
                            foreach(string handler in this._eventos_registrados[evObj.EventType][momento][evObj.FormTypeEx][id])
                            {
                                handler_existe = this.ExecEvent(handler, out BubbleEvent, new object[] { evObj, BubbleEvent });
                                if(!BubbleEvent) break;
                            }

                        }
                    }
                }
            }
            GC.Collect();
        }

        /// <summary>
        /// Handler global para controle de eventos de dados de formulários.
        /// By Labs - 03/2013
        /// </summary>
        /// <remarks>
        /// Será executado o método registrado para tratar exatamente o EVENTO ocorrido FORM,
        /// tal como registrado anteriormente via this.registerEvent
        /// </remarks>
        private void FormDataEvents(ref SAPbouiCOM.BusinessObjectInfo evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if(this.debugShowEvents)
            {
                this.StatusAlerta("Evento [" + (evObj.BeforeAction == true ? "BEFORE" : "AFTER") + "] " + evObj.EventType + ": " + evObj.FormTypeEx);

            }

            // Registra o Form do evento
            this._lastFormByEvent = null;
            try
            {
                this._lastFormByEvent = this.SBO_Application.Forms.Item(evObj.FormUID);
            } catch(Exception e) { }

            string momento = (evObj.BeforeAction ? "before" : "after");
            if(this._processFormEvents(evObj.EventType, momento, evObj.FormTypeEx))
            {
                foreach(string handler in this._eventos_registrados[evObj.EventType][momento][evObj.FormTypeEx][evObj.FormTypeEx])
                {
                    if(BubbleEvent)
                    {
                        this.ExecEvent(handler, out BubbleEvent, new object[] { evObj, BubbleEvent });
                        if(!BubbleEvent) break;
                    }
                }
            }
        }


        internal Boolean _processFormEvents(SAPbouiCOM.BoEventTypes EventType, string momento, string FormTypeEx)
        {
            Boolean res = false;

            // Verifica se estamos interessados no evento em questão:
            if(this._eventos_registrados.ContainsKey(EventType))
            {

                // Verifica se o momento do evento nos interessa:
                if(this._eventos_registrados[EventType].ContainsKey(momento))
                {

                    // Verifica se o form foi registrado para o evento:
                    if(this._eventos_registrados[EventType][momento].ContainsKey(FormTypeEx))
                    {

                        // Verifica se o evento ocorreu em um ítem registrado:
                        if(this._eventos_registrados[EventType][momento][FormTypeEx].ContainsKey(FormTypeEx))
                        {

                            // Bão, aí num mais tem jeito, temos que trabalhar:
                            res = true;
                        }
                    }
                }
            }
            GC.Collect();

            return res;
        }

        /// <summary>
        /// Limpa os eventos de um form
        /// </summary>
        /// <param name="formId"></param>
        public void zClearFormEvents(string formId)
        {
            try
            {
                foreach(KeyValuePair<SAPbouiCOM.BoEventTypes, Dictionary<string, Dictionary<string, Dictionary<string, List<string>>>>> evType in this._eventos_registrados)
                {
                    foreach(KeyValuePair<string, Dictionary<string, Dictionary<string, List<string>>>> momento in evType.Value)
                    {
                        if(this._eventos_registrados[evType.Key][momento.Key].ContainsKey(formId))
                        {
                            this._eventos_registrados[evType.Key][momento.Key][formId] = new Dictionary<string, List<string>>();
                        }
                    }
                }
            } catch { }
        }

        /// <summary>
        /// Método privado responsável pela chamada de execução dos handlers de eventos 
        /// implementados nas classes dos add-ons.
        /// By Labs - 11/2012
        /// </summary>
        internal bool ExecEvent(string evName, out bool BubbleEvent, object[] evParams = null, System.Reflection.MethodInfo mi = null, object caller = null)
        {
            if(caller == null)
            {
                caller = this;
            }
            
            if(this.VerboseMode)
            {
                this.StatusInfo("Evento '" + evName + "': detectado.");
            }

            if(mi == null)
            {
                // Escopo do handler é em um form
                if(evParams[0] != null)
                {
                    try
                    {
                        dynamic evObj = evParams[0];
                        string FormUID = evObj.FormUID;

                        // Verifica se o handler é em um form ou userFields do addon
                        if(!String.IsNullOrEmpty(FormUID))
                        {

                            //  Pode ser um exec direto, tipo "onCreate"
                            try
                            {

                                // Em UserFields
                                if(((FastOneItemEvent)evParams[0]).userFieldsHandler != null)
                                {
                                    string userFieldClass = "UserFields";
                                    if(this.UserFieldClass.ContainsKey(FormUID))
                                    {
                                        userFieldClass = this.UserFieldClass[FormUID];
                                    }

                                    mi = Type.GetType(this.AddonInfo.Namespace + "." + userFieldClass + "," + this.AddonInfo.ExeName).GetMethod(evName);
                                    if(mi != null) caller = ((FastOneItemEvent)evParams[0]).userFieldsHandler;

                                    // Ou em Forms
                                } else
                                {
                                    mi = Type.GetType(this.AddonInfo.Namespace + "." + FormUID + "," + this.AddonInfo.ExeName).GetMethod(evName);
                                    if(mi != null) caller = this.FormList[FormUID];

                                }
                            } catch(Exception e)
                            {

                                // Form SAP? Poderia ser trocado por getForm e IsSystem
                                int n;
                                if(int.TryParse(evObj.FormTypeEx, out n))
                                {
                                    // Userfields em classe separada
                                    string userFieldClass = "UserFields";
                                    if(this.UserFieldClass.ContainsKey(evObj.FormTypeEx))
                                    {
                                        userFieldClass = this.UserFieldClass[evObj.FormTypeEx];
                                    }
                                    mi = Type.GetType(this.AddonInfo.Namespace + "." + userFieldClass + "," + this.AddonInfo.ExeName).GetMethod(evName);
                                    if(mi != null)
                                    {
                                        caller = (userFieldClass == "UserFields" 
                                            ? this.UserFields["UserFields"]
                                            : this.UserFields[evObj.FormTypeEx]
                                        ); // ((FastOneItemEvent)evParams[0]).userFieldsHandler;
                                    }
                                    //if(mi != null) caller = this.UserFields[evObj.FormTypeEx];

                                // Form normal do addon
                                } else
                                {
                                    mi = Type.GetType(this.AddonInfo.Namespace + "." + FormUID + "," + this.AddonInfo.ExeName).GetMethod(evName);
                                    if(mi != null) caller = this.FormList[FormUID];
                                }
                            }
                        }

                    } catch(Exception e)
                    {
                        //this.StatusErro(((System.Reflection.MethodBase)e.TargetSite).Name + ": " + e.Message);
                    }
                }
            }

            // Escopo do handler é global
            if(mi == null)
            {
                mi = this.GetType().GetMethod(evName);
            }

            bool handler_existe = (mi != null);
            if(handler_existe)
            {
                try
                {
                    mi.Invoke(caller, evParams);
                    if(this.VerboseMode)
                    {
                        this.StatusInfo("Evento '" + evName + "': tratamento personalizado no add-on executado.");
                    }

                } catch(Exception e)
                {
                    this.DesenvTimeError(e, " - Evento '" + evName + "'");
                }
            } else
            {
                if(this.VerboseMode)
                {
                    this.StatusInfo("Evento '" + evName + "': não possui tratamento personalizado no add-on.");
                }
            }
            GC.Collect();

            try
            {
                BubbleEvent = (evParams[1] != null ? (bool)evParams[1] : true);
            } catch(Exception e)
            {
                BubbleEvent = true;
            }

            // Retorna true se o handler existe no addon:
            return handler_existe;
        }
        
        /// <summary>
        /// Recupera o DocEntry com base no evento.
        /// </summary>
        /// <param name="evObj"></param>
        /// <returns></returns>
        public string GetEventObjectKey(SAPbouiCOM.BusinessObjectInfo evObj)
        {
            String docEntry = "";
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(evObj.ObjectKey);
                docEntry = xmlDoc.DocumentElement.FirstChild.ChildNodes[0].Value;
            } catch(Exception e)
            {
                this.DesenvTimeError(e);
            }

            return docEntry;
        }

        #endregion


        #region :: Menus

        /// <summary>
        /// Registra menus.
        /// By Labs - 11/2012
        /// </summary>
        /// <example>
        ///  Para se criar vários menus de uma só vez para o add-on basta seguir o exemplo abaixo:
        ///  <code>
        ///    this.registerMenus(new List<menuStruct>() {
        ///        new menuStruct()     {refUID = "43520",           UID = "ztMnuTShark",     Label = "TShark",      Type = SAPbouiCOM.BoMenuType.mt_POPUP, Image = "logo.bmp"},
        ///          new menuStruct()   {refUID = "ztMnuTShark",     UID = "ztMnuHelloWorld", Label = "Hello World", Type = SAPbouiCOM.BoMenuType.mt_POPUP},
        ///            new menuStruct() {refUID = "ztMnuHelloWorld", UID = "ztMnuHelloOne",   Label = "Diz Hello"},
        ///            new menuStruct() {refUID = "ztMnuHelloWorld", UID = "ztMnuHelloTwo",   Label = "Diz Olá"},
        ///            new menuStruct() {refUID = "ztMnuHelloWorld", UID = "ztMnuHelloTree",  Label = "Diz Holla"}
        ///    });
        /// </code>
        /// </example>
        /// <param name="menuStructList"></param>
        /// <param name="quiet"></param>
        public void RegisterMenus(List<menuStruct> menuStructList, bool reset = false)
        {
            this.StatusInfo(this.AddonInfo.Descricao + ": Registrando menus...");
            
            // Recupera o menu principal
            SAPbouiCOM.Form formCmdCenter = this.SBO_Application.Forms.GetFormByTypeAndCount(169, 1);
            string last_menu = "";
            try
            {
                formCmdCenter.Freeze(true);
                foreach(menuStruct item in menuStructList)
                {
                    last_menu = item.UID;
                    this.Menus.addMenuItem(item, reset);
                }

            } catch(Exception e)
            {
                this.DesenvTimeError(e, " - Erro registrando menu '" + last_menu + "'");

            } finally
            {
                formCmdCenter.Freeze(false);
                formCmdCenter.Update();
            }
        }

        #endregion


        #region :: User Tables

        /// <summary>
        /// Registra as tabelas de usuário passados em "tblList" e que estão declarados
        /// como função na classe "userClassName".
        /// </summary>
        /// <param name="tblList">Listagem de IDS das tabelas de usuário que possuem função criada em userClassName com o ID como nome e retornam a parametrização necessária.</param>
        /// <param name="userClassName">Classe onde estão definidos os parametros de registro, sendo por padrão "UserTables".</param>
        /// <param name="reset">Se true, remove tabelas antes de criar.</param>
        public void RegisterUserTables(List<string> tblList, string userClassName = "UserTables", bool reset = false)
        {
            this.StatusInfo(this.AddonInfo.Descricao + ": Validando tabelas...");

            // Reseta via XML
            if(!reset)
            {
                string r = this.Xml.UserTables.SelectSingleNode("fullReset").Attributes[0].Value;
                
                if(r.ToLower() == "true" || r == "1")
                {
                    this.ForceTableReset = true;
                }
            }

            try
            {

                // Instancia a classe de UserTables:
                Type dtsClassType = Type.GetType(this.AddonInfo.Namespace + "." + userClassName + "," + this.AddonInfo.ExeName);
                object dtsClass = Activator.CreateInstance(
                    dtsClassType, new object[] { }
                );
                System.Reflection.MethodInfo mi = null; 
                
                Dictionary<string, datasource> dtSources = new Dictionary<string, datasource>();
                foreach(string dtsId in tblList)
                {
                    try
                    {
                        mi = dtsClassType.GetMethod(dtsId);
                        datasource dts = (datasource)mi.Invoke(dtsClass, new object[] { });
                        if(dts != null)
                        {
                            dtSources.Add(dtsId, dts);
                        }
                    } catch (Exception e){
                        this.DesenvTimeError(e, " - Existe a função '" + dtsId + "' na classe '" + userClassName + 
                                                "'?\n - O namespace e nome do addon estão configurados corretamente em 'setInfo'?"
                        );
                        return;
                    }
                }


                if(reset || this.ForceTableReset)
                {
                    if(!this.oCompany.InTransaction)
                    {
                        this.oCompany.StartTransaction();
                    }

                    this.RemoveDataSources(tblList, dtSources, false, true, userClassName);
                    this.StatusInfo(this.AddonInfo.Descricao + ": Recriando tabelas...");
                    
                    if(this.oCompany.InTransaction)
                    {
                        this.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }

                if(!this.oCompany.InTransaction)
                {
                    this.oCompany.StartTransaction();
                }

                int t = 1;
                foreach(KeyValuePair<string, datasource> dts in dtSources)
                {
                    try
                    {
                        this.StatusInfo(this.AddonInfo.Descricao + ": Verificando tabelas do módulo '" + userClassName + "': '" + dts.Key + "' ("  + t + "/" + tblList.Count + ")");
                        this.DtSources.addDatasource(dts.Key, dts.Value, "", dtsClass);
                        t++;
                    
                    } catch(Exception e)
                    {
                        this.DesenvTimeError(e, " - Criando a tabela a função '" + dts.Key + "' na classe '" + userClassName + "'");
                        return;
                    }
                }
                
                if(this.oCompany.InTransaction)
                {
                    this.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }

            } catch(Exception e)
            {
                this.DesenvTimeError(e, " - Erro em registerUserTables\n - O namespace e nome do addon estão configurados corretamente em 'setInfo'?");

            } finally
            {
                this.Browser = this.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            }
        }

        /// <summary>
        /// Salva todas as tabelas de usuário em "tbList".
        /// By Labs - 09/2013
        /// </summary>
        /// <param name="tblList"></param>
        /// <param name="quiet"></param>
        public void SaveDataSources(string formId, List<string> tblList, bool quiet = false)
        {
            if(!quiet)
            {
                this.StatusInfo(this.AddonInfo.Descricao + ": Salvando dados...");
            }

            foreach(string dtsId in tblList)
            {
                this.DtSources.saveUserDataSource(dtsId, formId);
            }
        }

        /// <summary>
        /// Salva os dados de uma tabela de usuário.
        /// By Labs - 09/2013
        /// </summary>
        /// <param name="tblList"></param>
        /// <param name="quiet"></param>
        public void SaveDataSource(string formId, string dtsId, bool quiet = false)
        {
            if(!quiet)
            {
                this.StatusInfo(this.AddonInfo.Descricao + ": Salvando dados...");
            }

            this.DtSources.saveUserDataSource(dtsId, formId);
        }

        /// <summary>
        /// Remove todas as tabelas de usuário em "tbList".
        /// By Labs - 09/2013
        /// </summary>
        /// <param name="tblList"></param>
        /// <param name="quiet"></param>
        public void RemoveDataSources(List<string> tblList, Dictionary<string, datasource> dtSources, bool quiet = false, bool reverse = false, string mod = "" )
        {
            if(!quiet)
            {
                this.StatusInfo(this.AddonInfo.Descricao + ": Removendo tabelas...");
            }

            if(reverse)
            {
                for(int i = tblList.Count - 1; i >= 0; i--)
                {
                    this.StatusInfo(this.AddonInfo.Descricao + ": Removendo tabelas do módulo '" + mod + "': '" + tblList[i] + "' (" + (i+1) + "/" + tblList.Count + ")");
                    this.DtSources.removeDatasource(tblList[i], dtSources[tblList[i]]);
                }
            } else
            {
                foreach(string dtsId in tblList)
                {
                    this.DtSources.removeDatasource(dtsId, dtSources[dtsId]);
                }
            }
        }

        /// <summary>
        /// Remove uma tabela de usuário.
        /// By Labs - 09/2013
        /// </summary>
        /// <param name="tblList"></param>
        /// <param name="quiet"></param>
        public void RemoveDataSource(string dtsId, datasource dts, bool quiet = false)
        {
            if(!quiet)
            {
                this.StatusInfo(this.AddonInfo.Descricao + ": Removendo tabelas...");
            }

            this.DtSources.removeDatasource(dtsId, dts);
        }

        #endregion


        #region :: User Fields

        /// <summary>
        /// Registra customizações em Forms padrão SAP
        /// </summary>
        /// <param name="formList"></param>
        /// <param name="quiet"></param>
        public void RegisterUserFields(Dictionary<string, string> formList, bool quiet = false)
        {
            if(!quiet)
            {
                this.StatusInfo(this.AddonInfo.Descricao + ": Validando campos de usuário...");
            }

            foreach(KeyValuePair<string, string> userfield in formList)
            {
                try
                {
                    // Instancia classe
                    Type frmType = Type.GetType(this.AddonInfo.Namespace + "." + userfield.Value + "," + this.AddonInfo.ExeName);
                    this.UserFields.Add(userfield.Key, (UserFields)Activator.CreateInstance(frmType, new object[] { this }));

                    System.Reflection.MethodInfo mi = null;
                    mi = frmType.GetMethod("sapForm" + userfield.Key);
                    mi.Invoke(this.UserFields[userfield.Key], new object[] { });

                    // Acha Eventos declarados
                    MethodInfo[] myArrayMethodInfo = frmType.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);
                    for(int i = 0; i < myArrayMethodInfo.Length; i++)
                    {
                        MethodInfo myMethodInfo = (MethodInfo)myArrayMethodInfo[i];
                        try
                        {
                            this.UserFields[userfield.Key].EventMethods.Add(myMethodInfo.Name, userfield.Value);
                        } catch(Exception e)
                        {
                            // pode haver metodos com a mesmo nome e assinatura diferente
                        }
                    }

                    // Reseta via XML
                    string r = this.Xml.UserFields.SelectSingleNode("fullReset").Attributes[0].Value;
                    this.Xml.UserFields.SelectSingleNode("fullReset").Attributes[0].Value = "false";
                    this.Xml.Save();

                    if(this.ForceTableReset || (r.ToLower() == "true" || r == "1"))
                    {
                        this.UserFields[userfield.Key].recreate = true;
                    }

                    // Registra a galera:
                    this.UserFields[userfield.Key].registerUserFields();
                    this.UserFieldClass.Add(userfield.Key, userfield.Value);

                } catch(Exception e)
                {
                    this.DesenvTimeError(e, " - Existe a função 'sapForm" + userfield.Value + "' em '" + userfield.Key + "' ?");
                    return;
                }
            }

        }

        /// <summary>
        /// Registra as tabelas de usuário passados em "formList" e que estão declarados
        /// como função "sapFormXXXX" na classe "userClassName".
        /// </summary>
        /// <param name="formList">Listagem de IDS de forms SAP que possuem função criada em "UserFields.cs" com o formato sapFormXXXX, onde XXXX é o id do form no SAP,
        /// e onde estão definidos os parametros de criação do componente.</param>
        /// <param name="quiet">Se true, não exibe a mensagem de inicialização na barra de status.</param>
        public void RegisterUserFields(List<string> formList, bool quiet = false)
        {
            if(!quiet)
            {
                this.StatusInfo(this.AddonInfo.Descricao + ": Validando campos de usuário...");
            }

            // Instancia dinamicamente a classe de campos de usuário:
            string userClassName = "UserFields";
            try
            {
                this.UserFields[userClassName] = (UserFields)Activator.CreateInstance(
                    Type.GetType(this.AddonInfo.Namespace + "." + userClassName + "," + this.AddonInfo.ExeName), new object[] { this }
                );

                //this.UserFields = new TShark.UserFields(this);

                // Recupera os parâmetros
                System.Reflection.MethodInfo mi = null;
                foreach(string sapFormId in formList)
                {
                    try
                    {
                        Type frmType = Type.GetType(this.AddonInfo.Namespace + "." + userClassName + "," + this.AddonInfo.ExeName); 
                        mi = frmType.GetMethod("sapForm" + sapFormId);
                        //mi = this.UserFields.GetType().GetMethod("sapForm" + sapFormId);
                        mi.Invoke(this.UserFields[userClassName], new object[] { });

                        // Acha Eventos declarados
                        MethodInfo[] myArrayMethodInfo = frmType.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);
                        for(int i = 0; i < myArrayMethodInfo.Length; i++)
                        {
                            MethodInfo myMethodInfo = (MethodInfo)myArrayMethodInfo[i];
                            try
                            {
                                this.UserFields[userClassName].EventMethods.Add(myMethodInfo.Name, userClassName);
                            } catch(Exception e)
                            {
                                // pode haver metodos com a mesmo nome e assinatura diferente
                            }
                        }

                    } catch(Exception e)
                    {
                        this.DesenvTimeError(e, " - Existe a função 'sapForm" + sapFormId + "' em 'Userfields' ?");
                        return;
                    }
                }

                // Reseta via XML
                string r = this.Xml.UserFields.SelectSingleNode("fullReset").Attributes[0].Value;
                this.Xml.UserFields.SelectSingleNode("fullReset").Attributes[0].Value = "false";
                this.Xml.Save();

                if(r.ToLower() == "true" || r == "1")
                {
                    this.UserFields[userClassName].recreate = true;
                }

                // Registra a galera:
                this.UserFields[userClassName].registerUserFields();

            } catch(Exception e)
            {
                this.DesenvTimeError(e, " - Existe a classe Userfields no projeto?");
            }
        }

        /// <summary>
        /// Handler executado na abertura dos forms SAP vigiados.
        /// By Labs - 09/2013
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void UserFieldsOnCreateHandler(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form form = this.SBO_Application.Forms.GetForm(evObj.FormTypeEx, evObj.FormTypeCount);
            string usr_class = (this.UserFields.ContainsKey(evObj.FormTypeEx) ? evObj.FormTypeEx : "UserFields");
            
            this.UserFields[usr_class].SapForm = form;
            this.UserFields[usr_class].SystemFormSetup(ref evObj);
        }

        /// <summary>
        /// Handler executado no resize dos forms SAP vigiados.
        /// By Labs - 09/2013
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void SystemFormResizeHandler(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form form = this.SBO_Application.Forms.GetForm(evObj.FormTypeEx, evObj.FormTypeCount);
            string usr_class = (this.UserFields.ContainsKey(evObj.FormTypeEx) ? evObj.FormTypeEx : "UserFields");

            this.UserFields[usr_class].SapForm = form;
            this.UserFields[usr_class].SystemFormResize(ref evObj);
        }

        public void SystemFormOnActivateHandler(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form form = this.SBO_Application.Forms.GetForm(evObj.FormTypeEx, evObj.FormTypeCount);
            string usr_class = (this.UserFields.ContainsKey(evObj.FormTypeEx) ? evObj.FormTypeEx : "UserFields");

            this.UserFields[usr_class].SapForm = form;
        }

        #endregion


        #region :: Forms

        /// <summary>
        /// Carrega um form de um arquivo XML.
        /// By Labs - 12/2012
        /// </summary>
        public SAPbouiCOM.Form LoadForm(string FormID, string XMLFilePath)
        {

            // Result:
            SAPbouiCOM.Form SapForm = null;

            // Carrega o arquivo:
            if(this.LoadFromXML(XMLFilePath))
            {

                // Recupera o form:
                //  this.Addon.SBO_Application.Forms.Item.
                SapForm = this.SBO_Application.Forms.Item(FormID);

            }

            // Retorna:
            return SapForm;
        }

        /// <summary>
        /// Carrega os dados em um arquivo XML para o SAP.
        /// By Labs - 12/2012
        /// </summary>
        private bool LoadFromXML(string FileName, string sPath = "")
        {

            // Verifica se o add-on preencheu corretamente o execPath:
            if(this.ExecPath == "")
            {
                this.SBO_Application.MessageBox(
                    "Erro de Desenvolvimento: Não foi informado o execPath na inicialização\n" +
                    "Acrescente na inicialização da classe do add-on: this.execPath = System.IO.Directory.GetParent(Application.StartupPath).ToString();"
                );
                return false;
            }

            // Arquivo a ser carregado:
            string arq = (sPath != "" ? sPath : this.ExecPath) + @"\" + FileName;

            // Verifica se o arquivo existe:
            if(!System.IO.File.Exists(arq))
            {
                this.StatusErro("Arquivo não encontrado para carga: " + arq);
                return false;
            }

            // Carrega o arquivo:
            try
            {
                System.Xml.XmlDocument oXmlDoc = new System.Xml.XmlDocument();
                oXmlDoc.Load(arq);

                //  Carrega os dados:
                string strXML = oXmlDoc.InnerXml.ToString();
                this.SBO_Application.LoadBatchActions(ref strXML);

            } catch(Exception e)
            {
                if(this.VerboseMode)
                {
                    this.StatusErro("Erro na carga do arquivo: " + arq + e.Message);
                }
                return false;
            }

            // Retorna OK:
            return true;
        }


        /// <summary>
        /// Abre um form.
        /// </summary>
        /// <param name="FormClassName"></param>
        /// <param name="ExtraParams">A Classe Form DEVERÁ acrescentar na assinatura de construção: 'Dictionary<string, dynamic> ExtraParams'</param>
        /// <returns></returns>
        public Forms OpenForm(string FormClassName, Dictionary<string, dynamic> ExtraParams = null)
        {
            return this._OpenForm(FormClassName, ExtraParams: ExtraParams);
        }

        public Forms OpenFormUDOAdd(string FormClassName, Dictionary<string, dynamic> ExtraParams = null)
        {
            return this._OpenForm(FormClassName, ExtraParams: ExtraParams, UDOAddMode: true);
        }

        /// <summary>
        /// Abre um form, passando o atual como referencia.
        /// </summary>
        /// <param name="FormClassName"></param>
        /// <param name="FormOpenner"></param>
        /// <param name="ExtraParams">A Classe Form DEVERÁ acrescentar na assinatura de construção: 'Dictionary<string, dynamic> ExtraParams'</param>
        /// <returns></returns>
        public Forms OpenForm(string FormClassName, Forms FormOpenner, Dictionary<string, dynamic> ExtraParams = null)
        {
            return this._OpenForm(FormClassName, FormOpenner: FormOpenner, ExtraParams: ExtraParams);
        }

        public Forms OpenFormAdd(string FormClassName, Forms FormOpenner, Dictionary<string, dynamic> ExtraParams = null)
        {
            return this._OpenForm(FormClassName, FormOpenner: FormOpenner, ExtraParams: ExtraParams, AddMode: true);
        }

        public Forms OpenFormUDOAdd(string FormClassName, Forms FormOpenner, Dictionary<string, dynamic> ExtraParams = null)
        {
            return this._OpenForm(FormClassName, FormOpenner: FormOpenner, ExtraParams: ExtraParams, UDOAddMode: true);
        }

        /// <summary>
        /// Abre ou cria um formulário com base nas configurações em this.formList
        /// da class "formClassName" e vai para o último registro ou para o da chave
        /// fornecida em key_value. AddonKeys é definido automaticamente.
        /// By Labs - 11/2013
        /// </summary>
        /// <param name="FormClassName"></param>
        /// <param name="table"></param>
        /// <param name="table_key"></param>
        /// <param name="key_value"></param>
        /// <param name="FormOpenner"></param>
        /// <param name="ExtraParams">A Classe Form DEVERÁ acrescentar na assinatura de construção: 'Dictionary<string, dynamic> ExtraParams'</param>
        /// <returns></returns>
        public Forms OpenForm(string FormClassName, string table, string table_key, string key_value = "", Forms FormOpenner = null, Dictionary<string, dynamic> ExtraParams = null)
        {
            return this._OpenForm(FormClassName, Conditions: this._buildConditions(table, table_key, key_value), FormOpenner: FormOpenner, ExtraParams: ExtraParams, UDOFindCode: key_value);
        }

        public Forms OpenFormAdd(string FormClassName, string table, string table_key, string key_value = "", Forms FormOpenner = null, Dictionary<string, dynamic> ExtraParams = null)
        {
            return this._OpenForm(FormClassName, Conditions: this._buildConditions(table, table_key, key_value), FormOpenner: FormOpenner, ExtraParams: ExtraParams, AddMode: true);
        }

        public Forms OpenFormUDOAdd(string FormClassName, string table, string table_key, string key_value = "", Forms FormOpenner = null, Dictionary<string, dynamic> ExtraParams = null)
        {
            return this._OpenForm(FormClassName, Conditions: this._buildConditions(table, table_key, key_value), FormOpenner: FormOpenner, ExtraParams: ExtraParams, UDOAddMode: true);
        }

        private SAPbouiCOM.Conditions _buildConditions(string table, string table_key, string key_value = "")
        {
            // Define conditions para pegar última cotação:
            SAPbouiCOM.Conditions conditions = this.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
            SAPbouiCOM.Condition condition = conditions.Add();
            condition.Alias = table_key;
            condition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            if(!String.IsNullOrEmpty(key_value))
            {
                condition.CondVal = key_value;
            } else
            {
                condition.CondVal = this.DtSources.getMaxCode(table);
            }
            string k = table.Substring(1);
            if(!this.AddonKeys.ContainsKey(k))
            {
                this.AddonKeys.Add(k, condition.CondVal);
            } else
            {
                this.AddonKeys[k] = condition.CondVal;
            }
            return conditions;
        }


        /// <summary>
        /// Abre um form, passando o atual como referencia.
        /// </summary>
        /// <param name="FormClassName"></param>
        /// <param name="Conditions"></param>
        /// <param name="FormOpenner"></param>
        /// <param name="ExtraParams">A Classe Form DEVERÁ acrescentar na assinatura de construção: 'Dictionary<string, dynamic> ExtraParams'</param>
        /// <returns></returns>
        public Forms OpenForm(string FormClassName, SAPbouiCOM.Conditions Conditions, Forms FormOpenner = null, Dictionary<string, dynamic> ExtraParams = null)
        {
            return this._OpenForm(FormClassName, Conditions: Conditions, FormOpenner: FormOpenner, ExtraParams: ExtraParams);
        }
        public Forms OpenFormAdd(string FormClassName, SAPbouiCOM.Conditions Conditions, Forms FormOpenner = null, Dictionary<string, dynamic> ExtraParams = null)
        {
            return this._OpenForm(FormClassName, Conditions: Conditions, FormOpenner: FormOpenner, ExtraParams: ExtraParams, AddMode: true);
        }
        public Forms OpenFormUDOAdd(string FormClassName, SAPbouiCOM.Conditions Conditions, Forms FormOpenner = null, Dictionary<string, dynamic> ExtraParams = null)
        {
            return this._OpenForm(FormClassName, Conditions: Conditions, FormOpenner: FormOpenner, ExtraParams: ExtraParams, UDOAddMode: true);
        }

        /// <summary>
        /// Abre um form UDO e posiciona no Code especificado
        /// </summary>
        /// <param name="FormClassName"></param>
        /// <param name="UDOFindCode"></param>
        /// <param name="FormOpenner"></param>
        /// <param name="ExtraParams"></param>
        /// <returns></returns>
        public Forms OpenFormUDOFind(string FormClassName, string UDOFindCode, Forms FormOpenner = null, Dictionary<string, dynamic> ExtraParams = null)
        {
            return this._OpenForm(FormClassName, UDOFindCode: UDOFindCode, FormOpenner: FormOpenner, ExtraParams: ExtraParams);
        }

        /// <summary>
        /// Abre ou cria um formulário com base nas configurações em this.formList
        /// da class "formClassName". Neste caso, deve-se registrar manualmente o valor 
        /// de addonKeys.
        /// By Labs - 12/2012
        /// </summary>
        /// <param name="FormClassName"></param>
        /// <param name="Conditions"></param>
        /// <param name="FormOpenner"></param>
        /// <param name="ExtraParams"></param>
        /// <param name="ExtraParams">A Classe Form DEVERÁ acrescentar na assinatura de construção: 'Dictionary<string, dynamic> ExtraParams'</param>
        /// <returns>A instancia do form (FastOne Forms) criada</returns>
        internal Forms _OpenForm(string FormClassName, SAPbouiCOM.Conditions Conditions = null, Forms FormOpenner = null, 
            Dictionary<string, dynamic> ExtraParams = null, bool AddMode = false, bool UDOAddMode = false, string UDOFindCode = "")
        {

            #region :: Instancía a classe form

            if(!this.FormList.ContainsKey(FormClassName))
            {

                // Instancia dinamicamente o form com base em seu id:
                try
                {
                    Type frmType = Type.GetType(this.AddonInfo.Namespace + "." + FormClassName + "," + this.AddonInfo.ExeName);
                    this.FormList[FormClassName] = (Forms)Activator.CreateInstance(frmType, new object[] { this, ExtraParams });
                    this.FormList[FormClassName].Oppener = FormOpenner;

                    // Acha Eventos declarados
                    MethodInfo[] myArrayMethodInfo = frmType.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);
                    for(int i = 0; i < myArrayMethodInfo.Length; i++)
                    {
                        MethodInfo myMethodInfo = (MethodInfo)myArrayMethodInfo[i];
                        try
                        {
                            this.FormList[FormClassName].EventMethods.Add(myMethodInfo.Name, FormClassName);
                        } catch(Exception e)
                        {
                            // pode haver metodos com a mesmo nome e assinatura diferente
                        }
                    }

                    // Registra Eventos
                    this.FormList[FormClassName].registerFormEvents();

                } catch(Exception e)
                {
                    this.DesenvTimeError(e, " - Existe a classe '" + FormClassName + "' no projeto?\n - O namespace e o nome do addon foi informado corretamente em 'setInfo'?");
                    return null;
                }
            }

            #endregion

            // Reseta containers de matrix
            this.FormList[FormClassName].MatrixRefreshColumns = new Dictionary<string, Dictionary<string, List<string>>>();
            this.FormList[FormClassName].MatrixEmptyRows = new Dictionary<string, List<string>>();
            this.FormList[FormClassName].MatrixParams = new Dictionary<string, SetupMatrixParams>();
            this.FormList[FormClassName].UDOCode = UDOFindCode;

            // Valida se o form já está aberto
            this.FormList[FormClassName].SapForm = null;
            try
            {
                this.FormList[FormClassName].SapForm = this.SBO_Application.Forms.Item(FormClassName);
                if(this.VerboseMode)
                {
                    this.StatusErro("Form '" + FormClassName + "' já existente no SAP.");
                }

                // Traz o form pra frente
                this.FormList[FormClassName].SapForm.Select();


                if(!String.IsNullOrEmpty(UDOFindCode))
                {

                    SAPbouiCOM.Conditions conds = this.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                    SAPbouiCOM.Condition condition = conds.Add();
                    condition.Alias = "Code";
                    condition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    condition.CondVal = UDOFindCode;

                    SAPbouiCOM.DBDataSource dts = this.FormList[FormClassName].SapForm.DataSources.DBDataSources.Item(this.FormList[FormClassName].FormParams.MainDatasource);
                    dts.Query(conds);
                    if(dts.Size > 0)
                    {
                        this.FormList[FormClassName].SapForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    } else
                    {
                        this.FormList[FormClassName].SapForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    }
                }

                // Retorna o form
                return this.FormList[FormClassName];

            } catch(Exception e)
            {
                if(this.VerboseMode)
                {
                    this.StatusErro("Form '" + FormClassName + "': " + e.Message);
                }
            }

            // Cria um novo form SAP
            this.FormList[FormClassName].Status = FormStatus.frmCreating;
            if(ExtraParams != null)
            {
                this.FormList[FormClassName].ExtraParams = ExtraParams;
            }
            this.FormList[FormClassName].Oppener = FormOpenner;
            this.FormList[FormClassName].LoadingFromXML = false;

            // TODO: Implementar mensagem de UDO não pronto na instalação

            // Form armazenado em XML
            bool saveXML = true;
            string arq = this.ExecPath + @"\forms\" + FormClassName + ".xml";
            if((this.UseFormXML || this.FormList[FormClassName].FormParams.UseXML) && System.IO.File.Exists(arq))
            {
                this.FormList[FormClassName].LoadingFromXML = true;
                this.FormList[FormClassName].SapForm = this.LoadForm(FormClassName, @"forms\" + FormClassName + ".xml");
                if(this.FormList[FormClassName].SapForm == null)
                {
                    this.StatusErro("Não foi possível carregar o form a partir do XML");
                } else
                {
                    this.FormList[FormClassName].RebuildLinks(true, Conditions);
                }
                saveXML = false;

            } else
            {

                // Se ainda não existe, cria:
                if(this.FormList[FormClassName].SapForm == null)
                {
                    this.FormList[FormClassName].makeForm(true, Conditions);
                }
            }

            // Inicializa dados
            if(this.FormList[FormClassName].SapForm != null)
            {
                try
                {
                    this.FormList[FormClassName].Status = FormStatus.frmCreated;
                    this.FormList[FormClassName].SapForm.PaneLevel = 1;
                   // this.FormList[FormClassName].SapForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

                    // Se em modo de inserção:
                    this.FormList[FormClassName].InInsertMode = UDOAddMode || AddMode;
                    
                    // Se em modo de find
                    if(!String.IsNullOrEmpty(UDOFindCode))
                    {
                        this.FormList[FormClassName].UDOCode = UDOFindCode;
                        if(!String.IsNullOrEmpty(this.FormList[FormClassName].FormParams.BrowseByComp))
                        {
                            this.FormList[FormClassName].SapForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        }
                    }

                    if(!this.FormList[FormClassName].SapForm.Visible)
                    {
                        this.FormList[FormClassName].SapForm.Visible = true;
                    }
                } catch(Exception e)
                {
                    this.DesenvTimeError(e, "Criando o form " + FormClassName);
                } finally
                {
                    if(this.FormList[FormClassName].SapForm != null)
                    {
                        this.FormList[FormClassName].SapForm.Freeze(false);
                    }
                }

            } else
            {
                this.FormList[FormClassName].Status = FormStatus.frmNull;
            }

            this.FormList[FormClassName].LoadingFromXML = false;
            if(this.FormList[FormClassName].SapForm != null)
            {
                this.FormList[FormClassName].SapForm.Freeze(false);
            }

            // Form armazenado em XML
            if(saveXML && (this.UseFormXML || this.FormList[FormClassName].FormParams.UseXML) && this.FormList[FormClassName].SapForm != null)
            {
                if(!System.IO.Directory.Exists(this.ExecPath + @"\forms\"))
                {
                    System.IO.Directory.CreateDirectory(this.ExecPath + @"\forms\");
                }
                System.Xml.XmlDocument oXmlDoc = new System.Xml.XmlDocument();
                oXmlDoc.LoadXml(this.FormList[FormClassName].SapForm.GetAsXML());
                oXmlDoc.Save(arq);
            }

            // Retorna:
            GC.Collect();
            return this.FormList[FormClassName];
        }


        /// <summary>
        /// Recupera ou um form SAP padrão ou um form FastOne instanciado
        /// </summary>
        /// <param name="frmId">Id do Form</param>
        /// <returns>Retorna um form sap</returns>
        public SAPbouiCOM.Form getForm(string frmId = "", int frmCount = 0)
        {
            SAPbouiCOM.Form frm = null;

            if(!String.IsNullOrEmpty(frmId))
            {

                // Se passou inclusive com o frmCount, pega direitinho
                if(frmCount > 0)
                {
                    try
                    {
                        frm = this.SBO_Application.Forms.GetForm(frmId, frmCount);
                    } catch(Exception e) { }

                    // Senão tenta pelo id
                } else
                {
                    try
                    {
                        frm = this.SBO_Application.Forms.Item(frmId);
                    } catch(Exception e) {
                        try
                        {
                            frm = this.SBO_Application.Forms.GetForm(frmId, frmCount);
                        } catch(Exception er)
                        {
                            frm = null;
                        }
                    }
                }
            }

            // Se não rolou o form (ainda):
            if(frm == null)
            {
                frm = (this._lastFormByEvent != null
                        ? this._lastFormByEvent                  // Se tiver o form do ultimo evento, vai ele
                        : this.SBO_Application.Forms.ActiveForm  // Se não, vai o form ativo mesmo, e seja o que Deus quiser
                );
            }

            // Retorna o form:
            return frm;
        }

        #endregion



        #region :: Soap


        /// <summary>
        /// Efetua conexão com a SUPERBUY
        /// Monta o pacote de acordo com o xml recebido
        /// Monta o cabeçalho a ser enviado.
        /// Envia o pacote montado
        /// </summary>
        /// <param name="xml">Pacote montado em função acessória de regra de negócio</param>
        /// <param name="operacao">Identificador de operação que irá guiar o Gateway de callbacks no retorno do pacote.</param>
        private void _SoapSEND(string xml, object callback_func, Dictionary<string, string> headers = null, bool GET = true)
        {
            try
            {
                WebClient client = new WebClient();

                client.Encoding = Encoding.UTF8;

                // Headers
                client.Headers.Add(HttpRequestHeader.UserAgent, "Apache-HttpClient/4.1.1");
                client.Headers.Add(HttpRequestHeader.ContentType, "text/xml;charset=utf-8");
                if(headers != null)
                {
                    foreach(KeyValuePair<string, string> header in headers)
                    {
                        client.Headers.Add(header.Key, header.Value);
                    }
                }

                // Credenciais
                if(this.SoapConfig.useCredentials)
                {
                    client.Credentials = new NetworkCredential(this.SoapConfig.user, this.SoapConfig.pwd);
                }

                // Montando o pacote
                string xmlString =
                        "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:web=\"" + this.SoapConfig.host + "\"> " +
                           "<soapenv:Header/> " +
                           "<soapenv:Body> " +

                                xml +

                           "</soapenv:Body> " +
                        "</soapenv:Envelope>";

                // Envia o pacote
                if(GET)
                {
                    var callback = (Action<object, UploadStringCompletedEventArgs>)callback_func;
                    client.UploadStringCompleted += new UploadStringCompletedEventHandler(callback);
                    client.UploadStringAsync(new Uri(this.SoapConfig.host), xmlString);

                } else
                {
                    var callback = (Action<object, UploadDataCompletedEventArgs>)callback_func;
                    client.UploadDataCompleted += new UploadDataCompletedEventHandler(callback);
                    client.UploadDataAsync(new Uri(this.SoapConfig.host), System.Text.Encoding.ASCII.GetBytes(xmlString));
                }
                this.StatusInfo("Estabelecendo conexão... enviando pacote...");

            } catch(Exception e)
            {
                this.StatusErro("Erro ao estabelecer comunicação remota! " + e.Message);
            }
        }

        /// <summary>
        /// Executa uma comunicação SOAP utilizando o methodo GET do HTTP
        /// </summary>
        /// <param name="xml"></param>
        /// <param name="operacao"></param>
        public void SoapGET(string xml, Action<object, UploadStringCompletedEventArgs> callback_func, Dictionary<string, string> headers = null)
        {
            this._SoapSEND(xml, callback_func, headers);
        }

        /// <summary>
        /// Executa uma comunicação SOAP utilizando o methodo POST do HTTP
        /// </summary>
        /// <param name="xml"></param>
        /// <param name="operacao"></param>
        public void SoapPOST(string xml, Action<object, UploadStringCompletedEventArgs> callback_func, Dictionary<string, string> headers = null)
        {
            this._SoapSEND(xml, callback_func, headers, true);
        }

        #endregion




        #region Integração WebDriver

        /// <summary>
        /// Dá um ping no servidor
        /// </summary>
        /// <returns>Mensagem de ok ou erro</returns>
        public void Ping()
        {

            // Executa:
            this.StatusInfo("Verificando conexão com o server...");
            call call = new call();
            call.exec = "ping";
            string result = this.WebDriver.Call(ref call);

            // Verifica:
            if(!String.IsNullOrEmpty(result))
            {
                this.StatusErro(result);
            }

        }

        /// <summary>
        /// Callback do ping.
        /// </summary>
        /// <param name="result"></param>
        public void ping_Callback(dynamic result)
        {
            if(result.ContainsKey("ok"))
            {
                this.StatusInfo("Conexão ao server OK!");
            }
        }

        /// <summary>
        /// Lista mensagens vindas do server.
        /// </summary>
        /// <param name="result"></param>
        public bool showMessage_Callback(dynamic result)
        {
            bool ret = false;
            if(result.ContainsKey("mensagem"))
            {
                string tipo = result["mensagem"].ContainsKey("tipo")   ? result["mensagem"]["tipo"]   : "4";
                string tit  = result["mensagem"].ContainsKey("titulo") ? result["mensagem"]["titulo"] : "";
                string msg  = result["mensagem"].ContainsKey("msg")    ? result["mensagem"]["msg"]    : "";
                string desc = result["mensagem"].ContainsKey("desc")   ? result["mensagem"]["desc"]   : "";

                tit = "Cloud: " + tit + " - " + msg;
                desc = ":: " + desc;

                switch(tipo)
                {
                    case "1":
                        ret = true;
                        this.StatusErro(desc); this.StatusErro(tit);
                        break;

                    case "2":
                        ret = false;
                        this.StatusAlerta(desc); this.StatusAlerta(tit);
                        break;

                    default:
                        ret = false;
                        this.StatusInfo(desc); this.StatusInfo(tit);
                        break;
                }

                this.ShowMessage(tit + "\n" + desc);
            }
            return ret;
        }

        #endregion


        #region :: Utils

        /// <summary>
        /// Executa um SQL e alimenta this.Browser.
        /// </summary>
        /// <param name="sql"></param>
        public bool Select(string sql, bool quiet = false)
        {
            return this.DtSources.Select(sql, quiet);
        }
        public bool ExecSql(string sql, bool quiet = false)
        {
            return this.DtSources.Select(sql, quiet);
        }

        public string Truncate(string value, int maxLength)
        {
            if(!string.IsNullOrEmpty(value) && value.Length > maxLength)
            {
                return value.Substring(0, maxLength);
            }

            return value;
        }

        /// <summary>
        /// Pega um string "time" do SAP (945 => 09:45) e converte para datetime.
        /// By Labs - 09/2013
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public DateTime fromSAPToTime(string time)
        {
            DateTime d = new DateTime(1999, 1, 1);
            double[] res = this.parseSAPTime(time);
            d = d.AddHours(res[0]);
            d = d.AddMinutes(res[1]);
            return d;
        }

        public string ToSAPTime(string time)
        {
            return Regex.Replace(time, @"[^\d]", string.Empty); // (time.Split(':').ToArray()).Join('');
        }


        public string ToSAPDate(string date)
        {
            return Regex.Replace(date, @"[^\d]", string.Empty);
        }

        public string ToSAPDate(DateTime date)
        {
            return date.Year + "" 
                + (date.Month < 10 ? "0" : "") + date.Month + "" 
                + (date.Day < 10 ? "0" : "") + date.Day; 
        }

        /// <summary>
        /// Pega um string "date" do SAP (20130804 => 04/08/2013) e converte para DateTime.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public DateTime fromSAPToDate(string date)
        {
            int[] res = this.parseSAPDate(date);
            return new DateTime(res[2], res[1], res[0]);
        }

        /// <summary>
        /// Pega um string "date" do SAP (20130804 => 04/08/2013) e converte para um string de data.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public string fromSAPToDateStr(string date)
        {
            if(!String.IsNullOrEmpty(date))
            {
                int[] res = this.parseSAPDate(date);
                return (res[0] < 10 ? "0" : "") + res[0] + "/" +
                       (res[1] < 10 ? "0" : "") + res[1] + "/" + 
                        res[2];
            } else
            {
                return "";
            }
        }

        /// <summary>
        /// Pega um string "date" (04/08/2013) e converte para DateTime.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public DateTime ToDatetime(string date)
        {
            string[] tmp = date.Split('/');
            return new DateTime(Convert.ToInt32(tmp[2]), Convert.ToInt32(tmp[1]), Convert.ToInt32(tmp[0]));
        }
        public string ToIsoDateStr(string date)
        {
            string[] tmp = date.Split('/');
            return tmp[2] + "-" + tmp[1] + "-" + tmp[0];
        }

        /// <summary>
        /// Pega um string "date" do SAP (20130804 => 04/08/2013) e um string de "time" (945 => 09:45)
        /// e converte para DateTime.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public DateTime fromSAPToDateTime(string date, string time = "0")
        {
            int[] dt = this.parseSAPDate(date);
            DateTime d = new DateTime(dt[2], dt[1], dt[0]);
            
            double[] tm = this.parseSAPTime(time);
            d = d.AddHours(tm[0]);
            d = d.AddMinutes(tm[1]);
            
            return d;
        }



        /// <summary>
        /// Pega um string "time" do SAP (945 => 09:45) e quebra em hora e minuto.
        /// </summary>
        /// <param name="value"></param>
        /// <returns>Retorna um array double onde arr[0] = hora e arr[1] = minuto.</returns>
        public double[] parseSAPTime(string value){
            double[] res = new double[] { 0, 0 };
            switch(value.Length)
            {
                case 4:
                    res[0] = Convert.ToDouble(value.Substring(0, 2));
                    res[1] = Convert.ToDouble(value.Substring(2, 2));
                    break;

                case 3:
                    res[0] = Convert.ToDouble(value.Substring(0, 1));
                    res[1] = Convert.ToDouble(value.Substring(1, 2));
                    break;

                case 2:
                    res[1] = Convert.ToDouble(value.Substring(0, 2));
                    break;

                case 1:
                    res[1] = Convert.ToDouble(value.Substring(0, 1));
                    break;
            }
            return res;
        }

        /// <summary>
        /// Pega um string "date" do SAP (20130804 => 04/08/2013) e quebra em dia, mes e ano.
        /// </summary>
        /// <param name="value"></param>
        /// <returns>Retorna um array int onde arr[0] = dia, arr[1] = mês e arr[2] = ano.</returns>
        public int[] parseSAPDate(string value)
        {
            int[] res = new int[] { 1, 1, 1999 };
            switch(value.Length)
            {
                case 8:
                    res[2] = Convert.ToInt32(value.Substring(0, 4));
                    res[1] = Convert.ToInt32(value.Substring(4, 2));
                    res[0] = Convert.ToInt32(value.Substring(6, 2));
                    break;

                case 6:
                    res[2] = 2000 + Convert.ToInt32(value.Substring(0, 2));
                    res[1] = Convert.ToInt32(value.Substring(2, 2));
                    res[0] = Convert.ToInt32(value.Substring(4, 2));
                    break;
            }
            return res;
        }



        /// <summary>
        /// Retorna o numero de linhas do arquivo
        /// </summary>
        /// <param name="f"></param>
        /// <returns></returns>
        public int CountLinesInFile(string f)
        {
            int count = 0;
            using(System.IO.StreamReader r = new System.IO.StreamReader(f))
            {
                string line;
                while((line = r.ReadLine()) != null)
                {
                    count++;
                }
            }
            return count;
        }

        /// <summary>
        /// Permite abrir arquivo, manipulando a thread
        /// </summary>
        public string FileDialog(String filtro="")
        {
            Thread td = new Thread(() => this.OpenDlg(filtro));

            td.SetApartmentState(ApartmentState.STA);
            td.IsBackground = true;
            td.Start();
            td.Join();
            return (this._filename_);
        }

        internal String _filename_;

        /// <summary>
        /// Busca Arquivo
        /// </summary>
        internal void OpenDlg(String filtro)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            System.Diagnostics.Process[] process = null;
            try
            {
                process = Process.GetProcessesByName("SAP Business One");
                openFile.Multiselect = false;
                openFile.Filter = filtro;
                int filterIndex = 0;

                openFile.FilterIndex = filterIndex;
                openFile.RestoreDirectory = true;

                if(process.Length > 0)
                {
                    for(int i = 0; i < process.Length; i++)
                    {
                        WindowWrapper MyWindow = new WindowWrapper(process[i].MainWindowHandle);
                        DialogResult ret = openFile.ShowDialog(MyWindow);

                        if(ret == DialogResult.OK)
                        {
                            this._filename_ = openFile.FileName;
                           // openFile.Dispose();
                        } else
                        {
                            this._filename_ = "";
                           // openFile.Dispose();
                            System.Windows.Forms.Application.ExitThread();
                        }
                    }
                }
            } catch(Exception ex)
            {
                this.StatusAlerta(ex.Message);

            } finally
            {
                openFile.Dispose();
                System.Threading.Thread.CurrentThread.Abort();
            }
        }


        
        public dynamic ConvertFloat(dynamic value)
        {
            string tmp = Convert.ToString(value);
            return Convert.ToDecimal(tmp.Replace('.', ','));
        }

        #endregion

    }



    public struct DateTimeSpan
    {
        private readonly int years;
        private readonly int months;
        private readonly int days;
        private readonly int hours;
        private readonly int minutes;
        private readonly int seconds;
        private readonly int milliseconds;

        public DateTimeSpan(int years, int months, int days, int hours, int minutes, int seconds, int milliseconds)
        {
            this.years = years;
            this.months = months;
            this.days = days;
            this.hours = hours;
            this.minutes = minutes;
            this.seconds = seconds;
            this.milliseconds = milliseconds;
        }

        public int Years { get { return years; } }
        public int Months { get { return months; } }
        public int Days { get { return days; } }
        public int Hours { get { return hours; } }
        public int Minutes { get { return minutes; } }
        public int Seconds { get { return seconds; } }
        public int Milliseconds { get { return milliseconds; } }

        enum Phase { Years, Months, Days, Done }

        public static DateTimeSpan CompareDates(DateTime date1, DateTime date2)
        {
            if(date2 < date1)
            {
                var sub = date1;
                date1 = date2;
                date2 = sub;
            }

            DateTime current = date1;
            int years = 0;
            int months = 0;
            int days = 0;

            Phase phase = Phase.Years;
            DateTimeSpan span = new DateTimeSpan();

            while(phase != Phase.Done)
            {
                switch(phase)
                {
                    case Phase.Years:
                        if(current.AddYears(years + 1) > date2)
                        {
                            phase = Phase.Months;
                            current = current.AddYears(years);
                        } else
                        {
                            years++;
                        }
                        break;
                    case Phase.Months:
                        if(current.AddMonths(months + 1) > date2)
                        {
                            phase = Phase.Days;
                            current = current.AddMonths(months);
                        } else
                        {
                            months++;
                        }
                        break;
                    case Phase.Days:
                        if(current.AddDays(days + 1) > date2)
                        {
                            current = current.AddDays(days);
                            var timespan = date2 - current;
                            span = new DateTimeSpan(years, months, days, timespan.Hours, timespan.Minutes, timespan.Seconds, timespan.Milliseconds);
                            phase = Phase.Done;
                        } else
                        {
                            days++;
                        }
                        break;
                }
            }

            return span;
        }
    }
}
