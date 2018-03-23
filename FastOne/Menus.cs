using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TShark {

    /// <summary>
    /// Classe para a criação e operação de menus.
    /// By Labs - 10/2012
    /// </summary>
    public class Menus {

        /// <summary>
        /// Armazena menuIDs criados para evitar processamento 
        /// extra de eventos.
        /// By Labs - 11/2012
        /// </summary>
        internal List<string> _menu_ids;

        /// <summary>
        /// Armazena form ids que serão abertos automaticamente em onclick
        /// </summary>
        internal Dictionary<string, Dictionary<string, MenuOpenType>> MenuForms;

        /// <summary>
        /// Objeto SAP Application
        /// By Labs - 10/2012
        /// </summary>
        FastOne Addon;
        
        /// <summary>
        /// Inicializa a classe.
        /// By Labs - 10/2012
        /// </summary>
        /// <param name="SBO_App"></param>
        public Menus(FastOne addon) {

            // Referencia ao addon
            this.Addon = addon;

            // Lista de menus
            this._menu_ids = new List<string>();

            // Forms automaticos
            this.MenuForms = new Dictionary<string, Dictionary<string, MenuOpenType>>();
        }

        /// <summary>
        /// Insere apenas um ítem de menu.
        /// By Labs - 12/2012
        /// </summary>
        public SAPbouiCOM.MenuItem addMenuItem(menuStruct item, bool reset = false) {

            // Parâmetros de criação de menu:
            SAPbouiCOM.MenuCreationParams oMenuConfig = this.Addon.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

            // Posição, UID, String
            oMenuConfig.Position = item.Position;
            oMenuConfig.UniqueID = item.UID;
            oMenuConfig.String   = item.Label;
            oMenuConfig.Type     = item.Type;
            oMenuConfig.Enabled  = item.Enabled;
            oMenuConfig.Checked  = item.Checked;

            if(!String.IsNullOrEmpty(item.Image) && String.IsNullOrEmpty(item.ImgPath))
            {
                item.ImgPath = this.Addon.oCompany.BitMapPath;
            }

            // Image
            if (!String.IsNullOrEmpty(item.Image)) {
                oMenuConfig.Image = item.ImgPath + "//" + item.Image;
            } else {
                oMenuConfig.Image = "";
            }

            // Adiciona o menu:
            string refMenu = item.SAPParentId != MenusSAP.menuNone ? ((int)item.SAPParentId).ToString() : item.parentUID;
            SAPbouiCOM.MenuItem menuItem = this._AddMenuItem(refMenu, ref oMenuConfig, reset);

            // Ajusta form automatico
            if(menuItem != null && !String.IsNullOrEmpty(item.OpenForm) && !this.MenuForms.ContainsKey(item.UID))
            {
                this.MenuForms.Add(item.UID, new Dictionary<string, MenuOpenType>(){
                    {item.OpenForm, item.OpenFormType}
                });
            }

            // Retorna
            return menuItem;
        }

        /// <summary>
        /// Função interna para criação do ítem de menu no SAP.
        /// By Labs - 12/2012
        /// </summary>
        private SAPbouiCOM.MenuItem _AddMenuItem(string refMenu, ref SAPbouiCOM.MenuCreationParams oMenuConfig, bool reset = false) {
            SAPbouiCOM.MenuItem res = null;

            try {

                // Tenta inserir um item de menu:
                this._menu_ids.Add(oMenuConfig.UniqueID);

                // Pega o menu de referência:
                SAPbouiCOM.MenuItem oMenuRef = this.Addon.SBO_Application.Menus.Item(refMenu);

                // Reseta se o menu existe
                if(reset && oMenuRef.SubMenus.Exists(oMenuConfig.UniqueID))
                {
                    oMenuRef.SubMenus.RemoveEx(oMenuConfig.UniqueID);
                }

                // Menu já existe, recupera:
                if (oMenuRef.SubMenus.Exists(oMenuConfig.UniqueID)) {
                    res = oMenuRef.SubMenus.Item(oMenuConfig.UniqueID);

                // Senão, insere:
                } else {
                    res = oMenuRef.SubMenus.AddEx(oMenuConfig);
                }

                // Manda um "hello world":
                if (this.Addon.VerboseMode) {
                    this.Addon.StatusInfo("Acrescentado com sucesso o menu: " + oMenuConfig.String);
                }

            } catch (Exception e) {
                if (this.Addon.VerboseMode) {
                    this.Addon.StatusErro("Erro na criação do menu: " + oMenuConfig.String + " - " + e.Message);
                }
            }

            // Retorna item criado ou recuperado:
            return res;
        }

        /// <summary>
        /// Função interna para criação do ítem de menu no SAP.
        /// By Labs - 12/2012
        /// </summary>
        public bool removeMenu(string refMenu, string menuID) {
            bool res = true;
            try {

                // Remove o menu da listagem interna:
                this._menu_ids.Remove(menuID);

                // Pega o menu de referência:
                SAPbouiCOM.MenuItem oMenuRef = this.Addon.SBO_Application.Menus.Item(refMenu);

                // Verifica se o menu existe
                if (oMenuRef.SubMenus.Exists(menuID)) {
                    oMenuRef.SubMenus.RemoveEx(menuID);
                }

                // Manda um "hello world":
                if (this.Addon.VerboseMode) {
                    this.Addon.StatusInfo("Removido com sucesso o menu: " + menuID);
                }

            } catch (Exception e) {
                if (this.Addon.VerboseMode) {
                    this.Addon.StatusErro("Erro na remoção do menu: " + menuID + " - " + e.Message);
                }
                res = false;
            }

            // Retorna:
            return res;
        }

        public bool removeMenu(MenusSAP refMenu, string menuID)
        {
            return this.removeMenu(((int)refMenu).ToString(), menuID);
        }

    }
}
