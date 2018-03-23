using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Windows.Forms;
using TShark;
using System.Net;
using System.Collections.Specialized;
using System.Xml;

namespace ITOne
{
    /// <summary>
    /// IT-One Comissões de Vendas
    /// Primeira versão: 21/03/2015
    /// FastOne: 1.8.5
    /// By: Softlabs
    /// </summary>
    class Addon : TShark.FastOne
    {
        public int opprid   = 0;
        public int lineid   = 0;
        public int lead_ata = 0;
        public bool active  = false;
        public bool mostrar_menu = true;

        /***
         * Aceleração retroativa: neste caso o % de comissionamento considerado para o período inteiro será o % estabelecido na regra de configuração de meta.
         * Aceleração não retroativa: neste caso o % de comissionamento será % estabelecido na regra de comissonamento de meta, porém calculado apenas para novos registros
         * Sem aceleração: neste caso independente do % da regra, este não será aplicado. Considere esta opção como "inativo".
         ***/
        public Dictionary<string, string> tipos_aceleracao = new Dictionary<string, string>()
        {
            {"1","Aceleração retroativa"},
            {"2","Aceleração não retroativa"},
            {"3","Sem aceleração"},
        };
        
        public Addon(): base(Application.StartupPath, false)
        {
            this.AddonInfo.setInfo(
                1, 0, 0,                               // versão, release, revisão
                "ITOne",                                // nome APP
                "ITOne",                                // namespace
                "ITOne",                                // descrição
                "1.0.0 Alessandro - 10/01/2017"        // autor e data
            );

            // Antes de Iniciar
            this.AddOnInitialize();


            #region :: Criação de Menu

            // configuração no addon xml que fala se é uma versão especial para vendedores.
            // vendedores não podem ter acesso a configurações/relatórios
            XmlNode node_special_version = this.Xml.Data.SelectSingleNode("specialversion");
            if (node_special_version != null)
            {
                this.mostrar_menu = node_special_version.Attributes["value"].Value != "true";
            }

            //Registro de Menus
            string parent_menu = "43520";
            this.Menus.removeMenu(parent_menu, "mnuITOne");

            if( this.mostrar_menu )
            {
                this.RegisterMenus(new List<menuStruct>()
                {
                    new menuStruct(){parentUID = parent_menu, UID = "mnuITOne", Label = "IT One", Type = SAPbouiCOM.BoMenuType.mt_POPUP, Position = 14},

                    new menuStruct(){parentUID = "mnuITOne", UID = "mnuComissionamento",        Label = "Comissionamento", Type = SAPbouiCOM.BoMenuType.mt_POPUP},
                        new menuStruct(){parentUID = "mnuComissionamento", UID = "mnuUpdate",       Label = "SPLIT",  OpenForm = "FrmAtualizaValores"},
                        new menuStruct(){parentUID = "mnuComissionamento", UID = "mnuConfig",       Label = "Configurações",    OpenForm = "FrmListaConfigMetas"},

                    new menuStruct(){parentUID = "mnuITOne", UID = "mnuRelatorioFolder", Label = "Relatórios", Type = SAPbouiCOM.BoMenuType.mt_POPUP},
                        new menuStruct(){parentUID = "mnuRelatorioFolder", UID = "mnuRelatComiss",  Label = "Comissionamento",      OpenForm = "FrmRelatorio"},
                        new menuStruct(){parentUID = "mnuRelatorioFolder", UID = "mnuRelatEqupe",   Label = "Perfomance de Equipe", OpenForm = "FrmRelatorioEquipe"},
                        new menuStruct(){parentUID = "mnuRelatorioFolder", UID = "mnuRelatNao",  Label = "Aceleração Não Retroativa",      OpenForm = "FrmRelatorioAceleracao"},
                });
            }

            #endregion


            #region :: Registro de UserFields

            this.RegisterUserFields(new List<String>() {
                
                // CFL para Pedido de Venda na Aba de 'Etapas de Negócio' na Oportunidade de venda
                "10016",

                // Oportunidade de Vendas
                "320",

                // Pedido de Venda
                "139",

                // Cadastro de Colaboradores
                "60100",

                // Dev. Nota de Saída
                "179",

                // Nota Fiscal de Saída
                "133",

                // Contas a Receber
                "170"
            });

            #endregion


            #region :: Registro de Tabelas

            this.RegisterUserTables(new List<String>()
            {   
                "UPD_IT_METAS",
                "UPD_IT_COMISSAO",

                // UDO
                "UPD_IT_FUNCOES",

                // NoObject
                "UPD_IT_PARTICIP",
            }, "UserTables", false);

            #endregion

            
            
            // Após Inicializado
            this.AddOnInitialized();

            this.showDesenvTimeMsgs = false;
        }


        #region :: Métodos genéricos

        /// <summary>
        /// Reseta as propriedades da oportunidade que são utilizados no pedido.
        /// </summary>
        public void resetarDadosOppr()
        {
            this.opprid = 0;
            this.lineid = 0;
            this.lead_ata = 0;
            this.active = false;
        }

        /// <summary>
        /// Escreve em um arquivo .txt o sql gerado.
        /// </summary>
        public void SalvarSQLArquivoTXT(string sql)
        {
            try
            {
                string fullpath = System.Reflection.Assembly.GetEntryAssembly().Location.Replace("ITOne.exe", "") + "sql.txt";
                TextWriter tw = new StreamWriter(fullpath);
                tw.WriteLine(sql);
                tw.Close();
            }
            catch (Exception ex)
            {

            }
        }

        /// <summary>
        /// 
        /// </summary>
        public bool ExportarCSV(string nome_arquivo, SAPbouiCOM.DataTable dt, SAPbouiCOM.Grid grid)
        {
            bool ret = false;
            try
            {
                string folder = Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.System)) + "\\tmp";
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);

                string file = folder + "\\" + nome_arquivo;
                TextWriter tw = new StreamWriter(file, false, Encoding.UTF8);
                string seperador = ";";

                //cabeçalho do csv
                string cabecalho = "";
                List<int> colunas_invisiveis = new List<int>() { };
                for (int i = 0; i < grid.Columns.Count; i++)
                {
                    if (grid.Columns.Item(i).Visible)
                    {
                        cabecalho += grid.Columns.Item(i).TitleObject.Caption + seperador;
                    }
                    else
                    {
                        colunas_invisiveis.Add(i);
                    }
                }

                tw.WriteLine(cabecalho);

                string line = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        //se for valor de uma coluna invisível, então pula
                        if (colunas_invisiveis.Contains(j))
                            continue;

                        line += dt.GetValue(j, i) + seperador;
                    }

                    if (!String.IsNullOrEmpty(line))
                    {
                        tw.WriteLine(line);
                        line = "";
                    }
                }

                tw.Close();
                ret = true;
            }
            catch (Exception ex)
            {
                throw;
            }

            return ret;
        }

        #endregion
    }
}