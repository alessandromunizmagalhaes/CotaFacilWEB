using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TShark {

    /// <summary>
    /// Parametros para criação de fields.
    /// By Labs - 01/2013
    /// </summary>
    public class fieldParams {
        public string descricao; 
        public SAPbobsCOM.BoFieldTypes tipo;
        public SAPbobsCOM.BoFldSubTypes subtipo; 
        public int size;
        public string LinkedTable;
    }

    /// <summary>
    /// Parametros para alteração de fields.
    /// By Labs - 01/2013
    /// </summary>
    public class changeFieldParams {
        public string descricao;
        public int size;
    }

    /// <summary>
    /// Parametros para criação de indices
    /// </summary>
    public class keyParams {
        public List<string> fields;
        public SAPbobsCOM.BoYesNoEnum unique = SAPbobsCOM.BoYesNoEnum.tNO;
        public keyParams() {
            this.fields = new List<string>();
        }
    }

    /// <summary>
    /// Permite o versionamento de fields
    /// </summary>
    public class fieldsVersionamento {
        public Dictionary<string, fieldParams> novos;
        public List<string> remover;
        public Dictionary<string, changeFieldParams> alterar;

    }

    /// <summary>
    /// Permite o versionamento de indices
    /// </summary>
    public class keysVersionamento {
        public Dictionary<string, keyParams> novos;
        public List<string> remover;
        public Dictionary<string, keyParams> alterar;
    }

    /// <summary>
    /// Permite o versionamento de tabelas
    /// </summary>
    public class dtsVersionamento {
        public keysVersionamento keys;
        public fieldsVersionamento fields;
        public string SQLTableChange;
        public string onBeforeChangeFunc;
        public string onAfterChangeFunc;

        // Inicializa:
        public dtsVersionamento() {
            this.keys = new keysVersionamento();
            this.fields = new fieldsVersionamento();
        }
    }

    public class FindColumn {
        public string Alias;
        public string Description;
    }

    /// <summary>
    /// Parametros para registro de User Defined objects    
    /// By Labs - 01/2013
    /// </summary>
    public class UDO {
        public SAPbobsCOM.BoYesNoEnum CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
        public SAPbobsCOM.BoYesNoEnum CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
        public SAPbobsCOM.BoYesNoEnum CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
        public SAPbobsCOM.BoYesNoEnum CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;
        public SAPbobsCOM.BoYesNoEnum CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
        public SAPbobsCOM.BoYesNoEnum CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO;
        public SAPbobsCOM.BoYesNoEnum ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
        public SAPbobsCOM.BoYesNoEnum CanLog = SAPbobsCOM.BoYesNoEnum.tYES;
        public SAPbobsCOM.BoUDOObjType ObjectType;
        public string Name;
        public string TableName;
        public string MenuCaption;
        public List<string> ChildTables;
        public List<FindColumn> FindColumns;
        public bool remove_if_exists = false;
    }

    /// <summary>
    /// Armazena os datasources.
    /// By Labs - 01/2013
    /// </summary>
    public class datasource {
        public string id;
        public int versao;
        public string descricao;
        public string master_udo;
        public bool remove_if_exists = false;
        public bool ignoreTipo = false;
        public SAPbobsCOM.BoUTBTableType tipo;
        public Dictionary<string, fieldParams> fields;
        public Dictionary<string, keyParams> keys;
        public List<Dictionary<string, dynamic>> defRows;
        public UDO UDO;
        public Dictionary<int, dtsVersionamento> versoes;

        public string SQLTable;
    }
    
    /// <summary>
    /// Classe para a criação e operação de menus.
    /// By Labs - 10/2012
    /// </summary>
    public class Datasources {

        public Dictionary<string, datasource> dtsList;

        /// <summary>
        /// Objeto SAP Application
        /// By Labs - 10/2012
        /// </summary>
        FastOne Addon;

        private bool check_versao = false;
        
        /// <summary>
        /// Inicializa a classe.
        /// By Labs - 10/2012
        /// </summary>
        /// <param name="SBO_App"></param>
        public Datasources(FastOne addon) {

            // Referencia ao addon
            this.Addon = addon;

            //  this.removeDatasource("ZTH_VERSIONAMENTO");
            this.CreateVersionamento();
        }


        #region :: Versionamento

        internal void CreateVersionamento() 
        {

            // Cria tabela principal:
            this.addDatasource("ZTH_VERSIONAMENTO", new datasource()
            {
                id = "ZTH_VERSIONAMENTO",
                versao = 1,
                descricao = "Versionamento",
                SQLTable =
                   " CREATE TABLE [@ZTH_VERSIONAMENTO]( " +
                        " [code] [int] IDENTITY(1,1) NOT NULL, " +
                        " [tabela] [varchar](150) NULL, " +
                        " [addon]  [varchar](250) NULL, " +
                        " [versao] [int] NULL, " +
                        " [atualizacao] [datetime] NULL, " +
                        " [obs] [varchar](300) NULL, " +
                   "   CONSTRAINT [PK_ZTH_VERSIONAMENTO] PRIMARY KEY CLUSTERED ( [code] ASC ) " +
                   " ) "
            });
        }

        #endregion

        public void ReleaseBrowser()
        {
            if(this.Addon.Browser != null)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(this.Addon.Browser);
                    this.Addon.Browser = null;
                    GC.Collect();
                } catch { }
            }
        }

        /// <summary>
        /// Cria um datasource.
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="dts">Objeto de criação do datasource</param>
        /// <param name="udoParams">Parâmtros para criação do UDO</param>
        /// <param name="sql">Se informado, o SQL será executado ao final da criação do datasource</param>
        /// <returns></returns>
        public void addDatasource(string dtsId, datasource dts, string sql = "", object dtsClass = null) {
            int res = 0; string msg;
            bool criou = false;

            this.Addon.SBO_Application.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, true);

          /*  if(dtsId.Length > 19)
            {
                throw new Exception("Nome do tabela '" + dtsId + "' possui mais de 19 caracteres.");
            }
            */
            // Hack de criação de tabelas SQL padrão
            if(!string.IsNullOrEmpty(dts.SQLTable))
            {
                
                // Mandaram remover se existir?
                if(dts.remove_if_exists)
                {
                    this.removeDatasource(dts.id, dts, true);
                }

                SAPbobsCOM.Recordset rec = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                try
                {
                    rec.DoQuery(
                        " IF OBJECT_ID('[@" + dts.id + "]', 'U') IS NULL " + 
                             dts.SQLTable
                    );

                    // Versiona
                    if(dts.id != "ZTH_VERSIONAMENTO")
                    {
                        rec.DoQuery(
                            " INSERT INTO [@ZTH_VERSIONAMENTO] (tabela, versao, addon, atualizacao) " +
                            " VALUES (" +
                                "'" + dts.id + "', " +
                                "1, " +
                                "'" + this.Addon.AddonInfo.Descricao + " - " + this.Addon.AddonInfo.VersaoStr + "', " +
                                "GETDATE()" +
                            " ) "
                        );
                    }

                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, " criando tabela SQL padrão " + dts.id);
                } finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
                    rec = null;
                }

            // Cria datasource padrão SAP    
            } else
            {

                // Libera o recordset padrão:
                this.ReleaseBrowser();

                #region Acrescenta a tabela

                // Mandaram remover se existir?
                if(dts.remove_if_exists)
                {
                    this.removeDatasource(dtsId, dts);
                }

                // Nova tabela
                SAPbobsCOM.UserTablesMD tables = ((SAPbobsCOM.UserTablesMD)(this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));
                bool ja_existe = tables.GetByKey(dtsId);
                if(!ja_existe)
                {
                    
                    // Cria a nova
                    tables.TableName = dtsId;
                    tables.TableDescription = dts.descricao; 
                    if(!dts.ignoreTipo)
                    {
                        tables.TableType = dts.tipo;
                    }
                    res = tables.Add();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(tables);
                    GC.Collect();

                    // Verifica se deu erro 
                    if(res == 0)
                    {
                        ja_existe = true;
                        criou = true;

                        // Versiona
                        if(dts.id != "ZTH_VERSIONAMENTO")
                        {
                            SAPbobsCOM.Recordset rec = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            try
                            {
                                rec.DoQuery(
                                    " INSERT INTO [@ZTH_VERSIONAMENTO] (tabela, versao, addon, atualizacao) " +
                                    " VALUES (" +
                                        "'" + dts.id + "', " +
                                        "1, " +
                                        "'" + this.Addon.AddonInfo.Descricao + " - " + this.Addon.AddonInfo.VersaoStr + "', " +
                                        "GETDATE()" +
                                    " ) "
                                );

                            } catch(Exception e)
                            {
                                this.Addon.DesenvTimeError(e, " inserindo versionamento de " + dts.id);
                            } finally
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
                                rec = null;
                            }
                        }
                    } else
                    {
                        this.Addon.oCompany.GetLastError(out res, out msg);
                        if(this.Addon.showDesenvTimeMsgs)
                        {
                            this.Addon.ShowMessage(msg);
                        }
                    }
                }

                if(tables != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(tables);
                }
                GC.Collect();

                #endregion


                #region Acrescenta fields

                if(dts.fields != null)
                {
                    foreach(KeyValuePair<string, fieldParams> fld in dts.fields)
                    {
                        this.addField(dtsId, fld.Key, fld.Value);
                    }
                }

                #endregion


                #region Chaves

                if(dts.keys != null)
                {
                    foreach(KeyValuePair<string, keyParams> key in dts.keys)
                    {
                        this.addKey(dtsId, key.Key, key.Value);
                    }
                }

                #endregion


                #region Registra UDO

                if(dts.UDO != null)
                {
                    SAPbobsCOM.UserObjectsMD udoMD = null;
                    udoMD = ((SAPbobsCOM.UserObjectsMD)(this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));

                    // Verifica se já existe
                    string udoId = (!String.IsNullOrEmpty(dts.UDO.TableName) ? dts.UDO.TableName : dtsId) + "O";
                    bool existe = udoMD.GetByKey(udoId);

                    // Mandaram remover se existir?
                    if(existe && (dts.UDO.remove_if_exists || dts.remove_if_exists))
                    {
                        udoMD.Remove();
                        existe = false;
                    }

                    if(!existe)
                    {
                        udoMD.CanCancel = dts.UDO.CanCancel;
                        udoMD.CanClose = dts.UDO.CanClose;
                        udoMD.CanCreateDefaultForm = dts.UDO.CanCreateDefaultForm;
                        udoMD.CanDelete = dts.UDO.CanDelete;
                        udoMD.CanFind = dts.UDO.CanFind;
                        udoMD.CanLog = dts.UDO.CanLog;
                        udoMD.CanYearTransfer = dts.UDO.CanYearTransfer;
                        udoMD.ManageSeries = dts.UDO.ManageSeries;

                        udoMD.Code = udoId;
                        udoMD.Name = (!String.IsNullOrEmpty(dts.UDO.Name) ? dts.UDO.Name : udoId);
                        if(dts.UDO.ObjectType == 0)
                        {
                            if(dts.tipo == SAPbobsCOM.BoUTBTableType.bott_MasterData)
                            {
                                udoMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData;
                            } else
                            {
                                udoMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document;
                            }
                        } else
                        {
                            udoMD.ObjectType = dts.UDO.ObjectType;
                        }
                        if((dts.UDO.TableName != null) && (dts.UDO.TableName != ""))
                        {
                            udoMD.TableName = dts.UDO.TableName;
                        } else
                        {
                            udoMD.TableName = dtsId;
                        }

                        udoMD.MenuCaption = (!String.IsNullOrEmpty(dts.UDO.MenuCaption) ? dts.UDO.MenuCaption : dts.descricao);

                        if(dts.UDO.ChildTables != null)
                        {
                            foreach(string tb in dts.UDO.ChildTables)
                            {
                                udoMD.ChildTables.TableName = tb;
                                udoMD.ChildTables.Add();
                            }
                        }

                        if(dts.UDO.FindColumns != null)
                        {
                            foreach(FindColumn findCol in dts.UDO.FindColumns)
                            {
                                udoMD.FindColumns.ColumnAlias = findCol.Alias;
                                udoMD.FindColumns.ColumnDescription = findCol.Description;
                                udoMD.FindColumns.Add();
                            }
                        }

                        res = udoMD.Add();
                        if(res != 0)
                        {
                            this.Addon.oCompany.GetLastError(out res, out msg);
                            this.Addon.ShowMessage("FindColumns - " + msg);
                        }

                    }
                    if(udoMD != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(udoMD);
                    }
                    GC.Collect();
                }

                #endregion

                if(tables != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(tables);
                }
                GC.Collect();

            }


            #region :: Versionamento
                
            // Versionamento de tabelas
            int versao = 1;
            int top_versao = 1;
            if(/*!criou && */dts.id != "ZTH_VERSIONAMENTO" && dts.versoes != null)
            {
                SAPbobsCOM.Recordset rec = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                try
                {
                    rec.DoQuery("SELECT * FROM [@ZTH_VERSIONAMENTO] (nolock) WHERE tabela = '" + dts.id + "'");
                    if(rec.RecordCount > 0)
                    {
                        versao = rec.Fields.Item("versao").Value;
                    } else
                    {
                        rec.DoQuery(
                            " INSERT INTO [@ZTH_VERSIONAMENTO] (tabela, versao, addon, atualizacao) " +
                            " VALUES (" +
                                "'" + dts.id + "', " +
                                "'" + versao + "', " +
                                "'" + this.Addon.AddonInfo.Descricao + " - " + this.Addon.AddonInfo.VersaoStr + "', " +
                                "GETDATE()" +
                            " ) "
                        );
                    }

                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, " verificando versionamento de " + dts.id);
                } finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
                    rec = null;
                }
                    

                #region :: Versionamento

                // UserTables dtsClass = new UserTables();
                foreach(KeyValuePair<int, dtsVersionamento> v in dts.versoes)
                {
                    if(v.Key > versao)
                    {

                        if(v.Key > top_versao)
                        {
                            top_versao = v.Key;
                        }

                        bool change_ok = true;

                        #region :: Função beforeChange
                            
                        if(!String.IsNullOrEmpty(v.Value.onBeforeChangeFunc))
                        {
                            try
                            {
                                System.Reflection.MethodInfo mi = dtsClass.GetType().GetMethod(v.Value.onBeforeChangeFunc);
                                if(mi != null)
                                {
                                    change_ok = (bool)mi.Invoke(dtsClass, new object[] { this.Addon, v.Value });
                                }

                            } catch(Exception e)
                            {
                                this.Addon.DesenvTimeError(e, " - Existe a função '" + v.Value + "' na classe 'UserTables'?");
                                return;
                            }
                        }
                            
                        #endregion


                        if(change_ok)
                        {

                            #region :: Remove indices

                            if(v.Value.keys.remover != null)
                            {
                                SAPbobsCOM.UserKeysMD k = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys);
                                foreach(string r in v.Value.keys.remover)
                                {
                                    try
                                    {
                                        int rk = this.getKeyIndex(dtsId, r);
                                        if(rk > -1)
                                        {
                                            k.GetByKey(dtsId, rk);
                                            res = k.Remove();
                                            if(res != 0)
                                            {
                                                this.Addon.oCompany.GetLastError(out res, out msg);
                                                this.Addon.ShowMessage("Versionamento - Removendo indice " + r + ": " + msg);
                                            }
                                        }
                                    } catch(Exception e)
                                    {
                                        this.Addon.StatusErro("Erro atualizando versão de " + dtsId + ": " + e.Message);
                                    }
                                }
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(k);
                                GC.Collect();
                            }

                            #endregion

                            #region :: Altera indices

                            if(v.Value.keys.alterar != null)
                            {
                                foreach(KeyValuePair<string, keyParams> r in v.Value.keys.alterar)
                                {
                                    try
                                    {
                                        int rk = this.getKeyIndex(dtsId, r.Key);
                                        if(rk > -1)
                                        {
                                            SAPbobsCOM.UserKeysMD k = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys);
                                            k.GetByKey(dtsId, rk);

                                            // Remove o velho
                                            res = k.Remove();
                                            if(res != 0)
                                            {
                                                this.Addon.oCompany.GetLastError(out res, out msg);
                                                this.Addon.ShowMessage("Versionamento - Alterando indice " + r.Key + ": " + msg);
                                            }
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(k);
                                            GC.Collect();

                                            // Cria um novo
                                            this.addKey(dtsId, r.Key, r.Value);
                                        }
                                    } catch(Exception e)
                                    {
                                        this.Addon.StatusErro("Erro atualizando versão de " + dtsId + ": " + e.Message);
                                    }
                                }
                            }

                            #endregion

                            #region :: Cria novos indices

                            if(v.Value.keys.novos != null)
                            {
                                foreach(KeyValuePair<string, keyParams> r in v.Value.keys.novos)
                                {
                                    this.addKey(dtsId, r.Key, r.Value);
                                }
                            }

                            #endregion


                            #region :: Remove fields

                            if(v.Value.fields.remover != null)
                            {
                                SAPbobsCOM.UserFieldsMD f = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                                foreach(string r in v.Value.fields.remover)
                                {
                                    try
                                    {
                                        int rf = this.getFieldIndex("@" + dtsId, r);
                                        if(rf > -1)
                                        {
                                            f.GetByKey("@" + dtsId, rf);
                                            res = f.Remove();
                                            if(res != 0)
                                            {
                                                this.Addon.oCompany.GetLastError(out res, out msg);
                                                this.Addon.ShowMessage("Versionamento - Removendo campo " + r + ": " + msg);
                                            }
                                        }
                                    } catch(Exception e)
                                    {
                                        this.Addon.StatusErro("Erro atualizando versão de " + dtsId + ": " + e.Message);
                                    }
                                }
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(f);
                                GC.Collect();
                            }

                            #endregion

                            #region :: Altera field

                            if(v.Value.fields.alterar != null)
                            {
                                SAPbobsCOM.UserFieldsMD f = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                                foreach(KeyValuePair<string, changeFieldParams> r in v.Value.fields.alterar)
                                {
                                    try
                                    {
                                        int rf = this.getFieldIndex("@" + dtsId, r.Key);
                                        if(rf > -1)
                                        {
                                            if(f.GetByKey("@" + dtsId, rf))
                                            {
                                                if(!String.IsNullOrEmpty(r.Value.descricao))
                                                {
                                                    f.Description = r.Value.descricao;
                                                }
                                                if(r.Value.size != 0)
                                                {
                                                    //f.Size = r.Value.size;
                                                    f.EditSize = r.Value.size;
                                                }
                                                res = f.Update();
                                                if(res != 0)
                                                {
                                                    this.Addon.oCompany.GetLastError(out res, out msg);
                                                    this.Addon.ShowMessage("Versionamento - Alterando campo " + r.Key + ": " + msg);
                                                }
                                            }
                                        }
                                    } catch(Exception e)
                                    {
                                        this.Addon.StatusErro("Erro atualizando versão de " + dtsId + ": " + e.Message);
                                    }
                                }
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(f);
                                GC.Collect();
                            }

                            #endregion

                            #region :: Cria novos campos

                            if(v.Value.fields.novos != null)
                            {
                                foreach(KeyValuePair<string, fieldParams> r in v.Value.fields.novos)
                                {
                                    this.addField(dtsId, r.Key, r.Value);
                                }
                            }

                            #endregion


                            #region :: SQLTable

                            if(!String.IsNullOrEmpty(v.Value.SQLTableChange))
                            {
                                rec = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                try
                                {
                                    rec.DoQuery(v.Value.SQLTableChange);

                                } catch(Exception e)
                                {
                                    this.Addon.DesenvTimeError(e, " executando versionamento de " + dts.id);
                                } finally
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
                                    rec = null;
                                }
                            }

                            #endregion

                        }

                        #region :: Função afterChange

                        if(!String.IsNullOrEmpty(v.Value.onAfterChangeFunc))
                        {
                            try
                            {
                                System.Reflection.MethodInfo mi = dtsClass.GetType().GetMethod(v.Value.onAfterChangeFunc);
                                if(mi != null)
                                {
                                    mi.Invoke(dtsClass, new object[] { this.Addon, v.Value });
                                }

                            } catch(Exception e)
                            {
                                this.Addon.DesenvTimeError(e, " - Existe a função '" + v.Value + "' na classe 'UserTables'?");
                                return;
                            }
                        }

                        #endregion

                    }
                }
                    

                #endregion

                
                // Atualiza a versão da tabela
                if(versao != top_versao)
                {
                    versao = top_versao;
                    rec = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    try
                    {
                        rec.DoQuery(
                            "UPDATE [@ZTH_VERSIONAMENTO] " +
                            "   SET versao = '" + versao + "'" +
                            "      , addon = '" + this.Addon.AddonInfo.Descricao + " - " + this.Addon.AddonInfo.VersaoStr + "' " +
                            "      , atualizacao = GETDATE() " +
                            " WHERE tabela = '" + dts.id + "'"
                        );
                        versao = rec.RecordCount > 0 ? rec.Fields.Item("U_versao").Value : 1;

                    } catch(Exception e)
                    {
                        this.Addon.DesenvTimeError(e, " atualizando versionamento de " + dts.id);
                    } finally
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
                        rec = null;
                    }
                }
            }

            #endregion


            #region SQL extra

            if(sql != "")
            {
                this.Select(sql);
            }

            #endregion


            #region Rows default

            if(dts.defRows != null)
            {
                SAPbobsCOM.Recordset rec = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string new_code = "";
                try
                {
                    rec.DoQuery("SELECT count(*) total FROM [@" + dtsId + "] (nolock)");
                    int total = rec.Fields.Item("total").Value;
                    if(total == 0)
                    {

                        foreach(Dictionary<string, dynamic> defRow in dts.defRows)
                        {
                            if(dts.UDO != null)
                            {
                                string m = this.getNextCode(dtsId);
                                if(!defRow.ContainsKey("Name") || String.IsNullOrEmpty(defRow["Name"]))
                                {
                                    defRow["Name"] = m;
                                }
                                if(!defRow.ContainsKey("Code") || String.IsNullOrEmpty(defRow["Code"]))
                                {
                                    defRow["Code"] = m;
                                }

                                this.udoInsert(dtsId, defRow, out new_code);

                            } else
                            {
                                this.dtsInsert(dtsId, defRow);
                            }
                        }
                    }
                } catch(Exception e)
                {

                } 

                if (rec != null){
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
                    rec = null;
                }
            }

            #endregion
            
            GC.Collect();
        }

        /// <summary>
        /// Recupera o indice de um key (index) no SAP baseado em seu ID
        /// </summary>
        /// <param name="tbId">Nome da tabela COM "@", se houver</param>
        /// <param name="keyId">Id do index</param>
        /// <returns>Indice do key na tabela ou -1 se não encontrar</returns>
        public int getKeyIndex(string tbId, string keyId) {
            return this.getByIndex(tbId, keyId, false);
        }

        /// <summary>
        /// Recupera o indice de um field no SAP baseado em seu ID
        /// </summary>
        /// <param name="tbId">Nome da tabela COM "@", se houver</param>
        /// <param name="fldId">Id do field SEM o "U_"</param>
        /// <returns>Indice do field na tabela ou -1 se não encontrar</returns>
        public int getFieldIndex(string tbId, string fldId) {
            return this.getByIndex(tbId, fldId, true);
        }

        /// <summary>
        /// Recupera o indice de um elemento, field ou index, na tabela
        /// </summary>
        /// <param name="tbId"></param>
        /// <param name="fldId"></param>
        /// <param name="is_field"></param>
        /// <returns></returns>
        private int getByIndex(string tbId, string fldId, bool is_field) {
            int res = -1;
            SAPbobsCOM.Recordset rs = null;
            try {
                rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sql = (is_field ?
                    "select FieldID as res from cufd where TableId = '" + tbId + "' and AliasID = '" + fldId + "'" :
                    "select KeyId as res   from oukd where TableName = '" + tbId + "' and KeyName = '" + fldId + "'"
                );

                rs.DoQuery(sql);
                if (rs.RecordCount > 0) {
                    res = rs.Fields.Item("res").Value;
                }

                // Verifica erro
            } catch (Exception e) {
                this.Addon.StatusErro(((System.Reflection.MethodBase)e.TargetSite).Name + ": " + e.Message);

                // libera o objeto
            } finally {
                if (rs != null) {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                }
                GC.Collect();
            }

            // Retorna
            return res;
        }

        /// <summary>
        /// Acrescanta o campo interno de versionamento de tabelas
        /// </summary>
        /// <param name="tbId">Nome da tabela</param>
        private void addFieldVersao(string tbId) {
            fieldParams f = new fieldParams() {
                descricao = "Versão da tabela",
                tipo = SAPbobsCOM.BoFieldTypes.db_Numeric
            };
            this.addField(tbId, "tbl_vers", f);
        }

        /// <summary>
        /// Cria um field em uma tabela
        /// </summary>
        /// <param name="tblId"></param>
        /// <param name="fldId"></param>
        /// <param name="fldParams"></param>
        public void addField(string tblId, string fldId, fieldParams fldParams) {
            this._addField(tblId, fldId, fldParams, null, true);
        }

        public void addField(string tblId, string fldId, fieldParams fldParams, string values = null) {
            this._addField(tblId, fldId, fldParams, values, true);
        }

        public void addField(string tblId, string fldId, fieldParams fldParams, Dictionary<string, string> values = null) {
            this._addField(tblId, fldId, fldParams, values, false);
        }

        private void _addField(string tblId, string fldId, fieldParams fldParams, dynamic values = null, bool sql_values = true) {
            int res; string msg;

            if(fldId.Length > 18)
            {
                if(fldId.Substring(0, 4) != "ZTH_")
                {
                    throw new Exception("Nome do field '" + fldId + "' em '" + tblId + "' possui mais de 8 caracteres.");
                }
            }

            // Libera o recordset padrão:
            this.ReleaseBrowser();

            SAPbobsCOM.UserFieldsMD field = null;
            try {
                field = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                field.TableName = tblId;
                field.Name = fldId;
                field.Description = fldParams.descricao;
                field.Type = fldParams.tipo;
                field.SubType = fldParams.subtipo;

                if(fldParams.size > 0)
                {
                    field.EditSize = fldParams.size;
                }

                if (!String.IsNullOrEmpty(fldParams.LinkedTable)) {
                    field.LinkedTable = fldParams.LinkedTable;
                }

                if (values != null) {
                    if (sql_values) {
                        SAPbobsCOM.Recordset rec = null;
                        try {
                            rec = (SAPbobsCOM.Recordset)this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            rec.DoQuery(values);
                            rec.MoveFirst();
                            while (!rec.EoF) {
                                try {
                                    field.ValidValues.Description = rec.Fields.Item(1).Value;
                                    field.ValidValues.Value = rec.Fields.Item(0).Value;
                                    field.ValidValues.Add();

                                } catch (Exception e) {
                                    this.Addon.DesenvTimeError(e, "em addField " + fldId);
                                }
                                rec.MoveNext();
                            }

                        } catch (Exception e) {
                          
                        } finally {
                            if (rec != null) {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
                            }
                            GC.Collect();
                        }

                    } else {
                        foreach (KeyValuePair<string, string> val in values) {
                            field.ValidValues.Description = val.Value;
                            field.ValidValues.Value = val.Key;
                            field.ValidValues.Add();
                        }
                    }
                }

                res = field.Add();

                // Check for errors
                if (res != 0) {
                    this.Addon.oCompany.GetLastError(out res, out msg);
                    if (this.Addon.VerboseMode) {
                        this.Addon.ShowMessage(msg);
                    }
                }

            } catch (Exception e) {
                this.Addon.DesenvTimeError(e, "Acrescentando o field " + fldId + " em " + tblId);
                
            // libera o objeto
            } finally {
                if (field != null) {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(field);
                }
                GC.Collect();
            }
        }

        /// <summary>
        /// Remove um field de uma tabela
        /// </summary>
        /// <param name="tblId"></param>
        /// <param name="fldId"></param>
        /// <param name="fldParams"></param>
        public void removeField(string tblId, string fldId) {
            int res; string msg;

            // Libera o recordset padrão:
            this.ReleaseBrowser();

            SAPbobsCOM.UserFieldsMD field = null;
            try {
                if (fldId.Substring(0, 2) == "U_") {
                    fldId = fldId.Substring(2);
                }
                field = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                field.GetByKey(tblId, this.getFieldIndex(tblId, fldId));
                res = field.Remove();

                // Check for errors
                if (res != 0) {
                    this.Addon.oCompany.GetLastError(out res, out msg);
                    //if (this.Addon.ShowDesenvMsgs) {
                        this.Addon.ShowMessage(msg);
                    //}
                }

            } catch (Exception e) {
                this.Addon.DesenvTimeError(e, " em removeField " + fldId);

                // libera o objeto
            } finally {
                if (field != null) {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(field);
                }
                GC.Collect();
            }
        }

        /// <summary>
        /// Remove um field de uma tabela
        /// </summary>
        /// <param name="tblId"></param>
        /// <param name="fldId"></param>
        /// <param name="fldParams"></param>
        public void removeFields(string tblId)
        {
            int res; string msg; string fldId = "";

            // Libera o recordset padrão:
            this.ReleaseBrowser();

            SAPbobsCOM.UserFieldsMD field = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            SAPbobsCOM.Field fld = null;
            SAPbobsCOM.UserTable table = this.Addon.oCompany.UserTables.Item(tblId); // .GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
            try
            {
                table.GetByKey(tblId);
                for(int f = table.UserFields.Fields.Count-1; f > 0; f--)
                {
                    fld = table.UserFields.Fields.Item(f);
                    if(fld.Name.Substring(0, 2) == "U_") { 
                        if(field.GetByKey("@" + tblId, fld.FieldID))
                        {
                            res = field.Remove();
                            if(res != 0)
                            {
                                this.Addon.oCompany.GetLastError(out res, out msg);
                                //if (this.Addon.ShowDesenvMsgs) {
                                this.Addon.ShowMessage(msg);
                                //}
                            }
                        }
                    }
                }

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " em removeField " + fldId);

                // libera o objeto
            } finally
            {
                if(field != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(field);
                }
                GC.Collect();
            }
        }

        /// <summary>
        /// Acrescenta um index a tabela
        /// </summary>
        /// <param name="tblId">Tabela alvo</param>
        /// <param name="keyId">Id do indice</param>
        /// <param name="keys">Parametros para criação do indice</param>
        private void addKey(string tblId, string keyId, keyParams keys) {
            int res; string msg;

            // Libera o recordset padrão:
            this.ReleaseBrowser();

            SAPbobsCOM.UserKeysMD dtsKeys = null;
            try {
                dtsKeys = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys);
                dtsKeys.TableName = tblId;
                dtsKeys.KeyName = keyId; 
                bool first = true;
                foreach (string f in keys.fields) {
                    if (!first) {
                        dtsKeys.Elements.Add();
                    }
                    dtsKeys.Elements.ColumnAlias = f;
                    first = false;
                }
                dtsKeys.Unique = keys.unique;
                res = dtsKeys.Add();

                // Check for errors
                if (res != 0) {
                    this.Addon.oCompany.GetLastError(out res, out msg);
                    throw new Exception(msg);
                }

            } catch (Exception e) {
                //this.Addon.DesenvTimeError(e, " - Setando Key '" + keyId + "' em '" + tblId + "'");

                // libera o objeto
            } finally {
                if (dtsKeys != null) {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(dtsKeys);
                }
                GC.Collect();
            }
        }

        /// <summary>
        /// Remove um datasource.
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="tbId"></param>
        public void removeDatasource(string tbId, datasource dts, bool plainSQL = false) 
        {
            int res; string msg;
            this.Addon.SBO_Application.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, true);
            
            // Remove do versionamento
            SAPbobsCOM.Recordset rec = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                rec.DoQuery("DELETE FROM [@ZTH_VERSIONAMENTO] WHERE tabela = '" + tbId + "'");

            } catch(Exception e)
            {
                this.Addon.DesenvTimeError(e, " removendo versionamento de " + tbId);
            } finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
                rec = null;
            }

            SAPbobsCOM.UserTablesMD tables = ((SAPbobsCOM.UserTablesMD)(this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));
            bool sap_table = tables.GetByKey(tbId);
            if(!sap_table)
            {
                rec = null;
                try
                {
                    rec = (SAPbobsCOM.Recordset)this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rec.DoQuery(" IF OBJECT_ID('[@" + tbId + "]', 'U') IS NOT NULL DROP TABLE [@" + tbId + "]");
                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, " removendo a tabela " + tbId);
                } finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
                }
            } else
            {
                // Libera o recordset padrão:
                this.ReleaseBrowser();
                
                if(tables != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(tables);
                    GC.Collect();
                }

                // Remove UDO
                SAPbobsCOM.UserObjectsMD udoMD = ((SAPbobsCOM.UserObjectsMD)(this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));
                if(udoMD.GetByKey(tbId + "O"))
                {
                    res = udoMD.Remove();
                    if(res != 0)
                    {
                        this.Addon.oCompany.GetLastError(out res, out msg);
                        this.Addon.ShowMessage("UDO " + tbId + "O: " + msg);
                    }
                }

                // Libera o objeto
                if(udoMD != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(udoMD);
                }
                GC.Collect();

                // Remove keys
                /*if (dts.keys != null)
                {
                    int i = 0;
                    SAPbobsCOM.UserKeysMD dtsKeys = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys);
                    foreach(KeyValuePair<string, keyParams> key in dts.keys)
                    {
                        try {
                            if(dtsKeys.GetByKey("@" + tbId, i))
                            {
                                res = dtsKeys.Remove();
                                if(res != 0)
                                {
                                    this.Addon.oCompany.GetLastError(out res, out msg);
                                    //if (this.Addon.ShowDesenvMsgs) {
                                    this.Addon.ShowMessage(msg);
                                    //}
                                }
                            }
                        } catch (Exception e){

                        }
                        i++;
                    }
                    if (dtsKeys != null){
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(dtsKeys);
                        GC.Collect();
                    }
                }*/

                // Remove Fields
                /*SAPbobsCOM.UserTable table = this.Addon.oCompany.UserTables.Item(tbId); 
                SAPbobsCOM.UserFieldsMD field = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                try
                {
                    table.GetByKey(tbId);
                    for(int f = table.UserFields.Fields.Count-1; f > 0; f--)
                    {
                        if (field.GetByKey("@" + tbId, f)){
                            res = field.Remove();
                            if(res != 0)
                            {
                                this.Addon.oCompany.GetLastError(out res, out msg);
                                //if (this.Addon.ShowDesenvMsgs) {
                                this.Addon.ShowMessage(msg);
                                //}
                            }
                        }
                    }
                } catch(Exception e)
                {
                    this.Addon.DesenvTimeError(e, " em removeField " + tbId);
                }

                if (field != null){
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(field);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
                    GC.Collect();
                }*/

                // Remove a tabela
                tables = ((SAPbobsCOM.UserTablesMD)(this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));
                tables.GetByKey(tbId);
                res = tables.Remove();
                if(res != 0)
                {
                    this.Addon.oCompany.GetLastError(out res, out msg);
                    this.Addon.ShowMessage(tbId + ": " +  msg);
                }
                
                // Libera o objeto
                if(tables != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(tables);
                }
            }
            GC.Collect();
        }

        /// <summary>
        /// Executa inserções em UDOs.
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="tableId"></param>
        /// <param name="values"></param>
        /// <param name="new_code">Retorna o Code do registro inserido</param>
        /// <returns></returns>
        public bool udoInsert(string tableId, Dictionary<string, dynamic> values, out string new_code)
        {
            return this.udoExec(1, tableId, values, out new_code);
        }

        /// <summary>
        /// Executa updates em UDOs.
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="tableId"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public bool udoUpdate(string tableId, Dictionary<string, dynamic> values)
        {
            string blah = "";
            return this.udoExec(2, tableId, values, out blah);
        }

        /// <summary>
        /// Executa deletes em UDOs.
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="tableId"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public bool udoDelete(string tableId, Dictionary<string, dynamic> values)
        {
            string blah = "";
            return this.udoExec(3, tableId, values, out blah);
        }

        /// <summary>
        /// Apenas seta os valores, sem salvar.
        /// By Labs - 11/2013
        /// </summary>
        /// <param name="tableId"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public bool udoSetValues(string tableId, Dictionary<string, dynamic> values)
        {
            string blah = "";
            return this.udoExec(4, tableId, values, out blah);
        }

        /// <summary>
        /// Executa operações CUD em UDO.
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="op"></param>
        /// <param name="tableId"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        private bool udoExec(int op, string tableId, Dictionary<string, dynamic> values, out string new_code) {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;

            DateTime date;
            string[] format = new string[] { "dd/MM/yyyy HH:mm:ss" };

            new_code = "";
            string origTableId = tableId;
            bool ok = true;
            try {
                if(tableId.Substring(0, 1) == "@")
                {
                    tableId = tableId.Substring(1);
                }

                oCompanyService = this.Addon.oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService(tableId + "O");

                if (op == 3) {
                    oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    foreach (KeyValuePair<string, dynamic> val in values) {
                        try
                        {
                            string valor = Convert.ToString(val.Value);
                            /*if (DateTime.TryParseExact(valor, format, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.NoCurrentDateDefault, out date))
                            {
                                valor = date.ToString("yyyyMMdd");
                            }*/

                            oGeneralParams.SetProperty(val.Key, valor);
                        } catch(Exception e)
                        {
                            this.Addon.DesenvTimeError(e, "udoDelete: key = " + val.Key + " | value = " + val.Value);
                            op = -1;
                        }
                    }

                } else {
                    oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                    foreach (KeyValuePair<string, dynamic> val in values) {
                        try
                        {
                            string valor = Convert.ToString(val.Value);
                            /*if (DateTime.TryParseExact(valor, format, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.NoCurrentDateDefault, out date))
                            {
                                valor = date.ToString("yyyyMMdd");
                            }*/
                            oGeneralData.SetProperty(val.Key, valor);
                            
                        } catch(Exception e)
                        {
                            this.Addon.DesenvTimeError(e, "udoInsert | udoUpdate: key = " + val.Key + " | value = " + val.Value);
                            op = -1;
                        }
                    }
                }

                switch (op) {
                    case 1:
                        new_code = this.getNextCode(origTableId, 5);
                        oGeneralData.SetProperty("Code", new_code);
                        oGeneralParams = oGeneralService.Add(oGeneralData);
                        break;

                    case 2:
                        oGeneralService.Update(oGeneralData);
                        break;

                    case 3:
                        oGeneralService.Delete(oGeneralParams);
                        break;

                    case 4:
                        break;
                }

            } catch (Exception e) {
                ok = false;
                this.Addon.DesenvTimeError(e, "udoExec");

            } finally {
                if (oGeneralData != null) {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralData);
                }
                GC.Collect();
            }

            // Retorna
            return ok;
        }

        /// <summary>
        /// Executa inserções em tabelas filhas em UDOs.
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="masterTable"></param>
        /// <param name="childTable"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public bool udoChildInsert(string masterTable, string childTable, Dictionary<string, dynamic> values, string masterValue, string masterProp = "Code") {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralData oChild = null;
            SAPbobsCOM.GeneralDataCollection oChildren = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;

            bool ok = true;
            try {
                if(masterTable.Substring(0, 1) == "@") { masterTable = masterTable.Substring(1); }
                if(childTable.Substring(0, 1) == "@") { childTable = childTable.Substring(1); }

                oCompanyService = this.Addon.oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService(masterTable + "O");
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty(masterProp, masterValue);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                // Insere na tabela filha
                oChildren = oGeneralData.Child(childTable); 
                oChild = oChildren.Add();
                foreach (KeyValuePair<string, dynamic> val in values) {
                    oChild.SetProperty(val.Key, val.Value);
                }
                oGeneralService.Update(oGeneralData);


            } catch (Exception e) {
                ok = false;
                this.Addon.StatusErro(((System.Reflection.MethodBase)e.TargetSite).Name + ": " + e.Message);

            } finally {
                if (oGeneralData != null) {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralData);
                }
                GC.Collect();
            }

            // Retorna
            return ok;
        }


        /// <summary>
        /// Executa inserções em tabelas filhas em UDOs.
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="masterTable"></param>
        /// <param name="childTable"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public bool udoChildReplace(string masterTable, string masterCode, string childTable, List<Dictionary<string, dynamic>> rows)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralData oChild = null;
            SAPbobsCOM.GeneralDataCollection oChildren = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;

            bool ok = true;
            try
            {
                // Recupera UDO
                oCompanyService = this.Addon.oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService(masterTable + "O");
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("Code", masterCode);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                // Limpa a tabela filha
                oChildren = oGeneralData.Child(childTable);
                for(int c = (oChildren.Count - 1); c >= 0; c--)
                {
                    oChildren.Remove(c);
                }

                foreach(Dictionary<string, dynamic> row in rows)
                {
                    oChild = oChildren.Add();
                    oChild.SetProperty("Code", masterCode);
                    foreach(KeyValuePair<string, dynamic> val in row)
                    {
                        oChild.SetProperty(val.Key, val.Value);
                    }
                }
                oGeneralService.Update(oGeneralData);

            } catch(Exception e)
            {
                ok = false;
                this.Addon.StatusErro(((System.Reflection.MethodBase)e.TargetSite).Name + ": " + e.Message);

            } finally
            {
                if(oGeneralData != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralData);
                }
                GC.Collect();
            }

            // Retorna
            return ok;
        }


        public SAPbobsCOM.GeneralDataParams udoFind(string tableId, string field, object value) {
            SAPbobsCOM.GeneralService oGeneralService;
            SAPbobsCOM.GeneralData oGeneralData;
            SAPbobsCOM.GeneralDataParams oGeneralParams;
            SAPbobsCOM.CompanyService sCmp;

            try {
                sCmp = this.Addon.oCompany.GetCompanyService();
                oGeneralService = sCmp.GetGeneralService(tableId + "O");
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty(field, value);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                GC.Collect();
                return oGeneralParams;

            } catch (Exception e) {
                this.Addon.StatusErro(((System.Reflection.MethodBase)e.TargetSite).Name + ": " + e.Message);
                GC.Collect();
                return null;
            }
        }






        /// <summary>
        /// Executa um insert via usertable.
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="tableId"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public bool dtsInsert(string tableId, Dictionary<string, dynamic> values) {
            return this.dtsExec(1, tableId, values);
        }

        /// <summary>
        /// Executa um update via usertable.
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="tableId"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public bool dtsUpdate(string tableId, Dictionary<string, dynamic> values) {
            return this.dtsExec(2, tableId, values);
        }

        /// <summary>
        /// Executa um delete via usertable.
        /// By Labs - 03/2013
        /// </summary>
        /// <param name="tableId"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public bool dtsDelete(string tableId, Dictionary<string, dynamic> values) {
            return this.dtsExec(3, tableId, values);
        }

        /// <summary>
        /// Executa operaçõe via usertable 
        /// </summary>
        /// <param name="op">1 - Insert | 2 - Update | 3 - Delete</param>
        /// <param name="tableId"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        private bool dtsExec(int op, string tableId, Dictionary<string, dynamic> values) {
            int res = 0; string msg;
            bool ok = true;
            SAPbobsCOM.UserTable table = null;
            string origTableId = tableId;
            if(origTableId[0] != '@')
            {
                origTableId = "@" + origTableId;
            }
            tableId = origTableId.Substring(1);
            
            try {
                table = this.Addon.oCompany.UserTables.Item(tableId);

                string code = "", name = "";
                if (values.ContainsKey("Code")) {
                    code = values["Code"];
                }

                int old_op = op;
                if (String.IsNullOrEmpty(code)) {
                    code = this.getNextCode(origTableId, 4);
                    table.Code = code;
                    op = (op == 2 ? op = 1 : op);
                } else {
                    op = (op == 1 ? op = 2 : op);
                }
                name = code;

                if (op > 1) {
                    if (!table.GetByKey(code)) {
                        this.Addon.StatusAlerta("Registro não encontrado");
                        return false;
                    }
                }
                table.Name = name;


                var teste = table.TableName;

                int fLen = table.UserFields.Fields.Count;
                for (int i = 0; i < fLen; i++) {
                    string field = table.UserFields.Fields.Item(i).Name;
                    if (values.ContainsKey(field)) {
                        if (table.UserFields.Fields.Item(field).Type == SAPbobsCOM.BoFieldTypes.db_Date) {
                            try
                            {
                                DateTime d;
                                if(table.UserFields.Fields.Item(field).SubType == SAPbobsCOM.BoFldSubTypes.st_Time)
                                {
                                    d = this.Addon.fromSAPToTime(values[field]);
                                } else
                                {
                                    try
                                    {
                                        d = Convert.ToDateTime(values[field]);
                                    } catch(Exception e)
                                    {
                                        d = this.Addon.fromSAPToDate(values[field]);
                                    }
                                }
                                table.UserFields.Fields.Item(field).Value = d;
                            } catch(Exception e)
                            {
                                table.UserFields.Fields.Item(field).SetNullValue();
                            }
                        } else if(table.UserFields.Fields.Item(field).Type == SAPbobsCOM.BoFieldTypes.db_Alpha)
                        {
                            table.UserFields.Fields.Item(field).Value = Convert.ToString(values[field]);
                        } else {
                            table.UserFields.Fields.Item(field).Value = values[field];
                        }
                    } else {
                        table.UserFields.Fields.Item(field).SetNullValue();
                    }
                }

                switch (op) {
                    case 1:
                        res = table.Add();
                        break;

                    case 2:
                        res = table.Update();
                        break;

                    case 3:
                        res = table.Remove();
                        break;
                }

                // Check for errors
                if (res != 0) {
                    ok = false;
                    this.Addon.oCompany.GetLastError(out res, out msg);
                    if (this.Addon.VerboseMode) {
                        this.Addon.StatusErro(msg);
                    }
                }
                
                values["Code"] = code;
                values["Name"] = name;

            } catch (Exception e) {
                ok = false;
                this.Addon.StatusErro(((System.Reflection.MethodBase)e.TargetSite).Name + ": " + e.Message);

            } finally {
                if (table != null) {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
                }
                GC.Collect();
            }

            // retorna
            return ok;
        }

        /// <summary>
        /// Salva todas as alterações feitas em um DBDatasource.
        /// Utilizar apenas em tabelas NÃO UDOs.
        /// </summary>
        /// <param name="dtsId"></param>
        /// <param name="frmId"></param>
        /// <returns></returns>
        public bool saveUserDataSource(string dtsId, string frmId = "", int frmCount = 0) {
            bool res = true;
            try {
                SAPbouiCOM.Form form = this.Addon.getForm(frmId, frmCount); 
                SAPbouiCOM.DBDataSource dts = form.DataSources.DBDataSources.Item(dtsId);
                Dictionary<string, dynamic> values = new Dictionary<string, dynamic>();

                int fLen = dts.Fields.Count;
                string field;
                for (int s = 0; s < dts.Size; s++) {
                    dts.Offset = s;
                    for (int f = 0; f < fLen; f++) {
                        field = dts.Fields.Item(f).Name;
                        values[field] = dts.GetValue(field, s).Trim();
                    }
                    if (this.dtsInsert(dtsId, values)) {
                        try {
                            dts.SetValue("Name", s, values["Name"]);
                            dts.SetValue("Code", s, values["Code"]);
                        } catch (Exception e) {
                        }
                    }
                }

            } catch (Exception e) {
                this.Addon.StatusErro(((System.Reflection.MethodBase)e.TargetSite).Name + ": " + e.Message);
                res = false;
            }
            return res;
        }
        
        public bool saveUserDataSource(string dtsId, Forms FormClass)
        {
            bool res = true;
            try
            {
                SAPbouiCOM.Form form = FormClass.SapForm;
                SAPbouiCOM.DBDataSource dts = form.DataSources.DBDataSources.Item(dtsId);
                Dictionary<string, dynamic> values = new Dictionary<string, dynamic>();

                int fLen = dts.Fields.Count;
                string field;
                for(int s = 0; s < dts.Size; s++)
                {
                    dts.Offset = s;
                    for(int f = 0; f < fLen; f++)
                    {
                        field = dts.Fields.Item(f).Name;
                        values[field] = dts.GetValue(field, s).Trim();
                    }
                    if(this.dtsInsert(dtsId, values))
                    {
                        try
                        {
                            dts.SetValue("Name", s, values["Name"]);
                            dts.SetValue("Code", s, values["Code"]);
                        } catch(Exception e)
                        {
                        }
                    }
                }

            } catch(Exception e)
            {
                this.Addon.StatusErro(((System.Reflection.MethodBase)e.TargetSite).Name + ": " + e.Message);
                res = false;
            }
            return res;
        }

        /// <summary>
        /// Executa SQL.
        /// By Labs - 01/2013
        /// </summary>
        /// <param name="sql"></param>
        public bool execSql(string sql) {
            bool res = false;
            SAPbobsCOM.Recordset rs = null;
            try {
                rs = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(sql);
                res = true;

                // Verifica erro
            } catch (Exception e) {
                this.Addon.StatusErro(((System.Reflection.MethodBase)e.TargetSite).Name + ": " + e.Message);

                // libera o objeto
            } finally {
                if (rs != null) {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                }
                GC.Collect();
            }
            return res;
        }

        /// <summary>
        /// Executa um SQL e alimenta this.Addon.Browser.
        /// </summary>
        /// <param name="sql"></param>
        public bool Select(string sql, bool quiet = false) {
            bool ok = false;
            try {
                if(this.Addon.Browser == null)
                {
                    this.Addon.Browser = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                }
                this.Addon.Browser.DoQuery(sql);
                ok = true;

            // Verifica erro
            } catch (Exception e) {
                if(!quiet)
                {
                    this.Addon.StatusErro(((System.Reflection.MethodBase)e.TargetSite).Name + ": " + e.Message);
                } else
                {
                    throw;
                }
            }

            return ok;
        }
        
        /// <summary>
        /// Recupera o próximo valor utilizavel de "Code" em uma tabela.
        /// NÃO É UM PROCEDIMENTO INTEIRAMENTE SEGURO!!!
        /// </summary>
        /// <param name="table">Com o "@" se houver</param>
        /// <param name="padding">Número de padding</param>
        /// <param name="padchar">char para padding</param>
        /// <returns></returns>
        public string getNextCode(string table, int padding=0, char padchar='0')
        {
            int i = Convert.ToInt32(this.getMaxCode(table));
            return (++i).ToString().PadLeft(padding, padchar);
        }

        /// <summary>
        /// Recupera o próximo valor da coluna IDENTITY em uma tabela.
        /// NÃO É UM PROCEDIMENTO INTEIRAMENTE SEGURO!!!
        /// </summary>
        /// <param name="table"></param>
        /// <param name="padding"></param>
        /// <param name="padchar"></param>
        /// <returns></returns>
        public string getNextIdentityCode(string table, int padding = 0, char padchar = '0')
        {
            int i = Convert.ToInt32(this.getMaxCode(table, identity: true));
            return i.ToString().PadLeft(padding, padchar);
        }

        /// <summary>
        /// Recupera o próximo valor utilizavel de "Code" em uma tabela.
        /// NÃO É UM PROCEDIMENTO INTEIRAMENTE SEGURO!!!
        /// </summary>
        /// <param name="table">Com o "@" se houver</param>
        /// <returns></returns>
        public string getMaxCode(string table, int padding = 0, char padchar = '0', bool identity = false)
        {
            string res = "1"; int i;
            SAPbobsCOM.Recordset rec = null;
            try {
                rec = this.Addon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sql = "SELECT MAX(Convert(int, Code)) as next FROM [" + table + "] WITH (NOLOCK)";
                if(identity)
                {
                    sql = "SELECT (IDENT_CURRENT('" + table + "') + 1 ) as next";
                }
                rec.DoQuery(sql);

                res = rec.Fields.Item("next").Value + "";
                //res = i.ToString();

            } catch (Exception e) {
                this.Addon.StatusErro(((System.Reflection.MethodBase)e.TargetSite).Name + ": " + e.Message);
            } finally {
                if (rec != null) {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
                }
                GC.Collect();
            }
            return res.PadLeft(padding, padchar);
        }


    }

}
