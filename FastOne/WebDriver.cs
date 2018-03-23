using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Specialized;
using System.Net;
using System.IO;
using System.Reflection;
using System.Xml;

namespace TShark
{

    /* Custom function
     *  [
     *      "exec":"finalizarMovimentacao", 
     *       "from":["movimentacoes", "movimentacoes"],
     *       "sendToServer":{
     *           "values":{
     *              "finalizacao":true
     *           },
     *           "rowValuesFrom":[
     *              "dtsMovimentacoes"
     *           ],
     *           "fromDataset":{
     *               "dtsMovimentacoes":[
     *                   "movimentacoes_key"
     *               ]
     *           }
     *       }
     *  ]
     */
    public class call
    {
        public string exec;
        public string callback;
        public List<string> from;
        public Dictionary<string, dynamic> data; // sendToServer

        public call()
        {
            this.data = new Dictionary<string, dynamic>();
        }

        public void sendToServer(string key, dynamic value)
        {
            this.data[key] = value;
        }
    }

    public static class HtmlRemoval
    {
        /// <summary>
        /// Remove HTML from string with Regex.
        /// </summary>
        public static string StripTagsRegex(string source)
        {
            return Regex.Replace(source, "<.*?>", string.Empty);
        }

        /// <summary>
        /// Compiled regular expression for performance.
        /// </summary>
        static Regex _htmlRegex = new Regex("<.*?>", RegexOptions.Compiled);

        /// <summary>
        /// Remove HTML from string with compiled Regex.
        /// </summary>
        public static string StripTagsRegexCompiled(string source)
        {
            return _htmlRegex.Replace(source, string.Empty);
        }

        /// <summary>
        /// Remove HTML tags from string using char array.
        /// </summary>
        public static string StripTagsCharArray(string source)
        {
            char[] array = new char[source.Length];
            int arrayIndex = 0;
            bool inside = false;

            for(int i = 0; i < source.Length; i++)
            {
                char let = source[i];
                if(let == '<')
                {
                    inside = true;
                    continue;
                }
                if(let == '>')
                {
                    inside = false;
                    continue;
                }
                if(!inside)
                {
                    array[arrayIndex] = let;
                    arrayIndex++;
                }
            }
            return new string(array, 0, arrayIndex);
        }
    }

    /* Post
     *  [
     *      {
     *          "exec":"post",
     *          "from":["fiscal","ztf_bloco0_r0000"],
     *          "post":{
     *              "ins":{
     *                  "provider":"provZtfBloco0R0000Upd",
     *                  "row":{
     *                      "ztf_bloco0_r0000_key":"NEW_KEY",
     *                      "nome":"asdasd",
     *                      "uf":"as",
     *                      "cnpj":"asaa",
     *                      "cpf":"asd",
     *                      "ie":"",
     *                      "cod_mun":"asasd",
     *                      "im":"",
     *                      "suframa":"",
     *                      "ind_perfil":"",
     *                      "ind_ativ":"0"
     *                  },
     *                  "eventos":{
     *                      "onAfterUpdate":"nomeDeUmaFuncaoExistenteNaLib",
     *                      "onBeforeDelete":["package","modulo","nomeDeUmaFuncaoExistenteNoModulo"],
     *                  }
     *              },
     *	            "upd":{
     *	                "provider":"provZtfBloco0R0000Upd",
     *	                "row":{
     *	                    "ztf_bloco0_r0000_key":"2",
     *	                    "nome":"labs C",
     *	                    "uf":"yt",
     *	                    "cnpj":"76547865",
     *	                    "cpf":"87658765",
     *	                    "ie":"jhg123",
     *	                    "cod_mun":"",
     *	                    "im":"",
     *	                    "suframa":"sss",
     *	                    "ind_perfil":"",
     *	                    "ind_ativ":"0"
     *	                },
     *	                "eventos":{},
     *	                "changed":{
     *	                    "ie":{"_new":"jhg123","_old":"jhg"},
     *	                    "suframa":{"_new":"sss","_old":""}
     *	                }
     *              },
     *              "del":{
     *                  "provider":"provZtfBloco0R0000Upd",
     *                  "row":{
     *                      "ztf_bloco0_r0000_key":"2",
     *                      "nome":"labs C",
     *                      "uf":"yt",
     *                      "cnpj":"76547865",
     *                      "cpf":"87658765",
     *                      "ie":"jhg",
     *                      "cod_mun":"",
     *                      "im":"",
     *                      "suframa":"",
     *                      "ind_perfil":"",
     *                      "ind_ativ":"0"
     *                  },
     *                  "eventos":{}
     *              }
     *          }
     *      }
     *  ]      
     */
    public class post_row
    {
        public string provider;
        public Dictionary<string, string> row;
        public Dictionary<string, string> eventos;
    }

    public class post_data
    {
        public post_row ins;
        public post_row upd;
        public post_row del;

        public post_data()
        {
            this.ins = new post_row();
            this.upd = new post_row();
            this.del = new post_row();
        }
    }

    public class post : call
    {
        public post_data data;

        public post()
        {
            this.exec = "post";
            this.data = new post_data();
        }
    }

    class SQLOverwrite
    {
        public string fields;
        public string join;
        public string where;
        public string order;
        public string key;
        public bool overwrite_fields = true;
        public bool overwrite_join = true;
        public bool overwrite_where = true;
        public bool overwrite_order = true;

        public List<Dictionary<string, dynamic>> toServer;

        public SQLOverwrite()
        {
            this.toServer = new List<Dictionary<string, dynamic>>();
        }
    }


    public class WebDriver
    {
        public string host;
        public string solution;
        public string app;
        public string client;
        public string user;
        public string pwd;
        public List<Dictionary<string, dynamic>> call_result;
        public FastOne Addon;
        public Type refAppType;
        private Dictionary<string, MethodInfo> CallbackCache;

        /// <summary>
        /// Inicializa.
        /// </summary>
        /// <param name="addon"></param>
        public WebDriver(FastOne addon)
        {
            this.Addon = addon;
            this.refAppType = addon.GetType();
            this.CallbackCache = new Dictionary<string, MethodInfo>();
        }

        /// <summary>
        /// Seta a configuração de acesso através do addon.xml
        /// </summary>
        public bool SetupByXML()
        {
            bool res = false;
            XmlNode webdriver = this.Addon.Xml.Config.SelectSingleNode("webdriver");
            if(webdriver != null)
            {
                this.host = webdriver["host"].Attributes["value"].Value;
                this.user = webdriver["user"].Attributes["value"].Value;
                this.pwd = webdriver["pwd"].Attributes["value"].Value;
                res = (!String.IsNullOrEmpty(this.host)
                    && !String.IsNullOrEmpty(this.user)
                    && !String.IsNullOrEmpty(this.pwd)
                );
            }
            return res;
        }

        /// <summary>
        /// Executa um post para o server.
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public string ping()
        {
            Object[] s = new Object[1] { new post (){
                exec = "ping"
            }};
            string json = fastJSON.JSON.Instance.ToJSON(s);
            return this._call(json);
        }

        /// <summary>
        /// Executa um post para o server.
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public string Post(ref post data)
        {
            Object[] s = new Object[1] { data };
            string json = fastJSON.JSON.Instance.ToJSON(s);
            return this._call(json);
        }

        /// <summary>
        /// Executa uma função específica em uma lib no server.
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public string Call(ref call data)
        {
            Object[] s = new Object[1] { data };
            string json = fastJSON.JSON.Instance.ToJSON(s);
            return this._call(json);
        }

        public string CallSync(ref call data)
        {
            Object[] s = new Object[1] { data };
            string json = fastJSON.JSON.Instance.ToJSON(s);
            return this._call(json, true);
        }



        /// <summary>
        /// Executa a chamada ao server.
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private string _call(string data, bool sync = false)
        {
            string result = null;
            System.Uri url = new Uri(this.host); // + "/" + this.solution + "/" + this.app + "/" + this.client + "/init.php");

            try
            {

                // Conexão:
                WebClient client = new WebClient();
                client.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
                client.UploadValuesCompleted += new UploadValuesCompletedEventHandler(_callback);
                client.Headers.Add(HttpRequestHeader.Cookie, "PHPSESSID=SAP_SESSION");
                
                NameValueCollection pack = new NameValueCollection();
                pack.Add("packid", Guid.NewGuid().ToString("N"));
                pack.Add("call", data);

                // Chamada:
                if(sync)
                {
                    client.UploadValues(url, "POST", pack);
                } else
                {
                    client.UploadValuesAsync(url, "POST", pack);
                }

            } catch(Exception e)
            {
                result = "Exceção na chamada ao server: " + e.Message;
            }

            // Retorna:
            return result;
        }

        /// <summary>
        /// Retorna uma instancia de funcao de callback bem como 
        /// gerencia o cache de callbacks.
        /// </summary>
        /// <param name="callback"></param>
        /// <returns></returns>
        private MethodInfo _getCallback(string callback)
        {
            if(!this.CallbackCache.ContainsKey(callback))
            {
                this.CallbackCache[callback] = this.refAppType.GetMethod(callback);
            }

            return this.CallbackCache[callback];
        }

        /// <summary>
        /// Gateway de callbacks
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void _callback(object sender, UploadValuesCompletedEventArgs e)
        {

            // Retorno:
            try
            {
                string res = Encoding.ASCII.GetString(e.Result);
                if(String.IsNullOrEmpty(res))
                {
                    this._getCallback("StatusErro").Invoke(this.Addon, new object[] { "O servidor retornou uma resposta vazia" });

                } else
                {

                    // Tenta o parse:
                    dynamic result = null;
                    try
                    {
                        result = fastJSON.JSON.Instance.ToObject<dynamic>(res);

                    // Se falha, assumimos que voltou um string qualquer não JSON, então exibimos:
                    } catch(Exception ex)
                    {
                        result = null;
                        this._getCallback("ShowMessage").Invoke(this.Addon, new object[] { "Erro no servidor: " + HtmlRemoval.StripTagsRegexCompiled(res) });
                    }

                    try
                    {
                        if(result != null)
                        {

                            foreach(dynamic ret in result)
                            {

                                // Callback:
                                if(ret.ContainsKey("exec") && (ret["exec"] != "setGlobals"))
                                {
                                    string callback = (ret.ContainsKey("callback") ? ret["callback"] : "");
                                    if(String.IsNullOrEmpty(callback))
                                    {
                                        callback = ret["exec"] + "_Callback";
                                    }
                                    MethodInfo mi = this._getCallback(callback);
                                    if(null != mi)
                                    {
                                        mi.Invoke(this.Addon, new object[] { ret });
                                    } else
                                    {
                                        this._getCallback("ShowMessage").Invoke(this.Addon, new object[] { "Callback nao encontrado: " + callback });
                                    }
                                }
                            }
                        }

                    } catch(Exception ex)
                    {
                        string msg = (ex.InnerException == null ? ex.Message : ex.InnerException.Message);
                        this._getCallback("ShowMessage").Invoke(this.Addon, new object[] { "Erro de callback: " + msg });
                    }
                }

            } catch(Exception ex)
            {
                string msg = (ex.InnerException == null ? ex.Message : ex.InnerException.Message);
                this._getCallback("ShowMessage").Invoke(this.Addon, new object[] { "Erro de callback: " + msg });
            }
        }


        #region Mapeamento de objectos DI
        public string diObjectMapGetSource(string interopPath, string intfTypeToMap)
        {
            string res = "";
            string fields = "";
            string update = "\r\n                     ";
            int n = 0;

            // Mapeamento de classe 
            try
            {
                System.Reflection.Assembly oAsm = System.Reflection.Assembly.LoadFrom(interopPath + "\\Interop.SAPbobsCOM.dll");
                Type oType = oAsm.GetType("SAPbobsCOM." + intfTypeToMap);

                System.Reflection.PropertyInfo[] oProps = oType.GetProperties();
                foreach(System.Reflection.PropertyInfo oProp in oProps)
                {
                    string fld = oProp.Name;
                    string Type = oProp.PropertyType.Name;
                    string lbl = fld;
                    string comp = "inText";
                    string size = "100";

                    switch(oProp.PropertyType.Name)
                    {
                        case "Int32":
                            Type = "int";
                            comp = "inpInt";
                            break;

                        case "Double":
                            Type = "float";
                            comp = "inpFloat";
                            break;

                        case "DateTime":
                            Type = "datetime";
                            comp = "inpDateTime";
                            break;

                        case "Date":
                            Type = "date";
                            comp = "inpDate";
                            break;

                        case "Time":
                            Type = "time";
                            comp = "inpTime";
                            break;

                        case "String":
                            Type = "string";
                            comp = "inpText";
                            break;

                        case "BoYesNoEnum":
                            Type = "int";
                            comp = "selectSimNao";
                            break;

                        default:
                            Type = "string";
                            lbl = fld + "(" + oProp.PropertyType.Name + ")";
                            break;
                    }

                    if(oProp.PropertyType.IsPublic)
                    {
                        fields += "\r\n                        '" + fld + "' => array(";
                        fields += "\r\n                            0, " + size + ", \"" + Type + "\", \"\", \"" + comp + "\", \"" + lbl + ":\"";
                        fields += "\r\n                        ),";
                    }

                    if(oProp.CanWrite)
                    {
                        update += "\"" + fld + "\", ";
                        n++;
                        if(n == 3)
                        {
                            update += "\r\n                     ";
                            n = 0;
                        }
                    }
                }

            } catch(Exception e)
            {

            }

            // Monta
            res = "\r\n            'tipo' => \"sapdi\",";
            res += "\r\n            'base' => \"SAP_DI\",";
            res += "\r\n            'table' => \"\",";
            res += "\r\n            'object_Type' => \"\",";
            res += "\r\n            'business_object' => \"" + intfTypeToMap.Substring(1) + "\",";
            res += "\r\n            'metadata' => array(";
            res += "\r\n                'key' => \"\",";
            res += "\r\n                'fields' => array(";
            res += fields;
            res += "\r\n                            ),";
            res += "\r\n            ),";
            res += "\r\n            'updates' => array(";
            res += "\r\n                'default' => array(";
            res += update;
            res += "\r\n                            ),";
            res += "\r\n            ),";

            // Retorna
            return res;
        }

        public string diObjectMapUpdate(string interopPath, string intfTypeToMap, string tableToMap, ref SAPbobsCOM.SBObob sbo)
        {
            string res = "";
            string fields = "";
            string update = "\r\n                     ";
            int n = 0;

            // Mapeamento de classe
            try
            {

                // Pega a lista de propriedades do objeto:
                System.Reflection.Assembly oAsm = System.Reflection.Assembly.LoadFrom(interopPath + "\\Interop.SAPbobsCOM.dll");
                Type oType = oAsm.GetType("SAPbobsCOM." + intfTypeToMap);
                System.Reflection.PropertyInfo[] oProps = oType.GetProperties();
                var props = oProps.OrderBy(item => item.PropertyType.Name).ToArray();

                // Pega os fields disponíveis:
                SAPbobsCOM.Recordset rs = sbo.GetTableFieldList(tableToMap);


                foreach(System.Reflection.PropertyInfo oProp in oProps)
                {
                    string fld = oProp.Name;
                    string Type = oProp.PropertyType.Name;
                    string lbl = fld;
                    string comp = "inText";
                    string size = "100";

                    switch(oProp.PropertyType.Name)
                    {
                        case "Int32":
                            Type = "int";
                            comp = "inpInt";
                            break;

                        case "Double":
                            Type = "float";
                            comp = "inpFloat";
                            break;

                        case "DateTime":
                            Type = "datetime";
                            comp = "inpDateTime";
                            break;

                        case "Date":
                            Type = "date";
                            comp = "inpDate";
                            break;

                        case "Time":
                            Type = "time";
                            comp = "inpTime";
                            break;

                        case "String":
                            Type = "string";
                            comp = "inpText";
                            break;

                        case "BoYesNoEnum":
                            Type = "int";
                            comp = "selectSimNao";
                            break;

                        default:
                            Type = "string";
                            lbl = fld + "(" + oProp.PropertyType.Name + ")";
                            break;
                    }

                    if(oProp.PropertyType.IsPublic)
                    {
                        fields += "\r\n                        '" + fld + "' => array(";
                        fields += "\r\n                            0, " + size + ", \"" + Type + "\", \"\", \"" + comp + "\", \"" + lbl + ":\"";
                        fields += "\r\n                        ),";
                    }

                    if(oProp.CanWrite)
                    {
                        update += "\"" + fld + "\", ";
                        n++;
                        if(n == 3)
                        {
                            update += "\r\n                     ";
                            n = 0;
                        }
                    }
                }

            } catch(Exception e)
            {

            }

            // Monta
            res = "\r\n            'tipo' => \"sapdi\",";
            res += "\r\n            'base' => \"SAP_DI\",";
            res += "\r\n            'table' => \"\",";
            res += "\r\n            'object_Type' => \"\",";
            res += "\r\n            'business_object' => \"" + intfTypeToMap.Substring(1) + "\",";
            res += "\r\n            'metadata' => array(";
            res += "\r\n                'key' => \"\",";
            res += "\r\n                'fields' => array(";
            res += fields;
            res += "\r\n                            ),";
            res += "\r\n            ),";
            res += "\r\n            'updates' => array(";
            res += "\r\n                'default' => array(";
            res += update;
            res += "\r\n                            ),";
            res += "\r\n            ),";

            // Retorna
            return res;
        }

        #endregion

    }

}
