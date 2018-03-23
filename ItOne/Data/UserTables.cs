using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TShark;

namespace ITOne
{
    class UserTables : TShark.UserTables
    {

        public UserTables() : base() { }


        #region :: Configurações

        /// <summary>
        /// Tabela de Funções/Cargos
        /// </summary>
        /// <returns></returns>
        public datasource UPD_IT_FUNCOES()
        {
            return new datasource()
            {
                id = "UPD_IT_FUNCOES",
                descricao = "Funções",
                remove_if_exists = false,
                tipo = SAPbobsCOM.BoUTBTableType.bott_MasterData,
                UDO = new UDO()
                {
                    TableName = "UPD_IT_FUNCOES",
                    ChildTables = new List<string>() { 
                        "UPD_IT_METAS",
                        "UPD_IT_COMISSAO",
                    },
                },
                versao = 3,
                versoes = new Dictionary<int, dtsVersionamento>(){
                    {3, new dtsVersionamento(){
                      fields = new fieldsVersionamento(){
                        novos = new Dictionary<string,fieldParams>(){
                               {"comissao", new fieldParams(){
                                    descricao = "Comissão Base",
                                    tipo = SAPbobsCOM.BoFieldTypes.db_Float,
                                    subtipo = SAPbobsCOM.BoFldSubTypes.st_Price
                                }}, 
                            }
                        }  
                    }}
                },
                fields = new Dictionary<string, fieldParams>(){
                    
                    {"nome", new fieldParams(){
                        descricao = "Nome",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                        size = 200
                    }},
                    {"ativo", new fieldParams(){
                        descricao = "Ativo",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                        size = 3
                    }},
                    {"mapfield", new fieldParams(){
                        descricao = "Coluna mapeada",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                        size = 50
                    }},
                    {"categ", new fieldParams(){
                        descricao = "Categoria",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                        size = 5
                    }},
                    {"percent", new fieldParams(){
                        descricao = "Percent Ata",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Float,
                        subtipo = SAPbobsCOM.BoFldSubTypes.st_Percentage,
                    }},
                    {"obs", new fieldParams(){
                        descricao = "Observações",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Memo,
                    }},
                },
            };
        }

        /// <summary>
        /// Tabela de Metas de cada Função
        /// </summary>
        /// <returns></returns>
        public datasource UPD_IT_METAS()
        {
            return new datasource()
            {
                id = "UPD_IT_METAS",
                descricao = "Metas de Funções",
                remove_if_exists = false,
                tipo = SAPbobsCOM.BoUTBTableType.bott_MasterDataLines,
                versao = 3,
                versoes = new Dictionary<int,dtsVersionamento>(){
                    {3, new dtsVersionamento(){
                      fields = new fieldsVersionamento(){
                        novos = new Dictionary<string,fieldParams>(){
                               {"acelera", new fieldParams(){
                                    descricao = "Tipo Aceleração",
                                    tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                                    size = 5
                                }}, 
                            }
                        }  
                    }}
                },
                fields = new Dictionary<string, fieldParams>(){
                    
                    {"nome", new fieldParams(){
                        descricao = "Nome",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                        size = 200
                    }},
                    {"empid", new fieldParams(){
                        descricao = "EmpID Colaborador",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                    }},
                    {"dtinicio", new fieldParams(){
                        descricao = "Data Inicial",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Date,
                    }},
                    {"dtfim", new fieldParams(){
                        descricao = "Data Fim",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Date,
                    }},
                    {"meta", new fieldParams(){
                        descricao = "Meta",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Float,
                        subtipo = SAPbobsCOM.BoFldSubTypes.st_Price
                    }},
                },
            };
        }

        /// <summary>
        /// Tabela de Comissão para cada Função
        /// </summary>
        /// <returns></returns>
        public datasource UPD_IT_COMISSAO()
        {
            return new datasource()
            {
                id = "UPD_IT_COMISSAO",
                descricao = "Comissao por Função",
                remove_if_exists = false,
                tipo = SAPbobsCOM.BoUTBTableType.bott_MasterDataLines,
                versao = 3,
                versoes = new Dictionary<int, dtsVersionamento>(){
                    {3, new dtsVersionamento(){
                      fields = new fieldsVersionamento(){
                        novos = new Dictionary<string,fieldParams>(){
                               {"overcota", new fieldParams(){
                                    descricao = "OVer Cota",
                                    tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                                    size = 3
                                }}, 
                            }
                        }
                    }}
                },
                fields = new Dictionary<string, fieldParams>(){
                    {"teto", new fieldParams(){
                        descricao = "Teto Limite maximo",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                    }},
                    {"piso", new fieldParams(){
                        descricao = "Piso Limite minimo",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                    }},
                    {"comissao", new fieldParams(){
                        descricao = "Percentual Comissao",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Float,
                        subtipo = SAPbobsCOM.BoFldSubTypes.st_Percentage
                    }},
                },
            };
        }

        /*
        /// <summary>
        /// Tabela de Funções/Cargos
        /// </summary>
        /// <returns></returns>
        public datasource UPD_IT_CONFIG_ATA()
        {
            return new datasource()
            {
                id = "UPD_IT_CONFIG_ATA",
                descricao = "Config. Ata",
                remove_if_exists = false,
                tipo = SAPbobsCOM.BoUTBTableType.bott_MasterData,
                UDO = new UDO()
                {
                    TableName = "UPD_IT_CONFIG_ATA",
                },
                fields = new Dictionary<string, fieldParams>(){
                    
                    {"desc", new fieldParams(){
                        descricao = "Descricao",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                        size = 200
                    }},
                    {"percent", new fieldParams(){
                        descricao = "Percentual",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Float,
                        subtipo = SAPbobsCOM.BoFldSubTypes.st_Percentage,
                    }},
                    {"obs", new fieldParams(){
                        descricao = "Observações",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Memo,
                    }},
                },
            };
        }
        */

        #endregion


        #region :: Tabela relacionadas com ORDR

        /// <summary>
        /// Participantes e suas Funções na Venda
        /// </summary>
        /// <returns></returns>
        public datasource UPD_IT_PARTICIP()
        {
            return new datasource()
            {
                id = "UPD_IT_PARTICIP",
                descricao = "Participantes Venda",
                remove_if_exists = false,
                tipo = SAPbobsCOM.BoUTBTableType.bott_NoObject,
                fields = new Dictionary<string, fieldParams>(){
                    
                    {"docentry", new fieldParams(){
                        descricao = "DocEntry ORDR",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                    }},
                    {"funcao", new fieldParams(){
                        descricao = "Função",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                        size = 30
                    }},
                    {"empid", new fieldParams(){
                        descricao = "EmpID Colaborador",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Numeric,
                    }},
                    {"vlcom", new fieldParams(){
                        descricao = "Valor Comissão",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Float,
                        subtipo = SAPbobsCOM.BoFldSubTypes.st_Price
                    }},
                    {"percom", new fieldParams(){
                        descricao = "Percent Comissão",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Float,
                        subtipo = SAPbobsCOM.BoFldSubTypes.st_Percentage
                    }},
                    {"oculto", new fieldParams(){
                        descricao = "Comissao Oculta",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                        size = 3,
                    }},
                    // percentual de ata
                    {"ata", new fieldParams(){
                        descricao = "Percent Ata",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Float,
                        subtipo = SAPbobsCOM.BoFldSubTypes.st_Percentage
                    }},
                    //marca falando se este empregado exerce função de gerente de contas 
                    {"gerente", new fieldParams(){
                        descricao = "Flag Gerente de Contas",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                        size = 1
                    }},
                    
                    //marca falando se este empregado exerce função de gerente de contas 
                    //utilizado para o split.
                    {"split", new fieldParams(){
                        descricao = "Flag Gerente de Contas",
                        tipo = SAPbobsCOM.BoFieldTypes.db_Alpha,
                        size = 1,
                    }},
                },
            };
        }

        #endregion
    }
}