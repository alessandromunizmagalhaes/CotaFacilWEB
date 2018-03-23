using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TShark;
using System.IO;

namespace ITOne
{
    class _FrmCadastros : TShark.Forms
    {
        public string FieldPrincipal = "U_desc";
        public List<ColumnDefinition> MatrixColumns = new List<ColumnDefinition>();

        public SAPbouiCOM.Button btnSalvar;

        public _FrmCadastros(Addon addon, Dictionary<string, dynamic> ExtraParams = null): base(addon, ExtraParams)
        {
            this.FormParams = new FormParams()
            {
                Title = "",
                MainDatasource = "",

                Bounds = new Bounds()
                {
                    Top = 100,
                    Left = 500,
                    Width = 460,
                    Height = 380,
                },

                #region Layout de Componentes

                Linhas = new Dictionary<string, int>()
                {
                    {"matrix", 100},
                    {"space",55},{"btnAdd",20},{"btnRmv",20}
                },

                Buttons = new Dictionary<string, int>(){
                    {"btnSalvar",20}, {"space",80}
                },

                #endregion

                #region Propriedade dos Componentes

                Controls = new Dictionary<string, CompDefinition>()
                {
                    
                    #region :: Matrix
                    
                    {"matrix", new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_MATRIX,
                        Caption = "",
                        Height = 300,
                    }},
                    {"btnAdd" , new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Adicionar"
                    }},
                    {"btnRmv" , new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Remover"
                    }},

                    #endregion


                    #region :: Buttons 

                    {"btnSalvar" , new CompDefinition(){
                        Type = SAPbouiCOM.BoFormItemTypes.it_BUTTON,
                        Caption = "Fechar",
                    }},

                    #endregion

                }

                #endregion

            };
        }

        #region Métodos onCreate

        /// <summary>
        /// criando a matriz
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void matrixOnCreate(SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            this.Matrix = this.SetupMatrix(evObj.ItemUID, this.FormParams.MainDatasource, this.MatrixColumns);

            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS, this.FormId, "matrix", "matrixOnValidate");
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_KEY_DOWN, this.FormId, "matrix", "matrixOnKeyDown");
            this.Addon.RegisterEvent(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK, this.FormId, "matrix", "matrixOnDoubleClick");
        }

        #endregion


        #region Eventos do Formulário

        /// <summary>
        /// Abertura do form
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void _FrmCadastrosOnFormOpen(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this.btnSalvar = this.GetItem("btnSalvar").Specific;

            this.RefreshMatrix(ref this.SapForm, "matrix", this.FormParams.MainDatasource);
        }

        /// <summary>
        /// Abertura do form
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void _FrmCadastrosOnFormClose(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            //se salvou
            //if (this.SaveMatrixToServer("matrix", this.FormParams.MainDatasource))
            //{
            //    this.Addon.StatusInfo("Registros salvos com sucesso.");
            //}
        }

        #endregion


        #region Evento dos Componentes

        /// <summary>
        /// Garante a mudança de label no botão salvar
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void matrixOnValidate(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            this.btnSalvar.Caption = "Salvar";
        }

        /// <summary>
        /// Insere nova linha ao teclar ENTER
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void matrixOnKeyDown(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if(evObj.CharPressed == 13)
            {

                // ENTER em row vazia deleta o row
                string v = this.Matrix.GetCellSpecific(this.FieldPrincipal, evObj.Row).Value;
                if(String.IsNullOrEmpty(v))
                {
                    this.Matrix.SelectRow(this.Matrix.RowCount, true, false);
                    this.DeleteOnMatrix(this.Matrix);
                    if(this.Matrix.RowCount > 0)
                    {
                        this.Matrix.SetCellFocus(this.Matrix.RowCount, 1);
                    }

                    // ENTER em row preenchido acrescenta novo row
                } else
                {
                    this.btnAddOnClick(ref evObj, out BubbleEvent);
                }
            }

        }


        /// <summary>
        /// Acrescenta linha à matriz
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnAddOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            //adicionando uma linha.
            if(!this.MatrixEmptyLastRow(this.Matrix, this.FieldPrincipal))
            {
                this.InsertOnMatrix("matrix", new Dictionary<string, dynamic>() {
                    {this.FieldPrincipal, ""},
                }, true);
            }

            this.btnSalvar.Caption = "Salvar";
        }

        /// <summary>
        /// Acrescenta linha à matriz
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnRmvOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            //removendo registro
            this.DeleteMatrixOnServer("matrix", this.FormParams.MainDatasource);
        }

        /// <summary>
        /// Evento que salva os registros no banco
        /// </summary>
        /// <param name="evObj"></param>
        /// <param name="BubbleEvent"></param>
        public void btnSalvarOnClick(ref SAPbouiCOM.ItemEvent evObj, out bool BubbleEvent)
        {
            BubbleEvent = true;

            for(int r = this.Matrix.RowCount; r > 0; r--)
            {
                string teste = Matrix.GetCellSpecific(this.FieldPrincipal, r).Value;
                string code = Matrix.GetCellSpecific("Code", r).Value;
                if(String.IsNullOrEmpty(teste))
                {
                    // Se já tem código, deleta no server
                    try
                    {
                        if(!String.IsNullOrEmpty(code))
                        {
                            this.btnRmvOnClick(ref evObj, out BubbleEvent);

                            // Senão, deleta em client
                        } else
                        {
                            Matrix.DeleteRow(r);
                        }
                    } catch
                    {
                        Matrix.DeleteRow(r);
                    }
                }
            }
            Matrix.FlushToDataSource();

            if(this.btnSalvar.Caption == "Salvar")
            {
                if(this.Matrix.RowCount > 0 && this.SaveMatrixToServer("matrix", this.FormParams.MainDatasource))
                {
                    this.RefreshMatrix(ref this.SapForm, "matrix", this.FormParams.MainDatasource);
                    this.Addon.StatusInfo("Registros salvos com sucesso.");
                }
                this.btnSalvar.Caption = "Fechar";
            } else
            {
                this.btnFecharOnClick(ref evObj, out BubbleEvent);
            }
        }

        #endregion


        #region Regras de Negócio


        #endregion

    }
}
