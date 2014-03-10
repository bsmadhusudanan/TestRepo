using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using ADODB;
namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Recordset rs = new Recordset();
                rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient;
                ADODB.Fields resultFields = rs.Fields;
                System.Data.DataColumnCollection inColumns = null;
                //Provider=SQLOLEDB.1;Data Source=claysys087s;Integrated Security=SSPI;Initial Catalog=tempdb
                OleDbConnection conn = new OleDbConnection("Provider=SQLOLEDB.1;Data Source=claysys087s;Integrated Security=SSPI;Initial Catalog=testdb");
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    OleDbDataAdapter dap = new OleDbDataAdapter("select * from table1", conn);
                    DataTable ds = new DataTable();
                    dap.Fill(ds);
                    inColumns = ds.Columns;
                    foreach (DataColumn inColumn in inColumns)
                    {
                        resultFields.Append(inColumn.ColumnName
                            , TranslateType(inColumn.DataType)
                            , inColumn.MaxLength
                            , inColumn.AllowDBNull ? ADODB.FieldAttributeEnum.adFldIsNullable :
                                                     ADODB.FieldAttributeEnum.adFldUnspecified
                            , null);
                    }
                    rs.Open(System.Reflection.Missing.Value
            , System.Reflection.Missing.Value
            , ADODB.CursorTypeEnum.adOpenStatic
            , ADODB.LockTypeEnum.adLockOptimistic, 0);


                    foreach (DataRow dr in ds.Rows)
                    {
                        rs.AddNew(System.Reflection.Missing.Value,
                                      System.Reflection.Missing.Value);

                        for (int columnIndex = 0; columnIndex < inColumns.Count; columnIndex++)
                        {
                            resultFields[columnIndex].Value = dr[columnIndex];
                        }
                    }

                    
                    
                }
                Recordset rs1 = rs;

            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message.ToString());
            }

        }

        public ADODB.DataTypeEnum TranslateType(Type columnType)
{
    switch (columnType.UnderlyingSystemType.ToString())
    {
        case "System.Boolean":
            return ADODB.DataTypeEnum.adBoolean;

        case "System.Byte":
            return ADODB.DataTypeEnum.adUnsignedTinyInt;

        case "System.Char":
            return ADODB.DataTypeEnum.adChar;

        case "System.DateTime":
            return ADODB.DataTypeEnum.adDate;

        case "System.Decimal":
            return ADODB.DataTypeEnum.adCurrency;

        case "System.Double":
            return ADODB.DataTypeEnum.adDouble;

        case "System.Int16":
            return ADODB.DataTypeEnum.adSmallInt;

        case "System.Int32":
            return ADODB.DataTypeEnum.adInteger;

        case "System.Int64":
            return ADODB.DataTypeEnum.adBigInt;

        case "System.SByte":
            return ADODB.DataTypeEnum.adTinyInt;

        case "System.Single":
            return ADODB.DataTypeEnum.adSingle;

        case "System.UInt16":
            return ADODB.DataTypeEnum.adUnsignedSmallInt;

        case "System.UInt32":
            return ADODB.DataTypeEnum.adUnsignedInt;

        case "System.UInt64":
            return ADODB.DataTypeEnum.adUnsignedBigInt;

        case "System.String":
        default:
            return ADODB.DataTypeEnum.adVarChar;
    }
}
    }
}
