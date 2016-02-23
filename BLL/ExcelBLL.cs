// ****************************************
// FileName:ExcelBLL.cs
// Description:Excel业务,数据库支持类
// Tables:Many
// Author:Gavin && Burney
// Create Date:2014-06-01
// Revision History:
// ****************************************

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Text;
using System.Linq;

namespace BLL
{
    using DAL;
    using MySql.Data.MySqlClient;
    using Utils.Configuration;
    using Utils.Excel;
    using Utils.Xml;

    /// <summary>
    /// Excel业务,数据库支持类
    /// </summary>
    public static class ExcelBLL
    {
        #region 静态变量

        //不需要加冒号的字段类型
        private static readonly List<String> mNoneColonType;

        //是否已数据库表字段为准
        private static Boolean baseOnBD;

        //表名备注信息
        private static Dictionary<String, String> mTableComments = new Dictionary<String, String>();

        #endregion

        #region 初始化

        /// <summary>
        /// 静态构造函数
        /// </summary>
        static ExcelBLL()
        {
            //不需要添加符号的字段类型
            mNoneColonType = new List<String> { "bool", "boolean", "bit", "short", "tinyint", "byte", "int" };

            //是否以数据库字段为准
            baseOnBD = Convert.ToBoolean(ConfigurationHelper.AppSettings["BaseOnDB"]);
        }

        #endregion

        #region 数据库

        /// <summary>
        /// 设置数据库连接字符串
        /// </summary>
        /// <param name="dbConnection">数据库连接字符串</param>
        public static void SetDbConnection(String dbConnection)
        {
            ExcelDAL.DbConnectionString = dbConnection;
        }

        /// <summary>
        /// 检测数据库是否可用
        /// </summary>
        /// <returns>如果数据库可连接返回true; 否则false</returns>
        public static Boolean IsDataBaseAccess()
        {
            return ExcelDAL.IsDataBaseAccess();
        }

        /// <summary>
        /// 获取数据库表名列表
        /// </summary>
        /// <returns>表名列表</returns>
        public static List<String> GetTableNameList()
        {
            List<String> list = new List<String>();

            DataTable dt = ExcelDAL.GetTableNames();
            if (dt == null) return list;

            //循环构造List
            for (Int32 i = 0; i < dt.Rows.Count; i++)
            {
                list.Add(dt.Rows[i][0].ToString());
            }

            return list;
        }

        /// <summary>
        /// 获取表所有数据
        /// </summary>
        /// <param name="tableName">表名</param>
        /// <returns>表数据</returns>
        public static DataTable GetTableData(String tableName)
        {
            return ExcelDAL.GetTableData(tableName);
        }

        /// <summary>
        /// 获取表字段,备注,字段类型信息
        /// </summary>
        /// <param name="tableName">表名</param>
        /// <returns>以字段名为Key,备注和字段类型为Value的字典</returns>
        public static Dictionary<String, String[]> GetComments(String tableName)
        {
            DataTable dt = ExcelDAL.GetComments(tableName);

            Dictionary<String, String[]> dic = new Dictionary<String, String[]>();
            foreach (DataRow dr in dt.Rows)
            {
                dic.Add(dr[0].ToString(), new String[] { dr[1].ToString(), dr[2].ToString() });
            }

            return dic;
        }

        /// <summary>
        /// 更新template_checkinfo表的时间
        /// </summary>
        public static void UpdateCheckInfoTime()
        {
            try
            {
                ExcelDAL.UpdateCheckInfoTime();
            }
            catch (Exception ex)
            {
                Trace.Write("更新template_checkinfo失败! 失败信息: " + ex.Message);
            }
        }

        /// <summary>
        /// 获取上次b_表更新时间时间
        /// </summary>
        /// <returns></returns>
        public static DateTime GetUpdateCheckInfoTime()
        {
            DateTime lastUpdateTime = DateTime.MinValue;

            try
            {
                lastUpdateTime = ExcelDAL.GetUpdateCheckInfoTime();
            }
            catch (Exception ex)
            {
                Trace.Write("更新template_checkinfo失败! 失败信息: " + ex.Message);
            }

            return lastUpdateTime;
        }

        /// <summary>
        /// 插入数据库
        /// </summary>
        /// <param name="sqlList">sql语句集合</param>
        /// <param name="tableName">表名</param>
        /// <param name="excelFile">excel文件名</param>
        /// <param name="needCreateTable">是否需要创建表</param>
        /// <returns>受影响行数</returns>
        public static Int32 Insert(List<String> sqlList, String tableName, String excelFile, Boolean needCreateTable)
        {
            MySqlTransaction trans = null;
            MySqlConnection myConn = new MySqlConnection(ExcelDAL.DbConnectionString);

            Int32 insertedDataCount = 0;
            Int32 tempSqlCount = needCreateTable ? 2 : 1;
            String sql = String.Empty;

            try
            {
                myConn.Open();
                trans = myConn.BeginTransaction();

                for (Int32 i = 0; i < sqlList.Count; i++)
                {
                    sql = sqlList[i];

                    if (i < tempSqlCount)
                        MySqlHelper.ExecuteNonQuery(myConn, sql);
                    else
                        insertedDataCount += MySqlHelper.ExecuteNonQuery(myConn, sql);
                }

                trans.Commit();
            }
            catch (Exception ex)
            {
                trans.Rollback();
                insertedDataCount = 0;

                //记录异常sql
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("【数据异常】");
                sb.AppendLine("文件: " + excelFile);
                sb.AppendLine("表单: " + tableName);
                sb.AppendLine("异常信息: " + ex.Message);
                if (sql.Length < 200) sb.AppendLine("异常的sql语句: " + sql);

                Trace.Write(sb.ToString());
            }
            finally
            {
                myConn.Close();
            }

            return insertedDataCount;
        }

        #endregion

        #region 读取

        /// <summary>
        /// 读取表单数据
        /// </summary>
        /// <param name="excel">文档对象</param>
        /// <param name="sheetIndex">表单序号</param>
        /// <returns>表单数据</returns>
        public static DataTable TryRead(MoqikakaExcel excel, Int32 sheetIndex)
        {
            DataTable table = null;

            try
            {
                //读取表单数据
                table = excel.ReadAt(sheetIndex);
            }
            catch (Exception ex)
            {
                //记录异常日志信息
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("【表单读取异常】");
                sb.AppendLine("异常文件: " + excel.Path);
                sb.AppendLine("异常表单: " + excel.GetSheetName(sheetIndex));
                sb.AppendLine("异常信息: " + ex.Message);
                sb.AppendLine("StackTrace: " + ex.StackTrace);
                Trace.Write(sb.ToString());
            }

            return table;
        }

        /// <summary>
        /// 读取表单数据
        /// </summary>
        /// <param name="excel">文档对象</param>
        /// <param name="sheetName">表单名称</param>
        /// <returns>表单数据</returns>
        public static DataTable TryRead(MoqikakaExcel excel, String sheetName)
        {
            return TryRead(excel, excel.GetSheetIndex(sheetName));
        }

        /// <summary>
        /// 是否为没用的表单
        /// </summary>
        /// <param name="sheetName">表单名</param>
        /// <returns>是否有用</returns>
        public static Boolean IsUselessSheet(String sheetName)
        {
            return CommonBLL.MatchedChinese(sheetName) || sheetName.StartsWith("Sheet");
        }

        #endregion

        #region  组装导入SQL

        /// <summary>
        /// DataTable to SQL
        /// </summary>
        /// <param name="table">需要插入的数据</param>
        /// <returns>SQL语句集合</returns>
        public static List<String> GetSQL(DataTable table)
        {
            //以数据库为准
            if (baseOnBD) return ToSqlByDB(table);

            return ToSQL(table);
        }

        /// <summary>
        /// 将DataTable转换为sql语句集合 以数据库为准
        /// </summary>
        /// <param name="table">数据源</param>
        /// <returns>sql语句集合</returns>
        private static List<String> ToSqlByDB(DataTable table)
        {
            //默认插入数据库表名
            String tableName = table.TableName;

            List<String> dbColumnName = GetComments(tableName).Keys.ToList();

            //当Excel表的字段数量 <= 数据库表字段数量时,以Excel字段为准
            if (table.Columns.Count <= dbColumnName.Count || dbColumnName.Count == 0)
                return ToSQL(table);

            //以数据库字段为准
            List<String> excelColumns = new List<String>();
            foreach (var item in table.Rows[MoqikakaExcelSettings.NameRowNum - 1].ItemArray)
                excelColumns.Add(item.ToString());

            //验证并收集Excel表单列名对应数据库字段的序号
            List<Int32> dbColumnIndex = new List<Int32>();
            dbColumnName.ForEach(p =>
            {
                if (!excelColumns.Exists(q =>
                {
                    if (String.Equals(q, p, StringComparison.OrdinalIgnoreCase))
                    {
                        dbColumnIndex.Add(excelColumns.IndexOf(q));
                        return true;
                    };
                    return false;
                }))
                    throw new Exception(String.Format("{0}表单中找不到与表字段{1}对应的字段!", tableName, p));
            });

            return MakeUpSQL(table, dbColumnName, dbColumnIndex);
        }

        /// <summary>
        /// 将DataTable转换为sql语句集合
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <returns>sql语句集合</returns>
        public static List<String> ToSQL(DataTable dt)
        {
            List<String> sqlList = new List<String>();

            //默认插入数据库表名
            String tableName = dt.TableName;

            //若表不存在,则添加创建表的sql
            if (IfCreateTable())
                sqlList.Add(CreateTableSql(dt));

            //没有数据
            if (dt.Rows.Count <= MoqikakaExcelSettings.DataRowNum - 1) return sqlList;

            //构造字段名字符串
            String columsNames = MakeUpColumNames(dt, ref tableName);

            //清空数据
            sqlList.Add(String.Format(@"DELETE FROM {0};{1}", tableName, Environment.NewLine));

            MakeUpSQL(tableName, columsNames, MakeUpColumValues(dt), ref sqlList);

            return sqlList;
        }

        /// <summary>
        /// 构造插入SQL 以数据库字段为准
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="dbColumnName">数据库字段集合</param>
        /// <param name="dbColumnIndex">对应Excel表单字段的index</param>
        /// <returns>SQL</returns>
        private static List<String> MakeUpSQL(DataTable dt, List<String> dbColumnName, List<Int32> dbColumnIndex)
        {
            List<String> sqlList = new List<String>();

            //插入数据字段名
            String columns = String.Format("`{0}`", String.Join("`,`", dbColumnName.ToArray()));

            //非数据行数量
            Int32 tempRowNum = MoqikakaExcelSettings.SpecialRowList.Count - 1;

            //不需要加引号的数据列
            List<Int32> noneColonColumn = NoneColonIndex(dt);

            //根据数据库字段顺序,构造values
            List<String> dataValues = new List<String>();
            foreach (DataRow item in dt.Rows)
            {
                //排除非数据行
                if (tempRowNum-- > 0) continue;

                String columsValues = "(";

                foreach (var index in dbColumnIndex)
                {
                    //数字类型的值不加引号
                    if (!noneColonColumn.Contains(index))
                        columsValues += "'" + item[index] + "',";
                    else
                        columsValues += item[index] + ",";
                }

                //去掉末尾","
                columsValues = columsValues.TrimEnd(',') + "),";
                dataValues.Add(columsValues);
            }

            //若表不存在,则添加创建表的sql
            if (IfCreateTable())
                sqlList.Add(CreateTableSql(dt));

            //清空数据
            sqlList.Add(String.Format(@"DELETE FROM {0};{1}", dt.TableName, Environment.NewLine));

            MakeUpSQL(dt.TableName, columns, dataValues, ref sqlList);

            return sqlList;
        }

        /// <summary>
        /// 构造Insert语句的字段名字符串
        /// </summary>
        /// <param name="dt">数据表</param>
        /// <param name="tableName">表名</param>
        /// <returns>Insert语句的字段名字符串</returns>
        private static String MakeUpColumNames(DataTable dt, ref String tableName)
        {
            String columsNames = String.Empty; //存放插入列名

            //若Excel字段行存在,描述行不存在
            if (MoqikakaExcelSettings.DescRowNum == -1 && MoqikakaExcelSettings.NameRowNum != -1)
            {
                for (Int32 j = 0; j < dt.Columns.Count; j++)
                {
                    columsNames += '`' + dt.Columns[j].Caption + '`' + ",";
                }
                columsNames = columsNames.TrimEnd(',');
            }
            //若字段和描述均存在
            else if (MoqikakaExcelSettings.DescRowNum != -1 && MoqikakaExcelSettings.NameRowNum != -1)
            {
                columsNames = String.Join("`,`", dt.Rows[MoqikakaExcelSettings.NameRowNum - 1].ItemArray.Cast<String>().ToArray());

                columsNames = String.Format("`{0}`", columsNames);
            }
            //仅描述存在, 字段不存在 (以字段描述去读取映射文件中的相关映射字段名)
            else if (MoqikakaExcelSettings.DescRowNum != -1 && MoqikakaExcelSettings.NameRowNum == -1)
            {
                XMLHelper xmlHelper = new XMLHelper("TableMapping.xml");

                //若该表存在映射关系, 获取映射后的表名
                String tempName = xmlHelper.GetTableNameMapping(dt.TableName);

                if (tempName != null)
                {
                    //设置插入表名为映射后的表名
                    tableName = tempName;

                    //字段映射字典
                    Dictionary<String, String> columnMappings = xmlHelper.GetTableMappingInfo(dt.TableName);

                    if (columnMappings.Count > 0)
                    {
                        for (Int32 j = 0; j < dt.Columns.Count; j++)
                        {
                            //若该字段存在映射关系
                            if (columnMappings.ContainsKey(dt.Columns[j].Caption))
                            {
                                columsNames += '`' + columnMappings[dt.Columns[j].Caption] + '`' + ",";
                            }
                            else
                            {
                                throw new Exception(String.Format("[{0}]的映射字段不全: [{1}]字段没有相关数据库字段映射!", dt.TableName, dt.Columns[j].Caption));
                            }
                        }
                        columsNames = columsNames.TrimEnd(',');
                    }
                }
            }

            return columsNames;
        }

        /// <summary>
        /// 构造数据值字符串
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <returns>sql语句集合</returns>
        private static List<String> MakeUpColumValues(DataTable dt)
        {
            //收集数字类型的字段 (不加引号)
            List<Int32> noneColonColumnList = NoneColonIndex(dt);

            //收集所有数据值
            List<String> valuesList = new List<String>();
            for (Int32 i = MoqikakaExcelSettings.DataRowNum - 1; i < dt.Rows.Count; i++)
            {
                String columsValues = "(";

                for (Int32 j = 0; j < dt.Columns.Count; j++)
                {
                    //数字类型的值不加引号
                    if (!noneColonColumnList.Contains(j))
                        columsValues += "'" + dt.Rows[i][j] + "',";
                    else
                        columsValues += dt.Rows[i][j].ToString() != "" ? dt.Rows[i][j].ToString() + "," : "null,";
                }

                //去掉末尾","
                columsValues = columsValues.TrimEnd(',') + "),";

                valuesList.Add(columsValues);
            }

            return valuesList;
        }

        /// <summary>
        /// 不需要加''号的字段列序号
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <returns>不需要加''号的字段列序号</returns>
        private static List<Int32> NoneColonIndex(DataTable dt)
        {
            //获取所有不需要加引号的字段的列号 
            List<Int32> noneColonColumnList = new List<Int32>();

            //如果字段类型行 存在
            if (MoqikakaExcelSettings.TypeRowNum != -1 && MoqikakaExcelSettings.SpecialRowList.Count > 1)
            {
                for (Int32 j = 0; j < dt.Columns.Count; j++)
                {
                    if (mNoneColonType.Contains(dt.Rows[MoqikakaExcelSettings.TypeRowNum - 1][j].ToString().ToLower()))
                        noneColonColumnList.Add(j);
                }
            }

            return noneColonColumnList;
        }

        /// <summary>
        /// 组装sql语句
        /// </summary>
        /// <param name="tableName">表名</param>
        /// <param name="colmnNames">列名字符串</param>
        /// <param name="dataList">数据集合</param>
        /// <param name="sqlList">sql语句集合</param>
        private static void MakeUpSQL(String tableName, String colmnNames, List<String> dataList, ref List<String> sqlList)
        {
            StringBuilder sqlBuilder = new StringBuilder();

            //获取配置 每多少条数据构造为一个sql语句
            Int32 perNum = Convert.ToInt32(ConfigurationHelper.AppSettings["PerDataNumOneSQL"]);
            String insertFormater = String.Format("INSERT INTO {0} ({1}) VALUES ", tableName, colmnNames);

            sqlBuilder.AppendLine(insertFormater);

            Int32 num = 1;
            foreach (var item in dataList) //todo 重构
            {
                sqlBuilder.AppendLine(item);

                if (num >= perNum && dataList.LastIndexOf(item) < dataList.Count - 1)
                {
                    sqlList.Add(sqlBuilder.ToString().TrimEnd().TrimEnd(',') + ';');

                    sqlBuilder = new StringBuilder();
                    sqlBuilder.AppendLine(insertFormater);

                    num = 1;
                    continue;
                }

                num++;
            }

            sqlList.Add(sqlBuilder.ToString().TrimEnd().TrimEnd(',') + ';');
        }

        /// <summary>
        /// 提供创建表格的sql ,默认字段类型和长度为VARCHAR(40)
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <returns>创建表的sql语句</returns>
        private static String CreateTableSql(DataTable dt)
        {
            String colums = String.Empty;
            String colDefaultTypeAndLength = @"VARCHAR(256) NOT NULL";  //默认字段类型和长度
            String colDefaultComment = String.Empty;            //默认字段备注

            //Excel中存放数据库列名的行必须存在
            if (MoqikakaExcelSettings.NameRowNum == -1) return String.Empty;

            //如果存在字段描述行,则字段名存在的行号为  MoqikakaExcelSettings.DataColumnNameRowNum - 1
            if (MoqikakaExcelSettings.DescRowNum != -1)
            {
                //字段类型行存在时
                if (MoqikakaExcelSettings.TypeRowNum != -1)
                {
                    for (Int32 i = 0; i < dt.Columns.Count; i++)
                    {
                        //根据字段类型数据映射MySql字段类型和长度
                        String colTypeAndLength = ToDBColumnType(dt.Rows[MoqikakaExcelSettings.TypeRowNum - 1][i].ToString());

                        //PlayerId字段单独处理,GUID
                        if (dt.Rows[MoqikakaExcelSettings.NameRowNum - 1][i].ToString().ToLower() == "playerid")
                        {
                            colTypeAndLength = @"char(36) NOT NULL";
                        }
                        colums += String.Format(@"`{0}` {1} COMMENT '{2}',", ConvertIDToId(dt.Rows[MoqikakaExcelSettings.NameRowNum - 1][i].ToString()), colTypeAndLength, dt.Columns[i].Caption.Replace(';', ','));
                    }

                    //添加主键 (#开头的字段)
                    List<String> primaryKeyList = new List<String>();
                    foreach (DataColumn col in dt.Columns)
                    {
                        //以#描述开头 &&　非String字段
                        if (col.Caption.StartsWith("#") && dt.Rows[MoqikakaExcelSettings.TypeRowNum - 1][col].ToString().ToLower() != "string")
                            primaryKeyList.Add("`" + dt.Rows[MoqikakaExcelSettings.NameRowNum - 1][col] + "`");
                    }

                    if (primaryKeyList.Count > 0)
                        colums += String.Format("PRIMARY KEY ({0})", String.Join(",", primaryKeyList.ToArray()));
                }
                //否则为默认字段类型和长度
                else
                {
                    for (Int32 i = 0; i < dt.Columns.Count; i++)
                    {
                        String colTypeAndLength = colDefaultTypeAndLength;
                        //PlayerId字段单独处理,GUID
                        if (dt.Rows[MoqikakaExcelSettings.NameRowNum - 1][i].ToString().ToLower() == "playerid")
                        {
                            colTypeAndLength = @"char(36) NOT NULL";
                        }
                        colums += String.Format(@"`{0}` {1} COMMENT '{2}',", dt.Rows[MoqikakaExcelSettings.NameRowNum - 1][i].ToString(), colTypeAndLength, dt.Columns[i].Caption.Replace(';', ','));
                    }
                }

            }
            //字段描述行不存在时(备注) 取DataTable列名
            else
            {
                //字段类型行存在时,通过ChangeCSharpTypeToDefaultMySqlTypeLength转换成为MySql类型和长度
                if (MoqikakaExcelSettings.TypeRowNum != -1)
                {
                    for (Int32 i = 0; i < dt.Columns.Count; i++)
                    {
                        //如果colDefaultTypeAndLength为空,则根据c#字段类型判断,否则,为默认varchar(256)
                        String colTypeAndLength = ToDBColumnType(dt.Columns[i].Caption);

                        //PlayerId字段单独处理,GUID
                        if (dt.Columns[i].Caption.ToLower() == "playerid")
                        {
                            colTypeAndLength = @"char(36) NOT NULL";
                        }

                        colums += String.Format(@"`{0}` {1} COMMENT '{2}',", dt.Columns[i].Caption, colTypeAndLength, colDefaultComment);
                    }
                }
                //否则,均为字段和长度以及备注均为默认值
                else
                {
                    for (Int32 i = 0; i < dt.Columns.Count; i++)
                    {
                        String colTypeAndLength = colDefaultTypeAndLength;
                        //PlayerId字段单独处理,GUID
                        if (dt.Columns[i].Caption == "playerid")
                        {
                            colTypeAndLength = @"char(36) NOT NULL";
                        }
                        colums += String.Format(@"`{0}` {1} COMMENT '{2}',", dt.Columns[i].Caption, colTypeAndLength, colDefaultComment);
                    }
                }

            }

            colums = colums.TrimEnd(',');
            String sql = String.Format("CREATE TABLE IF NOT EXISTS `{0}` ({1}) ENGINE=InnoDB DEFAULT CHARSET=utf8", dt.TableName, colums);

            //添加表备注
            if (mTableComments.ContainsKey(dt.TableName.ToLower()))
            {
                sql += String.Format(" COMMENT='{0}'", mTableComments[dt.TableName.ToLower()]);
            }

            return sql + ";" + Environment.NewLine;
        }

        /// <summary>
        /// 将Excel的C#中字段类型转换为默认的Mysql数据库字段类型和长度
        /// </summary>
        /// <param name="cSharpType">C#字段类型</param>
        /// <returns></returns>
        private static String ToDBColumnType(String cSharpType)
        {
            String mySqlColumnType;

            switch (cSharpType.ToLower())
            {
                case "bit":
                    mySqlColumnType = "bit(1)";
                    break;
                case "bool":
                    mySqlColumnType = "boolean";
                    break;
                case "guid":
                    mySqlColumnType = "char(36)";
                    break;
                case "byte":
                case "tinyint":
                    mySqlColumnType = "tinyint(3) UNSIGNED";
                    break;
                case "decimal":
                    mySqlColumnType = "decimal(10,2) UNSIGNED";
                    break;
                case "double":
                    mySqlColumnType = "double UNSIGNED";
                    break;
                case "float":
                    mySqlColumnType = "float UNSIGNED";
                    break;
                case "numeric":
                    mySqlColumnType = "numeric UNSIGNED";
                    break;
                case "int":
                case "short":
                case "smallint":
                case "mediumint":
                case "integer":
                case "bigint":
                case "int16":
                case "int32":
                case "int64":
                    mySqlColumnType = "int(10) UNSIGNED";
                    break;
                case "string":
                case "text":
                case "char":
                case "varchar":
                case "tinytext":
                    mySqlColumnType = "VARCHAR(256)";
                    break;

                case "time":
                    mySqlColumnType = "time";
                    break;

                case "datetime":
                case "date":
                    mySqlColumnType = "datetime";
                    break;
                default:
                    mySqlColumnType = "VARCHAR(256)";
                    break;
            }

            return mySqlColumnType + " NOT NULL";
        }

        /// <summary>
        /// 转换字段名中的ID为Id
        /// </summary>
        /// <param name="src">源字符串</param>
        /// <returns>转换后的字符串</returns>
        private static String ConvertIDToId(String src)
        {
            return src.Replace("ID", "Id");
        }

        /// <summary>
        /// 是否需要创建表
        /// </summary>
        /// <returns>是否需要创建表</returns>
        public static Boolean IfCreateTable()
        {
            return MoqikakaExcelSettings.NameRowNum != -1 && ConfigurationHelper.AppSettings["CreateTable"] == "true";
        }

        /// <summary>
        /// 添加表备注
        /// </summary>
        /// <param name="sheetName">表(单)名</param>
        /// <param name="comment">备注</param>
        public static void AddTableComment(String sheetName, String comment)
        {
            mTableComments[sheetName.ToLower()] = comment.Trim();
        }

        /// <summary>
        /// 获取表备注
        /// </summary>
        /// <param name="sheetName">表(单)名</param>
        /// <returns>备注</returns>
        public static String GetTableComment(String sheetName)
        {
            if (!mTableComments.ContainsKey(sheetName.ToLower()))
            {
                return null;
            }

            return mTableComments[sheetName.ToLower()];
        }

        #endregion
    }
}