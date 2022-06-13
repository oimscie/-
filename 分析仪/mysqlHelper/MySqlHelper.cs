using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Gama.Util;
using static LogHelp;

namespace Gama.mysqlHelper
{
    public class MySqlHelper
    {
        /// <summary>
        /// 数据库连接头
        /// </summary>
        private readonly MySqlConnection Conn;
        /// <summary>
        /// 数据库连接句柄
        /// </summary>
        private readonly MySqlCommand Command;
        /// <summary>
        /// 读取头
        /// </summary>
        private MySqlDataReader Reader;
        /// <summary>
        /// 连接字符串
        /// </summary>
        private readonly string ConnStr = "Server=" + ConfigurationManager.AppSettings["ServerIp"] + ";Database=product;uid=" + DESUtil.DesDecrypt(ConfigurationManager.AppSettings["ServerUser"], "qwertyuiop") + ";pwd=" + DESUtil.DesDecrypt(ConfigurationManager.AppSettings["ServerPassWord"], "qwertyuiop") + ";SslMode=none;charset=utf8";
        public MySqlHelper()
        {
            Conn = new MySqlConnection(ConnStr);
            Command = new MySqlCommand
            {
                Connection = Conn
            };
        }
        /// <summary>
        /// 关闭连接
        /// </summary>
        public void Close()
        {
            try
            {
                Conn.Close();
            }
            catch
            {
            }
        }
        /// <summary>
        /// 打开连接
        /// </summary>
        /// <returns></returns>
        public bool Open()
        {
            try
            {
                Conn.Open();
                if (Conn.State == ConnectionState.Open)
                {
                    return true;
                }
                return false;
            }
            catch (Exception e)
            {
                LogHelper.WriteLog("mysql连接错误", e);
                return false;
            }
        }
        public bool CheckConn()
        {
            if (Conn.State == ConnectionState.Closed)
            {
                return Open();
            }
            if (Conn.State == ConnectionState.Broken || Conn.State == ConnectionState.Connecting)
            {
                Close();
                return Open();
            }
            if (Conn.State == ConnectionState.Executing || Conn.State == ConnectionState.Fetching)
            {
                return false;
            }
            return true;
        }
        /// <summary>
        /// select（多条返回）
        /// </summary>
        /// <param name="sql">SQL语句</param>
        /// <param name="fieldName">要查询的数据字段名称集合</param>
        /// <returns>List<Dictionary<string, string>>，以输入的字段名称为key值</returns>
        public List<Dictionary<string, string>> MultipleSelect(string sql, List<string> fieldName)
        {
            if (!CheckConn()) { return null; }
            try
            {
                Command.CommandText = sql;
                Reader = Command.ExecuteReader();
                List<Dictionary<string, string>> back = new List<Dictionary<string, string>>();
                int index = 0;
                while (Reader.Read())
                {
                    Dictionary<string, string> dic = new Dictionary<string, string>();
                    for (int i = 0; i < Reader.FieldCount; i++)
                    {
                        dic.Add(fieldName[index], Reader.GetString(fieldName[index]));
                        index++;
                    }
                    index = 0;
                    back.Add(dic);
                }
                Reader.Close();
                if (back.Count == 0) { return null; };
                return back;
            }
            catch (Exception e)
            {
                LogHelper.WriteLog("sql查询错误------" + sql + "------", e);
                return null;
            }
            finally
            {
                Close();
            }
        }

        /// <summary>
        ///  select（多条返回）
        /// </summary>
        /// <param name="sql">SQL语句</param>
        /// <param name="fieldName">要查询的数据字段名称</param>
        /// <returns>ArrayList</returns>
        public List<string> MultipleSelect(string sql, string fieldName, string none)
        {
            if (!CheckConn()) { return null; }
            try
            {
                Command.CommandText = sql;
                Reader = Command.ExecuteReader();
                List<string> back = new List<string>();
                while (Reader.Read())
                {
                    back.Add(Reader.GetString(fieldName));
                }
                Reader.Close();
                if (back.Count == 0) { return null; };
                return back;
            }
            catch (Exception e)
            {
                LogHelper.WriteLog("sql查询错误------" + sql + "------", e);
                return null;
            }
            finally
            {
                Close();
            }
        }

        /// <summary>
        ///  select（多条返回）
        /// </summary>
        /// <param name="sql">SQL语句</param>
        /// <param name="fieldName">要查询的数据字段名称</param>
        /// <returns>ArrayList</returns>
        public List<double> MultipleSelect(string sql, string fieldName)
        {
            if (!CheckConn()) { return null; }
            try
            {
                Command.CommandText = sql;
                Reader = Command.ExecuteReader();
                List<double> back = new List<double>();
                while (Reader.Read())
                {
                    back.Add(Reader.GetDouble(fieldName));
                }
                Reader.Close();
                if (back.Count == 0) { return null; };
                return back;
            }
            catch (Exception e)
            {
                LogHelper.WriteLog("sql查询错误------" + sql + "------", e);
                return null;
            }
            finally
            {
                Close();
            }
        }

        /// <summary>
        /// select（多条返回）
        /// </summary>
        /// <param name="sql">SQL语句</param>
        /// <param name="fieldName">要查询的数据字段名称</param>
        /// <returns>Dictionary<string, string>，以输入的字段名称为key值</returns>
        public Dictionary<string, string> SingleSelect(string sql, string fieldName)
        {
            if (!CheckConn()) { return null; }
            try
            {
                Command.CommandText = sql;
                Reader = Command.ExecuteReader();
                Dictionary<string, string> back = new Dictionary<string, string>();
                while (Reader.Read())
                {
                    back.Add(fieldName, Reader.GetString(fieldName));
                }
                Reader.Close();
                if (back.Count == 0) { return null; };
                return back;
            }
            catch (Exception e)
            {
                LogHelper.WriteLog("sql查询错误------" + sql + "------", e);
                return null;
            }
            finally
            {
                Close();
            }
        }

        /// <summary>
        /// select（单条返回）
        /// </summary>
        /// <param name="sql">SQL语句</param>
        /// <param name="fieldName">要查询的数据字段名称集合</param>
        /// <returns>Dictionary<string, string>，以输入的字段名称为key值</returns>
        public Dictionary<string, string> SingleSelect(string sql, string[] fieldName)
        {
            if (!CheckConn()) { return null; }
            try
            {
                Command.CommandText = sql;
                Reader = Command.ExecuteReader();
                Dictionary<string, string> back = new Dictionary<string, string>();
                if (Reader.Read())
                {
                    for (int i = 0; i < Reader.FieldCount; i++)
                    {
                        back.Add(fieldName[i], Reader.GetString(fieldName[i]));
                    }
                }
                Reader.Close();
                if (back.Count == 0) { return null; };
                return back;
            }
            catch (Exception e)
            {
                LogHelper.WriteLog("sql查询错误------" + sql + "------", e);
                return null;
            }
            finally
            {
                Close();
            }
        }

        /// <summary>
        /// 插入、修改、删除
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public int UpdOrInsOrdel(string sql)
        {
            if (!CheckConn()) { return 0; }
            try
            {
                Command.CommandText = sql;
                return Command.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                LogHelper.WriteLog("sql查询错误------" + sql + "------", e);
                return 0;
            }
            finally
            {
                Close();
            }

        }
        /// <summary>
        /// 查找目标是否存在，必须是select count（ID） as Count.....
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public int GetCount(string sql)
        {
            try
            {
                int back = 0;
                if (!CheckConn()) { return 0; }
                Command.CommandText = sql;
                Reader = Command.ExecuteReader();
                while (Reader.Read())
                {
                    back = Reader.GetInt32("Count");
                }
                Reader.Close();
                return back;
            }
            catch (Exception e)
            {
                LogHelper.WriteLog("sql查询错误------" + sql + "------", e);
                return 0;
            }
            finally
            {
                Close();
            }

        }
    }
}
