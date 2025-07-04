﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;

namespace NGMapping
{
    #region class----SQLite データベースの処理を管理するクラス
    // System.Data.SQLiteをNugetからインストールして使用
    public class SQLiteCon
    {
        public string DbPath { get; private set; }
        public bool IsConnected { get; private set; } = false; // データベース接続状態を示す
        #region コンストラクタ
        public SQLiteCon(string dbPath, List<TableInfo> tables, bool admin)
        {
            DbPath = dbPath;
            try
            {
                if (admin)
                {
                    // 管理モード
                    if (!File.Exists(DbPath))
                    {
                        // ファイルが存在しない場合、新規作成
                        IsConnected = CreateDatabase(tables);
                    }
                    else
                    {
                        // ファイルが存在する場合、構造のチェック＆修正
                        IsConnected = CheckAndUpdateDatabaseStructure(tables);
                    }
                }
                else
                {
                    // 通常モード
                    if (File.Exists(DbPath) && CheckDatabaseStructure(tables))
                    {
                        // ファイルが存在し、構造が一致している場合
                        IsConnected = true;
                    }
                    else
                    {
                        // ファイルがない、または構造不一致の場合
                        IsConnected = false;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"データベース操作中にエラーが発生しました: {ex.Message}");
                IsConnected = false;
            }
        }
        #endregion
        #region method----CreateDatabase
        private bool CreateDatabase(List<TableInfo> tables)
        {
            try
            {
                SQLiteConnection.CreateFile(DbPath); // データベースファイルの作成
                Console.WriteLine($"データベースファイル '{DbPath}' を作成しました。");

                using var connection = new SQLiteConnection($"Data Source={DbPath};Version=3;");
                connection.Open();

                foreach (var table in tables)
                {
                    var columnDefinitions = new List<string>();

                    foreach (var column in table.Columns)
                    {
                        var colDef = $"{column.Name} {column.GetSQLiteType()}";
                        if (!column.IsNullable) colDef += " NOT NULL";
                        if (column.IsPrimaryKey) colDef += " PRIMARY KEY";
                        if (column.IsAutoIncrement) colDef += " AUTOINCREMENT";

                        columnDefinitions.Add(colDef);
                    }

                    var sql = $"CREATE TABLE IF NOT EXISTS {table.TableName} ({string.Join(", ", columnDefinitions)});";

                    using var command = new SQLiteCommand(sql, connection);
                    command.ExecuteNonQuery();

                    Console.WriteLine($"テーブル '{table.TableName}' を作成しました.");
                }

                return true; // 成功
            }
            catch (Exception ex)
            {
                Console.WriteLine($"データベース作成中にエラーが発生しました: {ex.Message}");
                return false; // 失敗
            }
        }
        #endregion
        #region method----DB構造確認、修正
        private bool CheckAndUpdateDatabaseStructure(List<TableInfo> tables)
        {
            try
            {
                using var connection = new SQLiteConnection($"Data Source={DbPath};Version=3;");
                connection.Open();

                foreach (var table in tables)
                {
                    string sql = $"PRAGMA table_info({table.TableName});";
                    using var command = new SQLiteCommand(sql, connection);

                    using var reader = command.ExecuteReader();
                    var existingColumns = new List<string>();

                    while (reader.Read())
                    {
                        existingColumns.Add(reader["name"].ToString());
                    }

                    if (existingColumns.Count == 0)
                    {
                        // テーブルが存在しない場合、新規作成
                        Console.WriteLine($"テーブル '{table.TableName}' が存在しないため新規作成します。");
                        CreateTable(table, connection);
                    }
                    else
                    {
                        // テーブルが存在する場合、カラム構造をチェック
                        foreach (var column in table.Columns)
                        {
                            if (!existingColumns.Contains(column.Name))
                            {
                                // カラム不足の場合、追加する
                                Console.WriteLine($"テーブル '{table.TableName}' に不足しているカラム '{column.Name}' を追加します。");
                                AddColumn(table.TableName, column, connection);
                            }
                            else
                            {
                                // TODO: データ型の不一致の修正を実装
                                Console.WriteLine($"カラム '{column.Name}' の既存定義は確認済みです。");
                            }
                        }
                    }
                }

                return true; // 成功
            }
            catch (Exception ex)
            {
                Console.WriteLine($"データベース修正中にエラーが発生しました: {ex.Message}");
                return false; // 失敗
            }
        }
        #endregion
        #region method----DB構造確認
        private bool CheckDatabaseStructure(List<TableInfo> tables)
        {
            try
            {
                using var connection = new SQLiteConnection($"Data Source={DbPath};Version=3;");
                connection.Open();

                foreach (var table in tables)
                {
                    string sql = $"PRAGMA table_info({table.TableName});";
                    using var command = new SQLiteCommand(sql, connection);

                    using var reader = command.ExecuteReader();
                    var existingColumns = new List<string>();

                    while (reader.Read())
                    {
                        existingColumns.Add(reader["name"].ToString());
                    }

                    // テーブルがない、またはカラムが不足している場合は構造不一致
                    if (existingColumns.Count == 0 || table.Columns.Any(c => !existingColumns.Contains(c.Name)))
                    {
                        return false;
                    }
                }

                return true; // 構造一致
            }
            catch (Exception ex)
            {
                Console.WriteLine($"データベース構造確認中にエラーが発生しました: {ex.Message}");
                return false; // 失敗
            }
        }
        #endregion
        #region method----テーブル新規作成
        private void CreateTable(TableInfo table, SQLiteConnection connection)
        {
            var columnDefinitions = new List<string>();

            foreach (var column in table.Columns)
            {
                var colDef = $"{column.Name} {column.GetSQLiteType()}";
                if (!column.IsNullable) colDef += " NOT NULL";
                if (column.IsPrimaryKey) colDef += " PRIMARY KEY";
                if (column.IsAutoIncrement) colDef += " AUTOINCREMENT";

                columnDefinitions.Add(colDef);
            }

            var sql = $"CREATE TABLE {table.TableName} ({string.Join(", ", columnDefinitions)});";

            using var command = new SQLiteCommand(sql, connection);
            command.ExecuteNonQuery();

            Console.WriteLine($"テーブル '{table.TableName}' を作成しました。");
        }
        #endregion
        #region method----テーブルにカラム追加
        private void AddColumn(string tableName, ColumnInfo column, SQLiteConnection connection)
        {
            var sql = $"ALTER TABLE {tableName} ADD COLUMN {column.Name} {column.GetSQLiteType()}";

            using var command = new SQLiteCommand(sql, connection);
            command.ExecuteNonQuery();

            Console.WriteLine($"カラム '{column.Name}' をテーブル '{tableName}' に追加しました。");
        }
        #endregion
        #region method----GetData(sql文でDataTabe取得)
        public bool GetData(string query, out DataTable result)
        {
            result = new DataTable();

            try
            {
                using var connection = new SQLiteConnection($"Data Source={DbPath};Version=3;");
                connection.Open();

                using var command = new SQLiteCommand(query, connection);
                using var adapter = new SQLiteDataAdapter(command);

                adapter.Fill(result);
                return true; // 成功
            }
            catch (Exception ex)
            {
                Console.WriteLine($"データ取得中にエラーが発生しました: {ex.Message}");
                return false; // 失敗
            }
        }
        #endregion
        public bool GetData(List<string> sql, out DataTable[] results)
        {
            results = new DataTable[sql.Count];

            try
            {
                using var connection = new SQLiteConnection($"Data Source={DbPath};Version=3;");
                connection.Open();

                for (int i = 0; i < sql.Count; i++)
                {
                    results[i] = new DataTable();
                    using var command = new SQLiteCommand(sql[i], connection);
                    using var adapter = new SQLiteDataAdapter(command);
                    adapter.Fill(results[i]);
                }
                return true; // 成功
            }
            catch (Exception ex)
            {
                Console.WriteLine($"データ取得中にエラーが発生しました: {ex.Message}");
                return false; // 失敗
            }
        }


        #region method----InsertData(sql文でInsert)
        public bool InsertData(string query)
        {
            try
            {
                using var connection = new SQLiteConnection($"Data Source={DbPath};Version=3;");
                connection.Open();

                using var command = new SQLiteCommand(query, connection);
                command.ExecuteNonQuery();

                return true; // 成功
            }
            catch (Exception ex)
            {
                Console.WriteLine($"データ挿入中にエラーが発生しました: {ex.Message}");
                return false; // 失敗
            }
        }
        #endregion
        #region method----Excute        
        public bool Execute(List<string> sql)
        {
            using var connection = new SQLiteConnection($"Data Source={DbPath};Version=3;");
            connection.Open();

            using var transaction = connection.BeginTransaction(); // トランザクションを開始
            try
            {
                using var command = new SQLiteCommand();
                command.Connection=connection;
                command.Transaction =  transaction;
                // コマンド作成と実行
                for (int i = 0; i < sql.Count; i++)
                {
                    command.CommandText = sql[i];
                    command.ExecuteNonQuery();
                }
                transaction.Commit();// 成功時にトランザクションをコミット
                return true; // 成功
            }
            catch (Exception ex)
            {
                Console.WriteLine($"クエリ実行中にエラーが発生しました: {ex.Message}");

                // エラー時にトランザクションをロールバック
                try
                {
                    transaction.Rollback();
                }
                catch (Exception rollbackEx)
                {
                    Console.WriteLine($"ロールバック処理中にエラーが発生しました: {rollbackEx.Message}");
                }

                return false; // 失敗
            }
        }
        #endregion
        #region method----InsertとID取得
        public bool InsertDataAndGetId(string query,out int ID)
        {
            try
            {
                using var connection = new SQLiteConnection($"Data Source={DbPath};Version=3;");
                connection.Open();

                using var command = new SQLiteCommand(query, connection);
                command.ExecuteNonQuery();

                // オートインクリメントされたIDを取得
                using var idCommand = new SQLiteCommand("SELECT last_insert_rowid();", connection);
                var id = idCommand.ExecuteScalar();


                ID= id != null ? Convert.ToInt32(id) : -1; // IDを返す
                return true; // 成功
            }
            catch (Exception ex)
            {
                Console.WriteLine($"データ挿入中にエラーが発生しました: {ex.Message}");
                ID = -1; // 失敗時は-1を返す
                return false; // 失敗時はnullを返す
            }
        }
        #endregion
    }
    #endregion
    #region enum-----サポートするデータ型定義
    public enum DataType
    {
        Text,       // 文字列型（TEXT）
        Integer,    // 整数型（INTEGER）
        Double,     // 実数型（REAL）
        Boolean,    // 真偽型（BOOLEANやINTEGERとして使用）
        DateTime    // 日付時刻型（SQLiteではTEXT型として保存）
    }
    #endregion
    #region class----カラム情報を定義するクラス
    public class ColumnInfo(
        string name,
        DataType type,
        bool isNullable = true,
        bool isPrimaryKey = false,
        bool isAutoIncrement = false,
        int maxLength = 256)
    {
        public string Name { get; set; } = name;
        public DataType Type { get; set; } = type;
        public bool IsNullable { get; set; } = isNullable;
        public bool IsPrimaryKey { get; set; } = isPrimaryKey;
        public bool IsAutoIncrement { get; set; } = isAutoIncrement;
        public int MaxLength { get; set; } = type == DataType.Text ? maxLength : 0; // 文字列型以外の時はMaxLengthを無効化

        /// <summary>
        /// SQLite用のデータ型定義を取得
        /// </summary>
        /// <returns>SQLiteで使用するデータ型定義</returns>
        public string GetSQLiteType()
        {
            return Type switch
            {
                DataType.Text => $"TEXT({MaxLength})",  // 文字列型には最大長を指定
                DataType.Integer => "INTEGER",
                DataType.Double => "REAL",
                DataType.Boolean => "INTEGER",         // SQLiteでは真偽値をINTEGERで扱うのが一般的
                DataType.DateTime => "TEXT",           // SQLiteでは日時型としてTEXT型を使用しISO 8601形式で扱う
                _ => throw new ArgumentOutOfRangeException()
            };
        }
    }
    #endregion
    #region class----テーブル情報を定義するクラス
    public class TableInfo(string tableName, List<ColumnInfo> columns)
    {
        public string TableName { get; set; } = tableName;              // テーブル名
        public List<ColumnInfo> Columns { get; set; } = columns;       // カラム定義（ColumnInfoのリスト）
    }
    #endregion

    public static class dbCM
    {
        public static string MakeInsertSQL<T>(string tblName, Dictionary<string, T> dic)
        {
            // null や空文字チェック
            if (string.IsNullOrWhiteSpace(tblName))
                throw new ArgumentException("テーブル名は空にできません。", nameof(tblName));

            if (dic == null || dic.Count == 0)
                throw new ArgumentException("カラムと値のペアを含む辞書が空です。", nameof(dic));

            // 列名と値を順次追加するリスト
            List<string> columnNames = new List<string>();
            List<string> values = new List<string>();

            foreach (var kvp in dic)
            {
                // カラム名をリストに追加
                columnNames.Add(kvp.Key);

                // 値の型で処理を分ける
                if (kvp.Value is string strValue)
                {
                    values.Add($"'{strValue}'"); // シングルクォーテーションで囲む
                }
                else if (kvp.Value is int || kvp.Value is float || kvp.Value is double)
                {
                    values.Add($"{kvp.Value}"); // 数値型はそのまま追加
                }
                else if (kvp.Value is DateTime dateTimeValue)
                {
                    // 日時型はフォーマットしてシングルクォーテーションで囲む
                    values.Add($"'{dateTimeValue:yyyy-MM-dd HH:mm:ss}'");
                }
                else if (kvp.Value is bool boolValue)
                {
                    // bool型はtrueを1、falseを0として追加
                    values.Add($"{(boolValue ? 1 : 0)}");
                }
                else
                {
                    // サポートされていない型の場合例外をスロー
                    throw new ArgumentException($"型 '{typeof(T)}' はサポートされていません。");
                }
            }

            // SQL文を組み立てる (string.Joinを活用して列名と値をカンマ区切りでまとめる)
            string sql = $"INSERT INTO {tblName} ({string.Join(", ", columnNames)}) VALUES ({string.Join(", ", values)});";

            return sql;
        }

    }


}


