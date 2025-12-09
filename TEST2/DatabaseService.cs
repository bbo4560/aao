using Dapper;
using Npgsql;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace TEST2
{
    public class DatabaseService
    {
        private readonly string _connectionString;

        public DatabaseService(string connectionString)
        {
            _connectionString = connectionString;
        }

        private IDbConnection CreateConnection()
        {
            return new NpgsqlConnection(_connectionString);
        }

        public static void EnsureDatabaseExists(string connectionString)
        {
            var builder = new NpgsqlConnectionStringBuilder(connectionString);
            var originalDb = builder.Database;

            builder.Database = "postgres";

            using var connection = new NpgsqlConnection(builder.ToString());
            connection.Open();

            var exists = connection.ExecuteScalar<bool>(
                "SELECT 1 FROM pg_database WHERE datname = @name",
                new { name = originalDb });

            if (!exists)
            {
                connection.Execute($"CREATE DATABASE \"{originalDb}\"");
            }
        }

        public async Task InitializeDatabaseAsync()
        {
            using var connection = CreateConnection();

            string createSql = @"
        CREATE TABLE IF NOT EXISTS operation_records (
            Id SERIAL PRIMARY KEY,
            Time TIMESTAMP,
            PanelID BIGINT,
            LOTID BIGINT,
            CarrierID BIGINT
        );
        
        CREATE TABLE IF NOT EXISTS system_logs (
            Id SERIAL PRIMARY KEY,
            OperationTime TIMESTAMP,
            UserName TEXT,
            MachineName TEXT,
            OperationType TEXT,
            AffectedData TEXT,
            DetailDescription TEXT
        );";
            await connection.ExecuteAsync(createSql);

            try
            {
                var checkSql = "SELECT data_type FROM information_schema.columns WHERE table_name = 'operation_records' AND column_name = 'panelid'";
                var dataType = await connection.QueryFirstOrDefaultAsync<string>(checkSql);

                if (dataType != null && (dataType.Equals("text", StringComparison.OrdinalIgnoreCase) || dataType.Contains("character", StringComparison.OrdinalIgnoreCase)))
                {
                    string alterSql = @"
                ALTER TABLE operation_records 
                ALTER COLUMN PanelID TYPE BIGINT USING PanelID::bigint,
                ALTER COLUMN LOTID TYPE BIGINT USING LOTID::bigint,
                ALTER COLUMN CarrierID TYPE BIGINT USING CarrierID::bigint;";
                    await connection.ExecuteAsync(alterSql);
                }
            }
            catch
            {
            }
        }


        public async Task<IEnumerable<OperationRecord>> GetAllAsync()
        {
            using var connection = CreateConnection();
            string sql = "SELECT * FROM operation_records ORDER BY PanelID ASC, LOTID ASC, CarrierID ASC, Time ASC";
            return await connection.QueryAsync<OperationRecord>(sql);
        }

        public async Task<IEnumerable<OperationRecord>> GetFilteredAsync(
            string? panelId, string? lotId, string? carrierId,
            DateTime? date, string? timeInput)
        {
            using var connection = CreateConnection();
            var sql = "SELECT * FROM operation_records WHERE 1=1";
            var parameters = new DynamicParameters();

            if (!string.IsNullOrWhiteSpace(panelId))
            {
                sql += " AND PanelID = @PanelID";
                parameters.Add("PanelID", long.Parse(panelId));
            }

            if (!string.IsNullOrWhiteSpace(lotId))
            {
                sql += " AND LOTID = @LOTID";
                parameters.Add("LOTID", long.Parse(lotId));
            }

            if (!string.IsNullOrWhiteSpace(carrierId))
            {
                sql += " AND CarrierID = @CarrierID";
                parameters.Add("CarrierID", long.Parse(carrierId));
            }

            if (date.HasValue)
            {
                sql += " AND Time >= @StartOfDay AND Time < @NextDay";
                parameters.Add("StartOfDay", date.Value.Date);
                parameters.Add("NextDay", date.Value.Date.AddDays(1));
            }

            if (!string.IsNullOrWhiteSpace(timeInput))
            {
                sql += " AND TO_CHAR(Time, 'HH24:MI:SS') LIKE @TimePattern";
                parameters.Add("TimePattern", $"{timeInput}%");
            }

            sql += " ORDER BY PanelID ASC, LOTID ASC, CarrierID ASC, Time ASC";

            return await connection.QueryAsync<OperationRecord>(sql, parameters);
        }



        public async Task InsertAsync(OperationRecord record)
        {
            using var connection = CreateConnection();
            string sql = @"INSERT INTO operation_records (Time, PanelID, LOTID, CarrierID) 
                   VALUES (@Time, @PanelID, @LOTID, @CarrierID)";
            await connection.ExecuteAsync(sql, record);

            await InsertLogAsync(new SystemLog
            {
                OperationTime = DateTime.Now,
                UserName = Environment.UserName,
                MachineName = Environment.MachineName,
                OperationType = "新增",
                AffectedData = $"Panel ID : {record.PanelID}",
                DetailDescription = $"Time: {record.Time}\nLOTID: {record.LOTID}\nCarrierID: {record.CarrierID}"
            });
        }


        public async Task UpdateAsync(OperationRecord record)
        {
            using var connection = CreateConnection();

            string selectSql = "SELECT * FROM operation_records WHERE Id = @Id";
            var oldRecord = await connection.QueryFirstOrDefaultAsync<OperationRecord>(selectSql, new { record.Id });

            string updateSql = @"UPDATE operation_records 
                 SET Time = @Time, PanelID = @PanelID, LOTID = @LOTID, CarrierID = @CarrierID 
                 WHERE Id = @Id";
            await connection.ExecuteAsync(updateSql, record);

            string panelIdChange = oldRecord != null
                ? $"{oldRecord.PanelID} ➜ {record.PanelID}"
                : $"{record.PanelID}";

            await InsertLogAsync(new SystemLog
            {
                OperationTime = DateTime.Now,
                UserName = Environment.UserName,
                MachineName = Environment.MachineName,
                OperationType = "修改",
                AffectedData = $"Panel ID : {panelIdChange}",
                DetailDescription = $"Time: {record.Time}\nLOTID: {record.LOTID}\nCarrierID: {record.CarrierID}"
            });
        }


        public async Task DeleteAsync(int id)
        {
            using var connection = CreateConnection();

            string selectSql = "SELECT * FROM operation_records WHERE Id = @id";
            var record = await connection.QueryFirstOrDefaultAsync<OperationRecord>(selectSql, new { id });

            if (record != null)
            {
                string sql = "DELETE FROM operation_records WHERE Id = @id";
                await connection.ExecuteAsync(sql, new { id });

                await InsertLogAsync(new SystemLog
                {
                    OperationTime = DateTime.Now,
                    UserName = Environment.UserName,
                    MachineName = Environment.MachineName,
                    OperationType = "刪除",
                    AffectedData = $"Panel ID : {record.PanelID}",
                    DetailDescription = $"Time: {record.Time}\nLOT ID : {record.LOTID}\nCarrier ID : {record.CarrierID}"
                });
            }
        }


        public async Task DeleteBatchAsync(IEnumerable<int> ids)
        {
            var idList = ids.ToList();
            if (!idList.Any()) return;

            using var connection = CreateConnection();
            connection.Open();
            using var transaction = connection.BeginTransaction();

            try
            {
                string selectSql = "SELECT PanelID FROM operation_records WHERE Id = ANY(@Ids)";
                var deletedPanelIds = await connection.QueryAsync<long>(selectSql, new { Ids = idList }, transaction);

                string deletedIdsStr = deletedPanelIds.Count() > 50
                    ? string.Join(", ", deletedPanelIds.Take(50)) + "..."
                    : string.Join(", ", deletedPanelIds);

                string deleteSql = "DELETE FROM operation_records WHERE Id = ANY(@Ids)";
                int count = await connection.ExecuteAsync(deleteSql, new { Ids = idList }, transaction);

                transaction.Commit();

                await InsertLogAsync(new SystemLog
                {
                    OperationTime = DateTime.Now,
                    UserName = Environment.UserName,
                    MachineName = Environment.MachineName,
                    OperationType = "批量刪除",
                    AffectedData = $"{count} 筆資料",
                    DetailDescription = $"PanelID: {deletedIdsStr}"
                });
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }


        public async Task<IEnumerable<SystemLog>> GetLogsAsync()
        {
            using var connection = CreateConnection();
            string sql = "SELECT * FROM system_logs ORDER BY OperationTime DESC";
            return await connection.QueryAsync<SystemLog>(sql);
        }

        public async Task InsertLogAsync(SystemLog log)
        {
            using var connection = CreateConnection();
            string sql = @"
        INSERT INTO system_logs
            (OperationTime, UserName, MachineName, OperationType, AffectedData, DetailDescription) 
        VALUES
            (@OperationTime, @UserName, @MachineName, @OperationType, @AffectedData, @DetailDescription)";
            await connection.ExecuteAsync(sql, log);
        }


        public async Task ClearAllLogsAsync()
        {
            using var conn = CreateConnection();
            string sql = "TRUNCATE TABLE system_logs RESTART IDENTITY;";
            await conn.ExecuteAsync(sql);
        }

        public async Task ExportToExcelAsync(IEnumerable<OperationRecord> data, string filePath)
        {
            if (data == null || !data.Any()) throw new ArgumentException("無資料可匯出");

            ExcelPackage.License.SetNonCommercialPersonal("DemoUser");

            try
            {
                using var package = new ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add("Operations");

                worksheet.Cells[1, 1].Value = "Time";
                worksheet.Cells[1, 2].Value = "PanelID";
                worksheet.Cells[1, 3].Value = "LOTID";
                worksheet.Cells[1, 4].Value = "CarrierID";

                int row = 2;
                foreach (var item in data)
                {
                    worksheet.Cells[row, 1].Value = item.Time.ToString("yyyy/MM/dd HH:mm:ss");
                    worksheet.Cells[row, 2].Value = item.PanelID;
                    worksheet.Cells[row, 3].Value = item.LOTID;
                    worksheet.Cells[row, 4].Value = item.CarrierID;
                    row++;
                }

                worksheet.Cells.AutoFitColumns();
                File.WriteAllBytes(filePath, package.GetAsByteArray());

                await InsertLogAsync(new SystemLog
                {
                    OperationTime = DateTime.Now,
                    UserName = Environment.UserName,
                    MachineName = Environment.MachineName,
                    OperationType = "匯出 Excel",
                    AffectedData = $"{Path.GetFileName(filePath)}\n{row - 2} 筆記錄",
                    DetailDescription = $"檔案已儲存至: {Path.GetFullPath(filePath)}"
                });
            }
            catch (Exception ex)
            {
                await InsertLogAsync(new SystemLog
                {
                    OperationTime = DateTime.Now,
                    UserName = Environment.UserName,
                    MachineName = Environment.MachineName,
                    OperationType = "匯出失敗",
                    AffectedData = "N/A",
                    DetailDescription = $"路徑: {filePath}, 錯誤: {ex.Message}"
                });
                throw;
            }
        }


        public async Task<string> ImportFromExcelAsync(string filePath)
        {
            if (!File.Exists(filePath)) return "檔案不存在";
            ExcelPackage.License.SetNonCommercialPersonal("DemoUser");

            var newRecords = new List<OperationRecord>();
            int skipped = 0;
            int formatError = 0;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet?.Dimension == null) return "Excel 無資料";

                int rowCount = worksheet.Dimension.Rows;
                for (int row = 2; row <= rowCount; row++)
                {
                    var val1 = worksheet.Cells[row, 1].Text;
                    var val2 = worksheet.Cells[row, 2].Text;
                    var val3 = worksheet.Cells[row, 3].Text;
                    var val4 = worksheet.Cells[row, 4].Text;

                    if (string.IsNullOrWhiteSpace(val1) && string.IsNullOrWhiteSpace(val2) &&
                        string.IsNullOrWhiteSpace(val3) && string.IsNullOrWhiteSpace(val4))
                    {
                        continue;
                    }

                    try
                    {
                        var time = worksheet.Cells[row, 1].GetValue<DateTime?>();
                        var panelId = worksheet.Cells[row, 2].GetValue<long?>();
                        var lotId = worksheet.Cells[row, 3].GetValue<long?>();
                        var carrierId = worksheet.Cells[row, 4].GetValue<long?>();

                        if (time.HasValue && panelId.HasValue && lotId.HasValue && carrierId.HasValue)
                        {
                            newRecords.Add(new OperationRecord
                            {
                                Time = time.Value,
                                PanelID = panelId.Value,
                                LOTID = lotId.Value,
                                CarrierID = carrierId.Value
                            });
                        }
                        else
                        {
                            skipped++;
                        }
                    }
                    catch
                    {
                        formatError++;
                    }
                }
            }

            if (!newRecords.Any() && formatError == 0) return "無有效資料";

            using var connection = CreateConnection();
            connection.Open();

            var minTime = newRecords.Any() ? newRecords.Min(r => r.Time) : DateTime.MinValue;
            var maxTime = newRecords.Any() ? newRecords.Max(r => r.Time) : DateTime.MaxValue;

            var existingSet = new HashSet<string>();
            if (newRecords.Any())
            {
                string fetchSql = @"SELECT Time, PanelID, LOTID, CarrierID 
                            FROM operation_records 
                            WHERE Time BETWEEN @MinTime AND @MaxTime";

                var existingData = await connection.QueryAsync<OperationRecord>(fetchSql, new { MinTime = minTime, MaxTime = maxTime });

                existingSet = new HashSet<string>(existingData.Select(r =>
                    $"{r.Time:yyyyMMddHHmmss}|{r.PanelID}|{r.LOTID}|{r.CarrierID}"));
            }

            var recordsToInsert = new List<OperationRecord>();
            var currentBatchSet = new HashSet<string>();
            int duplicates = 0;
            int added = 0;

            foreach (var rec in newRecords)
            {
                string key = $"{rec.Time:yyyyMMddHHmmss}|{rec.PanelID}|{rec.LOTID}|{rec.CarrierID}";
                if (existingSet.Contains(key) || currentBatchSet.Contains(key))
                {
                    duplicates++;
                    skipped++;
                }
                else
                {
                    recordsToInsert.Add(rec);
                    currentBatchSet.Add(key);
                    added++;
                }
            }

            if (recordsToInsert.Any())
            {
                using var transaction = connection.BeginTransaction();
                try
                {
                    string insertSql = @"INSERT INTO operation_records (Time, PanelID, LOTID, CarrierID) 
                                 VALUES (@Time, @PanelID, @LOTID, @CarrierID)";
                    await connection.ExecuteAsync(insertSql, recordsToInsert, transaction);
                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            }

            string resultSummary = $"新增: {added}, 失敗: {formatError}, 略過: {skipped} (含重複)";

            await InsertLogAsync(new SystemLog
            {
                OperationTime = DateTime.Now,
                UserName = Environment.UserName,
                MachineName = Environment.MachineName,
                OperationType = "匯入",
                AffectedData = $"{Path.GetFileName(filePath)}\n{added} 筆",
                DetailDescription = $"路徑: {Path.GetFullPath(filePath)}\n結果: {resultSummary}"
            });


            return resultSummary;
        }

        public void ExportPanelRecordsToDb(string filePath, IEnumerable<OperationRecord> records)
        {
            SQLiteConnection.CreateFile(filePath);

            var connectionString = $"Data Source={filePath};Version=3;";
            using var connection = new SQLiteConnection(connectionString);
            connection.Open();

            using var transaction = connection.BeginTransaction();
            try
            {
                string createTableQuery = @"
            CREATE TABLE PanelRecords (
                Time TEXT,
                PanelID INTEGER,
                LOTID INTEGER,
                CarrierID INTEGER
            );";

                using (var createCmd = new SQLiteCommand(createTableQuery, connection))
                {
                    createCmd.ExecuteNonQuery();
                }

                string insertQuery = @"
            INSERT INTO PanelRecords (Time, PanelID, LOTID, CarrierID)
            VALUES (@Time, @PanelID, @LOTID, @CarrierID)";

                using (var insertCmd = new SQLiteCommand(insertQuery, connection))
                {
                    insertCmd.Parameters.Add(new SQLiteParameter("@Time"));
                    insertCmd.Parameters.Add(new SQLiteParameter("@PanelID"));
                    insertCmd.Parameters.Add(new SQLiteParameter("@LOTID"));
                    insertCmd.Parameters.Add(new SQLiteParameter("@CarrierID"));

                    foreach (var record in records)
                    {
                        insertCmd.Parameters["@Time"].Value = record.Time.ToString("yyyy-MM-dd HH:mm:ss");
                        insertCmd.Parameters["@PanelID"].Value = record.PanelID;
                        insertCmd.Parameters["@LOTID"].Value = record.LOTID;
                        insertCmd.Parameters["@CarrierID"].Value = record.CarrierID;

                        insertCmd.ExecuteNonQuery();
                    }
                }

                transaction.Commit();
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }
    }
}


