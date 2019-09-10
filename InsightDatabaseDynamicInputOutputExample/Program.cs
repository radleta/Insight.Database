using System.Linq;
using Insight.Database;
using Insight.Database.Providers;
using Insight.Database.Structure;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;

namespace InsightDatabaseDynamicInputOutputExample
{
    class Program
    {
        

        static void Main(string[] args)
        {            
			// Note: You need to set your environment variable to a valid connection string and restart VS.NET before debugging for this to work.
            string ConnectionString = Environment.GetEnvironmentVariable("InsightDatabaseDynamicInputOutputExample.ConnectionString") ?? throw new ApplicationException("Please define the InsightDatabaseDynamicInputOutputExample.ConnectionString environment variable.");

            const string SchemaSqlScript = @"
                IF EXISTS ( SELECT * 
                    FROM   sysobjects 
                    WHERE  id = object_id(N'[dbo].[sp_get_InsightDatabaseExample1]') 
                           and OBJECTPROPERTY(id, N'IsProcedure') = 1 )
                BEGIN
                   DROP PROCEDURE dbo.sp_get_InsightDatabaseExample1;
                END;
                GO;
                CREATE PROCEDURE dbo.sp_get_InsightDatabaseExample1
                AS
                BEGIN
                    SELECT 'StringValue' InputString,
                        GETDATE() InputDate;
                END;
                GO;

                IF EXISTS ( SELECT * 
                    FROM   sysobjects 
                    WHERE  id = object_id(N'[dbo].[sp_write_InsightDatabaseExample1]') 
                           and OBJECTPROPERTY(id, N'IsProcedure') = 1 )
                BEGIN
                   DROP PROCEDURE dbo.sp_write_InsightDatabaseExample1;
                END;
                GO;

                IF EXISTS(SELECT 1 FROM sys.types WHERE name = 'InsightDatabaseExample1InputType' AND is_table_type = 1 AND SCHEMA_ID('dbo') = schema_id)
                BEGIN
                   DROP TYPE dbo.InsightDatabaseExample1InputType;
                END;
                GO;

                CREATE TYPE dbo.InsightDatabaseExample1InputType AS TABLE (
                    InputString VARCHAR(50) NOT NULL,
                    InputDate DATETIME NOT NULL,
                    OutputValue INT NOT NULL
                );
                GO;

                CREATE PROCEDURE dbo.sp_write_InsightDatabaseExample1(
                    @Input dbo.InsightDatabaseExample1InputType READONLY
                )
                AS
                BEGIN
                    
                    SET NOCOUNT ON;

                    -- return the data so we can check it
					SELECT *
					FROM @Input I;

                    RETURN 0;

                END;
            ";
            
            // register the sql provider
            Insight.Database.SqlInsightDbProvider.RegisterProvider();

            // create all the schema
            foreach (var sql in SchemaSqlScript.Split(new string[] { "GO;" }, StringSplitOptions.RemoveEmptyEntries))
            {
                new SqlConnection(ConnectionString).ExecuteSql(sql);
            }

            // call the read procedure
			var outProperties = new Dictionary<string, Type>();
            var rows = new SqlConnection(ConnectionString).ExecuteAndAutoClose(
                c => c.CreateCommand(sql: "dbo.sp_get_InsightDatabaseExample1"),
                (cmd, r) =>
                {
					// get field information when needed
					// Note: You could cache these fields depending on the frequency of the changes to the procedure.
					for (var i = 0; i < r.FieldCount; i++)
					{
						var name = r.GetName(i);
						if (!outProperties.ContainsKey(name))
							outProperties.Add(name, r.GetFieldType(i));
					}

                    return ListReader<FastExpando>.Default.Read(cmd, r);
                },
				true);

            // modify the results by using then adding data
            foreach (var row in rows.Cast<IDictionary<string, object>>())
            {
                row.Add("OutputValue", 1);
            }

			// add the output types
			outProperties.Add("OutputValue", typeof(int));
			
			// build our dynamic type
			// HACK: You would need to cache this type based on the contents of the outProperties since its expensive.
			var outType = RuntimeTypeBuilder.CompileResultTypeInfo("MyAssembly", "MyModule", "MyType", outProperties).AsType();

			// create an array of our output type
			// the type of the array tells insight how to map it
			var outRows = Array.CreateInstance(outType, rows.Count);

			// convert each row to our output type
			// and store in the output array
			for (var i = 0; i < rows.Count; i++)
			{
				var row = rows[i];
				var outRow = ExpandoConvert.ChangeType(row, outType);
				outRows.SetValue(outRow, i);
			}

            // call the write procedure
            var returnedRows = new SqlConnection(ConnectionString).Query(sql: "dbo.sp_write_InsightDatabaseExample1", parameters: outRows);

			// now lets double check that all the data we sent was received and returned			
			Debug.Assert(returnedRows != null, "Expected to get a valid instance back");
			Debug.Assert(returnedRows.Count == 1, "Expected to get back one row");
			var inputRow = rows[0];
			var returnedRow = returnedRows[0];
			Debug.Assert((string)inputRow["InputString"] == (string)returnedRow["InputString"], "Expected InputString to match the input");
			Debug.Assert((DateTime)inputRow["InputDate"] == (DateTime)returnedRow["InputDate"], "Expected InputDate to match the input");
			Debug.Assert((int)inputRow["OutputValue"] == (int)returnedRow["OutputValue"], "Expected OutputValue to match the input");
        }
    }
}
