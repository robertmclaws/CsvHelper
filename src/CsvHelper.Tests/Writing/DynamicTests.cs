using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CsvHelper.Tests.Writing
{
	[TestClass]
	public class DynamicTests
	{
		[TestMethod]
		public void WriteDynamicExpandoObjectsTest()
		{
			using( var stream = new MemoryStream() )
			using( var reader = new StreamReader( stream ) )
			using( var writer = new StreamWriter( stream ) )
			using( var csv = new CsvWriter( writer ) )
			{
				var list = new List<dynamic>();
				dynamic obj = new ExpandoObject();
				obj.Id = 1;
				obj.Name = "one";
				list.Add( obj );

				csv.WriteRecords( list );
				writer.Flush();
				stream.Position = 0;

				var expected = "Id,Name\r\n";
				expected += "1,one\r\n";

				Assert.AreEqual( expected, reader.ReadToEnd() );
			}
		}

		[TestMethod]
		public void WriteDynamicExpandoObjectTest()
		{
			using( var stream = new MemoryStream() )
			using( var reader = new StreamReader( stream ) )
			using( var writer = new StreamWriter( stream ) )
			using( var csv = new CsvWriter( writer ) )
			{
				dynamic obj = new ExpandoObject();
				obj.Id = 1;
				obj.Name = "one";

				csv.WriteRecord( obj );
                csv.NextRecord();
                writer.Flush();
				stream.Position = 0;

				var expected = "Id,Name\r\n";
				expected += "1,one\r\n";

				Assert.AreEqual( expected, reader.ReadToEnd() );
			}
		}

        [TestMethod]
        public void ReadDynamicExpandoObjectsThenWriteBackTest()
        {
            List<dynamic> items = new List<dynamic>();

            using (var reader = File.OpenText(@"C:\Users\rober\Documents\Lead.csv"))
            {
                using (var csvReader = new CsvReader(reader))
                {
                    var records = csvReader.GetRecords<dynamic>().ToList();
                    items.AddRange(records);
                }
            }

            Assert.AreNotEqual(items.Count, 0);

            using (var output = File.OpenWrite(@"C:\Users\rober\Documents\Lead-Output.csv"))
            {
                using (var writer = new StreamWriter(output))
                {
                    using (var csvWriter = new CsvWriter(writer))
                    {
                        csvWriter.Configuration.QuoteAllFields = true;
                        csvWriter.WriteRecords(items);
                    }
                }
            }

            

        }

    }
}
