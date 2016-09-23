﻿// Copyright 2009-2015 Josh Close and Contributors
// This file is a part of CsvHelper and is dual licensed under MS-PL and Apache 2.0.
// See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
// http://csvhelper.com
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using System.Linq;
using System.Runtime.CompilerServices;
#if !NET_2_0
using System.Linq.Expressions;
#endif
#if !NET_2_0 && !NET_3_5 && !PCL
using System.Dynamic;
using Microsoft.CSharp.RuntimeBinder;
#endif

#pragma warning disable 649
#pragma warning disable 169

namespace CsvHelper
{
	/// <summary>
	/// Used to write CSV files.
	/// </summary>
	public class CsvWriter : ICsvWriter
	{
		private readonly bool leaveOpen;
		private bool disposed;
		private readonly List<string> currentRecord = new List<string>();
		private ICsvSerializer serializer;
		private bool hasHeaderBeenWritten;
		private bool hasRecordBeenWritten;
#if !NET_2_0
		private readonly Dictionary<Type, Delegate> typeActions = new Dictionary<Type, Delegate>();
#endif
		private readonly CsvConfiguration configuration;
		private bool hasExcelSeperatorBeenRead;
		private int row = 1;

		/// <summary>
		/// Gets the current row.
		/// </summary>
		public virtual int Row => row;

		/// <summary>
		/// Get the current record;
		/// </summary>
		public virtual List<string> CurrentRecord => currentRecord;

		/// <summary>
		/// Gets the configuration.
		/// </summary>
		public virtual CsvConfiguration Configuration => configuration;

		/// <summary>
		/// Creates a new CSV writer using the given <see cref="TextWriter" />.
		/// </summary>
		/// <param name="writer">The writer used to write the CSV file.</param>
		public CsvWriter( TextWriter writer ) : this( new CsvSerializer( writer, new CsvConfiguration() ), false ) { }

		/// <summary>
		/// Creates a new CSV writer using the given <see cref="TextWriter"/>.
		/// </summary>
		/// <param name="writer">The writer used to write the CSV file.</param>
		/// <param name="leaveOpen">true to leave the reader open after the CsvReader object is disposed, otherwise false.</param>
		public CsvWriter( TextWriter writer, bool leaveOpen ) : this( new CsvSerializer( writer, new CsvConfiguration() ), leaveOpen ) { }

		/// <summary>
		/// Creates a new CSV writer using the given <see cref="TextWriter"/>.
		/// </summary>
		/// <param name="writer">The <see cref="StreamWriter"/> use to write the CSV file.</param>
		/// <param name="configuration">The configuration.</param>
		public CsvWriter( TextWriter writer, CsvConfiguration configuration ) : this( new CsvSerializer( writer, configuration ), false ) { }

		/// <summary>
		/// Creates a new CSV writer using the given <see cref="ICsvSerializer"/>.
		/// </summary>
		/// <param name="serializer">The serializer.</param>
		public CsvWriter( ICsvSerializer serializer ) : this( serializer, false ) { }

		/// <summary>
		/// Creates a new CSV writer using the given <see cref="ICsvSerializer"/>.
		/// </summary>
		/// <param name="serializer">The serializer.</param>
		/// <param name="leaveOpen">true to leave the reader open after the CsvReader object is disposed, otherwise false.</param>
		public CsvWriter( ICsvSerializer serializer, bool leaveOpen )
		{
			if( serializer == null )
			{
				throw new ArgumentNullException( nameof( serializer ) );
			}

			if( serializer.Configuration == null )
			{
				throw new CsvConfigurationException( "The given serializer has no configuration." );
			}

			this.serializer = serializer;
			configuration = serializer.Configuration;
			this.leaveOpen = leaveOpen;
		}

		/// <summary>
		/// Writes a field that has already been converted to a
		/// <see cref="string"/> from an <see cref="ITypeConverter"/>.
		/// If the field is null, it won't get written. A type converter 
		/// will always return a string, even if field is null. If the 
		/// converter returns a null, it means that the converter has already
		/// written data, and the returned value should not be written.
		/// </summary>
		/// <param name="field">The converted field to write.</param>
		public virtual void WriteConvertedField( string field )
		{
			if( field == null )
			{
				return;
			}

			WriteField( field );
		}

		/// <summary>
		/// Writes the field to the CSV file. The field
		/// may get quotes added to it.
		/// When all fields are written for a record,
		/// <see cref="ICsvWriter.NextRecord" /> must be called
		/// to complete writing of the current record.
		/// </summary>
		/// <param name="field">The field to write.</param>
		public virtual void WriteField( string field )
		{
			var shouldQuote = configuration.QuoteAllFields;

			if( field != null && configuration.TrimFields )
			{
				field = field.Trim();
			}

			if( !configuration.QuoteNoFields && !string.IsNullOrEmpty( field ) )
			{
			    var hasQuote = field.Contains( configuration.QuoteString );
				
                if( shouldQuote
				    || hasQuote
				    || field[0] == ' '
				    || field[field.Length - 1] == ' '
				    || field.IndexOfAny( configuration.QuoteRequiredChars ) > -1
				    || ( configuration.Delimiter.Length > 1 && field.Contains( configuration.Delimiter ) ) )
				{
					shouldQuote = true;
				}
			}

			WriteField( field, shouldQuote );
		}

		/// <summary>
		/// Writes the field to the CSV file. This will
		/// ignore any need to quote and ignore the
		/// <see cref="CsvConfiguration.QuoteAllFields"/>
		/// and just quote based on the shouldQuote
		/// parameter.
		/// When all fields are written for a record,
		/// <see cref="ICsvWriter.NextRecord" /> must be called
		/// to complete writing of the current record.
		/// </summary>
		/// <param name="field">The field to write.</param>
		/// <param name="shouldQuote">True to quote the field, otherwise false.</param>
		public virtual void WriteField( string field, bool shouldQuote )
		{
            // All quotes must be doubled.       
			if( shouldQuote && !string.IsNullOrEmpty( field ) )
			{
				field = field.Replace( configuration.QuoteString, configuration.DoubleQuoteString );
			}

			if( configuration.UseExcelLeadingZerosFormatForNumerics && !string.IsNullOrEmpty( field ) && field[0] == '0' && field.All( Char.IsDigit ) )
			{
				field = "=" + configuration.Quote + field + configuration.Quote;
			}
            else if (shouldQuote)
            {
                field = configuration.Quote + field + configuration.Quote;
            }

			currentRecord.Add( field );
		}

		/// <summary>
		/// Writes the field to the CSV file.
		/// When all fields are written for a record,
		/// <see cref="ICsvWriter.NextRecord" /> must be called
		/// to complete writing of the current record.
		/// </summary>
		/// <typeparam name="T">The type of the field.</typeparam>
		/// <param name="field">The field to write.</param>
		public virtual void WriteField<T>( T field )
		{
			var type = field == null ? typeof( string ) : field.GetType();
			var converter = TypeConverterFactory.GetConverter( type );
			WriteField( field, converter );
		}

		/// <summary>
		/// Writes the field to the CSV file.
		/// When all fields are written for a record,
		/// <see cref="ICsvWriter.NextRecord" /> must be called
		/// to complete writing of the current record.
		/// </summary>
		/// <typeparam name="T">The type of the field.</typeparam>
		/// <param name="field">The field to write.</param>
		/// <param name="converter">The converter used to convert the field into a string.</param>
		public virtual void WriteField<T>( T field, ITypeConverter converter )
		{
			var type = field == null ? typeof( string ) : field.GetType();
			var propertyMapData = new CsvPropertyMapData( null )
			{
				TypeConverter = converter,
				TypeConverterOptions = { CultureInfo = configuration.CultureInfo }
			};
			propertyMapData.TypeConverterOptions = TypeConverterOptions.Merge( propertyMapData.TypeConverterOptions, configuration.TypeConverterOptionsFactory.GetOptions( type ) );

			var fieldString = converter.ConvertToString( field, this, propertyMapData );
			WriteConvertedField( fieldString );
		}

		/// <summary>
		/// Writes the field to the CSV file
		/// using the given <see cref="ITypeConverter"/>.
		/// When all fields are written for a record,
		/// <see cref="ICsvWriter.NextRecord" /> must be called
		/// to complete writing of the current record.
		/// </summary>
		/// <typeparam name="T">The type of the field.</typeparam>
		/// <typeparam name="TConverter">The type of the converter.</typeparam>
		/// <param name="field">The field to write.</param>
		public virtual void WriteField<T, TConverter>( T field )
		{
			var converter = TypeConverterFactory.GetConverter<TConverter>();
			WriteField( field, converter );
		}

	    /// <summary>
	    /// Ends writing of the current record and starts a new record. 
	    /// This needs to be called to serialize the row to the writer.
	    /// </summary>
	    public virtual void NextRecord()
		{
	        try
	        {
	            serializer.Write( currentRecord.ToArray() );
	            currentRecord.Clear();
	            row++;
	        }
	        catch( Exception ex )
	        {
	            var csvHelperException = ex as CsvHelperException ?? new CsvHelperException( "An unexpected error occurred.", ex );
	            ExceptionHelper.AddExceptionData( csvHelperException, Row, null, null, null, currentRecord.ToArray() );
	            throw csvHelperException;
	        }
		}

        /// <summary>
        /// Write the Excel seperator record.
        /// </summary>
        public virtual void WriteExcelSeparator()
		{
			if( hasHeaderBeenWritten )
			{
				throw new CsvWriterException( "The Excel seperator record must be the first record written in the file." );
			}

			if( hasRecordBeenWritten )
			{
				throw new CsvWriterException( "The Excel seperator record must be the first record written in the file." );
			}

			WriteField( "sep=" + configuration.Delimiter, false );
		}

	    /// <summary>
	    /// Writes a comment.
	    /// </summary>
	    /// <param name="comment">The comment to write.</param>
	    public virtual void WriteComment( string comment )
	    {
	        WriteField( configuration.Comment + comment );
	    }

#if !NET_2_0

        /// <summary>
        /// Writes the header record from the given properties.
        /// </summary>
        /// <typeparam name="T">The type of the record.</typeparam>
        public virtual void WriteHeader<T>()
		{
			WriteHeader( typeof( T ) );
		}

		/// <summary>
		/// Writes the header record from the given properties.
		/// </summary>
		/// <param name="type">The type of the record.</param>
		public virtual void WriteHeader( Type type )
		{
			if( type == null )
			{
				throw new ArgumentNullException( nameof( type ) );
			}

			if( !configuration.HasHeaderRecord )
			{
				throw new CsvWriterException( "Configuration.HasHeaderRecord is false. This will need to be enabled to write the header." );
			}

			if( hasHeaderBeenWritten )
			{
				throw new CsvWriterException( "The header record has already been written. You can't write it more than once." );
			}

			if( hasRecordBeenWritten )
			{
				throw new CsvWriterException( "Records have already been written. You can't write the header after writing records has started." );
			}

			if( type == typeof( Object ) )
			{
				return;
			}

			if( configuration.Maps[type] == null )
			{
				configuration.Maps.Add( configuration.AutoMap( type ) );
			}

			var properties = new CsvPropertyMapCollection();
			AddProperties( properties, configuration.Maps[type] );

			foreach( var property in properties )
			{
				if( CanWrite( property ) )
				{
					WriteField( property.Data.Names.FirstOrDefault() );
				}
			}

			hasHeaderBeenWritten = true;
		}

#if !NET_2_0 && !NET_3_5 && !PCL

		/// <summary>
		/// Writes the header record for the given dynamic object.
		/// </summary>
		/// <param name="record">The dynamic record to write.</param>
		public virtual void WriteDynamicHeader( IDynamicMetaObjectProvider record )
		{
			if( record == null )
			{
				throw new ArgumentNullException( nameof( record ) );
			}

			if( !configuration.HasHeaderRecord )
			{
				throw new CsvWriterException( "Configuration.HasHeaderRecord is false. This will need to be enabled to write the header." );
			}

			if( hasHeaderBeenWritten )
			{
				throw new CsvWriterException( "The header record has already been written. You can't write it more than once." );
			}

			if( hasRecordBeenWritten )
			{
				throw new CsvWriterException( "Records have already been written. You can't write the header after writing records has started." );
			}

			var provider = (IDynamicMetaObjectProvider)record;
			var metaObject = provider.GetMetaObject( Expression.Constant( provider ) );
			var names = metaObject.GetDynamicMemberNames();
			foreach( var name in names )
			{
				WriteField( name );
			}

			hasHeaderBeenWritten = true;
		}

#endif

		/// <summary>
		/// Writes the record to the CSV file.
		/// </summary>
		/// <typeparam name="T">The type of the record.</typeparam>
		/// <param name="record">The record to write.</param>
		public virtual void WriteRecord<T>( T record )
		{
#if !NET_2_0 && !NET_3_5 && !PCL
			var dynamicRecord = record as IDynamicMetaObjectProvider;
			if( dynamicRecord != null )
			{
				if( configuration.HasHeaderRecord && !hasHeaderBeenWritten )
				{
					WriteDynamicHeader( dynamicRecord );
				    NextRecord();
				}

				if( !typeActions.ContainsKey( dynamicRecord.GetType() ) )
				{
					CreateActionForDynamic( dynamicRecord );
				}
			}
#endif

			try
			{
				GetWriteRecordAction<T>()( record );
                hasRecordBeenWritten = true;
            }
            catch( Exception ex )
			{
				var csvHelperException = ex as CsvHelperException ?? new CsvHelperException( "An unexpected error occurred.", ex );
				ExceptionHelper.AddExceptionData( csvHelperException, Row, record.GetType(), null, null, currentRecord.ToArray() );

				throw csvHelperException;
			}
		}

		/// <summary>
		/// Writes the list of records to the CSV file.
		/// </summary>
		/// <param name="records">The list of records to write.</param>
		public virtual void WriteRecords<T>( IEnumerable<T> records )
		{
			try
			{
				if( configuration.HasExcelSeparator && !hasExcelSeperatorBeenRead )
				{
					WriteExcelSeparator();
                    NextRecord();
                    hasExcelSeperatorBeenRead = true;
				}

                // Write the header. If records is a List<dynamic>, the header won't be written.
                // This is because typeof( T ) = Object.

                var isDynamic = false;
 
#if !NET_2_0 && !NET_3_5 && !PCL
                var dynamicRecord = records.FirstOrDefault() as IDynamicMetaObjectProvider;
                isDynamic = dynamicRecord != null;
#endif
                var isPrimitive = typeof(T).GetTypeInfo().IsPrimitive;

                if ( configuration.HasHeaderRecord && !hasHeaderBeenWritten )
				{

#if !NET_2_0 && !NET_3_5 && !PCL
                    if ( dynamicRecord != null )
                    {
                        WriteDynamicHeader(dynamicRecord);
                    }
#endif

                    if ( !hasHeaderBeenWritten && !isPrimitive )
                    {
                        WriteHeader(typeof(T));
                    }

                        if ( hasHeaderBeenWritten )
                    {
                        NextRecord();
                    }
				}

				foreach( var record in records )
				{

#if !NET_2_0 && !NET_3_5 && !PCL
                    if (isDynamic)
                    {
                        if (!typeActions.ContainsKey( typeof(T) ) && !typeActions.ContainsKey( typeof(ExpandoObject) ) )
                        {
                            CreateActionForDynamic( record as IDynamicMetaObjectProvider );
                        }
                    }
#endif

                    try
					{
						GetWriteRecordAction( record.GetType() ).DynamicInvoke( record );
                    }
                    catch( TargetInvocationException ex )
					{
						throw ex.InnerException;
					}

					NextRecord();
				}
			}
			catch( Exception ex )
			{
				var csvHelperException = ex as CsvHelperException ?? new CsvHelperException( "An unexpected error occurred.", ex );
				ExceptionHelper.AddExceptionData( csvHelperException, Row, typeof(T), null, null, currentRecord.ToArray() );

				throw csvHelperException;
			}
		}

		/// <summary>
		/// Clears the record cache for the given type. After <see cref="ICsvWriterRow.WriteRecord{T}"/> is called the
		/// first time, code is dynamically generated based on the <see cref="CsvPropertyMapCollection"/>,
		/// compiled, and stored for the given type T. If the <see cref="CsvPropertyMapCollection"/>
		/// changes, <see cref="ICsvWriter.ClearRecordCache{T}"/> needs to be called to update the
		/// record cache.
		/// </summary>
		/// <typeparam name="T">The record type.</typeparam>
		public virtual void ClearRecordCache<T>()
		{
			ClearRecordCache( typeof( T ) );
		}

		/// <summary>
		/// Clears the record cache for the given type. After <see cref="ICsvWriterRow.WriteRecord{T}"/> is called the
		/// first time, code is dynamically generated based on the <see cref="CsvPropertyMapCollection"/>,
		/// compiled, and stored for the given type T. If the <see cref="CsvPropertyMapCollection"/>
		/// changes, <see cref="ICsvWriter.ClearRecordCache(System.Type)"/> needs to be called to update the
		/// record cache.
		/// </summary>
		/// <param name="type">The record type.</param>
		public virtual void ClearRecordCache( Type type )
		{
			typeActions.Remove( type );
		}

		/// <summary>
		/// Clears the record cache for all types. After <see cref="ICsvWriterRow.WriteRecord{T}"/> is called the
		/// first time, code is dynamically generated based on the <see cref="CsvPropertyMapCollection"/>,
		/// compiled, and stored for the given type T. If the <see cref="CsvPropertyMapCollection"/>
		/// changes, <see cref="ICsvWriter.ClearRecordCache()"/> needs to be called to update the
		/// record cache.
		/// </summary>
		public virtual void ClearRecordCache()
		{
			typeActions.Clear();
		}

		/// <summary>
		/// Adds the properties from the mapping. This will recursively
		/// traverse the mapping tree and add all properties for
		/// reference maps.
		/// </summary>
		/// <param name="properties">The properties to be added to.</param>
		/// <param name="mapping">The mapping where the properties are added from.</param>
		protected virtual void AddProperties( CsvPropertyMapCollection properties, CsvClassMap mapping )
		{
			properties.AddRange( mapping.PropertyMaps );
			foreach( var refMap in mapping.ReferenceMaps )
			{
				AddProperties( properties, refMap.Data.Mapping );
			}
		}

		/// <summary>
		/// Creates a property expression for the given property on the record.
		/// This will recursively traverse the mapping to find the property
		/// and create a safe property accessor for each level as it goes.
		/// </summary>
		/// <param name="recordExpression">The current property expression.</param>
		/// <param name="mapping">The mapping to look for the property to map on.</param>
		/// <param name="propertyMap">The property map to look for on the mapping.</param>
		/// <returns>An Expression to access the given property.</returns>
		protected virtual Expression CreatePropertyExpression( Expression recordExpression, CsvClassMap mapping, CsvPropertyMap propertyMap )
		{
			if( mapping.PropertyMaps.Any( pm => pm == propertyMap ) )
			{
				// The property is on this level.
				return Expression.Property( recordExpression, propertyMap.Data.Property );
			}

			// The property isn't on this level of the mapping.
			// We need to search down through the reference maps.
			foreach( var refMap in mapping.ReferenceMaps )
			{
				var wrapped = Expression.Property( recordExpression, refMap.Data.Property );
				var propertyExpression = CreatePropertyExpression( wrapped, refMap.Data.Mapping, propertyMap );
				if( propertyExpression == null )
				{
					continue;
				}

				if( refMap.Data.Property.PropertyType.GetTypeInfo().IsValueType )
				{
					return propertyExpression;
				}

				var nullCheckExpression = Expression.Equal( wrapped, Expression.Constant( null ) );

				var isValueType = propertyMap.Data.Property.PropertyType.GetTypeInfo().IsValueType;
				var isGenericType = isValueType && propertyMap.Data.Property.PropertyType.GetTypeInfo().IsGenericType;
				Type propertyType;
				if( isValueType && !isGenericType && !configuration.UseNewObjectForNullReferenceProperties )
				{
					propertyType = typeof( Nullable<> ).MakeGenericType( propertyMap.Data.Property.PropertyType );
					propertyExpression = Expression.Convert( propertyExpression, propertyType );
				}
				else
				{
					propertyType = propertyMap.Data.Property.PropertyType;
				}

				var defaultValueExpression = isValueType && !isGenericType
					? (Expression)Expression.New( propertyType )
					: Expression.Constant( null, propertyType );
				var conditionExpression = Expression.Condition( nullCheckExpression, defaultValueExpression, propertyExpression );
				return conditionExpression;
			}

			return null;
		}

#endif

		/// <summary>
		/// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
		/// </summary>
		/// <filterpriority>2</filterpriority>
		public void Dispose()
		{
			Dispose( !leaveOpen );
			GC.SuppressFinalize( this );
		}

		/// <summary>
		/// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
		/// </summary>
		/// <param name="disposing">True if the instance needs to be disposed of.</param>
		protected virtual void Dispose( bool disposing )
		{
			if( disposed )
			{
				return;
			}

			if( disposing )
			{
				serializer?.Dispose();
			}

			disposed = true;
			serializer = null;
		}

#if !NET_2_0

		/// <summary>
		/// Gets the action delegate used to write the custom
		/// class object to the writer.
		/// </summary>
		/// <typeparam name="T">The type of the custom class being written.</typeparam>
		/// <returns>The action delegate.</returns>
		protected virtual Action<T> GetWriteRecordAction<T>()
		{
			var type = typeof( T );
			if( !typeActions.ContainsKey( type ) )
			{
				CreateWriteRecordAction( type );
			}

			return (Action<T>)typeActions[type];
		}

		/// <summary>
		/// Gets the action delegate used to write the custom
		/// class object to the writer.
		/// </summary>
		/// <param name="type">The type of the custom class being written.</param>
		/// <returns>The action delegate.</returns>
		protected virtual Delegate GetWriteRecordAction( Type type )
		{
			if( !typeActions.ContainsKey( type ) )
			{
				CreateWriteRecordAction( type );
			}

			return typeActions[type];
		}

		/// <summary>
		/// Creates the write record action for the given type if it
		/// doesn't already exist.
		/// </summary>
		/// <param name="type">The type of the custom class being written.</param>
		protected virtual void CreateWriteRecordAction( Type type )
		{
			if( configuration.Maps[type] == null )
			{
				// We need to check again in case the header was not written.
				configuration.Maps.Add( configuration.AutoMap( type ) );
			}

			if( type.GetTypeInfo().IsPrimitive )
			{
				CreateActionForPrimitive( type );
			}
			else
			{
				CreateActionForObject( type );
			}
		}

		/// <summary>
		/// Creates the action for an object.
		/// </summary>
		/// <param name="type">The type of object to create the action for.</param>
		protected virtual void CreateActionForObject( Type type )
		{
			var recordParameter = Expression.Parameter( type, "record" );

			// Get a list of all the properties so they will
			// be sorted properly.
			var properties = new CsvPropertyMapCollection();
			AddProperties( properties, configuration.Maps[type] );

			if( properties.Count == 0 )
			{
				throw new CsvWriterException( $"No properties are mapped for type '{type.FullName}'." );
			}

			var delegates = new List<Delegate>();

			foreach( var propertyMap in properties )
			{
				if( !CanWrite( propertyMap ) )
				{
					continue;
				}

				if( propertyMap.Data.TypeConverter == null )
				{
					// Skip if the type isn't convertible.
					continue;
				}

				var fieldExpression = CreatePropertyExpression( recordParameter, configuration.Maps[type], propertyMap );

				var typeConverterExpression = Expression.Constant( propertyMap.Data.TypeConverter );
				if( propertyMap.Data.TypeConverterOptions.CultureInfo == null )
				{
					propertyMap.Data.TypeConverterOptions.CultureInfo = configuration.CultureInfo;
				}

				propertyMap.Data.TypeConverterOptions = TypeConverterOptions.Merge( propertyMap.Data.TypeConverterOptions, configuration.TypeConverterOptionsFactory.GetOptions( propertyMap.Data.Property.PropertyType ), propertyMap.Data.TypeConverterOptions );

				var method = propertyMap.Data.TypeConverter.GetType().GetMethod( "ConvertToString" );
				fieldExpression = Expression.Convert( fieldExpression, typeof( object ) );
				fieldExpression = Expression.Call( typeConverterExpression, method, fieldExpression, Expression.Constant( this ), Expression.Constant( propertyMap.Data ) );

				if( type.GetTypeInfo().IsClass )
				{
					var areEqualExpression = Expression.Equal( recordParameter, Expression.Constant( null ) );
					fieldExpression = Expression.Condition( areEqualExpression, Expression.Constant( string.Empty ), fieldExpression );
				}

				var writeFieldMethodCall = Expression.Call( Expression.Constant( this ), "WriteConvertedField", null, fieldExpression );

				var actionType = typeof( Action<> ).MakeGenericType( type );
				delegates.Add( Expression.Lambda( actionType, writeFieldMethodCall, recordParameter ).Compile() );
			}

			typeActions[type] = CombineDelegates( delegates );
		}

		/// <summary>
		/// Creates the action for a primitive.
		/// </summary>
		/// <param name="type">The type of primitive to create the action for.</param>
		protected virtual void CreateActionForPrimitive( Type type )
		{
			var recordParameter = Expression.Parameter( type, "record" );

			Expression fieldExpression = Expression.Convert( recordParameter, typeof( object ) );

			var typeConverter = TypeConverterFactory.GetConverter( type );
			var typeConverterExpression = Expression.Constant( typeConverter );
			var method = typeConverter.GetType().GetMethod( "ConvertToString" );

			var propertyMapData = new CsvPropertyMapData( null )
			{
				Index = 0,
				TypeConverter = typeConverter,
				TypeConverterOptions = { CultureInfo = configuration.CultureInfo }
			};
			propertyMapData.TypeConverterOptions = TypeConverterOptions.Merge( propertyMapData.TypeConverterOptions, configuration.TypeConverterOptionsFactory.GetOptions( type ) );

			fieldExpression = Expression.Call( typeConverterExpression, method, fieldExpression, Expression.Constant( this ), Expression.Constant( propertyMapData ) );
			fieldExpression = Expression.Call( Expression.Constant( this ), "WriteConvertedField", null, fieldExpression );

			var actionType = typeof( Action<> ).MakeGenericType( type );
			typeActions[type] = Expression.Lambda( actionType, fieldExpression, recordParameter ).Compile();
		}

#if !NET_2_0 && !NET_3_5 && !PCL

		/// <summary>
		/// Creates the action for a dynamic object.
		/// </summary>
		/// <param name="provider">The dynamic object.</param>
		protected virtual void CreateActionForDynamic( IDynamicMetaObjectProvider provider )
		{
			var type = provider.GetType();
			var parameterExpression = Expression.Parameter( typeof( object ), "record" );

			var metaObject = provider.GetMetaObject( parameterExpression );
			var propertyNames = metaObject.GetDynamicMemberNames();

			var delegates = new List<Delegate>();
			foreach( var propertyName in propertyNames )
			{
				var getMemberBinder = (GetMemberBinder)Microsoft.CSharp.RuntimeBinder.Binder.GetMember( 0, propertyName, type, new[] { CSharpArgumentInfo.Create( 0, null ) } );
				var getMemberMetaObject = metaObject.BindGetMember( getMemberBinder );
				var fieldExpression = Expression.Block( Expression.Label( CallSiteBinder.UpdateLabel ), getMemberMetaObject.Expression );
				var writeFieldMethodCall = Expression.Call( Expression.Constant( this ), "WriteField", new[] { typeof( object ) }, fieldExpression );
				var lambda = Expression.Lambda( writeFieldMethodCall, parameterExpression );
				delegates.Add( lambda.Compile() );
			}

			typeActions[type] = CombineDelegates( delegates );
		}

#endif

		/// <summary>
		/// Combines the delegates into a single multicast delegate.
		/// This is needed because Silverlight doesn't have the
		/// Delegate.Combine( params Delegate[] ) overload.
		/// </summary>
		/// <param name="delegates">The delegates to combine.</param>
		/// <returns>A multicast delegate combined from the given delegates.</returns>
		protected virtual Delegate CombineDelegates( IEnumerable<Delegate> delegates )
		{
			return delegates.Aggregate<Delegate, Delegate>( null, Delegate.Combine );
		}

		/// <summary>
		/// Checks if the property can be written.
		/// </summary>
		/// <param name="propertyMap">The property map that we are checking.</param>
		/// <returns>A value indicating if the property can be written.
		/// True if the property can be written, otherwise false.</returns>
		protected virtual bool CanWrite( CsvPropertyMap propertyMap )
		{
			var cantWrite =
				// Ignored properties.
				propertyMap.Data.Ignore ||
				// Properties that don't have a public getter
				// and we are honoring the accessor modifier.
				propertyMap.Data.Property.GetGetMethod() == null && !configuration.IgnorePrivateAccessor ||
				// Properties that don't have a getter at all.
				propertyMap.Data.Property.GetGetMethod( true ) == null;
			return !cantWrite;
		}

#endif

	}
}
