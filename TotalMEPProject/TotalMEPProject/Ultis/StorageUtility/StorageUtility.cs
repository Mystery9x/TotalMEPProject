using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using System;
using System.Collections.Generic;

namespace TotalMEPProject.Ultis.StorageUtility
{
    public class IntEntity : Entity
    {
        public int Value { get; set; }
    }

    public class StorageUtility
    {
        public static Guid m_CreatedFrom_Guild = Guid.Parse("1E5B6F62-B8B3-4A2F-9B06-DDD953D4D4BC");
        public static string m_CreatedFrom = "CreatedFrom";

        public static Guid m_MEP_HoLyUpDown_Guild = Guid.Parse("2E5B6F62-B8B3-4A2F-9B06-DDD953D4D4BC");
        public static string m_MEP_HoLyUpDown = "MEP_HoLyUpDown";

        public static Schema CreateSchema(Guid guid, string fieldName, Type type)
        {
            try
            {
                if (SchemaBuilder.GUIDIsValid(guid))
                {
                }

                SchemaBuilder schemaBuilder = new SchemaBuilder(guid);

                if (schemaBuilder.AcceptableName(fieldName))
                {
                }

                if (type == typeof(int))
                    schemaBuilder.SetSchemaName("Mark_IntSchema");
                else // if (type == typeof(string))
                    schemaBuilder.SetSchemaName("Mark_StrSchema");

                // Have to define the field name as string and
                // set the type using typeof method

                FieldBuilder fieldBuilder = schemaBuilder.AddSimpleField(fieldName, type);
                return schemaBuilder.Finish();
            }
            catch (System.Exception ex)
            {
                return null;
            }
        }

        public static object GetValue(Element element, Schema schema, string storageName, Type type)
        {
            try
            {
                if (element == null)
                    return null;
                var entity = element.GetEntity(schema);
                if (entity == null)
                    return null;

                object value = null;
                if (type == typeof(int))
                    value = entity.Get<int>(storageName);
                else
                    value = entity.Get<string>(storageName);

                if (value == null)
                {
                    return null;
                }
                return value;
            }
            catch (System.Exception ex)
            {
            }
            return null;
        }

        public static bool SetValue(Element element, Guid guid, string fieldName, Type type, object value)
        {
            if (element == null)
                return false;

            try
            {
                var schema = Autodesk.Revit.DB.ExtensibleStorage.Schema.Lookup(guid);
                if (schema == null)
                    return false;

                var entity = element.GetEntity(schema);
                if (entity == null || entity.Schema == null)
                    return false;

                if (type == typeof(IList<ElementId>))
                {
                    var list = (IList<ElementId>)value;
                    entity.Set<IList<ElementId>>(fieldName, list);
                }
                else if (type == typeof(string))
                {
                    var str = (string)value;
                    entity.Set<string>(fieldName, str);
                }
                else if (type == typeof(ElementId))
                {
                    var id = (ElementId)value;
                    entity.Set<ElementId>(fieldName, id);
                }
                else if (type == typeof(int))
                {
                    var iValue = (int)value;
                    entity.Set<int>(fieldName, iValue);
                }
                else if (type == typeof(bool))
                {
                    var iValue = (bool)value == true ? 1 : 0;
                    entity.Set<int>(fieldName, iValue);
                }
                else if (type == typeof(double))
                {
                    var iValue = (double)value;
                    entity.Set<double>(fieldName, iValue, DisplayUnitType.DUT_CUSTOM);
                }

                element.Document.Regenerate();

                element.SetEntity(entity);

                return true;
            }
            catch (System.Exception ex)
            {
                string mess = ex.Message;
                return false;
            }
        }

        public static bool AddEntity(Element element, Guid guid, string name, object value)
        {
            try
            {
                Type type = value.GetType();
                var schema = Schema.Lookup(guid);
                if (schema == null)
                {
                    schema = StorageUtility.CreateSchema(guid, name, type);
                }
                var entity = new Autodesk.Revit.DB.ExtensibleStorage.Entity(schema);

                if (type == typeof(int))
                    entity.Set(name, (int)value);
                else
                    entity.Set(name, (string)value);

                element.SetEntity(entity);

                return true;
            }
            catch (System.Exception ex)
            {
            }

            return false;
        }
    }
}