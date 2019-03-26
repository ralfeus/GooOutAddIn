using System;

namespace R.GoogleOutlookSync
{
    internal abstract class FieldHandlerAttribute : Attribute
    {
        internal Field Field { get; set; }
    }

    internal class FieldComparerAttribute : FieldHandlerAttribute
    {
        internal FieldComparerAttribute(Field field)
        {
            this.Field = field;
        }
    }

    internal class FieldGetterAttribute : FieldHandlerAttribute
    {
        internal FieldGetterAttribute(Field field)
        {
            this.Field = field;
        }
    }

    internal class FieldSetterAttribute : FieldHandlerAttribute
    {
        internal FieldSetterAttribute(Field field)
        {
            this.Field = field;
        }
    }

    internal class PublicAttribute : Attribute
    {
    }
}
