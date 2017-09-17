///
/// Author: Asad Zia
/// Version: 1.0
///

namespace BP_XML_ExceptionsParser
{
    /// <summary>
    /// Definition of constants
    /// </summary>
    public static class Config
    {
        /// <summary>
        /// TAG and ATTRIBUTE names for parsing the XML
        /// </summary>
        public const string SUBSHEET_ID_ATTRIBUTE = "subsheetid";
        public const string SUBSHEET_TAG = "subsheet";
        public const string PROCESS_TAG = "process";
        public const string OBJECT_TAG = "object";
        public const string NAME_ATTRIBUTE = "name";
        public const string TYPE_ATTRIBUTE = "type";
        public const string DETAIL_ATTRIBUTE = "detail";
        public const string STAGE_TAG = "stage";
        public const string NARRATIVE_TAG = "narrative";
        public const string EXCEPTION_TAG = "exception";
        public const string USE_CURRENT_ATTRIBUTE = "usecurrent";
        public const string EXCEPTION_TYPE_PRESERVED = "PRESERVED EXCEPTION TYPE";
        public const string EXCEPTION_DETAIL_PRESERVED = "PRESERVED EXCEPTION DETAIL";
        public const string EMPTY_STRING = "";

        /// <summary>
        /// The CSV header configuration
        /// </summary>
        public const string TYPE_HEADER = "Exception Type";
        public const string DETAILS_HEADER = "Exception Details";
        public const string PROCESS_HEADER = "Process Name";
        public const string PAGE_HEADER = "Page Name";
        public const string OBJECT_HEADER = "Object Name";
        public const string ACTION_HEADER = "Object Action";
    }
}
