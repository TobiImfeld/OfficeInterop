using System;

namespace WordVba
{
    public class WordVbaReferenceControl : WordVbaReference
    {
        public WordVbaReferenceControl()
        {
            ReferenceRecordID = 0x2F;
        }
        /// <summary>
        /// LibIdExternal 
        /// For more info check MS-OVBA 2.1.1.8 LibidReference and 2.3.4.2.2 PROJECTREFERENCES
        /// </summary>
        public string LibIdExternal { get; set; }

        /// <summary>
        /// LibIdTwiddled
        /// For more info check MS-OVBA 2.1.1.8 LibidReference and 2.3.4.2.2 PROJECTREFERENCES
        /// </summary>
        public string LibIdTwiddled { get; set; }

        public Guid OriginalTypeLib { get; set; }

        internal uint Cookie { get; set; }
    }
}
