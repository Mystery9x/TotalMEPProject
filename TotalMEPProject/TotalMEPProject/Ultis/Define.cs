namespace TotalMEPProject.Ultis
{
    public class Define
    {
        #region Default value

        public const double OffsetHangerDefaultValue = 50 / 304.8;

        #endregion Default value

        #region ClassNames

        public const string CmdLoginClassName = "TotalMEPProject.Commands.Login.CmdLogin";

        #endregion ClassNames

        #region RibbonTabNames

        public const string LoginLicense = "License Service";

        #endregion RibbonTabNames

        #region License Message

        public const string UnknowLicense = "Error of unknown cause";
        public const string NotExistKeyLicense = "Key does not exist";
        public const string ExpiredKeyLicense = "Expired key";
        public const string HardwareDiffLicense = "The trigger device and the requested device do not overlap";
        public const string RejectLicense = "License denied";

        #endregion License Message

        public const string GUID = "61177B23-0C77-4C7F-B381-55BB4C949A6C";
        public const string SchemaName = "TotalMEP";
        public const string VendorId = "CD50481A-BE7A-4816-A2DE-EA1F50A0812A";
    }
}