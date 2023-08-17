namespace TotalMEPProject.Ultis
{
    public class Define
    {
        #region Default value

        public const double OffsetHangerDefaultValue = 50 / 304.8;
        public const string ERR_NO_SET_ELBOW_FOR_PIPETYPE = "No set elbow for pipetype";
        public const string ERR_ANGLE_ELBOW = "Elbow Angle is out of Acceptable Range. Please choose other method";

        #endregion Default value

        #region ClassNames

        public const string CmdLoginClassName = "TotalMEPProject.Commands.Login.CmdLogin";
        public const string CmdChangeElevationClassName = "TotalMEPProject.Commands.TotalMEP.CmdChangeElevation";
        public const string CmdFastVerticalClassName = "TotalMEPProject.Commands.TotalMEP.CmdFastVertical";
        public const string CmdHolyUpdownClassName = "TotalMEPProject.Commands.TotalMEP.CmdHolyUpdown";
        public const string CmdMEPConnectionClassName = "TotalMEPProject.Commands.TotalMEP.CmdMEPConnection";
        public const string CmdPickLineClassName = "TotalMEPProject.Commands.TotalMEP.CmdPickLine";
        public const string CmdVerticalMEPClassName = "TotalMEPProject.Commands.TotalMEP.CmdVerticalMEP";
        public const string CmdDeleteUnionClassName = "TotalMEPProject.Commands.Duct.CmdDeleteUnion";
        public const string CmdDiffuserConnectionClassName = "TotalMEPProject.Commands.Duct.CmdDiffuserConnection";
        public const string CmdGrillOnDuctClassName = "TotalMEPProject.Commands.Duct.CmdGrillOnDuct";
        public const string CmdSplitDuctClassName = "TotalMEPProject.Commands.Duct.CmdSplitDuct";
        public const string CmdTapFlipClassName = "TotalMEPProject.Commands.Duct.CmdTapFlip";
        public const string Cmd2LevelSmartClassName = "TotalMEPProject.Commands.FireFighting.Cmd2LevelSmart";
        public const string CmdC234 = "TotalMEPProject.Commands.FireFighting.CmdC234";
        public const string CmdFlexSprinklerClassName = "TotalMEPProject.Commands.FireFighting.CmdFlexSprinkler";
        public const string CmdSprinklerDownrightClassName = "TotalMEPProject.Commands.FireFighting.CmdSprinklerDownright";
        public const string CmdSprinklerUprightClassName = "TotalMEPProject.Commands.FireFighting.CmdSprinklerUpright";
        public const string CmdBlockCADToFamilyClassName = "TotalMEPProject.Commands.Modify.CmdBlockCADToFamily";
        public const string CmdFittingRotationClassName = "TotalMEPProject.Commands.Modify.CmdFittingRotation";
        public const string CmdFlipSelectionClassName = "TotalMEPProject.Commands.Modify.CmdFlipSelection";
        public const string CmdMoveFittingClassName = "TotalMEPProject.Commands.Modify.CmdMoveFitting";
        public const string CmdSmartSelectionClassName = "TotalMEPProject.Commands.Modify.CmdSmartSelection";
        public const string CmdTapToTeeClassName = "TotalMEPProject.Commands.Modify.CmdTapToTee";
        public const string CmdTeeToTapClassName = "TotalMEPProject.Commands.Modify.CmdTeeToTap";
        public const string CmdChangeOpeningClassName = "TotalMEPProject.Commands.Opening.ChangeObject.CmdChangeOpening";
        public const string CmdChangeSleeveClassName = "TotalMEPProject.Commands.Opening.ChangeObject.CmdChangeSleeve";
        public const string CmdSettingChangeOpeningClassName = "TotalMEPProject.Commands.Opening.ChangeObject.CmdSettingChangeOpening";
        public const string CmdCreateOpeningClassName = "TotalMEPProject.Commands.Opening.CmdCreateOpening";
        public const string CmdCreateSleeveClassName = "TotalMEPProject.Commands.Opening.CmdCreateSleeve";
        public const string CmdDeleteAllOpeningClassName = "TotalMEPProject.Commands.Opening.DeleteOpening.CmdDeleteAll";
        public const string CmdDeleteBySelectionOpeningClassName = "TotalMEPProject.Commands.Opening.DeleteOpening.CmdDeleteBySelection";
        public const string CmdDeleteAllSleeveClassName = "TotalMEPProject.Commands.Opening.DeleteSleeve.CmdDeleteAll";
        public const string CmdDeleteBySelectionSleeveClassName = "TotalMEPProject.Commands.Opening.DeleteSleeve.CmdDeleteBySelection";
        public const string CmdCreateCouplingClassName = "TotalMEPProject.Commands.Plumping.CmdCreateCoupling";
        public const string CmdCreateEndCapClassName = "TotalMEPProject.Commands.Plumping.CmdCreateEndCap";
        public const string CmdCreateNippleClassName = "TotalMEPProject.Commands.Plumping.CmdCreateNipple";
        public const string CmdExtendPipeClassName = "TotalMEPProject.Commands.Plumping.CmdExtendPipe";
        public const string CmdSlopePipeConnectionClassName = "TotalMEPProject.Commands.Plumping.CmdSlopePipeConnection";
        public const string CmdVerticalWYEConnection1ClassName = "TotalMEPProject.Commands.Plumping.CmdVerticalWYEConnection1";
        public const string CmdVeticalTeeConnectionClassName = "TotalMEPProject.Commands.Plumping.CmdVeticalTeeConnection";

        #endregion ClassNames

        #region RibbonTabNames

        public const string LoginLicense = "License Service";
        public const string DuctRibbonTabName = "Duct";
        public const string FireFightingRibbonTabName = "FireFighting";
        public const string ModifyRibbonTabName = "Modify";
        public const string OpeningRibbonTabName = "Opening";
        public const string PlumpingRibbonTabName = "Plumping";
        public const string TotalMEPRibbonTabName = "TotalMEP";

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
        public const string ERR_CAN_NOT_CREATE_THIS_CONNECTION_FOR_THIS_CASE = "Can not create this connection for this case";
        public const string PIPE_CAST_IRON = "CAST IRON";
    }
}