using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Xml;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Collections;
using System.IO;
using System.Diagnostics;
using System.Collections.Specialized;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB.Structure;
using System.Security;
using System.Security.Permissions;

namespace TotalMEPProject.Ultis
{
    public class Global
    {
        public static Autodesk.Revit.ApplicationServices.Application RVTApp = null;

        public static Autodesk.Revit.UI.UIApplication UIApp = null;

        public static Autodesk.Revit.Creation.Application AppCreation = null;

        public static UIDocument UIDoc = null;

        public static string GroupNameParameter = "MyParameters";

        public static List<UniformatCode> LISTUNIFORMATCODE = new List<UniformatCode>();

        public static List<BuiltInCategory> LISTCATEGORYAUTOJOIN = new List<BuiltInCategory>
        {
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_Floors ,
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_StructuralFoundation,
            BuiltInCategory.OST_Ceilings,
            BuiltInCategory.OST_GenericModel,
        };

        public static List<BuiltInCategory> LISTCATEGORYLOCATIONMARK = new List<BuiltInCategory>
        {
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_CableTray,
            BuiltInCategory.OST_Parts,
            BuiltInCategory.OST_GenericModel,
            //BuiltInCategory.OST_Floors
        };

        public static List<BuiltInCategory> LISTBUILTINCATEGORYNOTUSER = new List<BuiltInCategory>()
        {
            BuiltInCategory.OST_Grids,
            BuiltInCategory.OST_Levels,
            BuiltInCategory.OST_Viewers,
            BuiltInCategory.OST_Views,
            BuiltInCategory.OST_Sheets,
            BuiltInCategory.OST_Dimensions,
            BuiltInCategory.OST_SpotElevations,
            BuiltInCategory.OST_Viewports,
            BuiltInCategory.OST_SpanDirectionSymbol,
            BuiltInCategory.OST_FootingSpanDirectionSymbol,
            BuiltInCategory.OST_TextNotes,
            BuiltInCategory.OST_RasterImages,
            BuiltInCategory.OST_ElevationMarks,
        };

        private static List<string> LISTIRGONECATEGORYNAME = new List<string>()
        {
            "Point Clouds",

            "Rebar Set Toggle",

            "Rebar Cover References",

            "Pipe Segments",

            "Panel Schedule Graphics",

            "Routing Preferences",

            "Pipe Color Fill",

            "Pipe Color Fill Legends",

            "Duct Color Fill",

            "Duct Color Fill Legends",

            "Revision Clouds",

            "Grid Heads",

            "Level Heads",

            "Scope Boxes",

            "Boundary Conditions",

            "Structural Load Cases",

            "Structural Internal Loads",

            "Structural Loads",

            "Stair Tread/Riser Numbers",

            "Stair Paths",

            "Adaptive Points",

            "Reference Points",

            "Schedule Graphics",

            "Color Fill Legends",

            "Callout Boundary",

            "Callout Heads",

            "Callouts",

            "Elevations",

            "Reference Planes",

            "Cameras",

            "Section Marks",

            "Contour Labels",

            "Analysis Display Style",

            "Analysis Results",

            "Render Regions",

            "Section Boxes",

            "Spot Slopes",

            "Spot Coordinates",

            "Displacement Path",

            "Section Line",

            "Sections",

            "View Reference",

            "Imports in Families",

            "Masking Region",

            "Matchline",

            "Plan Region",

            "Filled region",

            "Curtain Grids",

            "Guide Grid",

            "Reference Lines",

            "Lines",
        };
    }

    public enum ConfigType
    {
        SURFACE_MAKER = 0,
    }

    public enum ConfigItemType
    {
        WALL_SURFACE_TYPE = 0,
        COLUMN_SURFACE_TYPE,
    }

    public class Config
    {
        public ConfigType ConfigType = ConfigType.SURFACE_MAKER;
        public Dictionary<ConfigItemType, string> Items = new Dictionary<ConfigItemType, string>();

        public Config(ConfigType type)
        {
            ConfigType = type;
        }
    }

    public class UniformatCode
    {
        public string Code = string.Empty;
        public string Name = string.Empty;
        public int Level = -1;
        public BuiltInCategory BuiltInCategory = BuiltInCategory.INVALID;
    }

    public class WindowHandle : IWin32Window
    {
        private IntPtr m_hwnd;

        public WindowHandle(IntPtr h)
        {
            Debug.Assert(IntPtr.Zero != h,
              "EXPECTED NON-NULL WINDOW HANDLE");

            m_hwnd = h;
        }

        public IntPtr Handle
        {
            get
            {
                return m_hwnd;
            }
        }
    }
}