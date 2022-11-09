using Revit = Autodesk.Revit;

namespace TotalMEPProject.Request
{
    /// <summary>
    /// A list of requests the dialog has available
    /// </summary>
    public enum RequestId : int
    {
        None = -1,
        VerticalMEP,
        FastVertical,
    }

    public class Request
    {
        // Member variables

        #region Member Variables

        public Revit.DB.ElementId _selectedTag = Revit.DB.ElementId.InvalidElementId;

        // Storing the value as a plain Int makes using the interlocking mechanism simpler
        private int _request = (int)RequestId.None;

        #endregion Member Variables

        //Member Functions

        #region Member Functions

        public RequestId Take()
        {
            return (RequestId)System.Threading.Interlocked.Exchange(ref _request, (int)RequestId.None);
        }

        public void Make(RequestId request)
        {
            System.Threading.Interlocked.Exchange(ref _request, (int)request);
        }

        #endregion Member Functions
    }
}