using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace TotalMEPProject.Ultis
{
    internal class IgnoreWarning : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            IList<FailureMessageAccessor> fmas = failuresAccessor.GetFailureMessages();

            if (fmas.Count == 0)
                return FailureProcessingResult.Continue;

            bool isResolved = false;

            foreach (FailureMessageAccessor fma in fmas)
            {
                if (fma.HasResolutions())
                {
                    failuresAccessor.ResolveFailure(fma);
                    isResolved = true;
                }
            }

            failuresAccessor.DeleteAllWarnings();

            if (isResolved)
            {
                return FailureProcessingResult.ProceedWithCommit;
            }

            return FailureProcessingResult.Continue;
        }
    }
}