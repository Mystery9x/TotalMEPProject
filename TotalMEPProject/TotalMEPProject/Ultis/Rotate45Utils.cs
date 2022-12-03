using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace TotalMEPProject.Ultis
{
    public class Rotate45Utils
    {
        public static void ce(bool increase)
        {
            Transaction t = new Transaction(Global.UIDoc.Document, "a");
            t.Start();

            try
            {
                var pickedObj = Global.UIDoc.Selection.PickObject(ObjectType.Element, "Pick fittings: ");

                var fitting = Global.UIDoc.Document.GetElement(pickedObj) as FamilyInstance;

                double angle45 = 45;
                double angleInRadian45 = angle45 * Math.PI / 180.0;
                double angle = angleInRadian45;

                Connector mainConnect1 = null;
                Connector mainConnect2 = null;
                mc(fitting, null, out mainConnect1, out mainConnect2);
                if (mainConnect1 == null || mainConnect2 == null)
                {
                    string mess = string.Format("main connector 1 = {0}, main connector 2  = {1}", mainConnect1 == null ? "null" : "#null", mainConnect2 == null ? "null" : "#null");
                    MessageBox.Show(mess);
                    t.RollBack();
                    return;
                }

                Line axis = Line.CreateUnbound(mainConnect1.Origin, mainConnect1.CoordinateSystem.BasisZ);// Line.CreateBound(mainConnect1.Origin, mainConnect2.Origin);

                PartType partType = (fitting.MEPModel as MechanicalFitting).PartType;
                if (partType == PartType.Elbow)
                {
                    //Uu tien quay tai vi tri da connect voi thiet bi khac
                    if (mainConnect2.IsConnected)
                    {
                        axis = Line.CreateUnbound(mainConnect2.Origin, mainConnect2.CoordinateSystem.BasisZ);
                    }
                }

                ElementTransformUtils.RotateElement(
                  Global.UIDoc.Document, fitting.Id, axis, increase ? angle : -angle);
                Global.UIDoc.Document.Regenerate();
            }
            catch (System.Exception ex)
            {
            }

            t.Commit();
        }

        public static void mc(FamilyInstance fitting, XYZ vector, out Connector mainConnect1, out Connector mainConnect2)
        {
            mainConnect1 = null;
            mainConnect2 = null;

            PartType partType = (fitting.MEPModel as MechanicalFitting).PartType;
            if (partType == PartType.Tee && vector == null && fitting.MEPModel.ConnectorManager.Connectors.Size == 3)
            {
                //Main : hướng connector của 2 connector fai song song voi nhau (nguoc chieu nhau)

                foreach (Connector c1 in fitting.MEPModel.ConnectorManager.Connectors)
                {
                    foreach (Connector c2 in fitting.MEPModel.ConnectorManager.Connectors)
                    {
                        if (c1.Id == c2.Id)
                        {
                            continue;
                        }
                        else
                        {
                            var z1 = c1.CoordinateSystem.BasisZ;
                            var z2 = c2.CoordinateSystem.BasisZ;

                            if (Common.IsParallel(z1, z2, 0.0001) == true)
                            {
                                mainConnect1 = c1;
                                mainConnect2 = c2;
                                break;
                            }
                        }
                    }

                    if (mainConnect1 != null && mainConnect2 != null)
                        break;
                }
            }
            else
            {
                foreach (Connector con in fitting.MEPModel.ConnectorManager.Connectors)
                {
                    if (vector != null)
                    {
                        if (Common.IsParallel(vector, con.CoordinateSystem.BasisZ, 0.0001) == false)
                        {
                            continue;
                        }
                    }

                    if (mainConnect1 == null)
                        mainConnect1 = con;
                    else
                    {
                        mainConnect2 = con;
                        break;
                    }
                }
            }

            //Connect nao gan location of fitting thi do la 1

            var p = (fitting.Location as LocationPoint).Point;
            if (mainConnect1.Origin.DistanceTo(p) > mainConnect2.Origin.DistanceTo(p))
            {
                Connector temp = mainConnect1;
                mainConnect1 = mainConnect2;

                mainConnect2 = temp;
            }
        }
    }
}