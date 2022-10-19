using Newtonsoft.Json.Linq;
using System;
using System.Collections.Specialized;
using System.Net;
using System.Text;

namespace TotalMEPProject.Ultis
{
    internal class LicenseUtils
    {
        public static bool CheckLicense(bool isHasInternet, ref string errmessage)
        {
            try
            {
                if (isHasInternet)
                {
                    string licenseKey = Microsoft.VisualBasic.Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, "AddinMep", "LicenseKey", "");
                    string hardwareId = Microsoft.VisualBasic.Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, "AddinMep", "HardwareId", "");

                    if (licenseKey == null || hardwareId == null)
                        return false;
                    else
                    {
                        using (var client = new WebClient())
                        {
                            var values = new NameValueCollection();
                            values["key"] = licenseKey;
                            values["hardware"] = hardwareId;

                            var response = client.UploadValues("https://license-en.duyanh.me/api/product/license/key/verify", values);

                            string responseString = Encoding.Default.GetString(response);
                            var jsonString = JObject.Parse(responseString);

                            var hasErr = jsonString["error"];
                            if (hasErr != null)
                            {
                                string errCase = hasErr.ToString();
                                switch (errCase)
                                {
                                    case "UNKNOWN":
                                        errmessage = Define.UnknowLicense;
                                        break;

                                    case "KEY_INVALID":
                                        errmessage = Define.NotExistKeyLicense;
                                        break;

                                    case "KEY_EXPIRE":
                                        errmessage = Define.ExpiredKeyLicense;
                                        break;

                                    case "HARDWARE_DIFFERENT":
                                        errmessage = Define.HardwareDiffLicense;
                                        break;

                                    case "LICENSE_REJECT":
                                        errmessage = Define.RejectLicense;
                                        break;
                                }

                                return false;
                            }
                            else
                            {
                                bool isHasRemainTime = CheckRemainTime();
                                if (!isHasRemainTime)
                                    errmessage = Define.ExpiredKeyLicense;

                                return isHasRemainTime;
                            }
                        }
                    }
                }
                else
                {
                    string licenseKey = Microsoft.VisualBasic.Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, "AddinMep", "LicenseKey", "");
                    string hardwareId = Microsoft.VisualBasic.Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, "AddinMep", "HardwareId", "");

                    if (licenseKey == null || hardwareId == null)
                        return false;

                    bool isHasRemainTime = CheckRemainTime();
                    if (!isHasRemainTime)
                        errmessage = Define.ExpiredKeyLicense;

                    return isHasRemainTime;
                }
            }
            catch (Exception ex)
            {
                string mess = ex.Message;
                return false;
            }
        }

        private static bool CheckRemainTime()
        {
            string remainTime = Microsoft.VisualBasic.Interaction.GetSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, "AddinMep", "RemainLicense", "");

            if (remainTime == null)
                return false;

            if ((long)Convert.ToDouble(remainTime) > new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds())
                return true;

            return false;
        }

        public static bool RegistLicense(string key, string hardware, ref string errmessage)
        {
            try
            {
                using (var client = new WebClient())
                {
                    var values = new NameValueCollection();
                    values["key"] = key;
                    values["hardware"] = hardware;

                    var response = client.UploadValues("https://license-en.duyanh.me/api/product/license/key/verify", values);

                    string responseString = Encoding.Default.GetString(response);
                    var jsonString = JObject.Parse(responseString);

                    var hasErr = jsonString["error"];
                    if (hasErr != null)
                    {
                        string errCase = hasErr.ToString();
                        switch (errCase)
                        {
                            case "UNKNOWN":
                                errmessage = Define.UnknowLicense;
                                break;

                            case "KEY_INVALID":
                                errmessage = Define.NotExistKeyLicense;
                                break;

                            case "KEY_EXPIRE":
                                errmessage = Define.ExpiredKeyLicense;
                                break;

                            case "HARDWARE_DIFFERENT":
                                errmessage = Define.HardwareDiffLicense;
                                break;

                            case "LICENSE_REJECT":
                                errmessage = Define.RejectLicense;
                                break;
                        }

                        return false;
                    }
                    else
                    {
                        var remain = jsonString["remain"];

                        var timestampNow = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds();
                        var remaintimestamp = timestampNow + (long)remain;

                        Microsoft.VisualBasic.Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, "AddinMep", "LicenseKey", key);
                        Microsoft.VisualBasic.Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, "AddinMep", "RemainLicense", remaintimestamp.ToString());
                        Microsoft.VisualBasic.Interaction.SaveSetting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, "AddinMep", "HardwareId", hardware);

                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                string mess = ex.Message;
                return false;
            }
        }

        public static bool CheckForInternetConnection(int timeoutMs = 10000, string url = null)
        {
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(url);
                request.KeepAlive = false;
                request.Timeout = timeoutMs;
                using (var response = (HttpWebResponse)request.GetResponse())
                    return true;
            }
            catch
            {
                return false;
            }
        }
    }
}