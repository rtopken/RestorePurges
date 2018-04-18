using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Exchange.WebServices.Data;

namespace RestorePurges
{
    public class PropSet
    {
        public static bool DoProps(ref List<ExtendedPropertyDefinition> epdProps)
        {
            bool bSuccess = true;

            epdProps = new List<ExtendedPropertyDefinition>();
            ExtendedPropertyDefinition extProp = null;
            int iTag = 0;
            MapiPropertyType mpType = MapiPropertyType.Null;
            int Bound0 = rgProps.GetUpperBound(0);

            for (int i = 0; i < Bound0; i++)
            {
                mpType = GetPropType(rgProps[i, 1]);

                if (rgProps[i, 0] == "") // for string values that have no pTag
                {
                    string strGUID = GetGUIDFromSetID(rgProps[i, 2]);
                    string strProp = rgProps[i, 3];
                    extProp = new ExtendedPropertyDefinition(new Guid(strGUID), strProp, mpType);
                }
                else // all others that have pTags
                {
                    iTag = Convert.ToInt32(rgProps[i, 0], 16);

                    if (rgProps[i, 2] == "") // If Not a Named Prop
                    {
                        extProp = new ExtendedPropertyDefinition(iTag, mpType);
                    }
                    else // else it is a named prop
                    {
                        string strGUID = GetGUIDFromSetID(rgProps[i, 2]);
                        extProp = new ExtendedPropertyDefinition(new Guid(strGUID), iTag, mpType);
                    }
                }
                epdProps.Add(extProp);
            }
            return bSuccess;
        }

        private static string[,] rgProps = new string[,]
        {
            {"001A", "String", "", "PR_MESSAGE_CLASS"},
            {"0037", "String", "", "PR_SUBJECT_W"},
            {"0042", "String", "", "PR_SENT_REPRESENTING_NAME_W"},
            {"0065", "String", "", "PR_SENT_REPRESENTING_EMAIL_ADDRESS_W"},
            {"3008", "SystemTime", "", "PR_LAST_MODIFICATION_TIME"},
            {"3007", "SystemTime", "", "PR_CREATION_TIME"},
            {"0003", "Binary", "PSETID_Meeting", "PidLidGlobalObjectId"},
            {"0005", "Boolean", "PSETID_Meeting", "PidLidIsRecurring"},
            {"0023", "Binary", "PSETID_Meeting", "PidLidCleanGlobalObjectId"},
            {"820D", "SystemTime", "PSETID_Appointment", "dispidApptStartWhole"},
            {"820E", "SystemTime", "PSETID_Appointment", "dispidApptEndWhole"},
            {"8217", "Integer", "PSETID_Appointment", "dispidApptStateFlags"},
            {"8208", "String", "PSETID_Appointment", "dispidLocation"},
        };

        public static string[,] RgProps { get => rgProps; set => rgProps = value; }


        private static string[,] rgGUIDS = new string[,]
        {
            {"{11000E07-B51B-40D6-AF21-CAA85EDAB1D0}", "PSETID_CalendarAssistant" },
            {"{6ED8DA90-450B-101B-98DA-00AA003F1305}", "PSETID_Meeting" },
            {"{00062002-0000-0000-C000-000000000046}", "PSETID_Appointment" },
            {"{00020329-0000-0000-C000-000000000046}", "PS_PUBLIC_STRINGS" },
            {"{00062008-0000-0000-C000-000000000046}", "PSETID_Common" }
        };

        public static string GetPropNameFromTag(string strTag, string strSetID)
        {
            string strOut = "";
            int Bound0 = rgProps.GetUpperBound(0);

            for (int i = 0; i <= Bound0; i++)
            {
                if (strTag == rgProps[i, 0])
                {
                    if (strSetID == rgProps[i, 2])
                    {
                        strOut = rgProps[i, 3];
                        break;
                    }
                }
            }
            return strOut;
        }

        public static string GetSetIDFromGUID(string strGUID)
        {
            string strOut = "";
            int Bound0 = rgGUIDS.GetUpperBound(0);

            for (int i = 0; i <= Bound0; i++)
            {
                if (strGUID.ToUpper() == rgGUIDS[i, 0])
                {
                    strOut = rgGUIDS[i, 1];
                    break;
                }
            }
            return strOut;
        }

        public static string GetGUIDFromSetID(string strSetID)
        {
            string strOut = "";
            int Bound0 = rgGUIDS.GetUpperBound(0);

            for (int i = 0; i <= Bound0; i++)
            {
                if (strSetID == rgGUIDS[i, 1])
                {
                    strOut = rgGUIDS[i, 0];
                    break;
                }
            }
            return strOut;
        }

        public static MapiPropertyType GetPropType(string strProp)
        {
            MapiPropertyType mpType = MapiPropertyType.Null;

            switch (strProp.ToUpper())
            {
                case "INTEGER":
                    mpType = MapiPropertyType.Integer;
                    break;
                case "STRING":
                    mpType = MapiPropertyType.String;
                    break;
                case "BINARY":
                    mpType = MapiPropertyType.Binary;
                    break;
                case "LONG":
                    mpType = MapiPropertyType.Long;
                    break;
                case "BOOLEAN":
                    mpType = MapiPropertyType.Boolean;
                    break;
                case "SYSTEMTIME":
                    mpType = MapiPropertyType.SystemTime;
                    break;
                case "STRINGARRAY":
                    mpType = MapiPropertyType.StringArray;
                    break;

                default:
                    return MapiPropertyType.Null;
            }
            return mpType;
        }


        // Populate the property values for each of the props the app checks on.
        // Some tests require multiple props, so best to go ahead and just get them all first.
        public static string GetPropsLine(Appointment appt)
        {
            string strHexTag = "";
            string strPropName = "";
            string strSetID = "";
            string strGUID = "";
            string strValue = "";
            string strType = "";
            string strItemProps = "";

            string strSubject = "";
            string strOrganizerName = "";
            string strOrganizerAddr = "";
            string strMsgClass = "";
            string strLastModified = "";
            string strCreateTime = "";
            string strRecurring = "";
            string strStartWhole = "";
            string strEndWhole = "";
            string strApptStateFlags = "";
            string strLocation = "";
            string strGlobalObjID = "";
            string strCleanGlobalObjID = "";

            foreach (ExtendedProperty extProp in appt.ExtendedProperties)
            {
                // Get the Tag
                if (extProp.PropertyDefinition.Tag.HasValue)
                {
                    strHexTag = extProp.PropertyDefinition.Tag.Value.ToString("X4");
                }
                else if (extProp.PropertyDefinition.Id.HasValue)
                {
                    strHexTag = extProp.PropertyDefinition.Id.Value.ToString("X4");
                }

                // Get the SetID for named props
                if (extProp.PropertyDefinition.PropertySetId.HasValue)
                {
                    strGUID = extProp.PropertyDefinition.PropertySetId.Value.ToString("B");
                    strSetID = PropSet.GetSetIDFromGUID(strGUID);
                }

                // Get the Property Type
                strType = extProp.PropertyDefinition.MapiType.ToString();

                // Get the Prop Name
                strPropName = PropSet.GetPropNameFromTag(strHexTag, strSetID);

                // if it's binary then convert it to a string-ized binary - will be converted using MrMapi
                if (strType == "Binary")
                {
                    byte[] binData = extProp.Value as byte[];
                    strValue = GetStringFromBytes(binData);
                }
                else
                {
                    if (extProp.Value != null)
                    {
                        strValue = extProp.Value.ToString();
                    }
                }

                switch (strPropName)
                {
                    case "PR_SUBJECT_W":
                        {
                            strSubject = strValue;
                            break;
                        }
                    case "PR_SENT_REPRESENTING_NAME_W":
                        {
                            strOrganizerName = strValue;
                            break;
                        }
                    case "PR_SENT_REPRESENTING_EMAIL_ADDRESS_W":
                        {
                            strOrganizerAddr = strValue;
                            break;
                        }
                    case "PR_MESSAGE_CLASS":
                        {
                            strMsgClass = strValue;
                            break;
                        }
                    case "PR_LAST_MODIFICATION_TIME":
                        {
                            strLastModified = strValue;
                            break;
                        }
                    case "PR_CREATION_TIME":
                        {
                            strCreateTime = strValue;
                            break;
                        }
                    case "dispidRecurring":
                        {
                            strRecurring = strValue;
                            break;
                        }
                    case "dispidApptStartWhole":
                        {
                            strStartWhole = strValue;
                            break;
                        }
                    case "dispidApptEndWhole":
                        {
                            strEndWhole = strValue;
                            break;
                        }
                    case "dispidApptStateFlags":
                        {
                            strApptStateFlags = strValue;
                            break;
                        }
                    case "dispidLocation":
                        {
                            strLocation = strValue;
                            break;
                        }
                    case "PidLidGlobalObjectId":
                        {
                            strGlobalObjID = strValue;
                            break;
                        }
                    case "PidLidCleanGlobalObjectId":
                        {
                            strCleanGlobalObjID = strValue;
                            break;
                        }
                    default:
                        {
                            break;
                        }
                }
            }
            strItemProps = strGlobalObjID + "," + strSubject + "," + strStartWhole + "," + strEndWhole + "," + strOrganizerAddr + "," + strRecurring;
            return strItemProps;
        }

        // EWS does not return a string-ized hex blob, and need it for MrMapi conversion
        public static string GetStringFromBytes(byte[] bytes)
        {
            StringBuilder ret = new StringBuilder();
            foreach (byte b in bytes)
            {
                ret.Append(Convert.ToString(b, 16).PadLeft(2, '0'));
            }

            return ret.ToString().ToUpper();
        }
    }
}
