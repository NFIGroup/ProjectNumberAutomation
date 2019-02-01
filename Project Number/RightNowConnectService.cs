using Project_Number.RightNowService;
using RightNow.AddIns.AddInViews;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Project_Number
{
    class RightNowConnectService
    {
        private static RightNowConnectService _rightnowConnectService;
        private static object _sync = new object();
        private static RightNowSyncPortClient _rightNowClient;

        private RightNowConnectService()
        {

        }

        public static RightNowConnectService GetService(IGlobalContext _globalContext)
        {
            if (_rightnowConnectService != null)
            {
                return _rightnowConnectService;
            }

            try
            {
                lock (_sync)
                {

                    if (_rightnowConnectService == null)
                    {


                        // Initialize client with current interface soap url 
                        string url = _globalContext.GetInterfaceServiceUrl(ConnectServiceType.Soap);

                        EndpointAddress endpoint = new EndpointAddress(_globalContext.GetInterfaceServiceUrl(ConnectServiceType.Soap));

                        BasicHttpBinding binding = new BasicHttpBinding(BasicHttpSecurityMode.TransportWithMessageCredential);
                        binding.Security.Message.ClientCredentialType = BasicHttpMessageCredentialType.UserName;

                        // Optional depending upon use cases
                        binding.MaxReceivedMessageSize = 1024 * 1024;
                        binding.MaxBufferSize = 1024 * 1024;
                        binding.MessageEncoding = WSMessageEncoding.Mtom;

                        _rightNowClient = new RightNowSyncPortClient(binding, endpoint);

                        BindingElementCollection elements = _rightNowClient.Endpoint.Binding.CreateBindingElements();
                        elements.Find<SecurityBindingElement>().IncludeTimestamp = false;
                        _rightNowClient.Endpoint.Binding = new CustomBinding(elements);
                        WorkspaceAddIn._globalContext.PrepareConnectSession(_rightNowClient.ChannelFactory);

                        _rightnowConnectService = new RightNowConnectService();

                    }

                }
            }
            catch (Exception e)
            {
                _rightnowConnectService = null;

                //  WorkspaceAddIn.InfoLog(e.Message);
            }

            return _rightnowConnectService;
        }
        /// <summary>
        /// Return individual fields as per query
        /// </summary>
        /// <param name="ApplicationID"></param>
        /// <param name="Query"></param>
        /// <returns> array of string delimited by '~'</returns>
        private string[] GetRNData(string ApplicationID, string Query)
        {
            string[] rnData = null;
            ClientInfoHeader hdr = new ClientInfoHeader() { AppID = ApplicationID };

            byte[] output = null;
            CSVTableSet data = null;

            try
            {
                data = _rightNowClient.QueryCSV(hdr, Query, 50, "~", false, false, out output);
                string dataRow = String.Empty;
                if (data != null && data.CSVTables.Length > 0 && data.CSVTables[0].Rows.Length > 0)
                {
                    return data.CSVTables[0].Rows;
                }
            }
            catch (Exception ex)
            {
                //  WorkspaceAddIn.InfoLog(ex.Message);
            }
            return rnData;
        }

        /// <summary>
        /// Return  the results as a set of objects
        /// </summary>
        /// <param name="query"></param>
        /// <param name="objectTemplate"></param>
        /// <returns></returns>
        public RNObject[] GetRNDataObject(string query, GenericObject objectTemplate)
        {
            RNObject[] rnDataObject = null;
            QueryResultData[] data = null;
            ClientInfoHeader clientInfoHeader = new ClientInfoHeader { AppID = "Get RN Object" };
            try
            {
                data = _rightNowClient.QueryObjects(clientInfoHeader, query, new RNObject[] { objectTemplate }, 1000);
                if (data != null)
                {
                    return data[0].RNObjectsResult;
                }
            }
            catch (Exception ex)
            {
                //  WorkspaceAddIn.InfoLog("Exception in GetRNDataObject: " + ex.Message);
            }
            return rnDataObject;
        }







        public string[] CheckIfRefExists(int CustID, int ID, string refno)
        {
            refno = refno.Trim();
            string query = "SELECT p.ReferenceNumber, p.CustomFields.CO.nf_claim_nmbr FROM Incident p WHERE ID != " + ID + " AND p.CustomFields.CO.nf_claim_nmbr like '%" + refno + "%' AND organization =" + CustID;

            // string query = "SELECT p.ReferenceNumber, p.CustomFields.CO.nf_claim_nmbr FROM Incident p WHERE ID = 3404 ";
            //  MessageBox.Show(query);
            string[] resultSet = GetRNData("Ref Check", query);

            return resultSet;


        }


        public string[] GetITSNumber(int ITSID)
        {
            string query = "SELECT p.ITS_nmbr FROM CO.ITS p WHERE ID = " + ITSID;

            // string query = "SELECT p.ReferenceNumber, p.CustomFields.CO.nf_claim_nmbr FROM Incident p WHERE ID = 3404 ";
            //  MessageBox.Show(query);
           // string[] resultSet = GetRNData("ITS", query);
            string[] resultSet = GetRNData("ITS", query);
            return resultSet;


        }


        public string GetIncVIN(string IncVinID)
        {
            string query = "SELECT p.Bus FROM CO.Incident_VIN p WHERE ID = " + IncVinID;
            string busID = "";

            List<string> incvins = new List<string>();
            MessageBox.Show(query);

            string[] resultSet = GetRNData("IncVin", query);

            if (resultSet != null && resultSet.Length > 0)
            {
                incvins = resultSet.ToList();
                foreach (string v in incvins)
                {
                    busID = v.Split('~')[0];
                    break;
                }
            }
            return busID;
        }

        public string GetBus(string busID)
        {
            string query = "SELECT p.sales_release FROM CO.Bus p WHERE ID = " + busID;
            
            List<string> buses = new List<string>();
            List<string> srs = new List<string>();
            String SR = "";
            String SRID = "";
            MessageBox.Show(query);

            string[] resultSet = GetRNData("Bus", query);

            if (resultSet != null && resultSet.Length > 0)
            {
                buses = resultSet.ToList();
                foreach (string b in buses)
                {
                    SRID = b.Split('~')[0];
                    break;
                }
            }

            if (SRID != "")
            {
                query = "SELECT p.sr_nmbr FROM CO.SalesRelease p WHERE ID = " + SRID;
                resultSet = GetRNData("SR", query);

                if (resultSet != null && resultSet.Length > 0)
                {
                    srs = resultSet.ToList();
                    foreach (string s in srs)
                    {
                        SR = s.Split('~')[0];
                        break;
                    }
                }

            }
            return SR;
        }


        /// <summary>
        /// Method which is called to get value of a custom field of Incident record.
        /// </summary>
        /// <param name="packageName">The name of the package.</param>
        /// <param name="fieldName">The name of the custom field.</param>
        /// <returns>Value of the field</returns>
        public string GetFieldFromIncidentRecord(IIncident incidentRecord, string packageName, string fieldName)
        {
            string value = "";
            IList<ICustomAttribute> incCustomAttributes = incidentRecord.CustomAttributes;

            foreach (ICustomAttribute val in incCustomAttributes)
            {
                if (val.PackageName == packageName)//if package name matches
                {
                    if (val.GenericField.Name == packageName + "$" + fieldName)//if field matches
                    {
                        if (val.GenericField.DataValue.Value != null)
                        {
                            value = val.GenericField.DataValue.Value.ToString();
                            break;
                        }
                    }
                }
            }
            return value;
        }




        #region Miscellaneous




        /// <summary>
        /// Create string type data value
        /// </summary>
        /// <param name="val"></param>
        /// <returns> DataValue</returns>
        private DataValue createStringDataValue(string val)
        {
            DataValue dv = new DataValue();
            dv.Items = new Object[] { val };
            dv.ItemsElementName = new ItemsChoiceType[] { ItemsChoiceType.StringValue };  //Change this to the type of field
            return dv;
        }
        /// <summary>
        /// Create string type data value
        /// </summary>
        /// <param name="val"></param>
        /// <returns> DataValue</returns>
        private DataValue createDateDataValue(DateTime val)
        {
            DataValue dv = new DataValue();
            dv.Items = new Object[] { val };
            dv.ItemsElementName = new ItemsChoiceType[] { ItemsChoiceType.DateValue };  //Change this to the type of field
            return dv;
        }
        /// <summary>
        /// Create Boolean type data value
        /// </summary>
        /// <param name="val"></param>
        /// <returns> DataValue</returns>
        private DataValue createBooleanDataValue(Boolean val)
        {
            DataValue dv = new DataValue();
            dv.Items = new Object[] { val };
            dv.ItemsElementName = new ItemsChoiceType[] { ItemsChoiceType.BooleanValue };

            return dv;
        }

        /// <summary>
        /// Create integer type data value
        /// </summary>
        /// <param name="val"></param>
        /// <returns> DataValue</returns>
        private DataValue createIntegerDataValue(int val)
        {
            DataValue dv = new DataValue();
            dv.Items = new Object[] { val };
            dv.ItemsElementName = new ItemsChoiceType[] { ItemsChoiceType.IntegerValue };  //Change this to the type of field
            return dv;
        }

        /// <summary>
        /// Create GenericField object
        /// </summary>
        /// <param name="name">Name Of Generic Field</param>
        /// <param name="dataValue">Vlaue of generic field</param>
        /// <param name="type">Type of generic field</param>
        /// <returns> GenericField</returns>
        private GenericField createGenericField(string name, DataValue dataValue, DataTypeEnum type)
        {
            GenericField genericField = new GenericField();

            genericField.dataType = type;
            genericField.dataTypeSpecified = true;
            genericField.name = name;
            genericField.DataValue = dataValue;
            return genericField;
        }

        /// <summary>
        /// Create Named ID type Data Value for NamedID as input
        /// </summary>
        /// <param name="namedvalue"></param>
        /// <returns></returns>
        private DataValue createNamedID(NamedID namedvalue)
        {
            DataValue dv = new DataValue();
            dv.Items = new Object[] { namedvalue };
            dv.ItemsElementName = new ItemsChoiceType[] { ItemsChoiceType.NamedIDValue };
            return dv;
        }

        /// <summary>
        /// Create Named ID type data value for integer value as input
        /// </summary>
        /// <param name="idVal"></param>
        /// <returns> DataValue</returns>
        private DataValue createNamedIDDataValue(long idVal)
        {
            ID id = new ID();
            id.id = idVal;
            id.idSpecified = true;

            NamedID namedID = new NamedID();
            namedID.ID = id;

            DataValue dv = new DataValue();
            dv.Items = new Object[] { namedID };
            dv.ItemsElementName = new ItemsChoiceType[] { ItemsChoiceType.NamedIDValue };

            return dv;
        }
        #endregion
    }
}
